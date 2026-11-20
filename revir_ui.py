"""
Revir UI - Saf Python Masaüstü Arayüz ve Reaktif İşleme Motoru (revir_ui.py)

Harici web çerçevesi gerektirmeden, yalnızca Python standart
kütüphanesi (ThreadingHTTPServer) ve yerel masaüstü penceresi (pywebview) ile
çalışan hızlı, çevrimdışı ve deterministik arayüz motorudur.
"""
from __future__ import annotations

import base64
import copy
import datetime as dt_mod
import functools
import hashlib
import html as html_lib
import io
import json
import math
import mimetypes
import os
import re
import secrets
import socket
import threading
import time
import traceback
import urllib.parse
import webbrowser
from collections import OrderedDict
from contextlib import contextmanager
from dataclasses import dataclass, field
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd


# ==============================================================================
# 1. ÇEKİRDEK DURUM, ÖNBELLEK VE İSTİSNA TİPLERİ
# ==============================================================================

class _RerunException(BaseException):
	"""`except Exception:` blokları tarafından yutulmaması için BaseException'dan türer."""
	def __init__(self, scope: str = "app") -> None:
		super().__init__(scope)
		self.scope = scope or "app"


class SessionState(dict):
	"""Hem sözlük (`session_state["k"]`) hem nitelik (`session_state.k`) erişimi destekler."""
	def __getattr__(self, name: str) -> Any:
		try:
			return self[name]
		except KeyError as exc:
			raise AttributeError(name) from exc

	def __setattr__(self, name: str, value: Any) -> None:
		self[name] = value

	def __delattr__(self, name: str) -> None:
		try:
			del self[name]
		except KeyError as exc:
			raise AttributeError(name) from exc


class UploadedFile(io.BytesIO):
	"""Dosya yükleyici bileşeninden dönen bellek-içi dosya nesnesi."""
	def __init__(self, name: str, data: bytes, mime_type: str = "application/octet-stream", file_id: str = "") -> None:
		super().__init__(data)
		self.name = name
		self.type = mime_type or mimetypes.guess_type(name)[0] or "application/octet-stream"
		self.size = len(data)
		self.file_id = file_id or secrets.token_hex(6)

	def getvalue(self) -> bytes:
		return super().getvalue()


@dataclass
class _SelectionState:
	rows: list[int] = field(default_factory=list)
	cells: list[tuple[int, str]] = field(default_factory=list)
	columns: list[str] = field(default_factory=list)

	def get(self, key: str, default: Any = None) -> Any:
		return getattr(self, key, default)

	def __getitem__(self, key: str) -> Any:
		return getattr(self, key)


@dataclass
class _DataframeSelectionEvent:
	selection: _SelectionState = field(default_factory=_SelectionState)

	def get(self, key: str, default: Any = None) -> Any:
		if key == "selection":
			return self.selection
		return default

	def __getitem__(self, key: str) -> Any:
		if key == "selection":
			return self.selection
		raise KeyError(key)


@dataclass
class _ActiveDialog:
	title: str
	width: str
	on_dismiss: Callable | None
	func: Callable
	args: tuple
	kwargs: dict


# ==============================================================================
# 2. ÖNBELLEK DEKORATÖRLERİ (cache_data / cache_resource)
# ==============================================================================

def _make_cache_key(args: tuple, kwargs: dict) -> str:
	def _norm(obj: Any) -> Any:
		if obj is None or isinstance(obj, (bool, int, float, str)):
			return obj
		if isinstance(obj, Path):
			return ("Path", str(obj))
		if isinstance(obj, (dt_mod.date, dt_mod.datetime, dt_mod.time)):
			return ("DT", obj.isoformat())
		if isinstance(obj, dict):
			return ("dict", sorted((str(k), _norm(v)) for k, v in obj.items()))
		if isinstance(obj, (list, tuple)):
			return (type(obj).__name__, [_norm(x) for x in obj])
		if isinstance(obj, set):
			return ("set", sorted(_norm(x) for x in obj))
		if isinstance(obj, pd.DataFrame):
			try:
				return ("df", int(pd.util.hash_pandas_object(obj, index=True).sum()), list(obj.columns), obj.shape)
			except Exception:
				return ("df_repr", obj.shape, list(obj.columns), str(obj.values.tolist()[:20]))
		return (type(obj).__name__, repr(obj))

	payload = json.dumps([_norm(args), _norm(kwargs)], ensure_ascii=False, default=str)
	return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def cache_data(func: Callable | None = None, *, ttl: float | None = None, show_spinner: bool | str = True, max_entries: int = 256):
	def _decorator(fn: Callable) -> Callable:
		store: OrderedDict[str, tuple[float, Any]] = OrderedDict()
		lock = threading.RLock()

		@functools.wraps(fn)
		def _wrapper(*args, **kwargs):
			key = _make_cache_key(args, kwargs)
			now = time.monotonic()
			with lock:
				entry = store.get(key)
				if entry is not None:
					ts, val = entry
					if ttl is None or (now - ts) < float(ttl):
						store.move_to_end(key)
						return copy.deepcopy(val)
					store.pop(key, None)
			result = fn(*args, **kwargs)
			with lock:
				store[key] = (now, copy.deepcopy(result))
				store.move_to_end(key)
				while len(store) > max_entries:
					store.popitem(last=False)
			return result

		def _clear() -> None:
			with lock:
				store.clear()

		_wrapper.clear = _clear  # type: ignore[attr-defined]
		return _wrapper

	if func is not None:
		return _decorator(func)
	return _decorator


def cache_resource(func: Callable | None = None, *, ttl: float | None = None, show_spinner: bool | str = True):
	def _decorator(fn: Callable) -> Callable:
		store: dict[str, tuple[float, Any]] = {}
		lock = threading.RLock()

		@functools.wraps(fn)
		def _wrapper(*args, **kwargs):
			key = _make_cache_key(args, kwargs)
			now = time.monotonic()
			with lock:
				entry = store.get(key)
				if entry is not None:
					ts, val = entry
					if ttl is None or (now - ts) < float(ttl):
						return val
					store.pop(key, None)
			result = fn(*args, **kwargs)
			with lock:
				store[key] = (now, result)
			return result

		def _clear() -> None:
			with lock:
				store.clear()

		_wrapper.clear = _clear  # type: ignore[attr-defined]
		return _wrapper

	if func is not None:
		return _decorator(func)
	return _decorator


# ==============================================================================
# 3. SÜTUN YAPILANDIRMA SINIFLARI (column_config)
# ==============================================================================

@dataclass
class _ColumnConfig:
	kind: str
	label: str | None = None
	help: str | None = None
	width: str | int | None = None
	disabled: bool | None = None
	required: bool = False
	min_value: float | int | None = None
	max_value: float | int | None = None
	step: float | int | None = None
	format: str | None = None
	options: list[Any] | None = None


class _ColumnConfigNamespace:
	@staticmethod
	def TextColumn(label: str | None = None, *, width: str | int | None = None, help: str | None = None, disabled: bool | None = None, required: bool = False, default: Any = None, max_chars: int | None = None, validate: str | None = None) -> _ColumnConfig:
		return _ColumnConfig(kind="text", label=label, help=help, width=width, disabled=disabled, required=required)

	@staticmethod
	def NumberColumn(label: str | None = None, *, width: str | int | None = None, help: str | None = None, disabled: bool | None = None, required: bool = False, default: Any = None, format: str | None = None, min_value: float | int | None = None, max_value: float | int | None = None, step: float | int | None = None) -> _ColumnConfig:
		return _ColumnConfig(kind="number", label=label, help=help, width=width, disabled=disabled, required=required, min_value=min_value, max_value=max_value, step=step, format=format)

	@staticmethod
	def SelectboxColumn(label: str | None = None, *, width: str | int | None = None, help: str | None = None, disabled: bool | None = None, required: bool = False, default: Any = None, options: Sequence[Any] | None = None) -> _ColumnConfig:
		return _ColumnConfig(kind="selectbox", label=label, help=help, width=width, disabled=disabled, required=required, options=list(options) if options is not None else [])

	@staticmethod
	def CheckboxColumn(label: str | None = None, *, width: str | int | None = None, help: str | None = None, disabled: bool | None = None, required: bool = False, default: Any = None) -> _ColumnConfig:
		return _ColumnConfig(kind="checkbox", label=label, help=help, width=width, disabled=disabled, required=required)


column_config = _ColumnConfigNamespace()


# ==============================================================================
# 4. MARKDOWN VE ÇEVRİMDIŞI SVG GRAFİK İŞLEYİCİLERİ
# ==============================================================================

_INLINE_CODE_RE = re.compile(r"`([^`]+)`")
_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)\s]+)\)")
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)")
_STRIKE_RE = re.compile(r"~~(.+?)~~")


def _format_inline_md(text: str, escape_html: bool = True) -> str:
	if not text:
		return ""
	placeholders: list[str] = []

	def _stash(raw_html: str) -> str:
		idx = len(placeholders)
		placeholders.append(raw_html)
		return f"\x00PH{idx}\x00"

	if not escape_html:
		# HTML etiketlerini koru, dışındaki metinde Markdown işle
		def _keep_tag(m: re.Match) -> str:
			return _stash(m.group(0))
		text = re.sub(r"<[^>]+>", _keep_tag, text)
	else:
		text = html_lib.escape(text, quote=False)

	def _code_sub(m: re.Match) -> str:
		return _stash(f"<code>{m.group(1)}</code>")

	text = _INLINE_CODE_RE.sub(_code_sub, text)

	def _link_sub(m: re.Match) -> str:
		lbl = m.group(1)
		href = html_lib.escape(m.group(2), quote=True)
		return _stash(f'<a href="{href}" target="_blank" rel="noopener">{lbl}</a>')

	text = _LINK_RE.sub(_link_sub, text)
	text = _BOLD_RE.sub(r"<strong>\1</strong>", text)
	text = _STRIKE_RE.sub(r"<del>\1</del>", text)
	text = _ITALIC_RE.sub(r"<em>\1</em>", text)

	for idx, val in enumerate(placeholders):
		text = text.replace(f"\x00PH{idx}\x00", val)
	return text


def _render_markdown(body: str, unsafe_allow_html: bool = False) -> str:
	if body is None:
		return ""
	text = str(body)
	if not text.strip():
		return ""

	if unsafe_allow_html:
		stripped = text.lstrip()
		if stripped.startswith(("<div", "<span", "<table", "<p", "<h1", "<h2", "<h3", "<h4", "<h5", "<h6", "<ul", "<ol", "<hr", "<br", "<section", "<article", "<header", "<footer", "<details", "<svg", "<img", "<a ", "<iframe", "<script", "<style", "<link", "<!--")):
			return text

	lines = text.splitlines()
	out: list[str] = []
	in_code = False
	code_buf: list[str] = []
	in_ul = False
	in_ol = False
	in_table = False
	table_rows: list[list[str]] = []

	def _close_lists():
		nonlocal in_ul, in_ol
		if in_ul:
			out.append("</ul>")
			in_ul = False
		if in_ol:
			out.append("</ol>")
			in_ol = False

	def _flush_table():
		nonlocal in_table, table_rows
		if not in_table:
			return
		in_table = False
		if not table_rows:
			return
		html_parts = ['<div class="revir-md-table-wrap"><table class="revir-md-table">']
		for r_idx, row in enumerate(table_rows):
			if r_idx == 1 and all(re.match(r"^:?-{2,}:?$", c.strip()) for c in row if c.strip()):
				continue
			tag = "th" if r_idx == 0 else "td"
			cells = "".join(f"<{tag}>{_format_inline_md(c.strip(), escape_html=not unsafe_allow_html)}</{tag}>" for c in row)
			if r_idx == 0:
				html_parts.append(f"<thead><tr>{cells}</tr></thead><tbody>")
			else:
				html_parts.append(f"<tr>{cells}</tr>")
		html_parts.append("</tbody></table></div>")
		out.append("".join(html_parts))
		table_rows = []

	for raw_line in lines:
		line = raw_line.rstrip()
		if line.strip().startswith("```"):
			_close_lists()
			_flush_table()
			if not in_code:
				in_code = True
				code_buf = []
			else:
				in_code = False
				out.append(f"<pre class='revir-code'><code>{html_lib.escape(chr(10).join(code_buf))}</code></pre>")
				code_buf = []
			continue
		if in_code:
			code_buf.append(raw_line)
			continue

		s = line.strip()
		if not s:
			_close_lists()
			_flush_table()
			continue

		# Tablo satırı
		if s.startswith("|") and s.endswith("|") and s.count("|") >= 2:
			_close_lists()
			in_table = True
			parts = [c for c in s.strip("|").split("|")]
			table_rows.append(parts)
			continue
		else:
			_flush_table()

		# Yatay çizgi
		if re.match(r"^([-*_])\1{2,}$", s):
			_close_lists()
			out.append("<hr class='revir-hr'/>")
			continue

		# Başlıklar
		hm = re.match(r"^(#{1,6})\s+(.*)$", s)
		if hm:
			_close_lists()
			lvl = len(hm.group(1))
			out.append(f"<h{lvl} class='revir-h{lvl}'>{_format_inline_md(hm.group(2), escape_html=not unsafe_allow_html)}</h{lvl}>")
			continue

		# Alıntı
		if s.startswith("> "):
			_close_lists()
			out.append(f"<blockquote class='revir-quote'>{_format_inline_md(s[2:], escape_html=not unsafe_allow_html)}</blockquote>")
			continue

		# Madde işareti
		um = re.match(r"^[-*•]\s+(.*)$", s)
		if um:
			if in_ol:
				out.append("</ol>")
				in_ol = False
			if not in_ul:
				out.append("<ul class='revir-ul'>")
				in_ul = True
			out.append(f"<li>{_format_inline_md(um.group(1), escape_html=not unsafe_allow_html)}</li>")
			continue

		# Numaralı liste
		om = re.match(r"^\d+\.\s+(.*)$", s)
		if om:
			if in_ul:
				out.append("</ul>")
				in_ul = False
			if not in_ol:
				out.append("<ol class='revir-ol'>")
				in_ol = True
			out.append(f"<li>{_format_inline_md(om.group(1), escape_html=not unsafe_allow_html)}</li>")
			continue

		_close_lists()
		out.append(f"<p>{_format_inline_md(s, escape_html=not unsafe_allow_html)}</p>")

	_close_lists()
	_flush_table()
	if in_code and code_buf:
		out.append(f"<pre class='revir-code'><code>{html_lib.escape(chr(10).join(code_buf))}</code></pre>")

	return "\n".join(out)


def _render_svg_bar_chart(data: Any, height: int = 220, color: str = "#0ea5e9") -> str:
	"""Çevrimdışı etkileşimli SVG çubuk grafiği üretir."""
	labels: list[str] = []
	values: list[float] = []
	if isinstance(data, pd.Series):
		for k, v in data.items():
			labels.append(str(k))
			try:
				values.append(float(v) if pd.notna(v) else 0.0)
			except Exception:
				values.append(0.0)
	elif isinstance(data, pd.DataFrame) and not data.empty:
		num_cols = [c for c in data.columns if pd.api.types.is_numeric_dtype(data[c])]
		col = num_cols[0] if num_cols else data.columns[0]
		for idx, val in zip(data.index, data[col]):
			labels.append(str(idx))
			try:
				values.append(float(val) if pd.notna(val) else 0.0)
			except Exception:
				values.append(0.0)
	elif isinstance(data, dict):
		for k, v in data.items():
			labels.append(str(k))
			values.append(float(v or 0))

	if not labels:
		return '<div class="revir-chart-empty">Grafik için veri yok.</div>'

	w = 760
	h = max(180, int(height or 220))
	pad_l, pad_r, pad_t, pad_b = 46, 20, 20, 44
	plot_w = w - pad_l - pad_r
	plot_h = h - pad_t - pad_b
	max_v = max(values) if values else 1.0
	if max_v <= 0:
		max_v = 1.0
	y_top = math.ceil(max_v * 1.15)
	if y_top <= 0:
		y_top = 1

	parts = [f'<div class="revir-chart-card"><svg viewBox="0 0 {w} {h}" class="revir-svg-chart" style="height:{h}px;">']
	# Yatay kılavuz çizgileri
	ticks = 4
	for i in range(ticks + 1):
		yv = (y_top / ticks) * i
		py = pad_t + plot_h - (yv / y_top) * plot_h
		parts.append(f'<line x1="{pad_l}" y1="{py:.1f}" x2="{w - pad_r}" y2="{py:.1f}" class="revir-grid-line"/>')
		parts.append(f'<text x="{pad_l - 8}" y="{py + 4:.1f}" text-anchor="end" class="revir-axis-lbl">{int(round(yv))}</text>')

	n = len(labels)
	slot_w = plot_w / max(n, 1)
	bar_w = min(46.0, max(8.0, slot_w * 0.62))
	safe_col = html_lib.escape(str(color or "#0ea5e9"), quote=True)

	for i, (lbl, val) in enumerate(zip(labels, values)):
		bh = (val / y_top) * plot_h if y_top else 0.0
		bx = pad_l + i * slot_w + (slot_w - bar_w) / 2.0
		by = pad_t + plot_h - bh
		short_lbl = lbl[-5:] if len(lbl) == 10 and lbl[4] == "-" else lbl[:12]
		tip = html_lib.escape(f"{lbl}: {val:g}", quote=True)
		parts.append(
			f'<g class="revir-bar-group" data-tip="{tip}">'
			f'<rect x="{bx:.1f}" y="{by:.1f}" width="{bar_w:.1f}" height="{max(1.0, bh):.1f}" rx="4" fill="{safe_col}" class="revir-bar-rect"/>'
			f'<text x="{bx + bar_w / 2:.1f}" y="{by - 5:.1f}" text-anchor="middle" class="revir-bar-val">{val:g}</text>'
			f'</g>'
		)
		if n <= 16 or i % max(1, n // 12) == 0:
			parts.append(f'<text x="{bx + bar_w / 2:.1f}" y="{h - 14}" text-anchor="middle" class="revir-axis-lbl">{html_lib.escape(short_lbl)}</text>')

	parts.append("</svg></div>")
	return "".join(parts)


_PALETTE = ["#0d9488", "#2563eb", "#d97706", "#dc2626", "#7c3aed", "#059669", "#db2777", "#0284c7"]


def _parse_x_val(val: Any, is_temporal: bool) -> float | None:
	if val is None or (isinstance(val, float) and math.isnan(val)):
		return None
	if is_temporal:
		try:
			ts = pd.to_datetime(val)
			return float(ts.value) / 1e9
		except Exception:
			return None
	try:
		return float(val)
	except Exception:
		return None


def _format_tip_val(val: Any, fmt: str | None = None, is_temporal: bool = False) -> str:
	if val is None:
		return "—"
	if is_temporal:
		try:
			return pd.to_datetime(val).strftime("%Y-%m-%d")
		except Exception:
			return str(val)[:10]
	if fmt and isinstance(val, (int, float)):
		try:
			if fmt.startswith("+.") and fmt.endswith("f"):
				return f"{val:{fmt}}"
			if fmt.startswith(".") and fmt.endswith("f"):
				return f"{val:{fmt}}"
			return format(val, fmt)
		except Exception:
			pass
	if isinstance(val, float):
		return f"{val:.1f}" if abs(val - round(val)) > 1e-6 else f"{int(round(val))}"
	return str(val)


def _render_svg_altair_chart(chart: Any) -> str:
	"""Altair / Vega-Lite grafik nesnesini %100 çevrimdışı, etkileşimli SVG olarak çizer."""
	try:
		spec = chart.to_dict() if hasattr(chart, "to_dict") else (chart if isinstance(chart, dict) else {})
	except Exception as exc:
		return f'<div class="revir-chart-empty">Grafik işlenemedi: {html_lib.escape(str(exc))}</div>'

	datasets: dict[str, list[dict]] = spec.get("datasets") or {}
	layers = spec.get("layer") if isinstance(spec.get("layer"), list) else [spec]
	height = int(spec.get("height") or 340)
	w = 820
	h = max(240, height)
	pad_l, pad_r, pad_t, pad_b = 58, 26, 22, 48
	plot_w = w - pad_l - pad_r
	plot_h = h - pad_t - pad_b

	def _resolve_rows(layer_obj: dict) -> list[dict]:
		d = layer_obj.get("data") or spec.get("data") or {}
		if isinstance(d, dict):
			if "values" in d and isinstance(d["values"], list):
				return d["values"]
			if "name" in d and d["name"] in datasets:
				return datasets[d["name"]]
		return []

	is_temporal = False
	x_title = ""
	y_title = ""
	x_domain_fixed = None
	y_domain_fixed = None
	color_domain: list[str] = []
	color_range: list[str] = []
	legend_items: list[tuple[str, str]] = []
	legend_seen: set[str] = set()

	all_xs: list[float] = []
	all_ys: list[float] = []
	parsed_layers: list[dict] = []

	for lay in layers:
		rows = _resolve_rows(lay)
		enc = lay.get("encoding") or {}
		mark_raw = lay.get("mark")
		if isinstance(mark_raw, str):
			mark_type = mark_raw
			mark_cfg: dict[str, Any] = {}
		elif isinstance(mark_raw, dict):
			mark_type = str(mark_raw.get("type") or "line")
			mark_cfg = mark_raw
		else:
			mark_type = "line"
			mark_cfg = {}

		x_enc = enc.get("x") or {}
		y_enc = enc.get("y") or {}
		c_enc = enc.get("color") or {}
		t_enc = enc.get("tooltip") or []
		if isinstance(t_enc, dict):
			t_enc = [t_enc]

		x_field = x_enc.get("field")
		y_field = y_enc.get("field")
		c_field = c_enc.get("field")

		if x_enc.get("type") == "temporal":
			is_temporal = True
		if x_enc.get("title") and not x_title:
			x_title = str(x_enc["title"])
		if y_enc.get("title") and not y_title:
			y_title = str(y_enc["title"])
		if isinstance(x_enc.get("scale"), dict) and x_enc["scale"].get("domain") and x_domain_fixed is None:
			dom = x_enc["scale"]["domain"]
			if isinstance(dom, list) and len(dom) == 2:
				x_domain_fixed = (float(dom[0]), float(dom[1]))
		if isinstance(y_enc.get("scale"), dict) and y_enc["scale"].get("domain") and y_domain_fixed is None:
			dom = y_enc["scale"]["domain"]
			if isinstance(dom, list) and len(dom) == 2:
				y_domain_fixed = (float(dom[0]), float(dom[1]))
		if isinstance(c_enc.get("scale"), dict):
			if c_enc["scale"].get("domain"):
				color_domain = [str(x) for x in c_enc["scale"]["domain"]]
			if c_enc["scale"].get("range"):
				color_range = [str(x) for x in c_enc["scale"]["range"]]

		for r in rows:
			if x_field and x_field in r:
				xv = _parse_x_val(r.get(x_field), is_temporal)
				if xv is not None:
					all_xs.append(xv)
			if y_field and y_field in r:
				yv = _parse_x_val(r.get(y_field), False)
				if yv is not None:
					all_ys.append(yv)

		parsed_layers.append({
			"rows": rows,
			"mark_type": mark_type,
			"mark_cfg": mark_cfg,
			"x_field": x_field,
			"y_field": y_field,
			"c_field": c_field,
			"tooltips": t_enc,
		})

	if not all_ys:
		return '<div class="revir-chart-empty">Grafik çizmek için yeterli veri bulunamadı.</div>'

	if x_domain_fixed is not None:
		xmin, xmax = x_domain_fixed
	elif all_xs:
		xmin, xmax = min(all_xs), max(all_xs)
		if abs(xmax - xmin) < 1e-9:
			pad_x = 86400.0 if is_temporal else 1.0
			xmin -= pad_x
			xmax += pad_x
	else:
		xmin, xmax = 0.0, 1.0

	if y_domain_fixed is not None:
		ymin, ymax = y_domain_fixed
	else:
		ymin, ymax = min(all_ys), max(all_ys)
		span = ymax - ymin
		if span < 1e-9:
			ymin -= 1.0
			ymax += 1.0
		else:
			ymin -= span * 0.06
			ymax += span * 0.06

	def _sx(x: float) -> float:
		return pad_l + ((x - xmin) / (xmax - xmin)) * plot_w if xmax != xmin else pad_l + plot_w / 2.0

	def _sy(y: float) -> float:
		return pad_t + plot_h - ((y - ymin) / (ymax - ymin)) * plot_h if ymax != ymin else pad_t + plot_h / 2.0

	color_map: dict[str, str] = {}
	for i, d_lbl in enumerate(color_domain):
		if i < len(color_range):
			color_map[d_lbl] = color_range[i]

	def _color_for(series_key: str) -> str:
		if series_key not in color_map:
			color_map[series_key] = _PALETTE[len(color_map) % len(_PALETTE)]
		col = color_map[series_key]
		if series_key not in legend_seen:
			legend_seen.add(series_key)
			legend_items.append((series_key, col))
		return col

	def _build_tip(row: dict, tooltips: list[dict]) -> str:
		if not tooltips:
			return ""
		chunks = []
		for t in tooltips:
			fld = t.get("field")
			if fld and fld in row:
				lbl = t.get("title") or fld
				val = _format_tip_val(row.get(fld), t.get("format"), t.get("type") == "temporal")
				chunks.append(f"{lbl}: {val}")
		return " · ".join(chunks)

	svg: list[str] = [f'<svg viewBox="0 0 {w} {h}" class="revir-svg-chart" style="height:{h}px;">']

	# Izgara ve Y ekseni
	y_ticks = 5
	for i in range(y_ticks + 1):
		yv = ymin + (ymax - ymin) * (i / y_ticks)
		py = _sy(yv)
		svg.append(f'<line x1="{pad_l}" y1="{py:.1f}" x2="{w - pad_r}" y2="{py:.1f}" class="revir-grid-line"/>')
		lbl_txt = f"{yv:.0f}" if (ymax - ymin) >= 15 else f"{yv:.1f}"
		svg.append(f'<text x="{pad_l - 8}" y="{py + 4:.1f}" text-anchor="end" class="revir-axis-lbl">{lbl_txt}</text>')

	# X ekseni işaretleri
	x_ticks = 6
	for i in range(x_ticks + 1):
		xv = xmin + (xmax - xmin) * (i / x_ticks)
		px = _sx(xv)
		svg.append(f'<line x1="{px:.1f}" y1="{pad_t}" x2="{px:.1f}" y2="{pad_t + plot_h}" class="revir-grid-line" opacity="0.5"/>')
		if is_temporal:
			try:
				lbl_x = dt_mod.datetime.fromtimestamp(xv).strftime("%d.%m.%y")
			except Exception:
				lbl_x = ""
		else:
			lbl_x = f"{xv:.0f}"
		svg.append(f'<text x="{px:.1f}" y="{pad_t + plot_h + 18}" text-anchor="middle" class="revir-axis-lbl">{lbl_x}</text>')

	# Eksen başlıkları
	if x_title:
		svg.append(f'<text x="{pad_l + plot_w / 2:.1f}" y="{h - 6}" text-anchor="middle" class="revir-axis-title">{html_lib.escape(x_title)}</text>')
	if y_title:
		svg.append(f'<text x="15" y="{pad_t + plot_h / 2:.1f}" text-anchor="middle" transform="rotate(-90 15 {pad_t + plot_h / 2:.1f})" class="revir-axis-title">{html_lib.escape(y_title)}</text>')

	# Eksen çerçevesi
	svg.append(f'<line x1="{pad_l}" y1="{pad_t + plot_h}" x2="{w - pad_r}" y2="{pad_t + plot_h}" class="revir-axis-line"/>')
	svg.append(f'<line x1="{pad_l}" y1="{pad_t}" x2="{pad_l}" y2="{pad_t + plot_h}" class="revir-axis-line"/>')

	# Katmanları çiz
	for lay in parsed_layers:
		rows = lay["rows"]
		mtype = lay["mark_type"]
		mcfg = lay["mark_cfg"]
		xf, yf, cf = lay["x_field"], lay["y_field"], lay["c_field"]
		tooltips = lay["tooltips"]
		dash = mcfg.get("strokeDash")
		dash_attr = f' stroke-dasharray="{",".join(str(d) for d in dash)}"' if isinstance(dash, list) else ""
		stroke_w = float(mcfg.get("strokeWidth") or (2.0 if mtype == "line" else 1.5))

		if mtype == "rule":
			rcol = html_lib.escape(str(mcfg.get("color") or "#94a3b8"), quote=True)
			for r in rows:
				if yf and yf in r:
					yv = _parse_x_val(r.get(yf), False)
					if yv is not None:
						py = _sy(yv)
						svg.append(f'<line x1="{pad_l}" y1="{py:.1f}" x2="{w - pad_r}" y2="{py:.1f}" stroke="{rcol}" stroke-width="1.2"{dash_attr}/>')
			continue

		if mtype == "line":
			groups: OrderedDict[str, list[tuple[float, float, dict]]] = OrderedDict()
			for r in rows:
				xv = _parse_x_val(r.get(xf), is_temporal) if xf else None
				yv = _parse_x_val(r.get(yf), False) if yf else None
				if xv is None or yv is None:
					continue
				gk = str(r.get(cf)) if cf else "__single__"
				groups.setdefault(gk, []).append((xv, yv, r))
			for gk, pts in groups.items():
				pts.sort(key=lambda item: item[0])
				col = _color_for(gk) if cf else str(mcfg.get("color") or "#1e1b4b")
				path_d = " ".join(
					f"{'M' if idx == 0 else 'L'}{_sx(px):.1f},{_sy(py):.1f}"
					for idx, (px, py, _) in enumerate(pts)
				)
				svg.append(f'<path d="{path_d}" fill="none" stroke="{html_lib.escape(col, quote=True)}" stroke-width="{stroke_w:.1f}" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>')
				if mcfg.get("point"):
					for px, py, row in pts:
						tip = html_lib.escape(_build_tip(row, tooltips), quote=True)
						svg.append(f'<circle cx="{_sx(px):.1f}" cy="{_sy(py):.1f}" r="4.5" fill="{html_lib.escape(col, quote=True)}" stroke="#ffffff" stroke-width="1.5" class="revir-chart-pt" data-tip="{tip}"/>')
				elif tooltips and len(pts) <= 90:
					# WHO eğrilerinde örnekleme noktaları ile ipucu
					step = max(1, len(pts) // 24)
					for idx in range(0, len(pts), step):
						px, py, row = pts[idx]
						tip = html_lib.escape(_build_tip(row, tooltips), quote=True)
						svg.append(f'<circle cx="{_sx(px):.1f}" cy="{_sy(py):.1f}" r="4" fill="transparent" class="revir-chart-pt" data-tip="{tip}"/>')
			continue

		if mtype in ("circle", "point"):
			pcol = str(mcfg.get("color") or "#b91c1c")
			pstroke = str(mcfg.get("stroke") or "#ffffff")
			psize = float(mcfg.get("size") or 90)
			radius = max(4.0, min(9.0, math.sqrt(psize / math.pi)))
			for r in rows:
				xv = _parse_x_val(r.get(xf), is_temporal) if xf else None
				yv = _parse_x_val(r.get(yf), False) if yf else None
				if xv is None or yv is None:
					continue
				col = _color_for(str(r.get(cf))) if cf else pcol
				tip = html_lib.escape(_build_tip(r, tooltips), quote=True)
				svg.append(
					f'<circle cx="{_sx(xv):.1f}" cy="{_sy(yv):.1f}" r="{radius:.1f}" '
					f'fill="{html_lib.escape(col, quote=True)}" stroke="{html_lib.escape(pstroke, quote=True)}" '
					f'stroke-width="{stroke_w:.1f}" class="revir-chart-pt" data-tip="{tip}"/>'
				)

	svg.append("</svg>")

	legend_html = ""
	if legend_items:
		chips = "".join(
			f'<span class="revir-legend-item"><span class="revir-legend-dot" style="background:{html_lib.escape(c, quote=True)};"></span>{html_lib.escape(lbl)}</span>'
			for lbl, c in legend_items
		)
		legend_html = f'<div class="revir-chart-legend">{chips}</div>'

	return f'<div class="revir-chart-card">{"".join(svg)}{legend_html}</div>'


# ==============================================================================
# 5. RENDER BAĞLAMI VE KONTEYNER AĞACI
# ==============================================================================

class _AppSession:
	def __init__(self) -> None:
		self.lock = threading.RLock()
		self.session_state = SessionState()
		self.widget_values: dict[str, Any] = {}
		self.widget_meta: dict[str, dict[str, Any]] = {}
		self.callbacks: dict[str, tuple[Callable, tuple, dict]] = {}
		self.active_tabs: dict[str, int] = {}
		self.open_expanders: dict[str, bool] = {}
		self.sidebar_collapsed: bool = False
		self.active_dialog: _ActiveDialog | None = None
		self.uploaded_files: dict[str, list[UploadedFile]] = {}
		self.downloads: OrderedDict[str, tuple[bytes, str, str]] = OrderedDict()
		self.media: OrderedDict[str, tuple[bytes, str]] = OrderedDict()
		self.dataframe_selections: dict[str, dict[str, Any]] = {}
		self.editor_input_hashes: dict[str, str] = {}
		self.editor_meta: dict[str, dict[str, Any]] = {}
		self.form_defs: dict[str, dict[str, Any]] = {}
		self.page_title: str = "Revir Defteri"
		self.last_html_payload: dict[str, Any] = {}
		self.floating_chat_enabled: bool = False
		self.floating_chat_label: str = "KVKK Korumalı"
		self.floating_chat_handler: Callable | None = None
		self.chat_history: list[dict[str, str]] = []

	def register_download(self, data: bytes, file_name: str, mime: str) -> str:
		h = hashlib.sha1(data[:2048] + file_name.encode("utf-8", errors="ignore") + str(len(data)).encode()).hexdigest()[:16]
		token = f"dl_{h}"
		self.downloads[token] = (data, file_name, mime)
		self.downloads.move_to_end(token)
		while len(self.downloads) > 200:
			self.downloads.popitem(last=False)
		return token

	def register_media(self, data: bytes, mime: str) -> str:
		h = hashlib.sha1(data[:4096] + str(len(data)).encode()).hexdigest()[:16]
		token = f"img_{h}"
		self.media[token] = (data, mime)
		self.media.move_to_end(token)
		while len(self.media) > 250:
			self.media.popitem(last=False)
		return token


class _RenderContext:
	def __init__(self, session: _AppSession) -> None:
		self.session = session
		self.main_nodes: list[str] = []
		self.sidebar_nodes: list[str] = []
		self.dialog_nodes: list[str] = []
		self.styles_parts: list[str] = []
		self.scripts_parts: list[str] = []
		self.node_stack: list[list[str]] = [self.main_nodes]
		self.in_sidebar: bool = False
		self.in_dialog: bool = False
		self.active_form_id: str | None = None
		self.clicked_button: str | None = None
		self.submitted_form: str | None = None
		self.widget_counter: dict[str, int] = {}
		self.dialog_rendered_this_pass: bool = False
		self.balloons_triggered: bool = False
		self.used_form_keys: dict[str, list[tuple[str, bool]]] = {}

	def current_nodes(self) -> list[str]:
		return self.node_stack[-1]

	def emit(self, html_chunk: str) -> None:
		self.node_stack[-1].append(html_chunk)

	def next_auto_id(self, prefix: str, label: str = "") -> str:
		scope = "dlg" if self.in_dialog else ("sb" if self.in_sidebar else "main")
		base = f"{scope}:{prefix}:{label[:48]}"
		cnt = self.widget_counter.get(base, 0)
		self.widget_counter[base] = cnt + 1
		h = hashlib.md5(f"{base}#{cnt}".encode("utf-8")).hexdigest()[:12]
		return f"w_{prefix}_{h}"


_ACTIVE_SESSION = _AppSession()
_LOCAL_CTX = threading.local()


def _get_ctx() -> _RenderContext:
	ctx = getattr(_LOCAL_CTX, "ctx", None)
	if ctx is None:
		ctx = _RenderContext(_ACTIVE_SESSION)
		_LOCAL_CTX.ctx = ctx
	return ctx


class _SessionStateProxy:
	"""Her zaman aktif oturumun SessionState nesnesine yönlenir."""
	def _target(self) -> SessionState:
		ctx = getattr(_LOCAL_CTX, "ctx", None)
		return ctx.session.session_state if ctx is not None else _ACTIVE_SESSION.session_state

	def __getitem__(self, key: str) -> Any:
		return self._target()[key]

	def __setitem__(self, key: str, value: Any) -> None:
		self._target()[key] = value

	def __delitem__(self, key: str) -> None:
		del self._target()[key]

	def __contains__(self, key: object) -> bool:
		return key in self._target()

	def __iter__(self):
		return iter(self._target())

	def __len__(self) -> int:
		return len(self._target())

	def get(self, key: str, default: Any = None) -> Any:
		return self._target().get(key, default)

	def pop(self, key: str, *args) -> Any:
		return self._target().pop(key, *args)

	def keys(self):
		return self._target().keys()

	def values(self):
		return self._target().values()

	def items(self):
		return self._target().items()

	def setdefault(self, key: str, default: Any = None) -> Any:
		return self._target().setdefault(key, default)

	def update(self, *args, **kwargs) -> None:
		self._target().update(*args, **kwargs)

	def clear(self) -> None:
		self._target().clear()

	def __getattr__(self, name: str) -> Any:
		return getattr(self._target(), name)

	def __setattr__(self, name: str, value: Any) -> None:
		self._target()[name] = value

	def __repr__(self) -> str:
		return repr(self._target())


session_state: Any = _SessionStateProxy()


# ==============================================================================
# 6. KONTEYNER VE BİLEŞEN SINIFI
# ==============================================================================

def _format_df_cell(val: Any, col_cfg: _ColumnConfig | None) -> str:
	if val is None:
		return ""
	try:
		if isinstance(val, float) and math.isnan(val):
			return ""
		if pd.isna(val):
			return ""
	except Exception:
		pass
	if isinstance(val, bool):
		return "✓" if val else ""
	if col_cfg and col_cfg.kind == "number" and col_cfg.format and isinstance(val, (int, float)):
		try:
			return col_cfg.format % val
		except Exception:
			pass
	if isinstance(val, float):
		if abs(val - round(val)) < 1e-9:
			return str(int(round(val)))
		return f"{val:.2f}".rstrip("0").rstrip(".")
	return str(val)


def _df_fingerprint(df: pd.DataFrame) -> str:
	try:
		cols = [str(c) for c in df.columns]
		h = hashlib.md5()
		h.update(json.dumps(cols, ensure_ascii=False).encode("utf-8"))
		h.update(f"|{len(df)}|".encode("utf-8"))
		for row in df.itertuples(index=False, name=None):
			h.update(repr(row).encode("utf-8", errors="replace"))
		return h.hexdigest()
	except Exception:
		return str(id(df))


def _coerce_editor_value(raw_val: Any, col_cfg: _ColumnConfig | None, orig_dtype: Any) -> Any:
	if col_cfg and col_cfg.kind == "checkbox":
		if isinstance(raw_val, bool):
			return raw_val
		if isinstance(raw_val, str):
			return raw_val.strip().lower() in ("true", "1", "yes", "evet", "on")
		return bool(raw_val)

	if raw_val is None or (isinstance(raw_val, str) and raw_val.strip() == ""):
		if col_cfg and col_cfg.kind == "number":
			return None
		if col_cfg and col_cfg.kind in ("text", "selectbox"):
			return ""
		if pd.api.types.is_bool_dtype(orig_dtype):
			return False
		if pd.api.types.is_numeric_dtype(orig_dtype):
			return None
		return ""

	if col_cfg and col_cfg.kind == "number":
		s = str(raw_val).strip().replace(",", ".")
		try:
			num = float(s)
			if col_cfg.min_value is not None and num < float(col_cfg.min_value):
				num = float(col_cfg.min_value)
			if col_cfg.max_value is not None and num > float(col_cfg.max_value):
				num = float(col_cfg.max_value)
			if col_cfg.format == "%d" or (isinstance(col_cfg.step, int) and not isinstance(col_cfg.step, bool)):
				return int(round(num))
			return num
		except Exception:
			return raw_val

	if pd.api.types.is_bool_dtype(orig_dtype):
		return bool(raw_val)
	if pd.api.types.is_numeric_dtype(orig_dtype) and not (col_cfg and col_cfg.kind in ("text", "selectbox")):
		s = str(raw_val).strip().replace(",", ".")
		try:
			return float(s)
		except Exception:
			return raw_val
	return str(raw_val)


def _label_html(label: str, label_visibility: str = "visible", help_text: str | None = None, html_for: str = "") -> str:
	if label_visibility == "collapsed":
		return ""
	cls = "revir-field-label revir-sr-only" if label_visibility == "hidden" else "revir-field-label"
	help_badge = f' <span class="revir-help-icon" title="{html_lib.escape(help_text, quote=True)}">?</span>' if help_text else ""
	for_attr = f' for="{html_lib.escape(html_for, quote=True)}"' if html_for else ""
	return f'<label class="{cls}"{for_attr}>{_format_inline_md(label, escape_html=True)}{help_badge}</label>'


class _UIContainer:
	"""Her konteyner (sidebar, column, tab, expander, popover, form, root) tüm bileşen metotlarını destekler."""
	def __init__(self, target_list: list[str] | None = None, *, is_sidebar: bool = False, form_id: str | None = None, wrapper_open: str = "", wrapper_close: str = "") -> None:
		self._target_list = target_list
		self._is_sidebar = is_sidebar
		self._form_id = form_id
		self._wrapper_open = wrapper_open
		self._wrapper_close = wrapper_close
		self._inner_nodes: list[str] = []
		self._prev_sidebar = False
		self._prev_form_id: str | None = None
		self._entered = False

	def _nodes_for_direct_call(self) -> list[str]:
		if self._target_list is not None:
			return self._target_list
		return self._inner_nodes

	def __enter__(self) -> _UIContainer:
		ctx = _get_ctx()
		self._prev_sidebar = ctx.in_sidebar
		self._prev_form_id = ctx.active_form_id
		if self._is_sidebar:
			ctx.in_sidebar = True
			ctx.node_stack.append(ctx.sidebar_nodes)
		elif self._target_list is not None:
			ctx.node_stack.append(self._target_list)
		else:
			self._inner_nodes = []
			ctx.node_stack.append(self._inner_nodes)
		if self._form_id is not None:
			ctx.active_form_id = self._form_id
		self._entered = True
		return self

	def __exit__(self, exc_type, exc_val, exc_tb) -> None:
		ctx = _get_ctx()
		if self._entered:
			ctx.node_stack.pop()
			if not self._is_sidebar and self._target_list is None:
				body = _flatten_nodes(self._inner_nodes)
				ctx.emit(f"{self._wrapper_open}{body}{self._wrapper_close}")
			ctx.in_sidebar = self._prev_sidebar
			ctx.active_form_id = self._prev_form_id
			self._entered = False

	def _run_in_self(self, fn: Callable, *args, **kwargs):
		"""`col.text_input(...)` gibi doğrudan çağrıları kendi düğüm listesinde çalıştırır."""
		if self._target_list is None and not self._is_sidebar:
			return fn(*args, **kwargs)
		with self:
			return fn(*args, **kwargs)

	# --- Sayfa & Düzen ---
	def set_page_config(self, page_title: str = "Revir Defteri", page_icon: Any = None, layout: str = "wide", initial_sidebar_state: str = "expanded") -> None:
		ctx = _get_ctx()
		ctx.session.page_title = str(page_title or "Revir Defteri")

	def columns(self, spec: int | Sequence[int | float], *, gap: str = "small", vertical_alignment: str = "top") -> list[_UIContainer]:
		ctx = _get_ctx()
		if isinstance(spec, int):
			weights = [1.0] * max(1, spec)
		else:
			weights = [float(x) if float(x) > 0 else 1.0 for x in spec]
		col_lists: list[list[str]] = [[] for _ in weights]
		grid_tpl = " ".join(f"minmax(0, {w:g}fr)" for w in weights)
		gap_px = {"small": "14px", "medium": "22px", "large": "32px"}.get(gap, "14px")
		align_css = {"top": "start", "center": "center", "bottom": "end"}.get(vertical_alignment, "start")

		# Yer tutucu liste; render sırasında son HTML birleştirilir
		placeholder = _LazyColumnsNode(col_lists, grid_tpl, gap_px, align_css)
		ctx.current_nodes().append(placeholder)  # type: ignore[arg-type]
		return [_UIContainer(target_list=cl) for cl in col_lists]

	def container(self, *, border: bool = False, height: int | None = None) -> _UIContainer:
		cls = "revir-container revir-card" if border else "revir-container"
		style = f' style="height:{int(height)}px; overflow-y:auto;"' if height else ""
		return _UIContainer(wrapper_open=f'<div class="{cls}"{style}>', wrapper_close="</div>")

	def expander(self, label: str, expanded: bool = False) -> _UIContainer:
		ctx = _get_ctx()
		exp_id = ctx.next_auto_id("exp", label)
		is_open = ctx.session.open_expanders.get(exp_id, bool(expanded))
		open_attr = " open" if is_open else ""
		lbl_html = _format_inline_md(label, escape_html=True)
		return _UIContainer(
			wrapper_open=(
				f'<details class="revir-expander" data-exp-id="{exp_id}"{open_attr}>'
				f'<summary class="revir-expander-summary">{lbl_html}</summary>'
				f'<div class="revir-expander-body">'
			),
			wrapper_close="</div></details>",
		)

	def popover(self, label: str, *, help: str | None = None, use_container_width: bool = False) -> _UIContainer:
		ctx = _get_ctx()
		pop_id = ctx.next_auto_id("pop", label)
		btn_cls = "revir-btn revir-btn-secondary revir-popover-trigger" + (" revir-w-full" if use_container_width else "")
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		return _UIContainer(
			wrapper_open=(
				f'<div class="revir-popover-wrap" data-pop-id="{pop_id}">'
				f'<button type="button" class="{btn_cls}"{title_attr} onclick="window.__revirTogglePopover(\'{pop_id}\', event)">'
				f'{_format_inline_md(label, escape_html=True)} <span class="revir-caret">▾</span></button>'
				f'<div class="revir-popover-body" id="pop_{pop_id}">'
			),
			wrapper_close="</div></div>",
		)

	def tabs(self, tab_labels: Sequence[str]) -> list[_UIContainer]:
		ctx = _get_ctx()
		labels = [str(x) for x in tab_labels]
		tab_sig = "|".join(labels)
		tab_id = ctx.next_auto_id("tabs", tab_sig)
		active_idx = ctx.session.active_tabs.get(tab_id, 0)
		if active_idx < 0 or active_idx >= len(labels):
			active_idx = 0
		pane_lists: list[list[str]] = [[] for _ in labels]
		node = _LazyTabsNode(tab_id, labels, active_idx, pane_lists)
		ctx.current_nodes().append(node)  # type: ignore[arg-type]
		return [_UIContainer(target_list=pl) for pl in pane_lists]

	def form(self, key: str, *, clear_on_submit: bool = False, border: bool = True) -> _UIContainer:
		ctx = _get_ctx()
		fid = str(key)
		ctx.session.form_defs[fid] = {"clear_on_submit": bool(clear_on_submit)}
		ctx.used_form_keys[fid] = []
		cls = "revir-form revir-card" if border else "revir-form"
		return _UIContainer(
			form_id=fid,
			wrapper_open=f'<form class="{cls}" data-form-id="{html_lib.escape(fid, quote=True)}" onsubmit="return false;">',
			wrapper_close="</form>",
		)

	@contextmanager
	def spinner(self, text: str = "İşlem yapılıyor..."):
		yield

	# --- Metin & Görsel Gösterim ---
	def markdown(self, body: Any, *, unsafe_allow_html: bool = False) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.markdown(body, unsafe_allow_html=unsafe_allow_html)
			return
		ctx = _get_ctx()
		raw = str(body or "")
		if unsafe_allow_html and ("<style" in raw.lower() or raw.lstrip().lower().startswith("<link")):
			ctx.styles_parts.append(raw)
			return
		rendered = _render_markdown(raw, unsafe_allow_html=unsafe_allow_html)
		if rendered:
			ctx.emit(f'<div class="revir-md">{rendered}</div>')

	def caption(self, body: Any, *, unsafe_allow_html: bool = False) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.caption(body, unsafe_allow_html=unsafe_allow_html)
			return
		ctx = _get_ctx()
		txt = str(body or "")
		inner = _format_inline_md(txt, escape_html=not unsafe_allow_html)
		ctx.emit(f'<div class="revir-caption">{inner}</div>')

	def write(self, *args, **kwargs) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.write(*args, **kwargs)
			return
		ctx = _get_ctx()
		for arg in args:
			if arg is None:
				continue
			if isinstance(arg, str):
				if arg == "":
					ctx.emit('<div class="revir-spacer"></div>')
				else:
					self.markdown(arg)
			elif isinstance(arg, pd.DataFrame):
				self.dataframe(arg)
			elif isinstance(arg, (dict, list)):
				self.code(json.dumps(arg, ensure_ascii=False, indent=2, default=str), language="json")
			else:
				self.markdown(str(arg))

	def text(self, body: Any) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.text(body)
			return
		_get_ctx().emit(f'<pre class="revir-plain-text">{html_lib.escape(str(body or ""))}</pre>')

	def code(self, body: Any, language: str | None = None) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.code(body, language=language)
			return
		_get_ctx().emit(f'<pre class="revir-code"><code>{html_lib.escape(str(body or ""))}</code></pre>')

	def divider(self) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.divider()
			return
		_get_ctx().emit('<hr class="revir-hr"/>')

	def _alert(self, kind: str, icon: str, body: Any) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root._alert(kind, icon, body)
			return
		inner = _render_markdown(str(body or ""), unsafe_allow_html=False)
		_get_ctx().emit(f'<div class="revir-alert revir-alert-{kind}"><span class="revir-alert-icon">{icon}</span><div class="revir-alert-body">{inner}</div></div>')

	def success(self, body: Any) -> None:
		self._alert("success", "✅", body)

	def info(self, body: Any) -> None:
		self._alert("info", "ℹ️", body)

	def warning(self, body: Any) -> None:
		self._alert("warning", "⚠️", body)

	def error(self, body: Any) -> None:
		self._alert("error", "❌", body)

	def metric(self, label: str, value: Any, delta: Any = None, *, delta_color: str = "normal", help: str | None = None) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.metric(label, value, delta, delta_color=delta_color, help=help)
			return
		val_str = "—" if value is None else str(value)
		delta_html = ""
		if delta is not None and str(delta).strip() != "":
			d_str = str(delta)
			is_neg = d_str.strip().startswith("-")
			if delta_color == "off":
				d_cls = "revir-delta-neutral"
			elif delta_color == "inverse":
				d_cls = "revir-delta-pos" if is_neg else "revir-delta-neg"
			else:
				d_cls = "revir-delta-neg" if is_neg else "revir-delta-pos"
			delta_html = f'<div class="revir-metric-delta {d_cls}">{html_lib.escape(d_str)}</div>'
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		_get_ctx().emit(
			f'<div class="revir-metric"{title_attr}>'
			f'<div class="revir-metric-label">{_format_inline_md(label, escape_html=True)}</div>'
			f'<div class="revir-metric-value">{html_lib.escape(val_str)}</div>'
			f'{delta_html}</div>'
		)

	def progress(self, value: float | int, text: str | None = None) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.progress(value, text=text)
			return
		v = float(value or 0)
		pct = v * 100.0 if v <= 1.0 else v
		pct = max(0.0, min(100.0, pct))
		lbl = f'<div class="revir-progress-text">{html_lib.escape(text)}</div>' if text else ""
		_get_ctx().emit(
			f'<div class="revir-progress-wrap">{lbl}'
			f'<div class="revir-progress-track"><div class="revir-progress-fill" style="width:{pct:.1f}%"></div></div></div>'
		)

	def balloons(self) -> None:
		_get_ctx().balloons_triggered = True

	def image(self, image: Any, caption: str | None = None, *, width: int | None = None, use_container_width: bool = False) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.image(image, caption=caption, width=width, use_container_width=use_container_width)
			return
		ctx = _get_ctx()
		src = ""
		try:
			if isinstance(image, (str, Path)):
				p = Path(image)
				if p.exists() and p.is_file():
					mime = mimetypes.guess_type(str(p))[0] or "image/png"
					tok = ctx.session.register_media(p.read_bytes(), mime)
					src = f"/api/media/{tok}"
			elif isinstance(image, (bytes, bytearray)):
				tok = ctx.session.register_media(bytes(image), "image/png")
				src = f"/api/media/{tok}"
			elif hasattr(image, "save"):  # PIL.Image
				buf = io.BytesIO()
				image.save(buf, format="PNG")
				tok = ctx.session.register_media(buf.getvalue(), "image/png")
				src = f"/api/media/{tok}"
		except Exception:
			src = ""
		if not src:
			return
		style = "width:100%; height:auto;" if use_container_width else (f"width:{int(width)}px; height:auto;" if width else "max-width:100%; height:auto;")
		cap = f'<div class="revir-caption">{html_lib.escape(caption)}</div>' if caption else ""
		ctx.emit(f'<div class="revir-img-wrap"><img src="{src}" style="{style}" class="revir-img"/>{cap}</div>')

	def altair_chart(self, chart: Any, *, use_container_width: bool = True, **kwargs) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.altair_chart(chart, use_container_width=use_container_width)
			return
		_get_ctx().emit(_render_svg_altair_chart(chart))

	def bar_chart(self, data: Any, *, height: int = 220, color: str = "#0ea5e9", **kwargs) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.bar_chart(data, height=height, color=color)
			return
		_get_ctx().emit(_render_svg_bar_chart(data, height=height, color=color))

	# --- Etkileşimli Bileşenler ---
	def _register_widget_meta(self, wid: str, is_keyed: bool, meta: dict[str, Any], on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None) -> None:
		ctx = _get_ctx()
		meta["is_keyed"] = is_keyed
		ctx.session.widget_meta[wid] = meta
		if on_change is not None:
			ctx.session.callbacks[wid] = (on_change, tuple(args or ()), dict(kwargs or {}))
		if ctx.active_form_id is not None:
			ctx.used_form_keys.setdefault(ctx.active_form_id, []).append((wid, is_keyed))

	def button(self, label: str, *, key: str | None = None, help: str | None = None, on_click: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, type: str = "secondary", disabled: bool = False, use_container_width: bool = False) -> bool:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.button(label, key=key, help=help, on_click=on_click, args=args, kwargs=kwargs, type=type, disabled=disabled, use_container_width=use_container_width)
		ctx = _get_ctx()
		wid = str(key) if key is not None else ctx.next_auto_id("btn", label)
		if on_click is not None:
			ctx.session.callbacks[wid] = (on_click, tuple(args or ()), dict(kwargs or {}))
		btn_type = "primary" if type == "primary" else "secondary"
		cls = f"revir-btn revir-btn-{btn_type}" + (" revir-w-full" if use_container_width else "")
		dis_attr = " disabled" if disabled else ""
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		onclick = "" if disabled else f' onclick="window.__revirClick(\'{html_lib.escape(wid, quote=True)}\')"'
		ctx.emit(
			f'<div class="revir-widget-row">'
			f'<button type="button" class="{cls}" data-key="{html_lib.escape(wid, quote=True)}"{dis_attr}{title_attr}{onclick}>'
			f'{_format_inline_md(label, escape_html=True)}</button></div>'
		)
		return (not disabled) and (ctx.clicked_button == wid)

	def form_submit_button(self, label: str = "Kaydet", *, help: str | None = None, on_click: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, type: str = "secondary", disabled: bool = False, use_container_width: bool = False) -> bool:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.form_submit_button(label, help=help, on_click=on_click, args=args, kwargs=kwargs, type=type, disabled=disabled, use_container_width=use_container_width)
		ctx = _get_ctx()
		fid = ctx.active_form_id or "__no_form__"
		btn_type = "primary" if type == "primary" else "secondary"
		cls = f"revir-btn revir-btn-{btn_type}" + (" revir-w-full" if use_container_width else "")
		dis_attr = " disabled" if disabled else ""
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		onclick = "" if disabled else f' onclick="window.__revirSubmitForm(\'{html_lib.escape(fid, quote=True)}\')"'
		ctx.emit(
			f'<div class="revir-widget-row">'
			f'<button type="button" class="{cls}" data-submit-form="{html_lib.escape(fid, quote=True)}"{dis_attr}{title_attr}{onclick}>'
			f'{_format_inline_md(label, escape_html=True)}</button></div>'
		)
		submitted = (not disabled) and (ctx.submitted_form == fid)
		if submitted and on_click is not None:
			on_click(*tuple(args or ()), **dict(kwargs or {}))
		return submitted

	def download_button(self, label: str, data: Any, file_name: str = "dosya.bin", mime: str | None = None, *, key: str | None = None, help: str | None = None, type: str = "secondary", disabled: bool = False, use_container_width: bool = False) -> bool:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.download_button(label, data, file_name=file_name, mime=mime, key=key, help=help, type=type, disabled=disabled, use_container_width=use_container_width)
		ctx = _get_ctx()
		if isinstance(data, bytes):
			raw = data
		elif isinstance(data, bytearray):
			raw = bytes(data)
		elif isinstance(data, str):
			raw = data.encode("utf-8")
		elif hasattr(data, "getvalue"):
			v = data.getvalue()
			raw = v.encode("utf-8") if isinstance(v, str) else bytes(v)
		elif hasattr(data, "read"):
			v = data.read()
			raw = v.encode("utf-8") if isinstance(v, str) else bytes(v)
		else:
			raw = str(data).encode("utf-8")
		eff_mime = mime or mimetypes.guess_type(file_name)[0] or "application/octet-stream"
		token = ctx.session.register_download(raw, file_name, eff_mime)
		btn_type = "primary" if type == "primary" else "secondary"
		cls = f"revir-btn revir-btn-{btn_type}" + (" revir-w-full" if use_container_width else "")
		dis_attr = " disabled" if disabled else ""
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		safe_fn = json.dumps(file_name)
		onclick = "" if disabled else f" onclick='window.__revirDownload(\"{token}\", {html_lib.escape(safe_fn, quote=True)})'"
		ctx.emit(
			f'<div class="revir-widget-row">'
			f'<button type="button" class="{cls}"{dis_attr}{title_attr}{onclick}>'
			f'{_format_inline_md(label, escape_html=True)}</button></div>'
		)
		return False

	def link_button(self, label: str, url: str, *, help: str | None = None, type: str = "secondary", disabled: bool = False, use_container_width: bool = False) -> None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				_root.link_button(label, url, help=help, type=type, disabled=disabled, use_container_width=use_container_width)
			return
		btn_type = "primary" if type == "primary" else "secondary"
		cls = f"revir-btn revir-btn-{btn_type}" + (" revir-w-full" if use_container_width else "")
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		safe_url = json.dumps(str(url))
		onclick = "" if disabled else f" onclick='window.__revirOpenExternal({html_lib.escape(safe_url, quote=True)})'"
		_get_ctx().emit(
			f'<div class="revir-widget-row">'
			f'<button type="button" class="{cls}"{title_attr}{onclick}>'
			f'{_format_inline_md(label, escape_html=True)}</button></div>'
		)

	def text_input(self, label: str, value: str | None = "", *, max_chars: int | None = None, key: str | None = None, type: str = "default", help: str | None = None, placeholder: str | None = None, disabled: bool = False, label_visibility: str = "visible", on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None) -> str:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.text_input(label, value=value, max_chars=max_chars, key=key, type=type, help=help, placeholder=placeholder, disabled=disabled, label_visibility=label_visibility, on_change=on_change, args=args, kwargs=kwargs)
		ctx = _get_ctx()
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("txt", label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = "" if value is None else str(value)
			cur = ctx.session.session_state[wid]
			cur_str = "" if cur is None else str(cur)
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = "" if value is None else str(value)
			cur_str = str(ctx.session.widget_values[wid] or "")

		self._register_widget_meta(wid, is_keyed, {"kind": "text"}, on_change=on_change, args=args, kwargs=kwargs)
		inp_type = "password" if type == "password" else "text"
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		ph_attr = f' placeholder="{html_lib.escape(str(placeholder), quote=True)}"' if placeholder else ""
		dis_attr = " disabled" if disabled else ""
		max_attr = f' maxlength="{int(max_chars)}"' if max_chars else ""
		esc_wid = html_lib.escape(wid, quote=True)
		lbl = _label_html(label, label_visibility, help, html_for=f"inp_{wid}")
		ctx.emit(
			f'<div class="revir-field" data-widget="text_input">'
			f'{lbl}'
			f'<input id="inp_{esc_wid}" type="{inp_type}" class="revir-input" data-key="{esc_wid}" data-kind="text"{in_form_attr} '
			f'aria-label="{html_lib.escape(label, quote=True)}" value="{html_lib.escape(cur_str, quote=True)}"{ph_attr}{dis_attr}{max_attr} '
			f'oninput="window.__revirMarkDirty(\'{esc_wid}\', this.value)" '
			f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.value, this)" '
			f'onkeydown="if(event.key===\'Enter\'){{event.preventDefault();this.blur();}}"/>'
			f'</div>'
		)
		return cur_str

	def text_area(self, label: str, value: str | None = "", *, height: int | None = None, max_chars: int | None = None, key: str | None = None, help: str | None = None, placeholder: str | None = None, disabled: bool = False, label_visibility: str = "visible", on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None) -> str:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.text_area(label, value=value, height=height, max_chars=max_chars, key=key, help=help, placeholder=placeholder, disabled=disabled, label_visibility=label_visibility, on_change=on_change, args=args, kwargs=kwargs)
		ctx = _get_ctx()
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("txa", label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = "" if value is None else str(value)
			cur_str = "" if ctx.session.session_state[wid] is None else str(ctx.session.session_state[wid])
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = "" if value is None else str(value)
			cur_str = str(ctx.session.widget_values[wid] or "")

		self._register_widget_meta(wid, is_keyed, {"kind": "text"}, on_change=on_change, args=args, kwargs=kwargs)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		ph_attr = f' placeholder="{html_lib.escape(str(placeholder), quote=True)}"' if placeholder else ""
		dis_attr = " disabled" if disabled else ""
		h_style = f' style="min-height:{int(height)}px;"' if height else ""
		esc_wid = html_lib.escape(wid, quote=True)
		lbl = _label_html(label, label_visibility, help, html_for=f"inp_{wid}")
		ctx.emit(
			f'<div class="revir-field" data-widget="text_area">{lbl}'
			f'<textarea id="inp_{esc_wid}" class="revir-textarea" data-key="{esc_wid}" data-kind="text"{in_form_attr}{ph_attr}{dis_attr}{h_style} '
			f'oninput="window.__revirMarkDirty(\'{esc_wid}\', this.value)" '
			f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.value, this)">'
			f'{html_lib.escape(cur_str)}</textarea></div>'
		)
		return cur_str

	def number_input(self, label: str, *, min_value: int | float | None = None, max_value: int | float | None = None, value: Any = "min", step: int | float | None = None, format: str | None = None, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, placeholder: str | None = None, disabled: bool = False, label_visibility: str = "visible") -> int | float | None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.number_input(label, min_value=min_value, max_value=max_value, value=value, step=step, format=format, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, placeholder=placeholder, disabled=disabled, label_visibility=label_visibility)
		ctx = _get_ctx()
		is_float = any(isinstance(v, float) for v in (value, step, min_value, max_value) if v is not None and v != "min")
		if value == "min":
			default_val: int | float | None = min_value if min_value is not None else (0.0 if is_float else 0)
		elif value is None:
			default_val = None
		else:
			default_val = float(value) if is_float else int(value)

		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("num", label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = default_val
			cur_val = ctx.session.session_state[wid]
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = default_val
			cur_val = ctx.session.widget_values[wid]

		self._register_widget_meta(
			wid, is_keyed,
			{"kind": "number", "is_float": is_float, "min_value": min_value, "max_value": max_value, "allow_none": value is None},
			on_change=on_change, args=args, kwargs=kwargs,
		)
		disp = "" if cur_val is None else (f"{cur_val:g}" if isinstance(cur_val, float) else str(cur_val))
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		ph_attr = f' placeholder="{html_lib.escape(str(placeholder), quote=True)}"' if placeholder else ""
		dis_attr = " disabled" if disabled else ""
		esc_wid = html_lib.escape(wid, quote=True)
		lbl = _label_html(label, label_visibility, help, html_for=f"inp_{wid}")
		ctx.emit(
			f'<div class="revir-field" data-widget="number_input">{lbl}'
			f'<input id="inp_{esc_wid}" type="text" inputmode="decimal" class="revir-input" data-key="{esc_wid}" data-kind="number"{in_form_attr} '
			f'aria-label="{html_lib.escape(label, quote=True)}" value="{html_lib.escape(disp, quote=True)}"{ph_attr}{dis_attr} '
			f'oninput="window.__revirMarkDirty(\'{esc_wid}\', this.value)" '
			f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.value, this)" '
			f'onkeydown="if(event.key===\'Enter\'){{event.preventDefault();this.blur();}}"/></div>'
		)
		return cur_val

	def checkbox(self, label: str, value: bool = False, *, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, disabled: bool = False, label_visibility: str = "visible") -> bool:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.checkbox(label, value=value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled, label_visibility=label_visibility)
		return self._render_bool_widget("checkbox", label, value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled)

	def toggle(self, label: str, value: bool = False, *, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, disabled: bool = False, label_visibility: str = "visible") -> bool:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.toggle(label, value=value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled, label_visibility=label_visibility)
		return self._render_bool_widget("toggle", label, value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled)

	def _render_bool_widget(self, style_kind: str, label: str, value: bool, *, key: str | None, help: str | None, on_change: Callable | None, args: tuple | None, kwargs: dict | None, disabled: bool) -> bool:
		ctx = _get_ctx()
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id(style_kind, label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = bool(value)
			cur = bool(ctx.session.session_state[wid])
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = bool(value)
			cur = bool(ctx.session.widget_values[wid])

		self._register_widget_meta(wid, is_keyed, {"kind": "bool"}, on_change=on_change, args=args, kwargs=kwargs)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		chk_attr = " checked" if cur else ""
		dis_attr = " disabled" if disabled else ""
		title_attr = f' title="{html_lib.escape(help, quote=True)}"' if help else ""
		esc_wid = html_lib.escape(wid, quote=True)
		if style_kind == "toggle":
			ctx.emit(
				f'<label class="revir-toggle-wrap"{title_attr}>'
				f'<input type="checkbox" class="revir-toggle-input" data-key="{esc_wid}" data-kind="bool"{in_form_attr}{chk_attr}{dis_attr} '
				f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.checked, this)"/>'
				f'<span class="revir-toggle-track"><span class="revir-toggle-thumb"></span></span>'
				f'<span class="revir-check-text">{_format_inline_md(label, escape_html=True)}</span></label>'
			)
		else:
			ctx.emit(
				f'<label class="revir-check-wrap"{title_attr}>'
				f'<input type="checkbox" class="revir-check-input" data-key="{esc_wid}" data-kind="bool"{in_form_attr}{chk_attr}{dis_attr} '
				f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.checked, this)"/>'
				f'<span class="revir-check-text">{_format_inline_md(label, escape_html=True)}</span></label>'
			)
		return cur

	def selectbox(self, label: str, options: Iterable[Any], index: int | None = 0, *, format_func: Callable[[Any], str] = str, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, placeholder: str = "Seçiniz...", disabled: bool = False, label_visibility: str = "visible") -> Any:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.selectbox(label, options, index=index, format_func=format_func, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, placeholder=placeholder, disabled=disabled, label_visibility=label_visibility)
		ctx = _get_ctx()
		opts = list(options)
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("sel", label)
		default_val = opts[index] if (index is not None and 0 <= index < len(opts)) else None

		if is_keyed:
			if wid in ctx.session.session_state:
				cur = ctx.session.session_state[wid]
				if cur is not None and cur not in opts:
					cur = default_val
					ctx.session.session_state[wid] = cur
			else:
				cur = default_val
				ctx.session.session_state[wid] = cur
		else:
			if wid in ctx.session.widget_values:
				cur = ctx.session.widget_values[wid]
				if cur is not None and cur not in opts:
					cur = default_val
			else:
				cur = default_val
			ctx.session.widget_values[wid] = cur

		sel_idx: int | None = None
		if cur is not None:
			for i, o in enumerate(opts):
				if o == cur:
					sel_idx = i
					break

		labels_list = [str(format_func(o)) for o in opts]
		self._register_widget_meta(wid, is_keyed, {"kind": "selectbox", "options": opts}, on_change=on_change, args=args, kwargs=kwargs)

		esc_wid = html_lib.escape(wid, quote=True)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		lbl = _label_html(label, label_visibility, help)
		cur_display = labels_list[sel_idx] if (sel_idx is not None and sel_idx < len(labels_list)) else ""
		ph = placeholder or "Seçiniz..."
		nullable = (index is None) or (cur is None)
		opts_json = html_lib.escape(json.dumps(labels_list, ensure_ascii=False), quote=True)
		cur_idx_str = "" if sel_idx is None else str(sel_idx)
		dis_cls = " revir-disabled" if disabled else ""
		clear_btn = (
			f'<button type="button" class="revir-combo-clear" title="Temizle" onclick="window.__revirSelectClear(\'{esc_wid}\', event)">×</button>'
			if (nullable and sel_idx is not None and not disabled) else ""
		)
		ctx.emit(
			f'<div class="revir-field" data-widget="selectbox">{lbl}'
			f'<div class="revir-combo{dis_cls}" data-key="{esc_wid}" data-kind="selectbox" data-value="{cur_idx_str}" '
			f'data-options="{opts_json}" data-placeholder="{html_lib.escape(ph, quote=True)}"{in_form_attr}>'
			f'<div class="revir-combo-control" onclick="window.__revirOpenCombo(\'{esc_wid}\', event)">'
			f'<input type="text" class="revir-combo-search" value="{html_lib.escape(cur_display, quote=True)}" '
			f'placeholder="{html_lib.escape(ph, quote=True)}" {"disabled" if disabled else ""} '
			f'onfocus="this.select();window.__revirOpenCombo(\'{esc_wid}\', event)" '
			f'onblur="window.__revirBlurCombo(\'{esc_wid}\', this)" '
			f'oninput="window.__revirFilterCombo(\'{esc_wid}\', this.value)" autocomplete="off"/>'
			f'{clear_btn}<span class="revir-combo-arrow">▾</span></div>'
			f'<div class="revir-combo-menu" id="combo_{esc_wid}"></div>'
			f'</div></div>'
		)
		return cur

	def multiselect(self, label: str, options: Iterable[Any], default: Any = None, *, format_func: Callable[[Any], str] = str, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, placeholder: str = "Seçim yapın...", disabled: bool = False, label_visibility: str = "visible") -> list[Any]:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.multiselect(label, options, default=default, format_func=format_func, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, placeholder=placeholder, disabled=disabled, label_visibility=label_visibility)
		ctx = _get_ctx()
		opts = list(options)
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("msel", label)

		def _norm_defaults(val: Any) -> list[Any]:
			if val is None:
				return []
			if isinstance(val, (list, tuple, set)):
				return [x for x in val if x in opts]
			return [val] if val in opts else []

		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = _norm_defaults(default)
			else:
				ctx.session.session_state[wid] = _norm_defaults(ctx.session.session_state[wid])
			cur = list(ctx.session.session_state[wid])
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = _norm_defaults(default)
			else:
				ctx.session.widget_values[wid] = _norm_defaults(ctx.session.widget_values[wid])
			cur = list(ctx.session.widget_values[wid])

		sel_indices: list[int] = []
		for v in cur:
			for i, o in enumerate(opts):
				if o == v and i not in sel_indices:
					sel_indices.append(i)
					break

		labels_list = [str(format_func(o)) for o in opts]
		self._register_widget_meta(wid, is_keyed, {"kind": "multiselect", "options": opts}, on_change=on_change, args=args, kwargs=kwargs)

		esc_wid = html_lib.escape(wid, quote=True)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		lbl = _label_html(label, label_visibility, help)
		opts_json = html_lib.escape(json.dumps(labels_list, ensure_ascii=False), quote=True)
		sel_json = html_lib.escape(json.dumps(sel_indices), quote=True)
		chips = "".join(
			f'<span class="revir-chip">{html_lib.escape(labels_list[i])}'
			f'<button type="button" class="revir-chip-x" onclick="window.__revirMultiRemove(\'{esc_wid}\', {i}, event)">×</button></span>'
			for i in sel_indices
		)
		ph = "" if sel_indices else (placeholder or "Seçim yapın...")
		ctx.emit(
			f'<div class="revir-field" data-widget="multiselect">{lbl}'
			f'<div class="revir-multi" data-key="{esc_wid}" data-kind="multiselect" data-selected="{sel_json}" data-options="{opts_json}"{in_form_attr}>'
			f'<div class="revir-multi-control" onclick="window.__revirOpenMulti(\'{esc_wid}\', event)">'
			f'{chips}<input type="text" class="revir-multi-search" placeholder="{html_lib.escape(ph, quote=True)}" {"disabled" if disabled else ""} '
			f'onfocus="window.__revirOpenMulti(\'{esc_wid}\', event)" oninput="window.__revirFilterMulti(\'{esc_wid}\', this.value)" autocomplete="off"/>'
			f'</div><div class="revir-combo-menu" id="multi_{esc_wid}"></div></div></div>'
		)
		return cur

	def radio(self, label: str, options: Iterable[Any], index: int | None = 0, *, format_func: Callable[[Any], str] = str, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, disabled: bool = False, horizontal: bool = False, label_visibility: str = "visible") -> Any:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.radio(label, options, index=index, format_func=format_func, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled, horizontal=horizontal, label_visibility=label_visibility)
		ctx = _get_ctx()
		opts = list(options)
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("rad", label)
		default_val = opts[index] if (index is not None and 0 <= index < len(opts)) else (opts[0] if opts else None)

		if is_keyed:
			if wid in ctx.session.session_state and ctx.session.session_state[wid] in opts:
				cur = ctx.session.session_state[wid]
			else:
				cur = default_val
				ctx.session.session_state[wid] = cur
		else:
			if wid in ctx.session.widget_values and ctx.session.widget_values[wid] in opts:
				cur = ctx.session.widget_values[wid]
			else:
				cur = default_val
				ctx.session.widget_values[wid] = cur

		sel_idx = next((i for i, o in enumerate(opts) if o == cur), 0 if opts else -1)
		self._register_widget_meta(wid, is_keyed, {"kind": "radio", "options": opts}, on_change=on_change, args=args, kwargs=kwargs)

		esc_wid = html_lib.escape(wid, quote=True)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		lbl = _label_html(label, label_visibility, help)
		dis_attr = " disabled" if disabled else ""

		if ctx.in_sidebar and not horizontal:
			items = []
			for i, o in enumerate(opts):
				active_cls = " active" if i == sel_idx else ""
				txt = html_lib.escape(str(format_func(o)))
				items.append(
					f'<button type="button" class="revir-nav-item{active_cls}"{dis_attr} '
					f'onclick="window.__revirRadioPick(\'{esc_wid}\', {i}, this)">{txt}</button>'
				)
			ctx.emit(f'<div class="revir-nav-group" data-key="{esc_wid}" data-kind="radio" data-value="{sel_idx}"{in_form_attr}>{lbl}{"".join(items)}</div>')
		else:
			grp_cls = "revir-radio-group revir-radio-horiz" if horizontal else "revir-radio-group"
			items = []
			for i, o in enumerate(opts):
				chk = " checked" if i == sel_idx else ""
				txt = html_lib.escape(str(format_func(o)))
				items.append(
					f'<label class="revir-radio-opt">'
					f'<input type="radio" name="rad_{esc_wid}" value="{i}"{chk}{dis_attr} '
					f'onchange="window.__revirRadioPick(\'{esc_wid}\', {i}, this)"/>'
					f'<span>{txt}</span></label>'
				)
			ctx.emit(f'<div class="revir-field" data-widget="radio">{lbl}<div class="{grp_cls}" data-key="{esc_wid}" data-kind="radio" data-value="{sel_idx}"{in_form_attr}>{"".join(items)}</div></div>')
		return cur

	def date_input(self, label: str, value: Any = "today", *, min_value: dt_mod.date | None = None, max_value: dt_mod.date | None = None, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, disabled: bool = False, label_visibility: str = "visible") -> dt_mod.date | None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.date_input(label, value=value, min_value=min_value, max_value=max_value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled, label_visibility=label_visibility)
		ctx = _get_ctx()

		def _to_date(v: Any) -> dt_mod.date | None:
			if v is None or v == "":
				return None
			if v == "today":
				return dt_mod.date.today()
			if isinstance(v, dt_mod.datetime):
				return v.date()
			if isinstance(v, dt_mod.date):
				return v
			try:
				return dt_mod.date.fromisoformat(str(v)[:10])
			except Exception:
				return None

		default_d = _to_date(value)
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("date", label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = default_d
			cur = _to_date(ctx.session.session_state[wid])
			ctx.session.session_state[wid] = cur
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = default_d
			cur = _to_date(ctx.session.widget_values[wid])
			ctx.session.widget_values[wid] = cur

		self._register_widget_meta(wid, is_keyed, {"kind": "date"}, on_change=on_change, args=args, kwargs=kwargs)
		esc_wid = html_lib.escape(wid, quote=True)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		min_attr = f' min="{min_value.isoformat()}"' if isinstance(min_value, dt_mod.date) else ""
		max_attr = f' max="{max_value.isoformat()}"' if isinstance(max_value, dt_mod.date) else ""
		dis_attr = " disabled" if disabled else ""
		val_str = cur.isoformat() if isinstance(cur, dt_mod.date) else ""
		lbl = _label_html(label, label_visibility, help, html_for=f"inp_{wid}")
		ctx.emit(
			f'<div class="revir-field" data-widget="date_input">{lbl}'
			f'<input id="inp_{esc_wid}" type="date" class="revir-input" data-key="{esc_wid}" data-kind="date"{in_form_attr} '
			f'value="{val_str}"{min_attr}{max_attr}{dis_attr} '
			f'onchange="window.__revirFieldChange(\'{esc_wid}\', this.value, this)"/></div>'
		)
		return cur

	def time_input(self, label: str, value: Any = "now", *, key: str | None = None, help: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None, disabled: bool = False, label_visibility: str = "visible", step: Any = None) -> dt_mod.time | None:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.time_input(label, value=value, key=key, help=help, on_change=on_change, args=args, kwargs=kwargs, disabled=disabled, label_visibility=label_visibility, step=step)
		ctx = _get_ctx()

		def _to_time(v: Any) -> dt_mod.time | None:
			if v is None or v == "":
				return None
			if v == "now":
				now = dt_mod.datetime.now()
				return dt_mod.time(now.hour, now.minute)
			if isinstance(v, dt_mod.datetime):
				return dt_mod.time(v.hour, v.minute)
			if isinstance(v, dt_mod.time):
				return v
			try:
				parts = str(v).strip().split(":")
				return dt_mod.time(int(parts[0]), int(parts[1]))
			except Exception:
				return None

		default_t = _to_time(value)
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("time", label)
		if is_keyed:
			if wid not in ctx.session.session_state:
				ctx.session.session_state[wid] = default_t
			cur = _to_time(ctx.session.session_state[wid])
			ctx.session.session_state[wid] = cur
		else:
			if wid not in ctx.session.widget_values:
				ctx.session.widget_values[wid] = default_t
			cur = _to_time(ctx.session.widget_values[wid])
			ctx.session.widget_values[wid] = cur

		self._register_widget_meta(wid, is_keyed, {"kind": "time"}, on_change=on_change, args=args, kwargs=kwargs)
		esc_wid = html_lib.escape(wid, quote=True)
		in_form_attr = f' data-form="{html_lib.escape(ctx.active_form_id, quote=True)}"' if ctx.active_form_id else ""
		dis_attr = " disabled" if disabled else ""
		val_str = f"{cur.hour:02d}:{cur.minute:02d}" if isinstance(cur, dt_mod.time) else ""
		lbl = _label_html(label, label_visibility, help, html_for=f"inp_{wid}")
		ctx.emit(
			f'<div class="revir-field" data-widget="time_input">{lbl}'
			f'<input id="inp_{esc_wid}" type="time" class="revir-input" data-key="{esc_wid}" data-kind="time"{in_form_attr} '
			f'value="{val_str}"{dis_attr} onchange="window.__revirFieldChange(\'{esc_wid}\', this.value, this)"/></div>'
		)
		return cur

	def file_uploader(self, label: str, *, type: str | Sequence[str] | None = None, accept_multiple_files: bool = False, key: str | None = None, help: str | None = None, on_change: Callable | None = None, disabled: bool = False, label_visibility: str = "visible") -> Any:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.file_uploader(label, type=type, accept_multiple_files=accept_multiple_files, key=key, help=help, on_change=on_change, disabled=disabled, label_visibility=label_visibility)
		ctx = _get_ctx()
		is_keyed = key is not None
		wid = str(key) if is_keyed else ctx.next_auto_id("upl", label)
		self._register_widget_meta(wid, is_keyed, {"kind": "upload", "multiple": bool(accept_multiple_files)}, on_change=on_change)

		files = ctx.session.uploaded_files.get(wid) or []
		for f in files:
			try:
				f.seek(0)
			except Exception:
				pass

		accept_str = ""
		if type:
			exts = [type] if isinstance(type, str) else list(type)
			accept_str = ",".join(e if e.startswith(".") else f".{e}" for e in exts)
		accept_attr = f' accept="{html_lib.escape(accept_str, quote=True)}"' if accept_str else ""
		mult_attr = " multiple" if accept_multiple_files else ""
		dis_attr = " disabled" if disabled else ""
		esc_wid = html_lib.escape(wid, quote=True)
		lbl = _label_html(label, label_visibility, help)

		file_badges = ""
		if files:
			items = "".join(
				f'<div class="revir-uploaded-item"><span>📎 {html_lib.escape(f.name)} <small>({f.size / 1024:.1f} KB)</small></span>'
				f'<button type="button" class="revir-chip-x" onclick="window.__revirClearUpload(\'{esc_wid}\', \'{html_lib.escape(f.file_id, quote=True)}\')">×</button></div>'
				for f in files
			)
			file_badges = f'<div class="revir-uploaded-list">{items}</div>'

		ctx.emit(
			f'<div class="revir-field" data-widget="file_uploader">{lbl}'
			f'<div class="revir-dropzone" ondragover="event.preventDefault();this.classList.add(\'dragover\')" '
			f'ondragleave="this.classList.remove(\'dragover\')" '
			f'ondrop="window.__revirHandleDrop(event, \'{esc_wid}\', {str(bool(accept_multiple_files)).lower()})">'
			f'<div class="revir-dropzone-text">📁 Dosyayı buraya sürükleyin veya seçin'
			f'{f" <small>({html_lib.escape(accept_str)})</small>" if accept_str else ""}</div>'
			f'<label class="revir-btn revir-btn-secondary revir-upload-btn">Dosya Seç'
			f'<input type="file" style="display:none"{accept_attr}{mult_attr}{dis_attr} '
			f'onchange="window.__revirUploadFiles(this.files, \'{esc_wid}\', {str(bool(accept_multiple_files)).lower()})"/></label>'
			f'</div>{file_badges}</div>'
		)
		if accept_multiple_files:
			return list(files)
		return files[0] if files else None

	def dataframe(self, data: Any, *, width: Any = None, height: int | None = None, use_container_width: bool = False, hide_index: bool | None = None, column_order: Sequence[str] | None = None, column_config: dict[str, Any] | None = None, key: str | None = None, on_select: str = "ignore", selection_mode: Any = "multi-row") -> Any:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.dataframe(data, width=width, height=height, use_container_width=use_container_width, hide_index=hide_index, column_order=column_order, column_config=column_config, key=key, on_select=on_select, selection_mode=selection_mode)
		ctx = _get_ctx()
		if hasattr(data, "data") and isinstance(data.data, pd.DataFrame):
			df = data.data
		elif isinstance(data, pd.DataFrame):
			df = data
		elif isinstance(data, pd.Series):
			df = data.to_frame()
		elif data is None:
			df = pd.DataFrame()
		else:
			df = pd.DataFrame(data)

		df_key = str(key) if key is not None else ctx.next_auto_id("df", "")
		cfg_map = column_config or {}
		all_cols = list(column_order) if column_order else list(df.columns)
		visible_cols = [c for c in all_cols if c in df.columns and not (c in cfg_map and cfg_map[c] is None)]
		show_idx = not bool(hide_index) if hide_index is not None else False
		selectable = (on_select == "rerun")

		sel_state = ctx.session.dataframe_selections.get(df_key, {"rows": [], "cells": []}) if selectable else {"rows": [], "cells": []}
		sel_rows_set = set(sel_state.get("rows") or [])

		h_style = f'max-height:{int(height)}px;' if height else "max-height:460px;"
		esc_key = html_lib.escape(df_key, quote=True)

		th_parts: list[str] = []
		if show_idx:
			th_parts.append("<th>#</th>")
		for c in visible_cols:
			cc = cfg_map.get(c)
			lbl = cc.label if isinstance(cc, _ColumnConfig) and cc.label else str(c)
			hlp = f' title="{html_lib.escape(cc.help, quote=True)}"' if isinstance(cc, _ColumnConfig) and cc.help else ""
			th_parts.append(f'<th{hlp} onclick="window.__revirSortTable(this)">{html_lib.escape(lbl)} <span class="revir-sort-ind">⇅</span></th>')

		row_parts: list[str] = []
		for r_pos in range(len(df)):
			row_sel_cls = " revir-row-selected" if r_pos in sel_rows_set else ""
			tds: list[str] = []
			if show_idx:
				tds.append(f'<td class="revir-idx-cell">{html_lib.escape(str(df.index[r_pos]))}</td>')
			for c in visible_cols:
				val = df.iloc[r_pos][c]
				cc = cfg_map.get(c)
				txt = _format_df_cell(val, cc if isinstance(cc, _ColumnConfig) else None)
				is_dosya = str(c) in ("📂 Dosya", "Dosya", "📂")
				if selectable:
					safe_col = json.dumps(str(c))
					click_attr = f" onclick='window.__revirDfSelect(\"{esc_key}\", {r_pos}, {html_lib.escape(safe_col, quote=True)})'"
					cell_cls = " revir-cell-clickable revir-cell-action" if is_dosya else " revir-cell-clickable"
					if is_dosya:
						tds.append(f'<td class="{cell_cls}"{click_attr}><span class="revir-dosya-pill">{html_lib.escape(txt or "📂 Aç")}</span></td>')
					else:
						tds.append(f'<td class="{cell_cls}"{click_attr}>{html_lib.escape(txt)}</td>')
				else:
					tds.append(f"<td>{html_lib.escape(txt)}</td>")
			row_parts.append(f'<tr class="{row_sel_cls}">{"".join(tds)}</tr>')

		if not row_parts:
			colspan = max(1, len(visible_cols) + (1 if show_idx else 0))
			row_parts.append(f'<tr><td colspan="{colspan}" class="revir-empty-td">Kayıt bulunamadı.</td></tr>')

		ctx.emit(
			f'<div class="revir-table-wrap" style="{h_style}">'
			f'<table class="revir-table{" revir-table-selectable" if selectable else ""}">'
			f'<thead><tr>{"".join(th_parts)}</tr></thead>'
			f'<tbody>{"".join(row_parts)}</tbody></table></div>'
		)

		if selectable:
			return _DataframeSelectionEvent(
				selection=_SelectionState(
					rows=list(sel_state.get("rows") or []),
					cells=[tuple(c) for c in (sel_state.get("cells") or [])],
				)
			)
		return None

	def data_editor(self, data: Any, *, width: Any = None, height: int | None = None, use_container_width: bool = False, hide_index: bool | None = None, column_order: Sequence[str] | None = None, column_config: dict[str, Any] | None = None, num_rows: str = "fixed", disabled: bool | Sequence[str] = False, key: str | None = None, on_change: Callable | None = None, args: tuple | None = None, kwargs: dict | None = None) -> pd.DataFrame:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.data_editor(data, width=width, height=height, use_container_width=use_container_width, hide_index=hide_index, column_order=column_order, column_config=column_config, num_rows=num_rows, disabled=disabled, key=key, on_change=on_change, args=args, kwargs=kwargs)
		ctx = _get_ctx()
		base_df = data.copy() if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
		is_keyed = key is not None
		ed_key = str(key) if is_keyed else ctx.next_auto_id("ed", "")
		cfg_map = column_config or {}

		# Girdi DataFrame değiştiyse bayat düzenleme farklarını sıfırla (sonsuz silme döngüsünü önler)
		fp = _df_fingerprint(base_df)
		prev_fp = ctx.session.editor_input_hashes.get(ed_key)
		if prev_fp != fp:
			ctx.session.editor_input_hashes[ed_key] = fp
			ctx.session.session_state[ed_key] = {"edited_rows": {}, "added_rows": [], "deleted_rows": []}

		ed_state = ctx.session.session_state.get(ed_key)
		if not isinstance(ed_state, dict) or "edited_rows" not in ed_state:
			ed_state = {"edited_rows": {}, "added_rows": [], "deleted_rows": []}
			ctx.session.session_state[ed_key] = ed_state

		# Düzenlemeleri kopya DataFrame'e uygula
		out_df = base_df.copy()
		edited_rows: dict = ed_state.get("edited_rows") or {}
		deleted_rows: list[int] = [int(x) for x in (ed_state.get("deleted_rows") or []) if str(x).isdigit()]
		added_rows: list[dict] = list(ed_state.get("added_rows") or [])

		for r_key, col_changes in edited_rows.items():
			try:
				r_pos = int(r_key)
			except Exception:
				continue
			if 0 <= r_pos < len(out_df) and isinstance(col_changes, dict):
				idx_lbl = out_df.index[r_pos]
				for col_name, raw_v in col_changes.items():
					if col_name not in out_df.columns:
						continue
					cc = cfg_map.get(col_name)
					coerced = _coerce_editor_value(raw_v, cc if isinstance(cc, _ColumnConfig) else None, out_df[col_name].dtype)
					if coerced is None and pd.api.types.is_integer_dtype(out_df[col_name].dtype):
						out_df[col_name] = out_df[col_name].astype("float64")
					elif isinstance(coerced, str) and not pd.api.types.is_object_dtype(out_df[col_name].dtype) and not pd.api.types.is_string_dtype(out_df[col_name].dtype):
						out_df[col_name] = out_df[col_name].astype("object")
					elif isinstance(coerced, float) and pd.api.types.is_integer_dtype(out_df[col_name].dtype):
						out_df[col_name] = out_df[col_name].astype("float64")
					out_df.at[idx_lbl, col_name] = coerced

		if deleted_rows:
			keep_mask = [i not in set(deleted_rows) for i in range(len(out_df))]
			out_df = out_df.iloc[keep_mask].copy()

		if added_rows:
			norm_added = []
			for a_row in added_rows:
				rec = {}
				for c in base_df.columns:
					cc = cfg_map.get(c)
					if c in a_row:
						rec[c] = _coerce_editor_value(a_row[c], cc if isinstance(cc, _ColumnConfig) else None, base_df[c].dtype)
					elif isinstance(cc, _ColumnConfig) and cc.kind == "checkbox":
						rec[c] = False
					else:
						rec[c] = None
				norm_added.append(rec)
			add_df = pd.DataFrame(norm_added, columns=base_df.columns)
			out_df = pd.concat([out_df, add_df], ignore_index=True)

		# Görünür sütunları ve salt-okunur sütunları belirle
		all_cols = list(column_order) if column_order else list(base_df.columns)
		visible_cols = [c for c in all_cols if c in base_df.columns and not (c in cfg_map and cfg_map[c] is None)]
		disabled_all = (disabled is True)
		disabled_set = set(disabled) if isinstance(disabled, (list, tuple, set)) else set()
		is_dynamic = (num_rows == "dynamic")

		ctx.session.editor_meta[ed_key] = {
			"columns": list(base_df.columns),
			"base_len": len(base_df),
			"is_dynamic": is_dynamic,
		}
		if on_change is not None:
			ctx.session.callbacks[ed_key] = (on_change, tuple(args or ()), dict(kwargs or {}))

		# HTML Düzenlenebilir Tablo
		esc_key = html_lib.escape(ed_key, quote=True)
		h_style = f"max-height:{int(height)}px;" if height else "max-height:480px;"
		th_parts = []
		for c in visible_cols:
			cc = cfg_map.get(c)
			lbl = cc.label if isinstance(cc, _ColumnConfig) and cc.label else str(c)
			hlp = f' title="{html_lib.escape(cc.help, quote=True)}"' if isinstance(cc, _ColumnConfig) and cc.help else ""
			th_parts.append(f"<th{hlp}>{html_lib.escape(lbl)}</th>")
		if is_dynamic:
			th_parts.append('<th class="revir-ed-act-th">Sil</th>')

		# Görüntülenecek satırları (base + added, deleted hariç) hazırla
		render_rows: list[tuple[str, int, dict[str, Any]]] = []
		del_set = set(deleted_rows)
		for r_pos in range(len(base_df)):
			if r_pos in del_set:
				continue
			row_vals = {}
			edits_for_row = edited_rows.get(r_pos) or edited_rows.get(str(r_pos)) or {}
			for c in visible_cols:
				row_vals[c] = edits_for_row[c] if c in edits_for_row else base_df.iloc[r_pos][c]
			render_rows.append(("base", r_pos, row_vals))
		for a_idx, a_row in enumerate(added_rows):
			row_vals = {c: a_row.get(c, None) for c in visible_cols}
			render_rows.append(("added", a_idx, row_vals))

		tr_parts: list[str] = []
		for row_kind, row_id, row_vals in render_rows:
			tds: list[str] = []
			for c in visible_cols:
				cc = cfg_map.get(c)
				col_dis = disabled_all or (c in disabled_set) or (isinstance(cc, _ColumnConfig) and bool(cc.disabled))
				val = row_vals.get(c)
				safe_col = html_lib.escape(json.dumps(str(c)), quote=True)

				if col_dis:
					disp_txt = _format_df_cell(val, cc if isinstance(cc, _ColumnConfig) else None)
					tds.append(f'<td class="revir-ed-ro">{html_lib.escape(disp_txt)}</td>')
					continue

				cell_id = f"ed_{esc_key}_{row_kind}_{row_id}_{abs(hash(c)) % 99999}"
				if isinstance(cc, _ColumnConfig) and cc.kind == "checkbox" or (cc is None and pd.api.types.is_bool_dtype(base_df[c].dtype)):
					chk = " checked" if (val is True or str(val).lower() == "true") else ""
					tds.append(
						f'<td class="revir-ed-cell revir-ed-center">'
						f'<input id="{cell_id}" type="checkbox"{chk} '
						f"onchange='window.__revirEditorCell(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.checked, \"{cell_id}\")'/></td>"
					)
				elif isinstance(cc, _ColumnConfig) and cc.kind == "selectbox":
					cur_s = "" if (val is None or (isinstance(val, float) and math.isnan(val))) else str(val)
					opts_list = list(cc.options or [])
					opt_html = ['<option value=""></option>'] if not cc.required else []
					seen_o = set()
					for opt_v in opts_list:
						os_val = str(opt_v)
						seen_o.add(os_val)
						sel_a = " selected" if os_val == cur_s else ""
						opt_html.append(f'<option value="{html_lib.escape(os_val, quote=True)}"{sel_a}>{html_lib.escape(os_val)}</option>')
					if cur_s and cur_s not in seen_o:
						opt_html.append(f'<option value="{html_lib.escape(cur_s, quote=True)}" selected>{html_lib.escape(cur_s)}</option>')
					tds.append(
						f'<td class="revir-ed-cell">'
						f'<select id="{cell_id}" class="revir-ed-select" '
						f"onchange='window.__revirEditorCell(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.value, \"{cell_id}\")'>"
						f'{"".join(opt_html)}</select></td>'
					)
				else:
					is_num_col = isinstance(cc, _ColumnConfig) and cc.kind == "number"
					disp_v = _format_df_cell(val, None if is_num_col else (cc if isinstance(cc, _ColumnConfig) else None))
					imode = ' inputmode="decimal"' if is_num_col else ""
					tds.append(
						f'<td class="revir-ed-cell">'
						f'<input id="{cell_id}" type="text"{imode} class="revir-ed-input" '
						f'value="{html_lib.escape(disp_v, quote=True)}" '
						f"oninput='window.__revirEditorDirty(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.value)' "
						f"onchange='window.__revirEditorCell(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.value, \"{cell_id}\")' "
						f'onkeydown="if(event.key===\'Enter\'){{event.preventDefault();this.blur();}}"/></td>'
					)

			if is_dynamic:
				tds.append(
					f'<td class="revir-ed-cell revir-ed-center">'
					f'<button type="button" class="revir-ed-del-btn" title="Satırı Sil" '
					f'onclick="window.__revirEditorDeleteRow(\'{esc_key}\', \'{row_kind}\', {row_id})">🗑️</button></td>'
				)
			tr_parts.append(f"<tr>{''.join(tds)}</tr>")

		add_footer = (
			f'<div class="revir-ed-footer">'
			f'<button type="button" class="revir-btn revir-btn-secondary revir-btn-sm" '
			f'onclick="window.__revirEditorAddRow(\'{esc_key}\')">➕ Yeni Satır Ekle</button></div>'
			if is_dynamic else ""
		)
		ctx.emit(
			f'<div class="revir-editor-box" data-editor-key="{esc_key}">'
			f'<div class="revir-table-wrap" style="{h_style}">'
			f'<table class="revir-table revir-editor-table">'
			f'<thead><tr>{"".join(th_parts)}</tr></thead>'
			f'<tbody>{"".join(tr_parts)}</tbody></table></div>{add_footer}</div>'
		)
		return out_df


class _LazyColumnsNode:
	def __init__(self, col_lists: list[list[str]], grid_tpl: str, gap_px: str, align_css: str) -> None:
		self.col_lists = col_lists
		self.grid_tpl = grid_tpl
		self.gap_px = gap_px
		self.align_css = align_css

	def render(self) -> str:
		cols_html = "".join(f'<div class="revir-col">{_flatten_nodes(cl)}</div>' for cl in self.col_lists)
		return f'<div class="revir-cols" style="grid-template-columns:{self.grid_tpl}; gap:{self.gap_px}; align-items:{self.align_css};">{cols_html}</div>'


class _LazyTabsNode:
	def __init__(self, tab_id: str, labels: list[str], active_idx: int, pane_lists: list[list[str]]) -> None:
		self.tab_id = tab_id
		self.labels = labels
		self.active_idx = active_idx
		self.pane_lists = pane_lists

	def render(self) -> str:
		btns = []
		panes = []
		for i, (lbl, pl) in enumerate(zip(self.labels, self.pane_lists)):
			act = " active" if i == self.active_idx else ""
			btns.append(
				f'<button type="button" class="revir-tab-btn{act}" data-tab-group="{self.tab_id}" data-tab-idx="{i}" '
				f'onclick="window.__revirSwitchTab(\'{self.tab_id}\', {i})">{html_lib.escape(lbl)}</button>'
			)
			panes.append(f'<div class="revir-tab-pane{act}" data-tab-group="{self.tab_id}" data-tab-idx="{i}">{_flatten_nodes(pl)}</div>')
		return f'<div class="revir-tabs"><div class="revir-tab-bar">{"".join(btns)}</div><div class="revir-tab-content">{"".join(panes)}</div></div>'


def _flatten_nodes(nodes: list[Any]) -> str:
	parts: list[str] = []
	for item in nodes:
		if isinstance(item, str):
			parts.append(item)
		elif hasattr(item, "render"):
			parts.append(item.render())
	return "".join(parts)


# ==============================================================================
# 7. GÖMÜLÜ HTML BİLEŞENİ (components.v1.html) VE DİYALOG DEKORATÖRÜ
# ==============================================================================

class _ComponentsV1:
	@staticmethod
	def html(html_str: str, *, height: int = 150, width: int | None = None, scrolling: bool = False) -> None:
		ctx = _get_ctx()
		raw = str(html_str or "")
		if height == 0:
			# Yardımcı betikler doğrudan belgeye eklenir
			ctx.scripts_parts.append(raw)
			return
		w_style = f"width:{int(width)}px;" if width else "width:100%;"
		scr_attr = "yes" if scrolling else "no"
		srcdoc = html_lib.escape(raw, quote=True)
		ctx.emit(
			f'<div class="revir-iframe-wrap">'
			f'<iframe class="revir-iframe" style="{w_style} height:{int(height)}px;" scrolling="{scr_attr}" srcdoc="{srcdoc}"></iframe>'
			f'</div>'
		)


class _ComponentsNamespace:
	v1 = _ComponentsV1()


components = _ComponentsNamespace()


def dialog(title: str, *, width: str = "small", on_dismiss: Callable | None = None):
	def _decorator(func: Callable) -> Callable:
		@functools.wraps(func)
		def _wrapper(*args, **kwargs):
			ctx = _get_ctx()
			ctx.session.active_dialog = _ActiveDialog(
				title=title,
				width=width,
				on_dismiss=on_dismiss,
				func=func,
				args=args,
				kwargs=kwargs,
			)
			ctx.dialog_rendered_this_pass = True
			ctx.dialog_nodes = []
			prev_dlg = ctx.in_dialog
			prev_sb = ctx.in_sidebar
			prev_form = ctx.active_form_id
			ctx.in_dialog = True
			ctx.in_sidebar = False
			ctx.active_form_id = None
			ctx.node_stack.append(ctx.dialog_nodes)
			try:
				return func(*args, **kwargs)
			finally:
				ctx.node_stack.pop()
				ctx.in_dialog = prev_dlg
				ctx.in_sidebar = prev_sb
				ctx.active_form_id = prev_form
		return _wrapper
	return _decorator


def rerun(*, scope: str = "app") -> None:
	raise _RerunException(scope=scope)


def configure_floating_chat(enabled: bool, label: str = "KVKK Korumalı", handler: Callable | None = None) -> None:
	ctx = _get_ctx()
	ctx.session.floating_chat_enabled = bool(enabled)
	ctx.session.floating_chat_label = str(label or "KVKK Korumalı")
	if handler is not None:
		ctx.session.floating_chat_handler = handler


# Kök konteyner ve modül düzeyinde API dışa aktarımları
_root = _UIContainer()
sidebar = _UIContainer(is_sidebar=True)

set_page_config = _root.set_page_config
columns = _root.columns
container = _root.container
expander = _root.expander
popover = _root.popover
tabs = _root.tabs
form = _root.form
spinner = _root.spinner
markdown = _root.markdown
caption = _root.caption
write = _root.write
text = _root.text
code = _root.code
divider = _root.divider
success = _root.success
info = _root.info
warning = _root.warning
error = _root.error
metric = _root.metric
progress = _root.progress
balloons = _root.balloons
image = _root.image
altair_chart = _root.altair_chart
bar_chart = _root.bar_chart
button = _root.button
form_submit_button = _root.form_submit_button
download_button = _root.download_button
link_button = _root.link_button
text_input = _root.text_input
text_area = _root.text_area
number_input = _root.number_input
checkbox = _root.checkbox
toggle = _root.toggle
selectbox = _root.selectbox
multiselect = _root.multiselect
radio = _root.radio
date_input = _root.date_input
time_input = _root.time_input
file_uploader = _root.file_uploader
dataframe = _root.dataframe
data_editor = _root.data_editor


# ==============================================================================
# 8. EYLEM İŞLEME VE ÇALIŞTIRMA DÖNGÜSÜ
# ==============================================================================

def _convert_wire_value(wid: str, raw_val: Any, session: _AppSession) -> Any:
	meta = session.widget_meta.get(wid) or {}
	kind = meta.get("kind", "text")
	if kind == "text":
		return "" if raw_val is None else str(raw_val)
	if kind == "bool":
		return bool(raw_val)
	if kind == "number":
		if raw_val is None or str(raw_val).strip() == "":
			return None if meta.get("allow_none") else (0.0 if meta.get("is_float") else 0)
		try:
			num = float(str(raw_val).strip().replace(",", "."))
		except Exception:
			return session.session_state.get(wid, session.widget_values.get(wid))
		if meta.get("min_value") is not None and num < float(meta["min_value"]):
			num = float(meta["min_value"])
		if meta.get("max_value") is not None and num > float(meta["max_value"]):
			num = float(meta["max_value"])
		return float(num) if meta.get("is_float") else int(round(num))
	if kind in ("selectbox", "radio"):
		opts = meta.get("options") or []
		if raw_val is None or raw_val == "":
			return None
		try:
			idx = int(raw_val)
			if 0 <= idx < len(opts):
				return opts[idx]
		except Exception:
			pass
		return None
	if kind == "multiselect":
		opts = meta.get("options") or []
		if not isinstance(raw_val, list):
			return []
		out = []
		for item in raw_val:
			try:
				idx = int(item)
				if 0 <= idx < len(opts):
					out.append(opts[idx])
			except Exception:
				pass
		return out
	if kind == "date":
		if not raw_val:
			return None
		try:
			return dt_mod.date.fromisoformat(str(raw_val)[:10])
		except Exception:
			return None
	if kind == "time":
		if not raw_val:
			return None
		try:
			parts = str(raw_val).strip().split(":")
			return dt_mod.time(int(parts[0]), int(parts[1]))
		except Exception:
			return None
	return raw_val


def _store_widget_value(wid: str, py_val: Any, session: _AppSession) -> None:
	meta = session.widget_meta.get(wid) or {}
	if meta.get("is_keyed", True):
		session.session_state[wid] = py_val
	else:
		session.widget_values[wid] = py_val


def _apply_editor_mutation(session: _AppSession, ed_key: str, mutation: dict[str, Any]) -> None:
	ed_state = session.session_state.get(ed_key)
	if not isinstance(ed_state, dict) or "edited_rows" not in ed_state:
		ed_state = {"edited_rows": {}, "added_rows": [], "deleted_rows": []}
		session.session_state[ed_key] = ed_state

	op = mutation.get("op")
	if op == "cell":
		row_kind = mutation.get("row_kind", "base")
		row_id = int(mutation.get("row_id", 0))
		col = str(mutation.get("col", ""))
		val = mutation.get("val")
		if row_kind == "base":
			ed_state["edited_rows"].setdefault(row_id, {})[col] = val
		elif row_kind == "added":
			while len(ed_state["added_rows"]) <= row_id:
				ed_state["added_rows"].append({})
			ed_state["added_rows"][row_id][col] = val
	elif op == "add_row":
		ed_state["added_rows"].append({})
	elif op == "delete_row":
		row_kind = mutation.get("row_kind", "base")
		row_id = int(mutation.get("row_id", 0))
		if row_kind == "base":
			if row_id not in ed_state["deleted_rows"]:
				ed_state["deleted_rows"].append(row_id)
		elif row_kind == "added":
			if 0 <= row_id < len(ed_state["added_rows"]):
				ed_state["added_rows"].pop(row_id)


def _execute_render_cycle(app_fn: Callable[[], None], session: _AppSession, payload: dict[str, Any]) -> dict[str, Any]:
	with session.lock:
		action = payload.get("action") or "init"
		act_key = payload.get("key")
		act_val = payload.get("value")

		# 1. Bekleyen kirli metin alanları ve düzenleyici hücreleri önce senkronize edilir
		pending_inputs = payload.get("pending_inputs") or {}
		if isinstance(pending_inputs, dict):
			for p_key, p_raw in pending_inputs.items():
				py_v = _convert_wire_value(p_key, p_raw, session)
				_store_widget_value(p_key, py_v, session)
				if p_key in session.callbacks and p_key != act_key:
					cb, c_args, c_kwargs = session.callbacks[p_key]
					try:
						cb(*c_args, **c_kwargs)
					except _RerunException:
						pass

		pending_edits = payload.get("pending_editor_cells") or []
		if isinstance(pending_edits, list):
			for mut in pending_edits:
				if isinstance(mut, dict) and mut.get("key"):
					_apply_editor_mutation(session, str(mut["key"]), mut)

		clicked_button: str | None = None
		submitted_form: str | None = None

		# 2. Eylem tipini işle
		if action == "click" and act_key:
			clicked_button = str(act_key)
			if clicked_button in session.callbacks:
				cb, c_args, c_kwargs = session.callbacks[clicked_button]
				try:
					cb(*c_args, **c_kwargs)
				except _RerunException:
					clicked_button = None

		elif action == "change" and act_key:
			wid = str(act_key)
			py_v = _convert_wire_value(wid, act_val, session)
			_store_widget_value(wid, py_v, session)
			if wid in session.callbacks:
				cb, c_args, c_kwargs = session.callbacks[wid]
				try:
					cb(*c_args, **c_kwargs)
				except _RerunException:
					pass

		elif action == "form_submit":
			fid = str(payload.get("form_id") or "")
			submitted_form = fid
			form_vals = payload.get("form_values") or {}
			if isinstance(form_vals, dict):
				for f_key, f_raw in form_vals.items():
					py_v = _convert_wire_value(f_key, f_raw, session)
					_store_widget_value(f_key, py_v, session)

		elif action == "data_editor" and act_key:
			ed_key = str(act_key)
			if isinstance(act_val, dict):
				_apply_editor_mutation(session, ed_key, act_val)
			if ed_key in session.callbacks:
				cb, c_args, c_kwargs = session.callbacks[ed_key]
				try:
					cb(*c_args, **c_kwargs)
				except _RerunException:
					pass

		elif action == "df_select" and act_key:
			df_key = str(act_key)
			if isinstance(act_val, dict):
				session.dataframe_selections[df_key] = {
					"rows": [int(r) for r in (act_val.get("rows") or [])],
					"cells": [(int(c[0]), str(c[1])) for c in (act_val.get("cells") or []) if isinstance(c, (list, tuple)) and len(c) >= 2],
				}

		elif action == "close_dialog":
			old_dlg = session.active_dialog
			session.active_dialog = None
			if old_dlg is not None and callable(old_dlg.on_dismiss):
				try:
					old_dlg.on_dismiss()
				except _RerunException:
					pass

		elif action == "clear_upload" and act_key:
			wid = str(act_key)
			file_id = str(act_val or "")
			if file_id:
				session.uploaded_files[wid] = [f for f in (session.uploaded_files.get(wid) or []) if f.file_id != file_id]
			else:
				session.uploaded_files.pop(wid, None)

		elif action == "sidebar_toggle":
			session.sidebar_collapsed = bool(act_val)

		# 3. Uygulama fonksiyonunu rerun döngüsü ile yürüt
		last_ctx: _RenderContext | None = None
		error_trace: str | None = None

		for _pass in range(8):
			ctx = _RenderContext(session)
			ctx.clicked_button = clicked_button
			ctx.submitted_form = submitted_form
			_LOCAL_CTX.ctx = ctx
			session.callbacks.clear()

			rerun_req: _RerunException | None = None
			try:
				app_fn()
				if session.active_dialog is not None and not ctx.dialog_rendered_this_pass:
					dlg = session.active_dialog
					ctx.dialog_nodes = []
					ctx.in_dialog = True
					ctx.in_sidebar = False
					ctx.active_form_id = None
					ctx.node_stack.append(ctx.dialog_nodes)
					try:
						dlg.func(*dlg.args, **dlg.kwargs)
					finally:
						ctx.node_stack.pop()
						ctx.in_dialog = False
			except _RerunException as rexc:
				rerun_req = rexc
			except Exception:
				error_trace = traceback.format_exc()
			finally:
				# clear_on_submit=True olan form gönderildiyse alanlarını temizle
				if submitted_form and session.form_defs.get(submitted_form, {}).get("clear_on_submit"):
					for w_key, is_k in ctx.used_form_keys.get(submitted_form, []):
						if is_k:
							session.session_state.pop(w_key, None)
						else:
							session.widget_values.pop(w_key, None)
				_LOCAL_CTX.ctx = None

			last_ctx = ctx
			if error_trace is not None:
				break
			if rerun_req is not None:
				clicked_button = None
				submitted_form = None
				if rerun_req.scope != "fragment":
					session.active_dialog = None
				continue
			break

		assert last_ctx is not None
		main_html = _flatten_nodes(last_ctx.main_nodes)
		sidebar_html = _flatten_nodes(last_ctx.sidebar_nodes)
		dialog_html = _flatten_nodes(last_ctx.dialog_nodes) if session.active_dialog is not None else None
		if error_trace:
			main_html += (
				'<div class="revir-alert revir-alert-error"><span class="revir-alert-icon">❌</span>'
				f'<div class="revir-alert-body"><strong>Çalışma Zamanı Hatası</strong><pre class="revir-code">{html_lib.escape(error_trace)}</pre></div></div>'
			)

		result = {
			"title": session.page_title,
			"sidebar_html": sidebar_html,
			"main_html": main_html,
			"dialog_html": dialog_html,
			"dialog_title": session.active_dialog.title if session.active_dialog else None,
			"dialog_width": session.active_dialog.width if session.active_dialog else "large",
			"sidebar_collapsed": session.sidebar_collapsed,
			"balloons": last_ctx.balloons_triggered,
			"styles_html": "\n".join(last_ctx.styles_parts),
			"scripts_html": "\n".join(last_ctx.scripts_parts),
			"floating_chat_enabled": session.floating_chat_enabled,
			"floating_chat_label": session.floating_chat_label,
		}
		session.last_html_payload = result
		return result


# ==============================================================================
# 9. YEREL KABUK HTML / CSS / JS (Sıfır Harici Bağımlılık, Çevrimdışı)
# ==============================================================================

_SHELL_HTML = r"""<!DOCTYPE html>
<html lang="tr">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Revir Defteri</title>
<style id="revir-base-css">
:root {
	--ink: #0f172a;
	--teal: #0d9488;
	--teal-hover: #0f766e;
	--mint: #f0fdfa;
	--line: #e2e8f0;
	--card-bg: #ffffff;
	--sidebar-bg: #f8fafc;
	--bg-page: #f4f7f9;
	--input-bg: #ffffff;
	--input-text: #0f172a;
	--text-main: #0f172a;
	--text-heading: #0f172a;
	--text-muted: #475569;
}
*, *::before, *::after { box-sizing: border-box; }
html, body {
	margin: 0; padding: 0; height: 100%; width: 100%;
	font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
	font-size: 15px; line-height: 1.5;
	color: var(--text-main); background: var(--bg-page);
	overflow: hidden;
}
.revir-app {
	display: flex; height: 100vh; width: 100vw; overflow: hidden;
	background: var(--bg-page);
}
.revir-sidebar {
	width: 280px; min-width: 280px; max-width: 280px;
	height: 100vh; overflow-y: auto; overflow-x: hidden;
	background: var(--sidebar-bg);
	border-right: 1px solid var(--line);
	padding: 18px 16px 24px 16px;
	transition: margin-left 0.22s ease, opacity 0.2s ease;
	z-index: 20;
	flex-shrink: 0;
}
.revir-sidebar.collapsed {
	margin-left: -280px;
	opacity: 0;
	pointer-events: none;
}
.revir-sidebar-open-btn {
	position: fixed; top: 14px; left: 14px; z-index: 500;
	width: 38px; height: 38px; border-radius: 10px;
	border: 1.5px solid var(--teal); background: var(--card-bg); color: var(--teal);
	font-size: 18px; font-weight: 700; display: none;
	align-items: center; justify-content: center; cursor: pointer;
	box-shadow: 0 3px 12px rgba(0,0,0,0.12);
	transition: all 0.18s ease;
}
.revir-sidebar-open-btn:hover { background: var(--teal); color: #ffffff; transform: translateX(2px); }
.revir-sidebar-open-btn.visible { display: flex; }

.revir-main {
	flex: 1; height: 100vh; overflow-y: auto; overflow-x: hidden;
	padding: 24px 32px 48px 32px;
	scroll-behavior: auto;
}
.revir-main-inner { max-width: 1480px; margin: 0 auto; }

/* Tipografi & Markdown */
.revir-md p { margin: 0 0 8px 0; }
.revir-md p:last-child { margin-bottom: 0; }
.revir-h1 { font-size: 1.75rem; margin: 0 0 12px 0; font-weight: 800; color: var(--text-heading); }
.revir-h2 { font-size: 1.42rem; margin: 6px 0 10px 0; font-weight: 800; color: var(--text-heading); }
.revir-h3 { font-size: 1.18rem; margin: 10px 0 8px 0; font-weight: 700; color: var(--text-heading); }
.revir-h4 { font-size: 1.05rem; margin: 10px 0 6px 0; font-weight: 700; color: var(--text-heading); }
.revir-h5 { font-size: 0.96rem; margin: 8px 0 6px 0; font-weight: 700; color: var(--text-heading); }
.revir-h6 { font-size: 0.88rem; margin: 6px 0 4px 0; font-weight: 700; color: var(--text-heading); }
.revir-caption { font-size: 0.84rem; color: var(--text-muted); margin: 2px 0 8px 0; line-height: 1.45; }
.revir-spacer { height: 12px; }
.revir-hr { border: none; border-top: 1px solid var(--line); margin: 14px 0; }
.revir-code { background: rgba(15,23,42,0.06); border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12.5px; overflow-x: auto; white-space: pre-wrap; margin: 8px 0; }
.revir-plain-text { background: rgba(15,23,42,0.04); border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px; font-family: ui-monospace, monospace; font-size: 12.5px; max-height: 360px; overflow: auto; white-space: pre-wrap; margin: 8px 0; }
.revir-ul, .revir-ol { margin: 4px 0 10px 0; padding-left: 22px; }
.revir-quote { border-left: 3px solid var(--teal); padding: 4px 12px; margin: 8px 0; color: var(--text-muted); }

/* Düzen & Kartlar */
.revir-cols { display: grid; width: 100%; margin-bottom: 8px; }
.revir-col { min-width: 0; }
.revir-container { margin-bottom: 10px; }
.revir-card {
	background: var(--card-bg); border: 1px solid var(--line);
	border-radius: 12px; padding: 14px 16px;
	box-shadow: 0 2px 10px rgba(0,0,0,0.03);
	margin-bottom: 12px;
}
.revir-form { margin-bottom: 12px; }

/* Uyarılar */
.revir-alert {
	display: flex; align-items: flex-start; gap: 10px;
	padding: 11px 14px; border-radius: 10px; margin: 8px 0 10px 0;
	border: 1px solid transparent; font-size: 0.92rem;
}
.revir-alert-icon { flex-shrink: 0; line-height: 1.4; }
.revir-alert-body { flex: 1; min-width: 0; }
.revir-alert-success { background: #ecfdf5; border-color: #a7f3d0; color: #065f46; }
.revir-alert-info { background: #eff6ff; border-color: #bfdbfe; color: #1e40af; }
.revir-alert-warning { background: #fffbeb; border-color: #fde68a; color: #92400e; }
.revir-alert-error { background: #fef2f2; border-color: #fecaca; color: #991b1b; }

/* Metrikler */
.revir-metric {
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 12px;
	padding: 14px 16px; margin-bottom: 8px;
	box-shadow: 0 2px 8px rgba(0,0,0,0.04);
	transition: transform 0.18s ease, box-shadow 0.18s ease;
}
.revir-metric:hover { transform: translateY(-2px); box-shadow: 0 6px 18px rgba(0,0,0,0.07); }
.revir-metric-label { font-size: 13px; font-weight: 600; color: var(--text-muted); margin-bottom: 4px; }
.revir-metric-value { font-size: 24px; font-weight: 800; color: var(--text-heading); line-height: 1.2; }
.revir-metric-delta { font-size: 12.5px; font-weight: 600; margin-top: 4px; }
.revir-delta-pos { color: #059669; }
.revir-delta-neg { color: #dc2626; }
.revir-delta-neutral { color: var(--text-muted); }

/* Butonlar */
.revir-widget-row { margin-bottom: 8px; }
.revir-btn {
	display: inline-flex; align-items: center; justify-content: center; gap: 6px;
	padding: 8px 14px; border-radius: 8px; font-size: 0.9rem; font-weight: 600;
	cursor: pointer; border: 1px solid var(--teal); background: var(--card-bg); color: var(--teal);
	transition: all 0.16s ease; font-family: inherit; line-height: 1.3;
}
.revir-btn:hover:not(:disabled) { background: var(--mint); transform: translateY(-1px); box-shadow: 0 3px 10px rgba(0,0,0,0.08); }
.revir-btn:disabled { opacity: 0.5; cursor: not-allowed; }
.revir-btn-primary { background: var(--teal); color: #ffffff; border-color: var(--teal); }
.revir-btn-primary:hover:not(:disabled) { background: var(--teal-hover); color: #ffffff; box-shadow: 0 4px 12px rgba(13,148,136,0.32); }
.revir-btn-sm { padding: 5px 10px; font-size: 0.82rem; }
.revir-w-full { width: 100%; }

/* Yan Menü Navigasyonu */
.revir-nav-group { display: flex; flex-direction: column; gap: 3px; margin: 8px 0; }
.revir-nav-item {
	width: 100%; text-align: left; padding: 9px 12px; border-radius: 9px;
	border: 1px solid transparent; background: transparent; color: var(--text-main);
	font-size: 0.92rem; font-weight: 600; cursor: pointer; font-family: inherit;
	transition: all 0.15s ease;
}
.revir-nav-item:hover { background: rgba(13,148,136,0.10); color: var(--teal); }
.revir-nav-item.active { background: var(--teal); color: #ffffff !important; box-shadow: 0 2px 8px rgba(13,148,136,0.28); }
.revir-sidebar-close-btn {
	display: flex; align-items: center; justify-content: center; gap: 6px;
	padding: 8px 10px; margin-top: 10px; border-radius: 8px;
	border: 1px solid var(--line); background: var(--card-bg); color: var(--text-muted);
	font-size: 12.5px; font-weight: 600; cursor: pointer; user-select: none;
	transition: all 0.16s ease;
}
.revir-sidebar-close-btn:hover { background: var(--mint); border-color: var(--teal); color: var(--teal); }

/* Form Alanları */
.revir-field { margin-bottom: 10px; }
.revir-field-label { display: block; font-size: 0.85rem; font-weight: 600; color: var(--text-main); margin-bottom: 4px; }
.revir-sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0,0,0,0); }
.revir-help-icon {
	display: inline-flex; align-items: center; justify-content: center;
	width: 16px; height: 16px; border-radius: 50%; font-size: 11px; font-weight: 700;
	border: 1px solid var(--line); color: var(--text-muted); cursor: help; margin-left: 4px;
}
.revir-input, .revir-textarea {
	width: 100%; padding: 8px 11px; border-radius: 7px;
	border: 1px solid var(--line); background: var(--input-bg); color: var(--input-text);
	font-family: inherit; font-size: 0.92rem; transition: border-color 0.15s, box-shadow 0.15s;
}
.revir-input:focus, .revir-textarea:focus { outline: none; border-color: var(--teal); box-shadow: 0 0 0 3px rgba(13,148,136,0.16); }
.revir-textarea { resize: vertical; min-height: 76px; }

/* Checkbox, Toggle, Radio */
.revir-check-wrap, .revir-toggle-wrap {
	display: inline-flex; align-items: center; gap: 8px; cursor: pointer;
	margin: 4px 0 8px 0; user-select: none; font-size: 0.9rem;
}
.revir-check-input { width: 17px; height: 17px; accent-color: var(--teal); cursor: pointer; flex-shrink: 0; }
.revir-toggle-input { display: none; }
.revir-toggle-track {
	width: 38px; height: 22px; border-radius: 999px; background: #cbd5e1;
	position: relative; transition: background 0.18s; flex-shrink: 0;
}
.revir-toggle-thumb {
	width: 16px; height: 16px; border-radius: 50%; background: #ffffff;
	position: absolute; top: 3px; left: 3px; transition: transform 0.18s;
	box-shadow: 0 1px 3px rgba(0,0,0,0.2);
}
.revir-toggle-input:checked + .revir-toggle-track { background: var(--teal); }
.revir-toggle-input:checked + .revir-toggle-track .revir-toggle-thumb { transform: translateX(16px); }
.revir-radio-group { display: flex; flex-direction: column; gap: 6px; }
.revir-radio-horiz { flex-direction: row; flex-wrap: wrap; gap: 14px; }
.revir-radio-opt { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; font-size: 0.9rem; }
.revir-radio-opt input { accent-color: var(--teal); width: 16px; height: 16px; }

/* Aranabilir Selectbox & Multiselect */
.revir-combo, .revir-multi { position: relative; width: 100%; }
.revir-combo-control, .revir-multi-control {
	display: flex; flex-wrap: wrap; align-items: center; gap: 4px;
	min-height: 38px; padding: 4px 8px; border-radius: 7px;
	border: 1px solid var(--line); background: var(--input-bg); color: var(--input-text);
	cursor: text;
}
.revir-combo-control:focus-within, .revir-multi-control:focus-within { border-color: var(--teal); box-shadow: 0 0 0 3px rgba(13,148,136,0.16); }
.revir-combo-search, .revir-multi-search {
	flex: 1; min-width: 80px; border: none; outline: none; background: transparent;
	color: var(--input-text); font-family: inherit; font-size: 0.91rem; padding: 3px 2px;
}
.revir-combo-arrow { color: var(--text-muted); font-size: 12px; padding: 0 4px; pointer-events: none; }
.revir-combo-clear {
	border: none; background: transparent; color: var(--text-muted);
	cursor: pointer; font-size: 16px; line-height: 1; padding: 0 4px;
}
.revir-combo-clear:hover { color: #dc2626; }
.revir-combo-menu {
	display: none; position: absolute; left: 0; right: 0; top: calc(100% + 4px);
	max-height: 270px; overflow-y: auto; z-index: 950;
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 9px;
	box-shadow: 0 10px 30px rgba(0,0,0,0.16); padding: 4px;
}
.revir-combo-menu.open { display: block; }
.revir-combo-opt {
	padding: 7px 10px; border-radius: 6px; cursor: pointer; font-size: 0.89rem;
	display: flex; align-items: center; gap: 8px; color: var(--text-main);
}
.revir-combo-opt:hover { background: var(--mint); }
.revir-combo-opt.selected { background: rgba(13,148,136,0.14); font-weight: 700; color: var(--teal); }
.revir-chip {
	display: inline-flex; align-items: center; gap: 4px;
	background: rgba(13,148,136,0.14); color: var(--teal);
	border: 1px solid rgba(13,148,136,0.28); border-radius: 6px;
	padding: 2px 7px; font-size: 0.81rem; font-weight: 600;
}
.revir-chip-x { border: none; background: transparent; color: inherit; cursor: pointer; font-weight: 800; padding: 0 2px; font-size: 14px; line-height: 1; }

/* Sekmeler */
.revir-tabs { margin: 8px 0 14px 0; }
.revir-tab-bar { display: flex; flex-wrap: wrap; gap: 4px; border-bottom: 1.5px solid var(--line); margin-bottom: 14px; }
.revir-tab-btn {
	padding: 8px 15px; border: none; border-bottom: 2.5px solid transparent;
	background: transparent; color: var(--text-muted); font-family: inherit;
	font-size: 0.9rem; font-weight: 600; cursor: pointer; border-radius: 8px 8px 0 0;
	margin-bottom: -1.5px; transition: all 0.15s;
}
.revir-tab-btn:hover { color: var(--teal); background: rgba(13,148,136,0.06); }
.revir-tab-btn.active { color: var(--teal); border-bottom-color: var(--teal); background: rgba(13,148,136,0.08); font-weight: 700; }
.revir-tab-pane { display: none; }
.revir-tab-pane.active { display: block; }

/* Expander */
.revir-expander {
	border: 1px solid var(--line); border-radius: 10px;
	background: var(--card-bg); margin: 8px 0 12px 0; overflow: hidden;
}
.revir-expander-summary {
	padding: 10px 14px; font-weight: 600; font-size: 0.93rem; cursor: pointer;
	user-select: none; list-style: none; display: flex; align-items: center; justify-content: space-between;
}
.revir-expander-summary::after { content: "▾"; font-size: 13px; color: var(--text-muted); transition: transform 0.18s; }
.revir-expander[open] > .revir-expander-summary::after { transform: rotate(180deg); }
.revir-expander-body { padding: 4px 14px 14px 14px; border-top: 1px solid var(--line); }

/* Popover */
.revir-popover-wrap { position: relative; display: block; margin-bottom: 6px; }
.revir-popover-body {
	display: none; position: absolute; right: 0; top: calc(100% + 6px);
	min-width: 280px; max-width: 350px; max-height: 75vh; overflow-y: auto;
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 12px;
	padding: 14px; box-shadow: 0 12px 34px rgba(0,0,0,0.22); z-index: 980;
}
.revir-popover-body.open { display: block; }

/* Tablolar & Data Editor */
.revir-table-wrap {
	width: 100%; overflow: auto; border: 1px solid var(--line);
	border-radius: 10px; background: var(--card-bg); margin-bottom: 10px;
}
.revir-table { width: 100%; border-collapse: collapse; font-size: 0.86rem; }
.revir-table thead th {
	position: sticky; top: 0; z-index: 3;
	background: var(--sidebar-bg); color: var(--text-heading);
	font-weight: 700; text-align: left; padding: 9px 11px;
	border-bottom: 1.5px solid var(--line); white-space: nowrap; cursor: pointer; user-select: none;
}
.revir-sort-ind { opacity: 0.4; font-size: 11px; margin-left: 3px; }
.revir-table tbody td { padding: 7px 11px; border-bottom: 1px solid var(--line); vertical-align: middle; }
.revir-table tbody tr:hover td { background: rgba(13,148,136,0.05); }
.revir-row-selected td { background: rgba(13,148,136,0.13) !important; }
.revir-cell-clickable { cursor: pointer; }
.revir-dosya-pill {
	display: inline-block; padding: 3px 9px; border-radius: 6px;
	background: rgba(13,148,136,0.12); color: var(--teal); font-weight: 700;
	border: 1px solid rgba(13,148,136,0.3); font-size: 0.82rem; white-space: nowrap;
}
.revir-dosya-pill:hover { background: var(--teal); color: #ffffff; }
.revir-empty-td { text-align: center; color: var(--text-muted); padding: 20px !important; }

.revir-editor-table td.revir-ed-cell { padding: 3px 4px; }
.revir-editor-table td.revir-ed-ro { background: rgba(148,163,184,0.08); color: var(--text-muted); font-weight: 500; }
.revir-ed-center { text-align: center; }
.revir-ed-input, .revir-ed-select {
	width: 100%; min-width: 72px; padding: 5px 7px; border-radius: 5px;
	border: 1px solid transparent; background: transparent; color: var(--text-main);
	font-family: inherit; font-size: 0.85rem;
}
.revir-ed-input:hover, .revir-ed-select:hover { border-color: var(--line); background: var(--input-bg); }
.revir-ed-input:focus, .revir-ed-select:focus { outline: none; border-color: var(--teal); background: var(--input-bg); box-shadow: 0 0 0 2px rgba(13,148,136,0.18); }
.revir-ed-del-btn { border: none; background: transparent; cursor: pointer; font-size: 14px; padding: 3px 6px; border-radius: 5px; }
.revir-ed-del-btn:hover { background: #fee2e2; }
.revir-ed-footer { padding: 6px 0 4px 0; }

/* Dosya Yükleme */
.revir-dropzone {
	border: 1.5px dashed var(--line); border-radius: 10px; padding: 14px 16px;
	background: var(--card-bg); display: flex; align-items: center; justify-content: space-between;
	gap: 12px; transition: border-color 0.16s, background 0.16s;
}
.revir-dropzone.dragover { border-color: var(--teal); background: var(--mint); }
.revir-dropzone-text { font-size: 0.88rem; color: var(--text-muted); }
.revir-uploaded-list { margin-top: 6px; display: flex; flex-direction: column; gap: 4px; }
.revir-uploaded-item {
	display: flex; align-items: center; justify-content: space-between;
	background: rgba(13,148,136,0.08); border: 1px solid rgba(13,148,136,0.22);
	padding: 5px 10px; border-radius: 7px; font-size: 0.84rem;
}

/* İlerleme Çubuğu & Görsel & Iframe */
.revir-progress-wrap { margin: 6px 0 10px 0; }
.revir-progress-track { height: 9px; border-radius: 999px; background: var(--line); overflow: hidden; }
.revir-progress-fill { height: 100%; background: var(--teal); border-radius: 999px; transition: width 0.25s ease; }
.revir-img-wrap { margin: 6px 0; }
.revir-img { border-radius: 8px; display: block; }
.revir-iframe-wrap { margin: 8px 0; border: 1px solid var(--line); border-radius: 10px; overflow: hidden; background: #ffffff; }
.revir-iframe { border: none; display: block; }

/* Çevrimdışı SVG Grafikler */
.revir-chart-card {
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 12px;
	padding: 12px 14px; margin: 8px 0 12px 0; position: relative;
}
.revir-svg-chart { width: 100%; display: block; overflow: visible; }
.revir-grid-line { stroke: var(--line); stroke-width: 1; }
.revir-axis-line { stroke: var(--text-muted); stroke-width: 1.2; opacity: 0.55; }
.revir-axis-lbl { fill: var(--text-muted); font-size: 11px; }
.revir-axis-title { fill: var(--text-main); font-size: 12px; font-weight: 700; }
.revir-bar-val { fill: var(--text-main); font-size: 11px; font-weight: 700; }
.revir-bar-rect { transition: opacity 0.15s; cursor: pointer; }
.revir-bar-rect:hover { opacity: 0.82; }
.revir-chart-pt { cursor: pointer; }
.revir-chart-legend { display: flex; flex-wrap: wrap; gap: 12px; justify-content: center; margin-top: 8px; font-size: 12px; }
.revir-legend-item { display: inline-flex; align-items: center; gap: 5px; color: var(--text-main); }
.revir-legend-dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
.revir-chart-empty { padding: 18px; text-align: center; color: var(--text-muted); }
#revir-chart-tooltip {
	position: fixed; display: none; z-index: 99999; pointer-events: none;
	background: #0f172a; color: #f8fafc; padding: 7px 11px; border-radius: 8px;
	font-size: 12px; font-weight: 500; max-width: 340px;
	box-shadow: 0 8px 24px rgba(0,0,0,0.28);
}

/* Modal Diyalog */
.revir-modal-backdrop {
	position: fixed; inset: 0; z-index: 2000;
	background: rgba(15, 23, 42, 0.55); backdrop-filter: blur(4px);
	display: none; align-items: center; justify-content: center; padding: 20px;
}
.revir-modal-backdrop.open { display: flex; }
.revir-modal-card {
	background: var(--card-bg); color: var(--text-main);
	border: 1px solid var(--line); border-radius: 16px;
	box-shadow: 0 24px 64px rgba(0,0,0,0.34);
	width: 92vw; max-width: 1160px; max-height: 90vh;
	display: flex; flex-direction: column; overflow: hidden;
}
.revir-modal-card.small { max-width: 640px; }
.revir-modal-header {
	display: flex; align-items: center; justify-content: space-between;
	padding: 14px 20px; border-bottom: 1px solid var(--line);
	background: var(--sidebar-bg); flex-shrink: 0;
}
.revir-modal-title { font-size: 1.08rem; font-weight: 800; color: var(--text-heading); margin: 0; }
.revir-modal-close {
	width: 32px; height: 32px; border-radius: 8px; border: 1px solid var(--line);
	background: var(--card-bg); color: var(--text-main); font-size: 18px;
	cursor: pointer; display: flex; align-items: center; justify-content: center;
}
.revir-modal-close:hover { background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }
.revir-modal-body { padding: 18px 22px 24px 22px; overflow-y: auto; flex: 1; }

/* Bildirim Toast & Yükleme Çubuğu */
#revir-busy-bar {
	position: fixed; top: 0; left: 0; height: 3px; width: 0%;
	background: var(--teal); z-index: 9999; transition: width 0.18s ease, opacity 0.2s;
	opacity: 0; pointer-events: none;
}
#revir-busy-bar.active { width: 78%; opacity: 1; }
#revir-toast {
	position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 99999;
	background: #0f172a; color: #ffffff; padding: 10px 16px; border-radius: 10px;
	font-size: 13.5px; font-weight: 600; box-shadow: 0 10px 28px rgba(0,0,0,0.25);
	display: none;
}

/* Yüzen Gemini Sohbet Butonu & Konuşma Balonu */
#revir-fab-btn {
	position: fixed; bottom: 22px; right: 24px; z-index: 3400;
	height: 54px; min-width: 54px; padding: 0 18px; border-radius: 999px;
	border: 1.5px solid rgba(255,255,255,0.45);
	background: linear-gradient(135deg, var(--teal, #0d9488) 0%, #0284c7 100%);
	color: #ffffff; font-size: 14.5px; font-weight: 700; font-family: inherit;
	display: none; align-items: center; justify-content: center; gap: 8px;
	cursor: pointer; box-shadow: 0 10px 28px rgba(13,148,136,0.38);
	transition: transform 0.2s cubic-bezier(0.34,1.56,0.64,1), box-shadow 0.2s ease;
	user-select: none;
}
#revir-fab-btn.visible { display: inline-flex; }
#revir-fab-btn:hover { transform: translateY(-3px) scale(1.03); box-shadow: 0 14px 34px rgba(2,132,199,0.45); }
#revir-fab-btn .revir-fab-icon { font-size: 22px; line-height: 1; }

#revir-chat-bubble {
	position: fixed; bottom: 88px; right: 24px; z-index: 3450;
	width: 390px; height: 535px; max-width: calc(100vw - 32px); max-height: calc(100vh - 110px);
	background: var(--card-bg, #ffffff); color: var(--text-main, #0f172a);
	border: 1px solid var(--line, #e2e8f0); border-radius: 20px;
	box-shadow: 0 22px 60px rgba(15,23,42,0.28);
	display: none; flex-direction: column; overflow: hidden;
	transition: width 0.22s ease, height 0.22s ease;
}
#revir-chat-bubble.open { display: flex; }
#revir-chat-bubble.expanded { width: 550px; height: 675px; }
#revir-chat-bubble::after {
	content: ""; position: absolute; bottom: -8px; right: 30px;
	width: 16px; height: 16px; background: var(--card-bg, #ffffff);
	border-right: 1px solid var(--line, #e2e8f0); border-bottom: 1px solid var(--line, #e2e8f0);
	transform: rotate(45deg); pointer-events: none;
}
.revir-chat-hdr {
	display: flex; align-items: center; justify-content: space-between;
	padding: 12px 14px; background: linear-gradient(135deg, var(--teal, #0d9488) 0%, #0284c7 100%);
	color: #ffffff; flex-shrink: 0;
}
.revir-chat-hdr-left { display: flex; align-items: center; gap: 9px; min-width: 0; }
.revir-chat-avatar {
	width: 34px; height: 34px; border-radius: 10px;
	background: rgba(255,255,255,0.2); display: flex; align-items: center; justify-content: center;
	font-size: 19px; flex-shrink: 0;
}
.revir-chat-title-wrap { min-width: 0; }
.revir-chat-title { font-size: 14px; font-weight: 800; color: #ffffff !important; line-height: 1.2; }
.revir-chat-sub { font-size: 11px; color: rgba(255,255,255,0.88) !important; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.revir-chat-hdr-btns { display: flex; align-items: center; gap: 4px; }
.revir-chat-hbtn {
	width: 29px; height: 29px; border-radius: 8px; border: none;
	background: rgba(255,255,255,0.16); color: #ffffff !important;
	cursor: pointer; font-size: 13.5px; display: flex; align-items: center; justify-content: center;
	transition: background 0.15s;
}
.revir-chat-hbtn:hover { background: rgba(255,255,255,0.3); }
.revir-chat-hbtn.active { background: #ffffff; color: #0284c7 !important; box-shadow: 0 2px 6px rgba(0,0,0,0.18); }

.revir-chat-chips {
	display: flex; gap: 6px; overflow-x: auto; padding: 8px 12px;
	border-bottom: 1px solid var(--line, #e2e8f0); background: var(--sidebar-bg, #f8fafc);
	flex-shrink: 0; scrollbar-width: none;
}
.revir-chat-qchip {
	white-space: nowrap; padding: 4px 10px; border-radius: 999px;
	border: 1px solid var(--line, #e2e8f0); background: var(--card-bg, #ffffff);
	color: var(--teal, #0d9488); font-size: 11.5px; font-weight: 600; cursor: pointer;
	transition: all 0.15s; font-family: inherit;
}
.revir-chat-qchip:hover { background: var(--teal, #0d9488); color: #ffffff !important; border-color: var(--teal, #0d9488); }

.revir-chat-msgs {
	flex: 1; overflow-y: auto; padding: 12px 14px;
	display: flex; flex-direction: column; gap: 10px;
	background: var(--bg-page, #f4f7f9);
}
.revir-cmsg {
	max-width: 88%; padding: 9px 13px; border-radius: 14px;
	font-size: 13.2px; line-height: 1.48; word-wrap: break-word; position: relative;
}
.revir-cmsg p { margin: 0 0 6px 0; }
.revir-cmsg p:last-child { margin-bottom: 0; }
.revir-cmsg ul, .revir-cmsg ol { margin: 4px 0 6px 0; padding-left: 18px; }
.revir-cmsg-user {
	align-self: flex-end; background: var(--teal, #0d9488); color: #ffffff !important;
	border-bottom-right-radius: 4px; box-shadow: 0 2px 8px rgba(13,148,136,0.22);
}
.revir-cmsg-user * { color: #ffffff !important; }
.revir-cmsg-ai {
	align-self: flex-start; background: var(--card-bg, #ffffff); color: var(--text-main, #0f172a);
	border: 1px solid var(--line, #e2e8f0); border-bottom-left-radius: 4px;
	box-shadow: 0 2px 6px rgba(0,0,0,0.04);
}
.revir-cmsg-tools {
	display: flex; align-items: center; justify-content: flex-end; gap: 6px;
	margin-top: 6px; padding-top: 4px; border-top: 1px dashed var(--line, #e2e8f0);
}
.revir-cmsg-tbtn {
	border: none; background: transparent; color: var(--text-muted, #64748b);
	font-size: 11px; font-weight: 600; cursor: pointer; padding: 2px 6px; border-radius: 5px;
}
.revir-cmsg-tbtn:hover { background: rgba(13,148,136,0.12); color: var(--teal, #0d9488); }

.revir-chat-files {
	display: none; flex-wrap: wrap; gap: 5px; padding: 6px 12px;
	background: var(--sidebar-bg, #f8fafc); border-top: 1px solid var(--line, #e2e8f0);
	max-height: 76px; overflow-y: auto;
}
.revir-chat-files.has-files { display: flex; }
.revir-chat-fchip {
	display: inline-flex; align-items: center; gap: 5px;
	background: rgba(13,148,136,0.12); color: var(--teal, #0d9488);
	border: 1px solid rgba(13,148,136,0.3); border-radius: 7px;
	padding: 2px 8px; font-size: 11.5px; font-weight: 600;
}

.revir-chat-composer {
	display: flex; align-items: flex-end; gap: 6px;
	padding: 9px 11px; background: var(--card-bg, #ffffff);
	border-top: 1px solid var(--line, #e2e8f0); flex-shrink: 0;
}
.revir-chat-ibtn {
	width: 34px; height: 34px; border-radius: 9px;
	border: 1px solid var(--line, #e2e8f0); background: var(--sidebar-bg, #f8fafc);
	color: var(--text-main, #0f172a); font-size: 15px; cursor: pointer;
	display: flex; align-items: center; justify-content: center; flex-shrink: 0;
	transition: all 0.15s;
}
.revir-chat-ibtn:hover { border-color: var(--teal, #0d9488); color: var(--teal, #0d9488); background: var(--mint, #f0fdfa); }
.revir-chat-ibtn.recording {
	background: #ef4444 !important; color: #ffffff !important; border-color: #dc2626 !important;
	animation: revirPulse 1.1s infinite;
}
@keyframes revirPulse {
	0% { box-shadow: 0 0 0 0 rgba(239,68,68,0.55); }
	70% { box-shadow: 0 0 0 9px rgba(239,68,68,0); }
	100% { box-shadow: 0 0 0 0 rgba(239,68,68,0); }
}
#revir-chat-inp {
	flex: 1; min-height: 34px; max-height: 92px; padding: 7px 10px;
	border-radius: 9px; border: 1px solid var(--line, #e2e8f0);
	background: var(--input-bg, #ffffff); color: var(--input-text, #0f172a);
	font-family: inherit; font-size: 13px; resize: none; outline: none; line-height: 1.38;
}
#revir-chat-inp:focus { border-color: var(--teal, #0d9488); box-shadow: 0 0 0 2px rgba(13,148,136,0.16); }
#revir-chat-send-btn {
	background: var(--teal, #0d9488); color: #ffffff !important; border-color: var(--teal, #0d9488);
	font-weight: 800;
}
#revir-chat-send-btn:hover { background: var(--teal-hover, #0f766e); }
</style>
<div id="revir-dynamic-styles">__REVIR_INIT_STYLES__</div>
</head>
<body>
<div id="revir-busy-bar"></div>
<button id="revir-sidebar-open" class="revir-sidebar-open-btn" type="button" title="Menüyü Aç" onclick="window.__revirSetSidebar(false)">❯</button>
<div class="revir-app">
	<aside id="revir-sidebar" class="revir-sidebar">__REVIR_INIT_SIDEBAR__</aside>
	<main id="revir-main" class="revir-main">
		<div id="revir-main-inner" class="revir-main-inner">__REVIR_INIT_MAIN__</div>
	</main>
</div>
<div id="revir-modal-backdrop" class="revir-modal-backdrop" onclick="if(event.target===this) window.__revirCloseDialog()">
	<div id="revir-modal-card" class="revir-modal-card" role="dialog" aria-modal="true">
		<div class="revir-modal-header">
			<h3 id="revir-modal-title" class="revir-modal-title"></h3>
			<button type="button" class="revir-modal-close" title="Kapat" onclick="window.__revirCloseDialog()">✕</button>
		</div>
		<div id="revir-modal-body" class="revir-modal-body"></div>
	</div>
</div>
<div id="revir-chart-tooltip"></div>
<div id="revir-toast"></div>

<!-- Yüzen Gemini Sohbet Butonu ve Konuşma Balonu -->
<button id="revir-fab-btn" type="button" title="Gemini Klinik Asistan ile Konuş" onclick="window.__revirToggleChatBubble()">
	<span class="revir-fab-icon">🤖</span>
	<span>Gemini Asistan</span>
</button>

<div id="revir-chat-bubble" ondragover="event.preventDefault()" ondrop="window.__revirChatDropFiles(event)">
	<div class="revir-chat-hdr">
		<div class="revir-chat-hdr-left">
			<div class="revir-chat-avatar">🤖</div>
			<div class="revir-chat-title-wrap">
				<div class="revir-chat-title">Gemini Klinik Asistan</div>
				<div id="revir-chat-sub" class="revir-chat-sub">🛡️ KVKK Korumalı · Ses & Dosya Destekli</div>
			</div>
		</div>
		<div class="revir-chat-hdr-btns">
			<button id="revir-chat-tts-btn" type="button" class="revir-chat-hbtn" title="Sesli Yanıt Okuma (Aç / Kapat)" onclick="window.__revirToggleChatTTS()">🔇</button>
			<button type="button" class="revir-chat-hbtn" title="Sohbeti Temizle" onclick="window.__revirClearChat()">🗑️</button>
			<button type="button" class="revir-chat-hbtn" title="Pencereyi Büyüt / Küçült" onclick="window.__revirExpandChat()">⤢</button>
			<button type="button" class="revir-chat-hbtn" title="Küçült" onclick="window.__revirToggleChatBubble()">✕</button>
		</div>
	</div>
	<div class="revir-chat-chips">
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Seçili kişinin sağlık, büyüme ve revir profilini kısaca özetle.')">🩺 Seçili Kişiyi Özetle</button>
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Bugünkü revir başvurularını ve dikkat edilmesi gerekenleri özetle.')">📊 Bugünkü Revir Özeti</button>
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Seçili kişinin kronik durumuna göre okul içi acil müdahale adımları nelerdir?')">🚨 Acil Müdahale Rehberi</button>
		<button type="button" class="revir-chat-qchip" onclick="document.getElementById('revir-chat-file-inp').click()">📎 Tahlil / Belge Yükle</button>
	</div>
	<div id="revir-chat-messages" class="revir-chat-msgs">
		<div class="revir-cmsg revir-cmsg-ai">
			<p>Merhaba! Ben <strong>Gemini Klinik Asistan</strong>. 🩺</p>
			<p>Bana yazılı soru sorabilir, <strong>📎 Dosya Ekle</strong> ile tahlil/PDF/görsel/Excel yükleyebilir veya <strong>🎙️ Mikrofon</strong> tuşuyla sesli konuşabilirsiniz.</p>
		</div>
	</div>
	<div id="revir-chat-files" class="revir-chat-files"></div>
	<div class="revir-chat-composer">
		<label class="revir-chat-ibtn" title="Dosya / Tahlil / Görsel Ekle (.pdf, .png, .jpg, .xlsx, .docx, .csv)">
			📎
			<input id="revir-chat-file-inp" type="file" multiple style="display:none" accept=".pdf,.png,.jpg,.jpeg,.webp,.xlsx,.xls,.csv,.docx,.txt,.mp3,.wav,.webm,.m4a,.ogg" onchange="window.__revirChatAddFiles(this.files);this.value='';"/>
		</label>
		<button id="revir-chat-mic-btn" type="button" class="revir-chat-ibtn" title="Sesli Konuş (Başlat / Gönder)" onclick="window.__revirToggleVoiceMic()">🎙️</button>
		<textarea id="revir-chat-inp" rows="1" placeholder="Mesaj yazın, dosya sürükleyin veya 🎙️ konuşun..." onkeydown="if(event.key==='Enter' && !event.shiftKey){event.preventDefault();window.__revirSendChat();}"></textarea>
		<button id="revir-chat-send-btn" type="button" class="revir-chat-ibtn" title="Gönder" onclick="window.__revirSendChat()">➤</button>
	</div>
</div>

<script>
(function() {
	var CSRF = "__REVIR_CSRF__";
	var pendingInputs = {};
	var pendingEditorCells = [];
	var inflight = false;
	var queuedReq = null;
	var lastFocusedEditorId = null;

	function showToast(msg) {
		var el = document.getElementById("revir-toast");
		if (!el) return;
		el.textContent = msg;
		el.style.display = "block";
		clearTimeout(el.__t);
		el.__t = setTimeout(function() { el.style.display = "none"; }, 3200);
	}

	function setBusy(on) {
		var bar = document.getElementById("revir-busy-bar");
		if (!bar) return;
		if (on) bar.classList.add("active");
		else bar.classList.remove("active");
	}

	function sendAction(payload, opts) {
		opts = opts || {};
		if (inflight) {
			queuedReq = { payload: payload, opts: opts };
			return;
		}
		inflight = true;
		setBusy(true);

		var mainEl = document.getElementById("revir-main");
		var dlgBody = document.getElementById("revir-modal-body");
		var savedMainScroll = mainEl ? mainEl.scrollTop : 0;
		var savedDlgScroll = dlgBody ? dlgBody.scrollTop : 0;
		var savedSbScroll = (document.getElementById("revir-sidebar") || {}).scrollTop || 0;

		payload.pending_inputs = Object.assign({}, pendingInputs);
		payload.pending_editor_cells = pendingEditorCells.slice();
		pendingInputs = {};
		pendingEditorCells = [];

		fetch("/api/action", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify(payload)
		})
		.then(function(r) { return r.json(); })
		.then(function(data) {
			applyRender(data, savedMainScroll, savedDlgScroll, savedSbScroll, opts.refocusId);
		})
		.catch(function(err) {
			console.error("Revir UI hata:", err);
			showToast("Bağlantı hatası: " + err);
		})
		.finally(function() {
			inflight = false;
			setBusy(false);
			if (queuedReq) {
				var nxt = queuedReq;
				queuedReq = null;
				sendAction(nxt.payload, nxt.opts);
			}
		});
	}

	function applyRender(data, mainScroll, dlgScroll, sbScroll, refocusId) {
		if (!data) return;
		if (data.title) document.title = data.title;

		var dynStyles = document.getElementById("revir-dynamic-styles");
		if (dynStyles && dynStyles.dataset.hash !== String((data.styles_html || "").length) + ":" + (data.styles_html || "").slice(0, 64)) {
			dynStyles.innerHTML = data.styles_html || "";
			dynStyles.dataset.hash = String((data.styles_html || "").length) + ":" + (data.styles_html || "").slice(0, 64);
		}

		var sb = document.getElementById("revir-sidebar");
		var sbBtn = document.getElementById("revir-sidebar-open");
		if (sb) {
			sb.innerHTML = data.sidebar_html || "";
			if (data.sidebar_collapsed) {
				sb.classList.add("collapsed");
				if (sbBtn) sbBtn.classList.add("visible");
			} else {
				sb.classList.remove("collapsed");
				if (sbBtn) sbBtn.classList.remove("visible");
			}
			sb.scrollTop = sbScroll;
		}

		var mainInner = document.getElementById("revir-main-inner");
		if (mainInner) mainInner.innerHTML = data.main_html || "";

		var mainEl = document.getElementById("revir-main");
		if (mainEl) mainEl.scrollTop = mainScroll;

		var backdrop = document.getElementById("revir-modal-backdrop");
		var card = document.getElementById("revir-modal-card");
		var mTitle = document.getElementById("revir-modal-title");
		var mBody = document.getElementById("revir-modal-body");
		if (data.dialog_html !== null && data.dialog_html !== undefined) {
			if (mTitle) mTitle.textContent = data.dialog_title || "";
			if (mBody) {
				mBody.innerHTML = data.dialog_html;
				mBody.scrollTop = dlgScroll;
			}
			if (card) {
				card.classList.toggle("small", data.dialog_width === "small");
			}
			if (backdrop) backdrop.classList.add("open");
		} else {
			if (backdrop) backdrop.classList.remove("open");
			if (mBody) mBody.innerHTML = "";
		}

		// Expander durumlarını dinle
		document.querySelectorAll("details.revir-expander").forEach(function(det) {
			det.addEventListener("toggle", function() {
				fetch("/api/ui_state", {
					method: "POST",
					headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
					body: JSON.stringify({ kind: "expander", id: det.dataset.expId, value: det.open })
				}).catch(function(){});
			});
		});

		// Yardımcı betikleri çalıştır (Smart Input UX vb.)
		if (data.scripts_html) {
			var tmp = document.createElement("div");
			tmp.innerHTML = data.scripts_html;
			tmp.querySelectorAll("script").forEach(function(sc) {
				try {
					var fn = new Function(sc.textContent || "");
					fn();
				} catch (e) {}
			});
		}

		// Konfeti animasyonu
		if (data.balloons) launchConfetti();

		// Yüzen Gemini sohbet butonu görünürlüğü ve alt başlığı
		var fabBtn = document.getElementById("revir-fab-btn");
		var chatBubble = document.getElementById("revir-chat-bubble");
		var chatSub = document.getElementById("revir-chat-sub");
		if (fabBtn) {
			if (data.floating_chat_enabled) {
				fabBtn.classList.add("visible");
			} else {
				fabBtn.classList.remove("visible");
				if (chatBubble) chatBubble.classList.remove("open");
			}
		}
		if (chatSub && data.floating_chat_label) {
			chatSub.textContent = data.floating_chat_label;
		}

		// Düzenleyici odağını koru
		if (refocusId) {
			var el = document.getElementById(refocusId);
			if (el && el.focus) {
				try { el.focus(); } catch (e) {}
			}
		}
	}

	function launchConfetti() {
		var colors = ["#0d9488", "#2563eb", "#f59e0b", "#ef4444", "#10b981", "#ec4899"];
		for (var i = 0; i < 36; i++) {
			var dot = document.createElement("div");
			dot.style.cssText = "position:fixed;z-index:99999;pointer-events:none;width:10px;height:10px;border-radius:50%;top:-12px;"
				+ "left:" + (Math.random() * 100) + "vw;background:" + colors[i % colors.length]
				+ ";transition:transform 1.8s ease-in, opacity 1.8s ease;";
			document.body.appendChild(dot);
			(function(d) {
				setTimeout(function() {
					d.style.transform = "translateY(" + (window.innerHeight + 60) + "px) rotate(" + (Math.random() * 540) + "deg)";
					d.style.opacity = "0";
				}, 20);
				setTimeout(function() { d.remove(); }, 1950);
			})(dot);
		}
	}

	// Genel pencere olayları
	window.__revirMarkDirty = function(key, val) {
		pendingInputs[key] = val;
	};

	window.__revirFieldChange = function(key, val, el) {
		if (el && el.dataset && el.dataset.form) {
			return; // Form içi değişiklikler gönderimde iletilir
		}
		delete pendingInputs[key];
		sendAction({ action: "change", key: key, value: val });
	};

	window.__revirClick = function(key) {
		sendAction({ action: "click", key: key });
	};

	window.__revirRadioPick = function(key, idx, el) {
		var grp = el && el.closest ? el.closest("[data-kind='radio']") : null;
		if (grp) grp.dataset.value = String(idx);
		if (grp && grp.dataset && grp.dataset.form) return;
		sendAction({ action: "change", key: key, value: idx });
	};

	window.__revirSubmitForm = function(formId) {
		var form = document.querySelector('form[data-form-id="' + CSS.escape(formId) + '"]');
		var vals = {};
		if (form) {
			form.querySelectorAll("[data-key][data-kind]").forEach(function(el) {
				var k = el.dataset.key;
				var kind = el.dataset.kind;
				if (kind === "bool") vals[k] = !!el.checked;
				else if (kind === "selectbox") vals[k] = el.dataset.value === "" ? null : parseInt(el.dataset.value, 10);
				else if (kind === "multiselect") {
					try { vals[k] = JSON.parse(el.dataset.selected || "[]"); } catch(e) { vals[k] = []; }
				}
				else if (kind === "radio") vals[k] = parseInt(el.dataset.value || "0", 10);
				else vals[k] = el.value;
			});
		}
		sendAction({ action: "form_submit", form_id: formId, form_values: vals });
	};

	// Aranabilir Selectbox
	function renderComboMenu(key, filterTxt) {
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("combo_" + key);
		if (!wrap || !menu) return;
		var opts = JSON.parse(wrap.dataset.options || "[]");
		var curVal = wrap.dataset.value;
		var q = (filterTxt || "").toLocaleLowerCase("tr-TR").trim();
		var html = [];
		for (var i = 0; i < opts.length; i++) {
			var txt = String(opts[i]);
			if (q && txt.toLocaleLowerCase("tr-TR").indexOf(q) === -1) continue;
			var sel = String(i) === curVal ? " selected" : "";
			html.push('<div class="revir-combo-opt' + sel + '" onmousedown="window.__revirSelectPick(\'' + key.replace(/'/g, "\\'") + '\', ' + i + ', event)">' + escHtml(txt) + '</div>');
			if (html.length >= 200) break;
		}
		if (!html.length) html.push('<div class="revir-combo-opt" style="color:var(--text-muted);cursor:default;">Sonuç bulunamadı</div>');
		menu.innerHTML = html.join("");
	}

	window.__revirOpenCombo = function(key, ev) {
		if (ev) ev.stopPropagation();
		closeAllMenus("combo_" + key);
		var menu = document.getElementById("combo_" + key);
		if (!menu) return;
		renderComboMenu(key, "");
		menu.classList.add("open");
	};

	window.__revirFilterCombo = function(key, val) {
		var menu = document.getElementById("combo_" + key);
		if (menu && !menu.classList.contains("open")) menu.classList.add("open");
		renderComboMenu(key, val);
	};

	window.__revirBlurCombo = function(key, inp) {
		setTimeout(function() {
			var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
			if (!wrap || !inp) return;
			var opts = JSON.parse(wrap.dataset.options || "[]");
			var v = wrap.dataset.value;
			inp.value = (v !== "" && v !== undefined && opts[parseInt(v, 10)] !== undefined) ? opts[parseInt(v, 10)] : "";
		}, 140);
	};

	window.__revirSelectPick = function(key, idx, ev) {
		if (ev) { ev.preventDefault(); ev.stopPropagation(); }
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("combo_" + key);
		if (menu) menu.classList.remove("open");
		if (!wrap) return;
		var opts = JSON.parse(wrap.dataset.options || "[]");
		wrap.dataset.value = String(idx);
		var inp = wrap.querySelector(".revir-combo-search");
		if (inp) inp.value = opts[idx] || "";
		if (wrap.dataset.form) return;
		sendAction({ action: "change", key: key, value: idx });
	};

	window.__revirSelectClear = function(key, ev) {
		if (ev) { ev.preventDefault(); ev.stopPropagation(); }
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		if (!wrap) return;
		wrap.dataset.value = "";
		var inp = wrap.querySelector(".revir-combo-search");
		if (inp) inp.value = "";
		if (wrap.dataset.form) return;
		sendAction({ action: "change", key: key, value: null });
	};

	// Multiselect
	function renderMultiMenu(key, filterTxt) {
		var wrap = document.querySelector('.revir-multi[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("multi_" + key);
		if (!wrap || !menu) return;
		var opts = JSON.parse(wrap.dataset.options || "[]");
		var sel = JSON.parse(wrap.dataset.selected || "[]");
		var q = (filterTxt || "").toLocaleLowerCase("tr-TR").trim();
		var html = [];
		for (var i = 0; i < opts.length; i++) {
			var txt = String(opts[i]);
			if (q && txt.toLocaleLowerCase("tr-TR").indexOf(q) === -1) continue;
			var isSel = sel.indexOf(i) !== -1;
			html.push('<div class="revir-combo-opt' + (isSel ? " selected" : "") + '" onmousedown="window.__revirMultiToggle(\'' + key.replace(/'/g, "\\'") + '\', ' + i + ', event)">'
				+ '<span>' + (isSel ? "☑" : "☐") + '</span><span>' + escHtml(txt) + '</span></div>');
			if (html.length >= 200) break;
		}
		if (!html.length) html.push('<div class="revir-combo-opt" style="color:var(--text-muted);">Sonuç bulunamadı</div>');
		menu.innerHTML = html.join("");
	}

	window.__revirOpenMulti = function(key, ev) {
		if (ev) ev.stopPropagation();
		closeAllMenus("multi_" + key);
		var menu = document.getElementById("multi_" + key);
		if (!menu) return;
		renderMultiMenu(key, "");
		menu.classList.add("open");
	};

	window.__revirFilterMulti = function(key, val) {
		var menu = document.getElementById("multi_" + key);
		if (menu && !menu.classList.contains("open")) menu.classList.add("open");
		renderMultiMenu(key, val);
	};

	window.__revirMultiToggle = function(key, idx, ev) {
		if (ev) { ev.preventDefault(); ev.stopPropagation(); }
		var wrap = document.querySelector('.revir-multi[data-key="' + CSS.escape(key) + '"]');
		if (!wrap) return;
		var sel = JSON.parse(wrap.dataset.selected || "[]");
		var pos = sel.indexOf(idx);
		if (pos === -1) sel.push(idx);
		else sel.splice(pos, 1);
		wrap.dataset.selected = JSON.stringify(sel);
		if (wrap.dataset.form) {
			renderMultiMenu(key, "");
			rebuildMultiChips(wrap, key, sel);
			return;
		}
		sendAction({ action: "change", key: key, value: sel });
	};

	window.__revirMultiRemove = function(key, idx, ev) {
		if (ev) { ev.preventDefault(); ev.stopPropagation(); }
		window.__revirMultiToggle(key, idx, ev);
	};

	function rebuildMultiChips(wrap, key, sel) {
		var opts = JSON.parse(wrap.dataset.options || "[]");
		var ctrl = wrap.querySelector(".revir-multi-control");
		var inp = wrap.querySelector(".revir-multi-search");
		if (!ctrl || !inp) return;
		ctrl.querySelectorAll(".revir-chip").forEach(function(c) { c.remove(); });
		sel.forEach(function(i) {
			var sp = document.createElement("span");
			sp.className = "revir-chip";
			sp.innerHTML = escHtml(opts[i] || "") + '<button type="button" class="revir-chip-x" onclick="window.__revirMultiRemove(\'' + key.replace(/'/g, "\\'") + '\', ' + i + ', event)">×</button>';
			ctrl.insertBefore(sp, inp);
		});
	}

	// Sekmeler (Anlık geçiş + arka plan kalıcılığı)
	window.__revirSwitchTab = function(groupId, idx) {
		document.querySelectorAll('.revir-tab-btn[data-tab-group="' + groupId + '"]').forEach(function(b) {
			b.classList.toggle("active", parseInt(b.dataset.tabIdx, 10) === idx);
		});
		document.querySelectorAll('.revir-tab-pane[data-tab-group="' + groupId + '"]').forEach(function(p) {
			p.classList.toggle("active", parseInt(p.dataset.tabIdx, 10) === idx);
		});
		fetch("/api/ui_state", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({ kind: "tab", id: groupId, value: idx })
		}).catch(function(){});
	};

	// Popover
	window.__revirTogglePopover = function(popId, ev) {
		if (ev) ev.stopPropagation();
		var el = document.getElementById("pop_" + popId);
		if (!el) return;
		var wasOpen = el.classList.contains("open");
		closeAllMenus();
		if (!wasOpen) el.classList.add("open");
	};

	function closeAllMenus(exceptId) {
		document.querySelectorAll(".revir-combo-menu.open, .revir-popover-body.open").forEach(function(m) {
			if (m.id !== exceptId) m.classList.remove("open");
		});
	}

	// Dataframe seçimi
	window.__revirDfSelect = function(key, rowIdx, colName) {
		sendAction({ action: "df_select", key: key, value: { rows: [rowIdx], cells: [[rowIdx, colName]] } });
	};

	// İstemci tarafı tablo sıralama
	window.__revirSortTable = function(th) {
		var tbl = th.closest("table");
		if (!tbl || tbl.classList.contains("revir-editor-table")) return;
		var tbody = tbl.querySelector("tbody");
		if (!tbody) return;
		var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
		var asc = th.dataset.sortDir !== "asc";
		th.dataset.sortDir = asc ? "asc" : "desc";
		var rows = Array.prototype.slice.call(tbody.querySelectorAll("tr"));
		rows.sort(function(a, b) {
			var av = (a.children[idx] ? a.children[idx].textContent : "").trim();
			var bv = (b.children[idx] ? b.children[idx].textContent : "").trim();
			var an = parseFloat(av.replace(",", "."));
			var bn = parseFloat(bv.replace(",", "."));
			if (!isNaN(an) && !isNaN(bn) && String(an).length === av.replace(",", ".").length && String(bn).length === bv.replace(",", ".").length) {
				return asc ? (an - bn) : (bn - an);
			}
			return asc ? av.localeCompare(bv, "tr") : bv.localeCompare(av, "tr");
		});
		rows.forEach(function(r) { tbody.appendChild(r); });
	};

	// Data Editor
	window.__revirEditorDirty = function(key, rowKind, rowId, col, val) {
		pendingEditorCells.push({ key: key, op: "cell", row_kind: rowKind, row_id: rowId, col: col, val: val });
	};

	window.__revirEditorCell = function(key, rowKind, rowId, col, val, cellId) {
		pendingEditorCells = pendingEditorCells.filter(function(m) {
			return !(m.key === key && m.row_kind === rowKind && m.row_id === rowId && m.col === col);
		});
		sendAction(
			{ action: "data_editor", key: key, value: { op: "cell", row_kind: rowKind, row_id: rowId, col: col, val: val } },
			{ refocusId: cellId }
		);
	};

	window.__revirEditorAddRow = function(key) {
		sendAction({ action: "data_editor", key: key, value: { op: "add_row" } });
	};

	window.__revirEditorDeleteRow = function(key, rowKind, rowId) {
		sendAction({ action: "data_editor", key: key, value: { op: "delete_row", row_kind: rowKind, row_id: rowId } });
	};

	// Dosya Yükleme
	window.__revirUploadFiles = function(fileList, key, multiple) {
		if (!fileList || !fileList.length) return;
		var arr = Array.prototype.slice.call(fileList);
		if (!multiple) arr = [arr[0]];
		var promises = arr.map(function(f) {
			return new Promise(function(resolve, reject) {
				var reader = new FileReader();
				reader.onload = function() {
					var b64 = String(reader.result || "").split(",")[1] || "";
					resolve({ name: f.name, type: f.type || "application/octet-stream", data_b64: b64 });
				};
				reader.onerror = reject;
				reader.readAsDataURL(f);
			});
		});
		setBusy(true);
		Promise.all(promises).then(function(encoded) {
			return fetch("/api/upload", {
				method: "POST",
				headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
				body: JSON.stringify({ key: key, multiple: !!multiple, files: encoded })
			});
		}).then(function() {
			sendAction({ action: "init" });
		}).catch(function(err) {
			setBusy(false);
			showToast("Dosya yükleme hatası: " + err);
		});
	};

	window.__revirHandleDrop = function(ev, key, multiple) {
		ev.preventDefault();
		ev.currentTarget.classList.remove("dragover");
		if (ev.dataTransfer && ev.dataTransfer.files) {
			window.__revirUploadFiles(ev.dataTransfer.files, key, multiple);
		}
	};

	window.__revirClearUpload = function(key, fileId) {
		sendAction({ action: "clear_upload", key: key, value: fileId });
	};

	// Yerel Masaüstü Dosya İndirme (Native Save Dialog destekli)
	window.__revirDownload = function(token, fileName) {
		if (window.pywebview && window.pywebview.api && window.pywebview.api.save_download) {
			window.pywebview.api.save_download(token, fileName).then(function(res) {
				if (res && res.saved) showToast("✅ Kaydedildi: " + res.path);
				else if (res && res.error) showToast("❌ İndirme hatası: " + res.error);
			}).catch(function() {
				fallbackDownload(token, fileName);
			});
		} else {
			fallbackDownload(token, fileName);
		}
	};

	function fallbackDownload(token, fileName) {
		var a = document.createElement("a");
		a.href = "/api/download/" + encodeURIComponent(token);
		a.download = fileName || "dosya";
		document.body.appendChild(a);
		a.click();
		a.remove();
	}

	window.__revirOpenExternal = function(url) {
		fetch("/api/open_external", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({ url: url })
		}).catch(function(){});
	};

	// Diyalog & Yan Menü
	window.__revirCloseDialog = function() {
		sendAction({ action: "close_dialog" });
	};

	window.__revirSetSidebar = function(collapsed) {
		var sb = document.getElementById("revir-sidebar");
		var btn = document.getElementById("revir-sidebar-open");
		if (sb) sb.classList.toggle("collapsed", !!collapsed);
		if (btn) btn.classList.toggle("visible", !!collapsed);
		fetch("/api/ui_state", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({ kind: "sidebar", value: !!collapsed })
		}).catch(function(){});
	};

	function escHtml(s) {
		return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
	}

	// Genel tıklama & klavye delegasyonu
	document.addEventListener("click", function(ev) {
		var closeSb = ev.target && ev.target.closest ? ev.target.closest(".revir-sidebar-close-btn") : null;
		if (closeSb) { window.__revirSetSidebar(true); return; }
		if (!ev.target.closest(".revir-combo") && !ev.target.closest(".revir-multi") && !ev.target.closest(".revir-popover-wrap")) {
			closeAllMenus();
		}
	});

	document.addEventListener("keydown", function(ev) {
		if (ev.key === "Escape") {
			closeAllMenus();
			var bd = document.getElementById("revir-modal-backdrop");
			if (bd && bd.classList.contains("open")) window.__revirCloseDialog();
		}
	});

	// SVG Grafik Hover Tooltip
	document.addEventListener("mousemove", function(ev) {
		var tt = document.getElementById("revir-chart-tooltip");
		if (!tt) return;
		var target = ev.target && ev.target.closest ? ev.target.closest("[data-tip]") : null;
		if (target && target.getAttribute("data-tip")) {
			tt.textContent = target.getAttribute("data-tip");
			tt.style.display = "block";
			var lx = Math.min(window.innerWidth - 360, ev.clientX + 14);
			var ly = Math.max(10, ev.clientY - 38);
			tt.style.left = lx + "px";
			tt.style.top = ly + "px";
		} else {
			tt.style.display = "none";
		}
	});

	// İlk yükleme (Sunucu tarafından önceden işlenmiş veri varsa anında uygula)
	var chatFiles = [];
	var chatTtsEnabled = false;
	var chatBusy = false;
	var voiceRec = null;
	var mediaRecorder = null;
	var audioChunks = [];
	var isRecordingVoice = false;

	window.__revirToggleChatBubble = function() {
		var b = document.getElementById("revir-chat-bubble");
		if (!b) return;
		b.classList.toggle("open");
		if (b.classList.contains("open")) {
			var inp = document.getElementById("revir-chat-inp");
			if (inp) setTimeout(function() { inp.focus(); }, 60);
		}
	};

	window.__revirExpandChat = function() {
		var b = document.getElementById("revir-chat-bubble");
		if (b) b.classList.toggle("expanded");
	};

	window.__revirToggleChatTTS = function() {
		chatTtsEnabled = !chatTtsEnabled;
		var btn = document.getElementById("revir-chat-tts-btn");
		if (btn) {
			btn.textContent = chatTtsEnabled ? "🔊" : "🔇";
			btn.classList.toggle("active", chatTtsEnabled);
		}
		if (!chatTtsEnabled && window.speechSynthesis) {
			try { window.speechSynthesis.cancel(); } catch (e) {}
		}
		showToast(chatTtsEnabled ? "🔊 Sesli yanıt okuma açıldı" : "🔇 Sesli yanıt okuma kapatıldı");
	};

	window.__revirSpeakText = function(txt) {
		if (!window.speechSynthesis || !txt) return;
		try {
			window.speechSynthesis.cancel();
			var clean = String(txt).replace(/[*#_`~>-]+/g, " ").replace(/\s+/g, " ").trim();
			var utt = new SpeechSynthesisUtterance(clean);
			utt.lang = "tr-TR";
			utt.rate = 1.04;
			var voices = window.speechSynthesis.getVoices() || [];
			for (var i = 0; i < voices.length; i++) {
				if (voices[i].lang && voices[i].lang.toLowerCase().indexOf("tr") === 0) {
					utt.voice = voices[i];
					break;
				}
			}
			window.speechSynthesis.speak(utt);
		} catch (e) {}
	};

	function renderChatFileChips() {
		var box = document.getElementById("revir-chat-files");
		if (!box) return;
		if (!chatFiles.length) {
			box.classList.remove("has-files");
			box.innerHTML = "";
			return;
		}
		box.classList.add("has-files");
		box.innerHTML = chatFiles.map(function(f, idx) {
			return '<span class="revir-chat-fchip">📎 ' + escHtml(f.name)
				+ ' <button type="button" class="revir-chip-x" onclick="window.__revirRemoveChatFile(' + idx + ')">×</button></span>';
		}).join("");
	}

	window.__revirRemoveChatFile = function(idx) {
		chatFiles.splice(idx, 1);
		renderChatFileChips();
	};

	window.__revirChatAddFiles = function(fileList) {
		if (!fileList || !fileList.length) return;
		Array.prototype.slice.call(fileList).forEach(function(f) {
			if (chatFiles.length >= 5) return;
			var reader = new FileReader();
			reader.onload = function() {
				var b64 = String(reader.result || "").split(",")[1] || "";
				chatFiles.push({ name: f.name, type: f.type || "application/octet-stream", data_b64: b64 });
				renderChatFileChips();
			};
			reader.readAsDataURL(f);
		});
	};

	window.__revirChatDropFiles = function(ev) {
		ev.preventDefault();
		if (ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files.length) {
			window.__revirChatAddFiles(ev.dataTransfer.files);
		}
	};

	function appendChatMsg(role, htmlContent, plainText) {
		var box = document.getElementById("revir-chat-messages");
		if (!box) return null;
		var div = document.createElement("div");
		div.className = "revir-cmsg " + (role === "user" ? "revir-cmsg-user" : "revir-cmsg-ai");
		div.innerHTML = htmlContent;
		if (role === "ai" && plainText) {
			var tools = document.createElement("div");
			tools.className = "revir-cmsg-tools";
			var spkBtn = document.createElement("button");
			spkBtn.type = "button";
			spkBtn.className = "revir-cmsg-tbtn";
			spkBtn.textContent = "🔊 Dinle";
			spkBtn.onclick = function() { window.__revirSpeakText(plainText); };
			var cpBtn = document.createElement("button");
			cpBtn.type = "button";
			cpBtn.className = "revir-cmsg-tbtn";
			cpBtn.textContent = "📋 Kopyala";
			cpBtn.onclick = function() {
				if (navigator.clipboard) navigator.clipboard.writeText(plainText);
				showToast("📋 Yanıt panoya kopyalandı");
			};
			tools.appendChild(spkBtn);
			tools.appendChild(cpBtn);
			div.appendChild(tools);
		}
		box.appendChild(div);
		box.scrollTop = box.scrollHeight;
		return div;
	}

	window.__revirQuickPrompt = function(txt) {
		var inp = document.getElementById("revir-chat-inp");
		if (inp) inp.value = txt;
		window.__revirSendChat();
	};

	window.__revirClearChat = function() {
		var box = document.getElementById("revir-chat-messages");
		if (box) {
			box.innerHTML = '<div class="revir-cmsg revir-cmsg-ai"><p>Sohbet geçmişi temizlendi. Size nasıl yardımcı olabilirim?</p></div>';
		}
		if (window.speechSynthesis) {
			try { window.speechSynthesis.cancel(); } catch (e) {}
		}
		fetch("/api/gemini_chat_clear", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: "{}"
		}).catch(function(){});
	};

	window.__revirSendChat = function(voicePayload) {
		if (chatBusy) return;
		var inp = document.getElementById("revir-chat-inp");
		var msgTxt = inp ? inp.value.trim() : "";
		var filesToSend = chatFiles.slice();
		var hasVoice = !!(voicePayload && voicePayload.data_b64);

		if (!msgTxt && !filesToSend.length && !hasVoice) return;

		if (inp) inp.value = "";
		chatFiles = [];
		renderChatFileChips();

		var badges = [];
		filesToSend.forEach(function(f) { badges.push("📎 " + escHtml(f.name)); });
		if (hasVoice) badges.push("🎙️ Sesli Mesaj");
		var userHtml = "";
		if (badges.length) {
			userHtml += '<div style="font-size:11.5px;opacity:0.9;margin-bottom:4px;">' + badges.join(" · ") + "</div>";
		}
		if (msgTxt) {
			userHtml += "<p>" + escHtml(msgTxt) + "</p>";
		}
		var userBubble = appendChatMsg("user", userHtml || "<p>🎙️ Sesli mesaj gönderildi...</p>", "");
		var waitBubble = appendChatMsg("ai", '<p style="color:var(--text-muted);">⏳ <em>Gemini yanıt hazırlıyor...</em></p>', "");
		chatBusy = true;

		fetch("/api/gemini_chat", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({
				message: msgTxt,
				attachments: filesToSend,
				voice_audio_b64: hasVoice ? voicePayload.data_b64 : "",
				voice_audio_mime: hasVoice ? (voicePayload.mime || "audio/webm") : ""
			})
		})
		.then(function(r) { return r.json(); })
		.then(function(res) {
			if (waitBubble) waitBubble.remove();
			if (res.transcribed && userBubble && !msgTxt) {
				userBubble.innerHTML = '<div style="font-size:11.5px;opacity:0.9;margin-bottom:4px;">🎙️ Sesli Mesaj</div><p>' + escHtml(res.transcribed) + "</p>";
			}
			if (res.ok) {
				appendChatMsg("ai", res.reply_html || ("<p>" + escHtml(res.reply_text || "") + "</p>"), res.reply_text || "");
				if (chatTtsEnabled || hasVoice) {
					window.__revirSpeakText(res.reply_text || "");
				}
			} else {
				appendChatMsg("ai", '<p style="color:#dc2626;">❌ ' + escHtml(res.error || "Yanıt alınamadı.") + "</p>", "");
			}
		})
		.catch(function(err) {
			if (waitBubble) waitBubble.remove();
			appendChatMsg("ai", '<p style="color:#dc2626;">❌ Bağlantı hatası: ' + escHtml(String(err)) + "</p>", "");
		})
		.finally(function() {
			chatBusy = false;
		});
	};

	// Sesli Konuşma: Web Speech API + MediaRecorder (Gemini Multimodal Audio) Çift Motorlu
	window.__revirToggleVoiceMic = function() {
		var micBtn = document.getElementById("revir-chat-mic-btn");
		if (isRecordingVoice) {
			stopVoiceCapture();
			return;
		}
		var SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
		if (SpeechRec && !window.__revirForceMediaRecorder) {
			try {
				voiceRec = new SpeechRec();
				voiceRec.lang = "tr-TR";
				voiceRec.interimResults = true;
				voiceRec.continuous = false;
				var finalTranscript = "";
				isRecordingVoice = true;
				if (micBtn) { micBtn.classList.add("recording"); micBtn.textContent = "⏹️"; }
				showToast("🎙️ Dinleniyor... Konuşmanız bitince ⏹️ tuşuna basın.");

				voiceRec.onresult = function(ev) {
					var interim = "";
					for (var i = ev.resultIndex; i < ev.results.length; i++) {
						if (ev.results[i].isFinal) finalTranscript += ev.results[i][0].transcript;
						else interim += ev.results[i][0].transcript;
					}
					var inp = document.getElementById("revir-chat-inp");
					if (inp) inp.value = (finalTranscript + " " + interim).trim();
				};
				voiceRec.onerror = function() {
					window.__revirForceMediaRecorder = true;
					isRecordingVoice = false;
					if (micBtn) { micBtn.classList.remove("recording"); micBtn.textContent = "🎙️"; }
					startMediaRecorderCapture();
				};
				voiceRec.onend = function() {
					var wasRec = isRecordingVoice;
					isRecordingVoice = false;
					if (micBtn) { micBtn.classList.remove("recording"); micBtn.textContent = "🎙️"; }
					var inp = document.getElementById("revir-chat-inp");
					if (wasRec && inp && inp.value.trim()) {
						chatTtsEnabled = true;
						var ttsBtn = document.getElementById("revir-chat-tts-btn");
						if (ttsBtn) { ttsBtn.textContent = "🔊"; ttsBtn.classList.add("active"); }
						window.__revirSendChat();
					}
				};
				voiceRec.start();
				return;
			} catch (e) {
				window.__revirForceMediaRecorder = true;
			}
		}
		startMediaRecorderCapture();
	};

	function startMediaRecorderCapture() {
		var micBtn = document.getElementById("revir-chat-mic-btn");
		if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
			showToast("⚠️ Bu cihazda mikrofon erişimi desteklenmiyor.");
			return;
		}
		navigator.mediaDevices.getUserMedia({ audio: true }).then(function(stream) {
			audioChunks = [];
			var mime = "audio/webm";
			if (window.MediaRecorder && MediaRecorder.isTypeSupported) {
				if (MediaRecorder.isTypeSupported("audio/webm")) mime = "audio/webm";
				else if (MediaRecorder.isTypeSupported("audio/mp4")) mime = "audio/mp4";
				else if (MediaRecorder.isTypeSupported("audio/wav")) mime = "audio/wav";
			}
			mediaRecorder = new MediaRecorder(stream);
			isRecordingVoice = true;
			if (micBtn) { micBtn.classList.add("recording"); micBtn.textContent = "⏹️"; }
			showToast("🎙️ Ses kaydediliyor... Göndermek için ⏹️ tuşuna basın.");

			mediaRecorder.ondataavailable = function(e) {
				if (e.data && e.data.size > 0) audioChunks.push(e.data);
			};
			mediaRecorder.onstop = function() {
				stream.getTracks().forEach(function(t) { t.stop(); });
				isRecordingVoice = false;
				if (micBtn) { micBtn.classList.remove("recording"); micBtn.textContent = "🎙️"; }
				if (!audioChunks.length) return;
				var blob = new Blob(audioChunks, { type: mime });
				var reader = new FileReader();
				reader.onload = function() {
					var b64 = String(reader.result || "").split(",")[1] || "";
					if (b64) {
						chatTtsEnabled = true;
						var ttsBtn = document.getElementById("revir-chat-tts-btn");
						if (ttsBtn) { ttsBtn.textContent = "🔊"; ttsBtn.classList.add("active"); }
						window.__revirSendChat({ data_b64: b64, mime: mime });
					}
				};
				reader.readAsDataURL(blob);
			};
			mediaRecorder.start();
		}).catch(function(err) {
			isRecordingVoice = false;
			if (micBtn) { micBtn.classList.remove("recording"); micBtn.textContent = "🎙️"; }
			showToast("⚠️ Mikrofon izni alınamadı: " + err.message);
		});
	}

	function stopVoiceCapture() {
		var micBtn = document.getElementById("revir-chat-mic-btn");
		if (voiceRec) {
			try { voiceRec.stop(); } catch (e) {}
			voiceRec = null;
		}
		if (mediaRecorder && mediaRecorder.state !== "inactive") {
			try { mediaRecorder.stop(); } catch (e) {}
			mediaRecorder = null;
		}
		isRecordingVoice = false;
		if (micBtn) { micBtn.classList.remove("recording"); micBtn.textContent = "🎙️"; }
	}

	var INITIAL_DATA = __REVIR_INITIAL_JSON__;
	function bootInitial() {
		if (INITIAL_DATA && INITIAL_DATA.main_html) {
			applyRender(INITIAL_DATA, 0, 0, 0);
		} else {
			sendAction({ action: "init" });
		}
	}
	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", bootInitial);
	} else {
		bootInitial();
	}
})();
</script>
</body>
</html>
"""


# ==============================================================================
# 10. GÖMÜLÜ YEREL HTTP SUNUCUSU (Sadece 127.0.0.1)
# ==============================================================================

class LocalUIServer:
	"""Python süreç içi yerel arayüz sunucusu (alt süreç veya harici çerçeve gerektirmez)."""
	def __init__(self, app_fn: Callable[[], None], host: str = "127.0.0.1", port: int = 0) -> None:
		self.app_fn = app_fn
		self.host = host
		self.session = _ACTIVE_SESSION
		self.csrf_token = secrets.token_urlsafe(24)
		self._shell_bytes = _SHELL_HTML.replace("__REVIR_CSRF__", self.csrf_token).encode("utf-8")
		self._httpd = ThreadingHTTPServer((host, int(port)), self._build_handler())
		self._httpd.daemon_threads = True
		self.port = int(self._httpd.server_address[1])
		self.url = f"http://{self.host}:{self.port}"
		self._thread: threading.Thread | None = None

	def _build_handler(self):
		server_ref = self

		class _Handler(BaseHTTPRequestHandler):
			protocol_version = "HTTP/1.1"

			def log_message(self, format: str, *args) -> None:
				return  # Sessiz yerel masaüstü sunucusu

			def _send_bytes(self, status: int, content_type: str, body: bytes, extra_headers: dict[str, str] | None = None) -> None:
				self.send_response(status)
				self.send_header("Content-Type", content_type)
				self.send_header("Content-Length", str(len(body)))
				self.send_header("Cache-Control", "no-store")
				if extra_headers:
					for k, v in extra_headers.items():
						self.send_header(k, v)
				self.end_headers()
				self.wfile.write(body)

			def _read_json(self) -> dict[str, Any]:
				length = int(self.headers.get("Content-Length") or 0)
				raw = self.rfile.read(length) if length > 0 else b"{}"
				return json.loads(raw.decode("utf-8")) if raw else {}

			def _check_csrf(self) -> bool:
				return self.headers.get("X-Revir-Token") == server_ref.csrf_token

			def do_GET(self) -> None:
				parsed = urllib.parse.urlparse(self.path)
				path = parsed.path
				if path == "/" or path == "/index.html":
					try:
						init_res = _execute_render_cycle(server_ref.app_fn, server_ref.session, {"action": "init"})
						init_json = json.dumps(init_res, ensure_ascii=False).replace("</", "<\\/")
						init_styles = init_res.get("styles_html") or ""
						init_sb = init_res.get("sidebar_html") or ""
						init_main = init_res.get("main_html") or ""
					except Exception:
						init_json = "null"
						init_styles = ""
						init_sb = ""
						init_main = ""
					page_html = (
						_SHELL_HTML
						.replace("__REVIR_CSRF__", server_ref.csrf_token)
						.replace("__REVIR_Init_STYLES__" if False else "__REVIR_INIT_STYLES__", init_styles)
						.replace("__REVIR_INIT_SIDEBAR__", init_sb)
						.replace("__REVIR_INIT_MAIN__", init_main)
						.replace("__REVIR_INITIAL_JSON__", init_json)
					)
					self._send_bytes(HTTPStatus.OK, "text/html; charset=utf-8", page_html.encode("utf-8"))
					return
				if path == "/health":
					self._send_bytes(HTTPStatus.OK, "text/plain; charset=utf-8", b"ok")
					return
				if path.startswith("/api/media/"):
					token = urllib.parse.unquote(path[len("/api/media/"):])
					item = server_ref.session.media.get(token)
					if item is None:
						self._send_bytes(HTTPStatus.NOT_FOUND, "text/plain", b"Not Found")
						return
					self._send_bytes(HTTPStatus.OK, item[1], item[0])
					return
				if path.startswith("/api/download/"):
					token = urllib.parse.unquote(path[len("/api/download/"):])
					item = server_ref.session.downloads.get(token)
					if item is None:
						self._send_bytes(HTTPStatus.NOT_FOUND, "text/plain", b"Not Found")
						return
					data, fname, mime = item
					safe_ascii = re.sub(r"[^A-Za-z0-9._-]+", "_", fname) or "dosya"
					utf8_fn = urllib.parse.quote(fname)
					self._send_bytes(
						HTTPStatus.OK,
						mime,
						data,
						{"Content-Disposition": f"attachment; filename=\"{safe_ascii}\"; filename*=UTF-8''{utf8_fn}"},
					)
					return
				self._send_bytes(HTTPStatus.NOT_FOUND, "text/plain", b"Not Found")

			def do_POST(self) -> None:
				if not self._check_csrf():
					self._send_bytes(HTTPStatus.FORBIDDEN, "application/json", b'{"error":"invalid_token"}')
					return
				parsed = urllib.parse.urlparse(self.path)
				path = parsed.path

				if path == "/api/action":
					try:
						payload = self._read_json()
						res = _execute_render_cycle(server_ref.app_fn, server_ref.session, payload)
						body = json.dumps(res, ensure_ascii=False).encode("utf-8")
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", body)
					except Exception as exc:
						err_body = json.dumps({"error": str(exc)}).encode("utf-8")
						self._send_bytes(HTTPStatus.INTERNAL_SERVER_ERROR, "application/json", err_body)
					return

				if path == "/api/ui_state":
					try:
						payload = self._read_json()
						kind = payload.get("kind")
						with server_ref.session.lock:
							if kind == "tab":
								server_ref.session.active_tabs[str(payload.get("id"))] = int(payload.get("value", 0))
							elif kind == "expander":
								server_ref.session.open_expanders[str(payload.get("id"))] = bool(payload.get("value", False))
							elif kind == "sidebar":
								server_ref.session.sidebar_collapsed = bool(payload.get("value", False))
						self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					except Exception:
						self._send_bytes(HTTPStatus.BAD_REQUEST, "application/json", b'{"ok":false}')
					return

				if path == "/api/upload":
					try:
						payload = self._read_json()
						wid = str(payload.get("key") or "")
						multiple = bool(payload.get("multiple", False))
						raw_files = payload.get("files") or []
						new_objs: list[UploadedFile] = []
						for rf in raw_files:
							name = str(rf.get("name") or "dosya")
							mime = str(rf.get("type") or "application/octet-stream")
							raw_bytes = base64.b64decode(rf.get("data_b64") or "")
							new_objs.append(UploadedFile(name=name, data=raw_bytes, mime_type=mime))
						with server_ref.session.lock:
							if multiple:
								existing = list(server_ref.session.uploaded_files.get(wid) or [])
								server_ref.session.uploaded_files[wid] = existing + new_objs
							else:
								server_ref.session.uploaded_files[wid] = new_objs[:1]
						self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					except Exception as exc:
						self._send_bytes(HTTPStatus.BAD_REQUEST, "application/json", json.dumps({"error": str(exc)}).encode("utf-8"))
					return

				if path == "/api/open_external":
					try:
						payload = self._read_json()
						url = str(payload.get("url") or "").strip()
						if url.startswith(("http://", "https://", "mailto:", "tel:")):
							webbrowser.open(url)
						self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					except Exception:
						self._send_bytes(HTTPStatus.BAD_REQUEST, "application/json", b'{"ok":false}')
					return

				if path == "/api/gemini_chat_clear":
					with server_ref.session.lock:
						server_ref.session.chat_history.clear()
					self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					return

				if path == "/api/gemini_chat":
					try:
						payload = self._read_json()
						handler = server_ref.session.floating_chat_handler
						if not callable(handler):
							self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": "Gemini sohbet işleyicisi hazır değil."}).encode("utf-8"))
							return
						res_dict = handler(payload, server_ref.session.chat_history)
						if res_dict.get("ok"):
							u_txt = str(res_dict.get("transcribed") or payload.get("message") or "").strip()
							r_txt = str(res_dict.get("reply_text") or "").strip()
							if u_txt:
								server_ref.session.chat_history.append({"role": "user", "text": u_txt})
							if r_txt:
								server_ref.session.chat_history.append({"role": "assistant", "text": r_txt})
							res_dict["reply_html"] = _render_markdown(r_txt, unsafe_allow_html=False)
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", json.dumps(res_dict, ensure_ascii=False).encode("utf-8"))
					except Exception as exc:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": str(exc)}).encode("utf-8"))
					return

				self._send_bytes(HTTPStatus.NOT_FOUND, "application/json", b'{"error":"not_found"}')

		return _Handler

	def start(self) -> str:
		self._thread = threading.Thread(target=self._httpd.serve_forever, kwargs={"poll_interval": 0.25}, daemon=True)
		self._thread.start()
		return self.url

	def stop(self) -> None:
		try:
			self._httpd.shutdown()
			self._httpd.server_close()
		except Exception:
			pass

	def get_download_item(self, token: str) -> tuple[bytes, str, str] | None:
		return self.session.downloads.get(token)


def reset_session() -> None:
	"""Test veya oturum sıfırlama için aktif oturumu temizler."""
	global _ACTIVE_SESSION
	_ACTIVE_SESSION = _AppSession()


def run_action_headless(app_fn: Callable[[], None], payload: dict[str, Any] | None = None) -> dict[str, Any]:
	"""Otomatik doğrulama ve testler için başsız render döngüsü çalıştırır."""
	return _execute_render_cycle(app_fn, _ACTIVE_SESSION, payload or {"action": "init"})
