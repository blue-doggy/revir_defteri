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
	mode: str = "modal"


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


_ALL_DATA_CACHES: list[Callable[[], None]] = []


def clear_all_caches() -> None:
	"""Tüm @ui.cache_data önbelleklerini tek seferde temizler."""
	for clr in list(_ALL_DATA_CACHES):
		try:
			clr()
		except Exception:
			pass


def clear_ephemeral_widget_values() -> None:
	"""Anahtarsız (key=None) widget önbelleklerini ve tablo düzenleyici hash'lerini temizler."""
	sessions = [_ACTIVE_SESSION]
	try:
		ctx = getattr(_LOCAL_CTX, "ctx", None)
		if ctx is not None and getattr(ctx, "session", None) is not None and ctx.session is not _ACTIVE_SESSION:
			sessions.append(ctx.session)
	except Exception:
		pass
	for sess in sessions:
		try:
			with sess.lock:
				sess.widget_values.clear()
				sess.editor_input_hashes.clear()
		except Exception:
			pass


def clear_dataframe_selection(key: str | None = None) -> None:
	"""Tek seferlik tablo/hücre seçimini temizler (ör. '📂 Aç' ile modal açıldıktan sonra sonraki tıklamalarda yeniden tetiklenmeyi önler)."""
	sessions = [_ACTIVE_SESSION]
	try:
		ctx = getattr(_LOCAL_CTX, "ctx", None)
		if ctx is not None and getattr(ctx, "session", None) is not None and ctx.session is not _ACTIVE_SESSION:
			sessions.append(ctx.session)
	except Exception:
		pass
	for sess in sessions:
		try:
			with sess.lock:
				if key is None:
					sess.dataframe_selections.clear()
				else:
					sess.dataframe_selections.pop(str(key), None)
		except Exception:
			pass


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

		_ALL_DATA_CACHES.append(_clear)
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
	is_nominal = False
	nominal_labels: list[str] = []
	nominal_index: dict[str, int] = {}
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

	def _resolve_x_coord(raw_v: Any, temporal_flag: bool, nominal_flag: bool) -> float | None:
		if raw_v is None or (isinstance(raw_v, float) and math.isnan(raw_v)):
			return None
		if nominal_flag:
			key_s = str(raw_v)
			if key_s not in nominal_index:
				nominal_index[key_s] = len(nominal_labels)
				nominal_labels.append(key_s)
			return float(nominal_index[key_s])
		return _parse_x_val(raw_v, temporal_flag)

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

		x_t = x_enc.get("type")
		if x_t == "temporal":
			is_temporal = True
		elif x_t in ("nominal", "ordinal"):
			is_nominal = True
		if x_enc.get("title") and not x_title:
			x_title = str(x_enc["title"])
		if y_enc.get("title") and not y_title:
			y_title = str(y_enc["title"])
		if isinstance(x_enc.get("scale"), dict) and x_enc["scale"].get("domain") and x_domain_fixed is None:
			dom = x_enc["scale"]["domain"]
			if isinstance(dom, list) and len(dom) == 2:
				try:
					x_domain_fixed = (float(dom[0]), float(dom[1]))
				except Exception:
					pass
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
				xv = _resolve_x_coord(r.get(x_field), is_temporal, is_nominal)
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
	if is_nominal and nominal_labels:
		n_nom = len(nominal_labels)
		step_n = max(1, math.ceil(n_nom / 6.0))
		tick_indices = list(range(0, n_nom, step_n))
		if (n_nom - 1) not in tick_indices:
			tick_indices.append(n_nom - 1)
		for ti in tick_indices:
			px = _sx(float(ti))
			svg.append(f'<line x1="{px:.1f}" y1="{pad_t}" x2="{px:.1f}" y2="{pad_t + plot_h}" class="revir-grid-line" opacity="0.5"/>')
			raw_l = nominal_labels[ti]
			short_l = raw_l[5:16] if len(raw_l) >= 16 and raw_l[4] == "-" else (raw_l[:13] + "…" if len(raw_l) > 14 else raw_l)
			svg.append(f'<text x="{px:.1f}" y="{pad_t + plot_h + 18}" text-anchor="middle" class="revir-axis-lbl">{html_lib.escape(short_l)}</text>')
	else:
		x_ticks = 6
		short_time_span = is_temporal and ((xmax - xmin) <= 86400.0 * 4.5)
		for i in range(x_ticks + 1):
			xv = xmin + (xmax - xmin) * (i / x_ticks)
			px = _sx(xv)
			svg.append(f'<line x1="{px:.1f}" y1="{pad_t}" x2="{px:.1f}" y2="{pad_t + plot_h}" class="revir-grid-line" opacity="0.5"/>')
			if is_temporal:
				try:
					lbl_x = dt_mod.datetime.fromtimestamp(xv).strftime("%d.%m %H:%M" if short_time_span else "%d.%m.%y")
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

		if mtype == "bar":
			bcol_default = str(mcfg.get("color") or "#0ea5e9")
			n_bars = max(1, len(rows))
			bw = max(10.0, min(42.0, (plot_w / max(n_bars, 2)) * 0.56))
			y_base = _sy(max(ymin, 0.0)) if ymin < 0.0 else (pad_t + plot_h)
			for r in rows:
				xv = _resolve_x_coord(r.get(xf), is_temporal, is_nominal) if xf else None
				yv = _parse_x_val(r.get(yf), False) if yf else None
				if xv is None or yv is None:
					continue
				col = _color_for(str(r.get(cf))) if cf else bcol_default
				cx = _sx(xv)
				cy = _sy(yv)
				by = min(cy, y_base)
				bh = max(2.0, abs(y_base - cy))
				tip = html_lib.escape(_build_tip(r, tooltips), quote=True)
				svg.append(f'<rect x="{cx - bw / 2.0:.1f}" y="{by:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="4" fill="{html_lib.escape(col, quote=True)}" class="revir-bar-rect" data-tip="{tip}"/>')
			continue

		if mtype == "line":
			groups: OrderedDict[str, list[tuple[float, float, dict]]] = OrderedDict()
			for r in rows:
				xv = _resolve_x_coord(r.get(xf), is_temporal, is_nominal) if xf else None
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
				xv = _resolve_x_coord(r.get(xf), is_temporal, is_nominal) if xf else None
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
		self.dialog_stack: list[_ActiveDialog] = []
		self.uploaded_files: dict[str, list[UploadedFile]] = {}
		self.downloads: OrderedDict[str, tuple[Any, str, str]] = OrderedDict()
		self.media: OrderedDict[str, tuple[bytes, str]] = OrderedDict()
		self.dataframe_selections: dict[str, dict[str, Any]] = {}
		self.editor_input_hashes: dict[str, str] = {}
		self.editor_meta: dict[str, dict[str, Any]] = {}
		self.form_defs: dict[str, dict[str, Any]] = {}
		self.page_title: str = "Revir Defteri"
		self.last_html_payload: dict[str, Any] = {}
		self.box_layout_enabled: bool = True
		self.floating_chat_enabled: bool = False
		self.floating_chat_title: str = "Gemini Klinik Asistan"
		self.floating_chat_label: str = "KVKK Korumalı"
		self.floating_chat_chips: list[dict[str, str]] | None = None
		self.floating_chat_handler: Callable | None = None
		self.floating_chat_export_handler: Callable | None = None
		self.floating_chat_write_handler: Callable | None = None
		self.pending_chat_writes: dict[str, dict[str, Any]] = {}
		self.chat_history: list[dict[str, str]] = []
		self.topbar_html: str = ""
		self.topbar_action_handler: Callable | None = None
		self.cmd_search_handler: Callable | None = None
		self.cmd_exec_handler: Callable | None = None

	@property
	def active_dialog(self) -> _ActiveDialog | None:
		return self.dialog_stack[-1] if self.dialog_stack else None

	@active_dialog.setter
	def active_dialog(self, value: _ActiveDialog | None) -> None:
		if value is None:
			self.dialog_stack.clear()
		else:
			self.push_dialog(value)

	def push_dialog(self, dlg: _ActiveDialog) -> None:
		if self.dialog_stack and self.dialog_stack[-1].func is dlg.func:
			self.dialog_stack[-1] = dlg
		else:
			self.dialog_stack.append(dlg)

	def pop_dialog(self) -> _ActiveDialog | None:
		return self.dialog_stack.pop() if self.dialog_stack else None

	def register_download(self, data: bytes, file_name: str, mime: str) -> str:
		h = hashlib.sha1(data[:2048] + file_name.encode("utf-8", errors="ignore") + str(len(data)).encode()).hexdigest()[:16]
		token = f"dl_{h}"
		self.downloads[token] = (data, file_name, mime)
		self.downloads.move_to_end(token)
		while len(self.downloads) > 260:
			self.downloads.popitem(last=False)
		return token

	def register_lazy_download(self, factory: Callable[[], Any], file_name: str, mime: str, stable_key: str = "") -> str:
		"""Dosya içeriğini yalnızca kullanıcı indir butonuna tıkladığında üreten tembel indirme kaydı."""
		seed = f"lazy:{stable_key}:{file_name}:{mime}:{id(factory)}".encode("utf-8", errors="ignore")
		token = f"ldl_{hashlib.sha1(seed).hexdigest()[:16]}"
		self.downloads[token] = (factory, file_name, mime)
		self.downloads.move_to_end(token)
		while len(self.downloads) > 260:
			self.downloads.popitem(last=False)
		return token

	def resolve_download_item(self, token: str) -> tuple[bytes, str, str] | None:
		entry = self.downloads.get(token)
		if entry is None:
			return None
		payload, fname, mime = entry
		if callable(payload):
			produced = payload()
			if isinstance(produced, bytes):
				raw = produced
			elif isinstance(produced, bytearray):
				raw = bytes(produced)
			elif isinstance(produced, str):
				raw = produced.encode("utf-8")
			elif hasattr(produced, "getvalue"):
				v = produced.getvalue()
				raw = v.encode("utf-8") if isinstance(v, str) else bytes(v)
			elif hasattr(produced, "read"):
				v = produced.read()
				raw = v.encode("utf-8") if isinstance(v, str) else bytes(v)
			else:
				raw = str(produced or "").encode("utf-8")
			return (raw, fname, mime)
		return (payload, fname, mime)

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


def _parse_editor_number(raw_val: Any) -> float:
	if isinstance(raw_val, (int, float)) and not isinstance(raw_val, bool):
		return float(raw_val)
	s = re.sub(r"[₺\s]", "", str(raw_val or "").replace("TL", "").replace("tl", "")).strip()
	if not s:
		raise ValueError("empty")
	if "." in s and "," in s:
		if s.rfind(",") > s.rfind("."):
			s = s.replace(".", "").replace(",", ".")
		else:
			s = s.replace(",", "")
	elif "," in s:
		if s.count(",") > 1 or re.match(r"^-?[1-9]\d{0,2}(?:,\d{3})+$", s):
			s = s.replace(",", "")
		else:
			s = s.replace(",", ".")
	elif "." in s:
		if s.count(".") > 1 or re.match(r"^-?[1-9]\d{0,2}(?:\.\d{3})+$", s):
			s = s.replace(".", "")
	return float(s)


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
		try:
			num = _parse_editor_number(raw_val)
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
		try:
			return _parse_editor_number(raw_val)
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

	def columns(self, spec: int | Sequence[int | float], *, gap: str = "small", vertical_alignment: str = "top", sticky_first_col: bool = False, sticky_last_col: bool = False, class_name: str = "") -> list[_UIContainer]:
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
		placeholder = _LazyColumnsNode(col_lists, grid_tpl, gap_px, align_css, sticky_first_col=sticky_first_col, sticky_last_col=sticky_last_col, class_name=class_name)
		ctx.current_nodes().append(placeholder)  # type: ignore[arg-type]
		return [_UIContainer(target_list=cl) for cl in col_lists]

	def container(self, *, border: bool = False, height: int | None = None, class_name: str = "") -> _UIContainer:
		cls = "revir-container revir-card" if border else "revir-container"
		if str(class_name or "").strip():
			cls = f"{cls} {html_lib.escape(str(class_name).strip(), quote=True)}"
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
				f'<button type="button" class="revir-pop-close" title="Kapat (ESC)" onclick="window.__revirClosePopovers(event)">✕</button>'
			),
			wrapper_close="</div></div>",
		)

	def tabs(self, tab_labels: Sequence[str], *, orientation: str = "horizontal", lazy: bool = False) -> _TabContainers:
		ctx = _get_ctx()
		labels = [str(x) for x in tab_labels]
		tab_sig = f"{orientation}|" + "|".join(labels)
		tab_id = ctx.next_auto_id("tabs", tab_sig)
		forced_lbl = ctx.session.session_state.get("_force_tab_label")
		if isinstance(forced_lbl, str) and forced_lbl in labels:
			ctx.session.active_tabs[tab_id] = labels.index(forced_lbl)
			ctx.session.session_state.pop("_force_tab_label", None)
		active_idx = ctx.session.active_tabs.get(tab_id, 0)
		if active_idx < 0 or active_idx >= len(labels):
			active_idx = 0
		pane_lists: list[list[str]] = [[] for _ in labels]
		node = _LazyTabsNode(tab_id, labels, active_idx, pane_lists, orientation=orientation, lazy=lazy)
		ctx.current_nodes().append(node)  # type: ignore[arg-type]
		return _TabContainers([_UIContainer(target_list=pl) for pl in pane_lists], active_idx, lazy=lazy)

	def sidebar_nav_groups(
		self,
		groups: Sequence[tuple[str, Sequence[dict[str, Any]]]],
		*,
		key: str = "main_menu_page",
		default_value: str | None = None,
	) -> str:
		"""Gruplanmış ve ikon-rayı (rail) moduna daraltılabilen yan menü navigasyonu."""
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.sidebar_nav_groups(groups, key=key, default_value=default_value)
		ctx = _get_ctx()
		flat_opts: list[str] = []
		for _g_title, items in groups:
			for it in items:
				flat_opts.append(str(it["value"]))
		wid = str(key)
		cur = ctx.session.session_state.get(wid)
		if cur not in flat_opts:
			cur = default_value if (default_value in flat_opts) else (flat_opts[0] if flat_opts else "")
			ctx.session.session_state[wid] = cur
		sel_idx = flat_opts.index(cur) if cur in flat_opts else 0
		self._register_widget_meta(wid, True, {"kind": "radio", "options": flat_opts})
		esc_wid = html_lib.escape(wid, quote=True)

		chunks: list[str] = [f'<nav class="revir-nav-grouped" aria-label="Ana gezinme" data-key="{esc_wid}" data-kind="radio" data-value="{sel_idx}">']
		running_idx = 0
		for g_title, items in groups:
			chunks.append(f'<div class="revir-nav-sec-hdr">{html_lib.escape(str(g_title))}</div>')
			for it in items:
				val = str(it["value"])
				ico = str(it.get("icon") or "•")
				lbl = str(it.get("label") or val)
				badge = it.get("badge")
				act_cls = " active" if running_idx == sel_idx else ""
				tip = html_lib.escape(lbl, quote=True)
				badge_html = (
					f'<span class="revir-nav-badge">{html_lib.escape(str(badge))}</span>'
					if badge is not None and str(badge) != "" and str(badge) != "0"
					else ""
				)
				aria_current = ' aria-current="page"' if running_idx == sel_idx else ""
				chunks.append(
					f'<button type="button" class="revir-nav-item revir-nav-rail-item{act_cls}" title="{tip}" aria-label="{tip}"{aria_current} '
					f'onclick="window.__revirRadioPick(\'{esc_wid}\', {running_idx}, this)">'
					f'<span class="revir-nav-ico">{html_lib.escape(ico)}</span>'
					f'<span class="revir-nav-txt">{html_lib.escape(lbl)}</span>'
					f'{badge_html}</button>'
				)
				running_idx += 1
		chunks.append("</nav>")
		ctx.emit("".join(chunks))
		return str(cur)

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
		eff_mime = mime or mimetypes.guess_type(file_name)[0] or "application/octet-stream"
		if callable(data):
			token = ctx.session.register_lazy_download(data, file_name, eff_mime, stable_key=str(key or label))
		else:
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
			f'onkeydown="window.__revirComboKey(\'{esc_wid}\', event)" '
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
			f'onfocus="window.__revirOpenMulti(\'{esc_wid}\', event)" onkeydown="window.__revirMultiKey(\'{esc_wid}\', event)" oninput="window.__revirFilterMulti(\'{esc_wid}\', this.value)" autocomplete="off"/>'
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

	def dataframe(self, data: Any, *, width: Any = None, height: int | None = None, use_container_width: bool = False, hide_index: bool | None = None, column_order: Sequence[str] | None = None, column_config: dict[str, Any] | None = None, key: str | None = None, on_select: str = "ignore", selection_mode: Any = "multi-row", sticky_cols: int | None = None) -> Any:
		if self._target_list is not None or self._is_sidebar:
			with self:
				return _root.dataframe(data, width=width, height=height, use_container_width=use_container_width, hide_index=hide_index, column_order=column_order, column_config=column_config, key=key, on_select=on_select, selection_mode=selection_mode, sticky_cols=sticky_cols)
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

		# Geniş tablolarda ilk aksiyon/isim sütununu otomatik yapışkan yap
		if sticky_cols is not None:
			eff_sticky = max(0, min(2, int(sticky_cols)))
		elif len(visible_cols) >= 6 and visible_cols and str(visible_cols[0]) in ("📂 Dosya", "Dosya", "📂"):
			eff_sticky = 1
		else:
			eff_sticky = 0

		sel_state = ctx.session.dataframe_selections.get(df_key, {"rows": [], "cells": []}) if selectable else {"rows": [], "cells": []}
		sel_rows_set = set(sel_state.get("rows") or [])

		h_style = f'max-height:{int(height)}px;' if height else "max-height:460px;"
		esc_key = html_lib.escape(df_key, quote=True)

		th_parts: list[str] = []
		if show_idx:
			th_parts.append("<th>#</th>")
		for c_idx, c in enumerate(visible_cols):
			cc = cfg_map.get(c)
			lbl = cc.label if isinstance(cc, _ColumnConfig) and cc.label else str(c)
			hlp = f' title="{html_lib.escape(cc.help, quote=True)}"' if isinstance(cc, _ColumnConfig) and cc.help else ""
			st_cls = f' class="revir-sticky-th revir-sticky-c{c_idx}"' if c_idx < eff_sticky else ""
			th_parts.append(f'<th{st_cls}{hlp} onclick="window.__revirSortTable(this)">{html_lib.escape(lbl)} <span class="revir-sort-ind">⇅</span></th>')

		row_parts: list[str] = []
		for r_pos in range(len(df)):
			row_sel_cls = " revir-row-selected" if r_pos in sel_rows_set else ""
			tds: list[str] = []
			if show_idx:
				tds.append(f'<td class="revir-idx-cell">{html_lib.escape(str(df.index[r_pos]))}</td>')
			for c_idx, c in enumerate(visible_cols):
				val = df.iloc[r_pos][c]
				cc = cfg_map.get(c)
				txt = _format_df_cell(val, cc if isinstance(cc, _ColumnConfig) else None)
				is_dosya = str(c) in ("📂 Dosya", "Dosya", "📂")
				st_td = f" revir-sticky-td revir-sticky-c{c_idx}" if c_idx < eff_sticky else ""
				if selectable:
					safe_col = json.dumps(str(c))
					click_attr = f" onclick='window.__revirDfSelect(\"{esc_key}\", {r_pos}, {html_lib.escape(safe_col, quote=True)})'"
					cell_cls = ("revir-cell-clickable revir-cell-action" if is_dosya else "revir-cell-clickable") + st_td
					if is_dosya:
						tds.append(f'<td class="{cell_cls}"{click_attr}><span class="revir-dosya-pill">{html_lib.escape(txt or "📂 Aç")}</span></td>')
					else:
						tds.append(f'<td class="{cell_cls}"{click_attr}>{html_lib.escape(txt)}</td>')
				else:
					cls_attr = f' class="{st_td.strip()}"' if st_td else ""
					tds.append(f"<td{cls_attr}>{html_lib.escape(txt)}</td>")
			row_parts.append(f'<tr class="{row_sel_cls}">{"".join(tds)}</tr>')

		if not row_parts:
			colspan = max(1, len(visible_cols) + (1 if show_idx else 0))
			row_parts.append(f'<tr><td colspan="{colspan}" class="revir-empty-td">Kayıt bulunamadı.</td></tr>')

		ctx.emit(
			f'<div class="revir-table-outer" data-tbl-key="{esc_key}" data-tbl-explicit="{1 if key is not None else 0}">'
			f'<div class="revir-table-wrap" style="{h_style}">'
			f'<table class="revir-table{" revir-table-selectable" if selectable else ""}">'
			f'<thead><tr>{"".join(th_parts)}</tr></thead>'
			f'<tbody>{"".join(row_parts)}</tbody></table></div></div>'
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
		for vis_r_idx, (row_kind, row_id, row_vals) in enumerate(render_rows):
			tds: list[str] = []
			for c_idx, c in enumerate(visible_cols):
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
						f'<input id="{cell_id}" type="checkbox"{chk} data-ed-row="{vis_r_idx}" data-ed-col="{c_idx}" '
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
						f'<select id="{cell_id}" class="revir-ed-select" data-ed-row="{vis_r_idx}" data-ed-col="{c_idx}" '
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
						f'data-ed-row="{vis_r_idx}" data-ed-col="{c_idx}" '
						f'value="{html_lib.escape(disp_v, quote=True)}" '
						f"oninput='window.__revirEditorDirty(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.value)' "
						f"onchange='window.__revirEditorCell(\"{esc_key}\", \"{row_kind}\", {row_id}, {safe_col}, this.value, \"{cell_id}\")' "
						f'onkeydown="if(event.key===\'Enter\'){{event.preventDefault();window.__revirEditorEnter(this);}}"/></td>'
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
			f'<div class="revir-table-outer" data-tbl-key="{esc_key}" data-tbl-explicit="{1 if is_keyed else 0}">'
			f'<div class="revir-table-wrap" style="{h_style}">'
			f'<table class="revir-table revir-editor-table">'
			f'<thead><tr>{"".join(th_parts)}</tr></thead>'
			f'<tbody>{"".join(tr_parts)}</tbody></table></div></div>{add_footer}</div>'
		)
		return out_df


class _LazyColumnsNode:
	def __init__(self, col_lists: list[list[str]], grid_tpl: str, gap_px: str, align_css: str, *, sticky_first_col: bool = False, sticky_last_col: bool = False, class_name: str = "") -> None:
		self.col_lists = col_lists
		self.grid_tpl = grid_tpl
		self.gap_px = gap_px
		self.align_css = align_css
		self.sticky_first_col = sticky_first_col
		self.sticky_last_col = sticky_last_col
		self.class_name = str(class_name or "").strip()

	def render(self) -> str:
		col_parts: list[str] = []
		last_i = len(self.col_lists) - 1
		for idx, cl in enumerate(self.col_lists):
			is_st = (self.sticky_first_col and idx == 0) or (self.sticky_last_col and idx == last_i)
			cls = "revir-col revir-sticky-col" if is_st else "revir-col"
			col_parts.append(f'<div class="{cls}">{_flatten_nodes(cl)}</div>')
		extra = f" {html_lib.escape(self.class_name, quote=True)}" if self.class_name else ""
		return f'<div class="revir-cols revir-cols-{len(self.col_lists)}{extra}" style="grid-template-columns:{self.grid_tpl}; gap:{self.gap_px}; align-items:{self.align_css};">{"".join(col_parts)}</div>'


class _TabContainers(list[_UIContainer]):
	def __init__(self, containers: list[_UIContainer], active_index: int, *, lazy: bool) -> None:
		super().__init__(containers)
		self.active_index = active_index
		self.lazy = lazy

	def render(self, index: int, renderer: Callable[[], Any]) -> None:
		if index < 0 or index >= len(self):
			raise IndexError("Sekme indeksi geçersiz.")
		if self.lazy and index != self.active_index:
			return
		with self[index]:
			renderer()


class _LazyTabsNode:
	def __init__(
		self,
		tab_id: str,
		labels: list[str],
		active_idx: int,
		pane_lists: list[list[str]],
		*,
		orientation: str = "horizontal",
		lazy: bool = False,
	) -> None:
		self.tab_id = tab_id
		self.labels = labels
		self.active_idx = active_idx
		self.pane_lists = pane_lists
		self.orientation = orientation
		self.lazy = lazy

	def render(self) -> str:
		btns = []
		panes = []
		tablist_id = html_lib.escape(self.tab_id, quote=True)
		group_attr = html_lib.escape(self.tab_id, quote=True)
		group_js = html_lib.escape(json.dumps(self.tab_id), quote=True)
		orientation = "vertical" if self.orientation == "vertical" else "horizontal"
		lazy_attr = ' data-rerender-on-tab="true"' if self.lazy else ""
		for i, (lbl, pl) in enumerate(zip(self.labels, self.pane_lists)):
			act = " active" if i == self.active_idx else ""
			selected = i == self.active_idx
			button_id = html_lib.escape(f"{self.tab_id}_tab_{i}", quote=True)
			pane_id = html_lib.escape(f"{self.tab_id}_pane_{i}", quote=True)
			btns.append(
				f'<button type="button" class="revir-tab-btn{act}" data-tab-group="{group_attr}" data-tab-idx="{i}" '
				f'id="{button_id}" role="tab" aria-selected="{str(selected).lower()}" aria-controls="{pane_id}" '
				f'tabindex="{0 if selected else -1}" '
				f'onclick="window.__revirSwitchTab({group_js}, {i})" '
				f'onkeydown="window.__revirTabKey(event, {group_js}, {i}, \'{orientation}\')">{html_lib.escape(lbl)}</button>'
			)
			panes.append(
				f'<div class="revir-tab-pane{act}" data-tab-group="{group_attr}" data-tab-idx="{i}" '
				f'id="{pane_id}" role="tabpanel" aria-labelledby="{button_id}" aria-hidden="{str(not selected).lower()}" '
				f'tabindex="0">{_flatten_nodes(pl)}</div>'
			)
		if self.orientation == "vertical":
			return (
				f'<div class="revir-tabs revir-tabs-vertical"{lazy_attr}>'
				f'<div class="revir-tab-bar-v" id="{tablist_id}" role="tablist" aria-orientation="vertical">{"".join(btns)}</div>'
				f'<div class="revir-tab-content-v">{"".join(panes)}</div></div>'
			)
		if self.orientation == "segmented":
			return (
				f'<div class="revir-tabs revir-tabs-seg"{lazy_attr}>'
				f'<div class="revir-tab-bar-seg" id="{tablist_id}" role="tablist" aria-orientation="horizontal">{"".join(btns)}</div>'
				f'<div class="revir-tab-content">{"".join(panes)}</div></div>'
			)
		return f'<div class="revir-tabs"{lazy_attr}><div class="revir-tab-bar" id="{tablist_id}" role="tablist" aria-orientation="horizontal">{"".join(btns)}</div><div class="revir-tab-content">{"".join(panes)}</div></div>'


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


def dialog(title: str, *, width: str = "small", mode: str = "modal", on_dismiss: Callable | None = None):
	def _decorator(func: Callable) -> Callable:
		@functools.wraps(func)
		def _wrapper(*args, **kwargs):
			ctx = _get_ctx()
			ctx.session.push_dialog(_ActiveDialog(
				title=title,
				width=width,
				on_dismiss=on_dismiss,
				func=func,
				args=args,
				kwargs=kwargs,
				mode=mode,
			))
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


def close_dialog(*, all_dialogs: bool = False) -> None:
	ctx = _get_ctx()
	if all_dialogs:
		ctx.session.dialog_stack.clear()
	else:
		ctx.session.pop_dialog()


def rerun(*, scope: str = "app") -> None:
	raise _RerunException(scope=scope)


def configure_floating_chat(
	enabled: bool,
	label: str = "KVKK Korumalı",
	handler: Callable | None = None,
	export_handler: Callable | None = None,
	write_handler: Callable | None = None,
	chips: list[dict[str, str]] | None = None,
	title: str | None = None,
) -> None:
	ctx = _get_ctx()
	ctx.session.floating_chat_enabled = bool(enabled)
	ctx.session.floating_chat_label = str(label or "KVKK Korumalı")
	if title is not None:
		ctx.session.floating_chat_title = str(title or "Gemini Klinik Asistan")
	if chips is not None:
		ctx.session.floating_chat_chips = list(chips)
	if handler is not None:
		ctx.session.floating_chat_handler = handler
	if export_handler is not None:
		ctx.session.floating_chat_export_handler = export_handler
	if write_handler is not None:
		ctx.session.floating_chat_write_handler = write_handler


def register_pending_chat_write(proposal: dict[str, Any]) -> str:
	"""Yüzen sohbet balonundan onay bekleyen veri/belge yazma önerisini oturumda saklar."""
	token = f"cw_{secrets.token_urlsafe(10)}"
	with _ACTIVE_SESSION.lock:
		while len(_ACTIVE_SESSION.pending_chat_writes) >= 24:
			oldest = next(iter(_ACTIVE_SESSION.pending_chat_writes))
			_ACTIVE_SESSION.pending_chat_writes.pop(oldest, None)
		_ACTIVE_SESSION.pending_chat_writes[token] = proposal
	return token


def pop_pending_chat_write(token: str) -> dict[str, Any] | None:
	with _ACTIVE_SESSION.lock:
		return _ACTIVE_SESSION.pending_chat_writes.pop(str(token or ""), None)


def register_download(data: bytes, file_name: str, mime: str = "application/octet-stream") -> str:
	"""Herhangi bir işleyiciden (örn. Gemini Sohbet) indirilebilir dosya kaydı açar ve indirme token'ı döndürür."""
	token = f"ai_dl_{secrets.token_urlsafe(12)}"
	with _ACTIVE_SESSION.lock:
		ai_keys = [k for k in _ACTIVE_SESSION.downloads.keys() if str(k).startswith("ai_dl_")]
		while len(ai_keys) >= 40:
			old_k = ai_keys.pop(0)
			_ACTIVE_SESSION.downloads.pop(old_k, None)
		_ACTIVE_SESSION.downloads[token] = (bytes(data), str(file_name or "rapor.bin"), str(mime or "application/octet-stream"))
	return token


def configure_topbar(html_str: str, action_handler: Callable | None = None) -> None:
	ctx = _get_ctx()
	ctx.session.topbar_html = str(html_str or "")
	if action_handler is not None:
		ctx.session.topbar_action_handler = action_handler


def configure_box_layout(enabled: bool = True) -> None:
	ctx = _get_ctx()
	ctx.session.box_layout_enabled = bool(enabled)


def configure_command_palette(search_handler: Callable | None = None, exec_handler: Callable | None = None) -> None:
	ctx = _get_ctx()
	if search_handler is not None:
		ctx.session.cmd_search_handler = search_handler
	if exec_handler is not None:
		ctx.session.cmd_exec_handler = exec_handler


# Kök konteyner ve modül düzeyinde API dışa aktarımları
_root = _UIContainer()
sidebar = _UIContainer(is_sidebar=True)

set_page_config = _root.set_page_config
columns = _root.columns
container = _root.container
expander = _root.expander
popover = _root.popover
tabs = _root.tabs
sidebar_nav_groups = _root.sidebar_nav_groups
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
			close_all = bool(payload.get("all"))
			if close_all:
				while session.dialog_stack:
					old_dlg = session.pop_dialog()
					if old_dlg is not None and callable(old_dlg.on_dismiss):
						try:
							old_dlg.on_dismiss()
						except _RerunException:
							pass
			else:
				old_dlg = session.pop_dialog()
				if old_dlg is not None and callable(old_dlg.on_dismiss):
					try:
						old_dlg.on_dismiss()
					except _RerunException:
						pass

		elif action == "topbar_action":
			if callable(session.topbar_action_handler):
				try:
					session.topbar_action_handler(str(act_key or ""), act_val)
				except _RerunException:
					pass

		elif action == "cmd_exec":
			if callable(session.cmd_exec_handler):
				try:
					session.cmd_exec_handler(payload.get("item") or {})
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
					session.pop_dialog()
				continue
			break

		assert last_ctx is not None
		main_html = _flatten_nodes(last_ctx.main_nodes)
		sidebar_html = _flatten_nodes(last_ctx.sidebar_nodes)
		top_dlg = session.active_dialog
		dialog_html = _flatten_nodes(last_ctx.dialog_nodes) if top_dlg is not None else None
		if error_trace:
			main_html += (
				'<div class="revir-alert revir-alert-error"><span class="revir-alert-icon">❌</span>'
				f'<div class="revir-alert-body"><strong>Çalışma Zamanı Hatası</strong><pre class="revir-code">{html_lib.escape(error_trace)}</pre></div></div>'
			)

		result = {
			"title": session.page_title,
			"active_page": str(session.session_state.get("main_menu_page") or ""),
			"topbar_html": session.topbar_html,
			"sidebar_html": sidebar_html,
			"main_html": main_html,
			"dialog_html": dialog_html,
			"dialog_title": top_dlg.title if top_dlg else None,
			"dialog_width": top_dlg.width if top_dlg else "large",
			"dialog_mode": top_dlg.mode if top_dlg else "modal",
			"dialog_depth": len(session.dialog_stack),
			"sidebar_collapsed": session.sidebar_collapsed,
			"balloons": last_ctx.balloons_triggered,
			"styles_html": "\n".join(last_ctx.styles_parts),
			"scripts_html": "\n".join(last_ctx.scripts_parts),
			"box_layout_enabled": session.box_layout_enabled,
			"floating_chat_enabled": session.floating_chat_enabled,
			"floating_chat_title": session.floating_chat_title,
			"floating_chat_label": session.floating_chat_label,
			"floating_chat_chips": session.floating_chat_chips,
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
	--status-warning-bg: #fffbeb;
	--status-warning-text: #92400e;
	--status-warning-border: #fde68a;
	--status-info-bg: #eff6ff;
	--status-info-text: #1e40af;
	--status-info-border: #bfdbfe;
	--status-danger-bg: #fef2f2;
	--status-danger-text: #991b1b;
	--status-danger-border: #fecaca;
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
	width: 268px; min-width: 268px; max-width: 268px;
	height: 100vh; overflow-y: auto; overflow-x: hidden;
	background: var(--sidebar-bg);
	border-right: 1px solid var(--line);
	padding: 14px 13px 22px 13px;
	transition: width 0.2s ease, min-width 0.2s ease, max-width 0.2s ease, padding 0.2s ease;
	z-index: 20;
	flex-shrink: 0;
}
.revir-sidebar.collapsed {
	width: 66px; min-width: 66px; max-width: 66px;
	padding: 12px 7px 18px 7px;
}
.revir-sidebar.collapsed .revir-sb-expanded-only,
.revir-sidebar.collapsed .revir-nav-sec-hdr,
.revir-sidebar.collapsed .revir-nav-txt,
.revir-sidebar.collapsed .revir-nav-badge,
.revir-sidebar.collapsed .revir-caption,
.revir-sidebar.collapsed .revir-hr,
.revir-sidebar.collapsed .revir-toggle-wrap,
.revir-sidebar.collapsed .revir-card,
.revir-sidebar.collapsed .revir-cols,
.revir-sidebar.collapsed .revir-widget-row {
	display: none !important;
}
.revir-sidebar.collapsed .revir-nav-rail-item {
	justify-content: center;
	padding: 10px 0;
}
.revir-sidebar.collapsed .revir-nav-ico {
	font-size: 19px;
	margin: 0;
}
.revir-sidebar.collapsed .revir-sidebar-close-btn span:last-child { display: none; }
.revir-sidebar.collapsed .revir-sidebar-close-btn { padding: 8px 4px; }
.revir-sidebar-open-btn { display: none !important; }

.revir-workspace {
	flex: 1; min-width: 0; height: 100vh;
	display: flex; flex-direction: column; overflow: hidden;
}
.revir-topbar {
	height: 52px; min-height: 52px; flex-shrink: 0;
	display: flex; align-items: center; justify-content: space-between; gap: 12px;
	padding: 0 22px;
	background: var(--card-bg);
	border-bottom: 1px solid var(--line);
	z-index: 30;
	box-shadow: 0 1px 4px rgba(15, 23, 42, 0.04);
}
.revir-tb-left, .revir-tb-right { display: flex; align-items: center; gap: 8px; min-width: 0; }
.revir-tb-center { flex: 1; display: flex; justify-content: center; max-width: 440px; min-width: 160px; }
.revir-tb-icon-btn {
	width: 34px; height: 34px; border-radius: 9px;
	border: 1px solid var(--line); background: var(--sidebar-bg); color: var(--text-main);
	cursor: pointer; display: inline-flex; align-items: center; justify-content: center;
	font-size: 15px; transition: all 0.15s; flex-shrink: 0;
}
.revir-tb-icon-btn:hover { border-color: var(--teal); color: var(--teal); background: var(--mint); }
.revir-tb-crumb {
	font-size: 0.92rem; font-weight: 700; color: var(--text-heading);
	white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.revir-tb-crumb small { color: var(--text-muted); font-weight: 500; margin-left: 4px; }
.revir-tb-pill {
	display: inline-flex; align-items: center; gap: 5px;
	padding: 5px 10px; border-radius: 999px;
	border: 1px solid var(--line); background: var(--sidebar-bg); color: var(--text-main);
	font-size: 12px; font-weight: 700; cursor: pointer; font-family: inherit;
	white-space: nowrap; transition: all 0.15s;
}
.revir-tb-pill:hover { border-color: var(--teal); color: var(--teal); background: var(--mint); }
.revir-tb-pill.active { background: var(--teal); color: #ffffff !important; border-color: var(--teal); }
.revir-tb-pill.archive { background: #fef3c7; color: #92400e !important; border-color: #f59e0b; }
.revir-tb-pill.kvkk-on { background: #ecfdf5; color: #047857 !important; border-color: #10b981; }
.revir-tb-search-btn {
	width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 8px;
	padding: 6px 12px; border-radius: 10px;
	border: 1px solid var(--line); background: var(--sidebar-bg); color: var(--text-muted);
	font-size: 13px; font-weight: 500; cursor: pointer; font-family: inherit;
	transition: all 0.15s;
}
.revir-tb-search-btn:hover { border-color: var(--teal); color: var(--text-main); background: var(--card-bg); box-shadow: 0 2px 8px rgba(13,148,136,0.1); }
.revir-kbd {
	font-size: 11px; font-weight: 700; padding: 1px 6px; border-radius: 5px;
	border: 1px solid var(--line); background: var(--card-bg); color: var(--text-muted);
}

.revir-main {
	flex: 1; overflow-y: auto; overflow-x: auto;
	padding: 20px 28px 64px 28px;
	scroll-behavior: auto;
}
.revir-main-inner { max-width: 1480px; margin: 0 auto; }
.revir-sticky-col {
	position: sticky; top: 8px; align-self: start;
	max-height: calc(100vh - 104px); overflow-y: auto;
}

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
.revir-settings-status-grid {
	display: grid; grid-template-columns: repeat(4, minmax(0, 1fr));
	gap: 10px; margin-bottom: 14px;
}
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
.revir-status-chip {
	display: inline-flex; align-items: center; gap: 5px;
	padding: 3px 10px; border: 1px solid; border-radius: 999px;
	font-size: 12px; font-weight: 700; line-height: 1.4;
}
.revir-status-chip.warning {
	background: var(--status-warning-bg); color: var(--status-warning-text); border-color: var(--status-warning-border);
}
.revir-status-chip.info {
	background: var(--status-info-bg); color: var(--status-info-text); border-color: var(--status-info-border);
}
.revir-status-chip.danger {
	background: var(--status-danger-bg); color: var(--status-danger-text); border-color: var(--status-danger-border);
}

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

/* Yan Menü Navigasyonu (Gruplu & İkon Rayı Destekli) */
.revir-nav-group, .revir-nav-grouped { display: flex; flex-direction: column; gap: 2px; margin: 6px 0; }
.revir-nav-sec-hdr {
	font-size: 10.5px; font-weight: 800; letter-spacing: 0.06em; text-transform: uppercase;
	color: var(--text-muted); padding: 10px 10px 4px 10px; user-select: none;
}
.revir-nav-sec-hdr:first-child { padding-top: 2px; }
.revir-nav-item {
	width: 100%; text-align: left; padding: 8px 10px; border-radius: 9px;
	border: 1px solid transparent; background: transparent; color: var(--text-main);
	font-size: 0.89rem; font-weight: 600; cursor: pointer; font-family: inherit;
	transition: all 0.15s ease; display: flex; align-items: center; gap: 9px;
}
.revir-nav-ico { font-size: 16px; line-height: 1; flex-shrink: 0; width: 20px; text-align: center; }
.revir-nav-txt { flex: 1; min-width: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.revir-nav-badge {
	font-size: 11px; font-weight: 800; padding: 1px 7px; border-radius: 999px;
	background: rgba(13,148,136,0.14); color: var(--teal); flex-shrink: 0;
}
.revir-nav-item.active .revir-nav-badge { background: rgba(255,255,255,0.24); color: #ffffff !important; }
.revir-nav-item:hover { background: rgba(13,148,136,0.10); color: var(--teal); }
.revir-nav-item.active { background: var(--teal); color: #ffffff !important; box-shadow: 0 2px 8px rgba(13,148,136,0.28); }
.revir-nav-item.active * { color: #ffffff !important; }
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
	max-height: 270px; overflow-y: auto; z-index: 9990;
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 9px;
	box-shadow: 0 12px 32px rgba(15,23,42,0.22); padding: 4px;
}
.revir-combo-menu.revir-portaled {
	position: fixed !important; right: auto !important; z-index: 99990 !important;
}
.revir-combo-menu.open { display: block; }
.revir-combo-opt {
	padding: 7px 10px; border-radius: 6px; cursor: pointer; font-size: 0.89rem;
	display: flex; align-items: center; gap: 8px; color: var(--text-main);
}
.revir-combo-opt:hover, .revir-combo-opt.kbd-active { background: var(--mint); }
.revir-combo-opt.selected { background: rgba(13,148,136,0.14); font-weight: 700; color: var(--teal); }
.revir-combo-opt.selected.kbd-active { background: rgba(13,148,136,0.22); }
.revir-chip {
	display: inline-flex; align-items: center; gap: 4px;
	background: rgba(13,148,136,0.14); color: var(--teal);
	border: 1px solid rgba(13,148,136,0.28); border-radius: 6px;
	padding: 2px 7px; font-size: 0.81rem; font-weight: 600;
}
.revir-chip-x { border: none; background: transparent; color: inherit; cursor: pointer; font-weight: 800; padding: 0 2px; font-size: 14px; line-height: 1; }

/* Sekmeler (Yatay, Dikey & Segmented Hap) */
.revir-tabs { margin: 8px 0 14px 0; }
.revir-tab-bar { display: flex; flex-wrap: wrap; gap: 4px; border-bottom: 1.5px solid var(--line); margin-bottom: 14px; }
.revir-tab-btn {
	padding: 8px 14px; border: none; border-bottom: 2.5px solid transparent;
	background: transparent; color: var(--text-muted); font-family: inherit;
	font-size: 0.89rem; font-weight: 600; cursor: pointer; border-radius: 8px 8px 0 0;
	margin-bottom: -1.5px; transition: all 0.15s;
}
.revir-tab-btn:hover { color: var(--teal); background: rgba(13,148,136,0.06); }
.revir-tab-btn.active { color: var(--teal); border-bottom-color: var(--teal); background: rgba(13,148,136,0.08); font-weight: 700; }
.revir-tab-pane { display: none; }
.revir-tab-pane.active { display: block; }
.revir-tab-pane:not(.active),
.revir-tab-pane.revir-grid-zone:not(.active) { display: none !important; }

.revir-tabs-seg .revir-tab-bar-seg {
	display: inline-flex; flex-wrap: wrap; gap: 4px;
	padding: 4px; border-radius: 11px;
	background: var(--sidebar-bg); border: 1px solid var(--line);
	margin-bottom: 14px;
}
.revir-tabs-seg .revir-tab-btn {
	border: none; border-radius: 8px; margin-bottom: 0; padding: 6px 13px;
	font-size: 0.85rem; color: var(--text-muted);
}
.revir-tabs-seg .revir-tab-btn.active {
	background: var(--card-bg); color: var(--teal);
	box-shadow: 0 2px 6px rgba(15,23,42,0.08); font-weight: 700;
}

.revir-tabs-vertical {
	display: grid; grid-template-columns: 215px minmax(0, 1fr); gap: 18px; align-items: start;
}
.revir-tab-bar-v {
	display: flex; flex-direction: column; gap: 3px;
	position: sticky; top: 0;
	background: var(--sidebar-bg); border: 1px solid var(--line);
	border-radius: 12px; padding: 7px;
}
.revir-tab-bar-v .revir-tab-btn {
	width: 100%; text-align: left; border: none; border-radius: 8px;
	margin-bottom: 0; padding: 8px 11px; font-size: 0.86rem;
	display: flex; align-items: center; gap: 6px;
}
.revir-tab-bar-v .revir-tab-btn.active {
	background: var(--teal); color: #ffffff !important;
	box-shadow: 0 2px 8px rgba(13,148,136,0.24);
}
.revir-tab-content-v { min-width: 0; }
@media (max-width: 900px) {
	.revir-tabs-vertical { grid-template-columns: 1fr; }
	.revir-tab-bar-v { flex-direction: row; flex-wrap: wrap; position: static; }
	.revir-tab-bar-v .revir-tab-btn { width: auto; }
	.revir-cols-5, .revir-cols-6 { grid-template-columns: repeat(3, minmax(0, 1fr)) !important; }
	.revir-settings-status-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
	.revir-sticky-col { position: static; max-height: none; overflow: visible; }
}
@media (max-width: 760px) {
	.revir-cols-3, .revir-cols-4, .revir-cols-5, .revir-cols-6 {
		grid-template-columns: repeat(2, minmax(0, 1fr)) !important;
	}
	.revir-cols-7 { min-width: 620px; }
	.revir-sidebar {
		position: fixed; top: 0; left: 0; bottom: 0; height: 100dvh;
		width: min(268px, 84vw); min-width: min(268px, 84vw); max-width: min(268px, 84vw);
		transform: translateX(0); box-shadow: 12px 0 32px rgba(15, 23, 42, 0.18);
	}
	.revir-sidebar.collapsed {
		width: min(268px, 84vw); min-width: min(268px, 84vw); max-width: min(268px, 84vw);
		transform: translateX(-105%);
	}
	.revir-app:has(.revir-sidebar:not(.collapsed))::after {
		content: ""; position: fixed; inset: 0; z-index: 19; background: rgba(15, 23, 42, 0.35);
	}
	.revir-workspace { width: 100%; }
	.revir-topbar { height: auto; min-height: 52px; flex-wrap: wrap; gap: 6px; padding: 7px 12px; }
	.revir-tb-left { flex: 1 1 auto; }
	.revir-tb-center { order: 3; flex: 1 0 100%; min-width: 0; max-width: none; }
	.revir-tb-right { gap: 5px; }
	.revir-main { padding: 16px 14px 48px; }
	.revir-modal-backdrop { padding: 8px; }
	.revir-modal-card { width: 100%; max-height: 96dvh; }
	.revir-modal-body { padding: 12px 14px 56px; }
	.revir-sticky-col { position: static; max-height: none; overflow: visible; }
}
@media (max-width: 520px) {
	.revir-cols-2, .revir-cols-3, .revir-cols-4, .revir-cols-5, .revir-cols-6 {
		grid-template-columns: minmax(0, 1fr) !important;
	}
	.revir-settings-status-grid { grid-template-columns: minmax(0, 1fr); }
	.revir-tb-crumb { max-width: 34vw; }
	.revir-tb-center { flex-basis: 100%; }
	.revir-modal-hdr-left { min-width: 0; }
	.revir-modal-title { white-space: normal; }
}

/* Expander */
.revir-expander {
	border: 1px solid var(--line); border-radius: 11px;
	background: var(--card-bg); margin: 8px 0 12px 0; overflow: visible;
	transition: border-color 0.16s ease, box-shadow 0.16s ease;
}
.revir-expander:hover { border-color: rgba(13,148,136,0.45); }
.revir-expander[open] { border-color: rgba(13,148,136,0.38); box-shadow: 0 4px 14px rgba(15,23,42,0.06); }
.revir-expander-summary {
	padding: 10px 14px; font-weight: 600; font-size: 0.93rem; cursor: pointer;
	user-select: none; list-style: none; display: flex; align-items: center; justify-content: space-between;
	border-radius: 10px; transition: background 0.14s ease, color 0.14s ease;
}
.revir-expander-summary:hover { background: rgba(13,148,136,0.06); color: var(--text-heading); }
.revir-expander-summary:focus-visible { outline: 2px solid var(--teal); outline-offset: -2px; }
.revir-expander-summary::after {
	content: "▾"; font-size: 12px; color: var(--text-muted); transition: transform 0.18s ease, background 0.14s ease;
	width: 22px; height: 22px; border-radius: 999px; display: inline-flex; align-items: center; justify-content: center;
	background: rgba(148,163,184,0.12); flex-shrink: 0;
}
.revir-expander[open] > .revir-expander-summary { border-radius: 10px 10px 0 0; }
.revir-expander[open] > .revir-expander-summary::after { transform: rotate(180deg); background: rgba(13,148,136,0.14); color: var(--teal); }
.revir-expander-body { padding: 6px 14px 14px 14px; border-top: 1px solid var(--line); border-radius: 0 0 10px 10px; }
.revir-expander[open] > .revir-expander-body { animation: revirExpSlideIn 0.16s ease; }
@keyframes revirExpSlideIn {
	from { opacity: 0; transform: translateY(-4px); }
	to { opacity: 1; transform: translateY(0); }
}

/* Popover */
.revir-popover-wrap { position: relative; display: block; margin-bottom: 6px; }
.revir-popover-body {
	display: none; position: absolute; right: 0; top: calc(100% + 6px);
	min-width: 280px; max-width: min(370px, calc(100vw - 24px)); max-height: 75vh; overflow-y: auto;
	background: var(--card-bg); border: 1px solid var(--line); border-radius: 12px;
	padding: 16px 14px 14px 14px; box-shadow: 0 14px 36px rgba(15,23,42,0.22); z-index: 9900;
}
.revir-popover-body.align-left { right: auto; left: 0; }
.revir-popover-body.flip-up { top: auto; bottom: calc(100% + 6px); }
.revir-popover-body.open { display: block; animation: revirPopIn 0.14s ease; }
@keyframes revirPopIn {
	from { opacity: 0; transform: scale(0.97) translateY(-3px); }
	to { opacity: 1; transform: scale(1) translateY(0); }
}
.revir-pop-close {
	position: sticky; top: 0; float: right; margin: -6px -4px 4px 8px;
	width: 24px; height: 24px; border-radius: 6px; border: 1px solid var(--line);
	background: var(--sidebar-bg); color: var(--text-muted); font-size: 12px; font-weight: 700;
	cursor: pointer; display: inline-flex; align-items: center; justify-content: center; z-index: 2;
}
.revir-pop-close:hover { background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }

/* Tablolar & Data Editor */
.revir-table-outer {
	position: relative; display: block; width: 100%; max-width: 100%;
	margin-bottom: 10px; box-sizing: border-box;
}
.revir-table-outer > .revir-table-wrap,
.revir-table-outer > .revir-md-table-wrap {
	margin-bottom: 0 !important; width: 100%; box-sizing: border-box;
}
.revir-table-outer.revir-tbl-resizing {
	outline: 2px dashed var(--teal); outline-offset: 1px; border-radius: 10px;
	user-select: none !important;
}
.revir-table-outer.revir-tbl-resizing > .revir-table-wrap,
.revir-table-outer.revir-tbl-resizing > .revir-md-table-wrap {
	border-color: var(--teal); box-shadow: 0 8px 26px rgba(13,148,136,0.18);
}
.revir-tbl-handle-r, .revir-tbl-handle-l {
	position: absolute; top: 6px; bottom: 14px; width: 9px;
	cursor: ew-resize; z-index: 12; touch-action: none;
}
.revir-tbl-handle-r { right: -4px; }
.revir-tbl-handle-l { left: -4px; }
.revir-tbl-handle-r::after, .revir-tbl-handle-l::after {
	content: ""; position: absolute; top: 18%; bottom: 18%; left: 3px; width: 3px;
	border-radius: 999px; background: var(--teal); opacity: 0; transition: opacity 0.14s ease;
}
.revir-tbl-handle-r:hover::after, .revir-tbl-handle-l:hover::after,
.revir-table-outer.revir-tbl-resizing .revir-tbl-handle-r::after,
.revir-table-outer.revir-tbl-resizing .revir-tbl-handle-l::after { opacity: 0.78; }

.revir-tbl-handle-b {
	position: absolute; left: 14px; right: 16px; bottom: -4px; height: 9px;
	cursor: ns-resize; z-index: 12; touch-action: none;
}
.revir-tbl-handle-b::after {
	content: ""; position: absolute; left: 22%; right: 22%; top: 3px; height: 3px;
	border-radius: 999px; background: var(--teal); opacity: 0; transition: opacity 0.14s ease;
}
.revir-tbl-handle-b:hover::after,
.revir-table-outer.revir-tbl-resizing .revir-tbl-handle-b::after { opacity: 0.78; }

.revir-tbl-handle-sw {
	position: absolute; left: -2px; bottom: -2px; width: 14px; height: 14px;
	cursor: nesw-resize; z-index: 13; touch-action: none;
}
.revir-tbl-handle-se {
	position: absolute; right: 0; bottom: 0; width: 16px; height: 16px;
	cursor: nwse-resize; z-index: 14; touch-action: none;
	border-radius: 0 0 9px 0;
	background: linear-gradient(135deg, transparent 46%, rgba(13,148,136,0.45) 48%, transparent 56%, rgba(13,148,136,0.75) 76%);
	opacity: 0.58; transition: opacity 0.15s ease;
}
.revir-table-outer:hover .revir-tbl-handle-se,
.revir-table-outer.revir-tbl-resizing .revir-tbl-handle-se { opacity: 1; }

.revir-col-resizer {
	position: absolute; top: 0; right: 0; width: 7px; height: 100%;
	cursor: col-resize; z-index: 7; user-select: none; touch-action: none;
}
.revir-col-resizer::after {
	content: ""; position: absolute; top: 22%; bottom: 22%; right: 2px; width: 2px;
	border-radius: 2px; background: var(--teal); opacity: 0; transition: opacity 0.12s ease;
}
.revir-col-resizer:hover::after, .revir-col-resizer.active::after { opacity: 0.85; }

.revir-table-wrap, .revir-md-table-wrap {
	width: 100%; overflow: auto; border: 1px solid var(--line);
	border-radius: 10px; background: var(--card-bg); margin-bottom: 10px;
	box-sizing: border-box;
}
.revir-md-table-wrap { max-height: 460px; }
.revir-cmsg .revir-md-table-wrap { max-height: 280px; margin: 6px 0; }
.revir-table, .revir-md-table { width: 100%; border-collapse: collapse; font-size: 0.86rem; }
.revir-table-fixed-cols { table-layout: fixed; }
.revir-table-fixed-cols th, .revir-table-fixed-cols td { overflow: hidden; text-overflow: ellipsis; }
.revir-table-fixed-cols .revir-ed-input, .revir-table-fixed-cols .revir-ed-select { min-width: 0; }
.revir-table th, .revir-table td, .revir-md-table th, .revir-md-table td { box-sizing: border-box; }
.revir-table thead th, .revir-md-table thead th {
	position: sticky; top: 0; z-index: 3;
	background: var(--sidebar-bg); color: var(--text-heading);
	font-weight: 700; text-align: left; padding: 9px 11px;
	border-bottom: 1.5px solid var(--line); white-space: nowrap; cursor: pointer; user-select: none;
}
.revir-md-table thead th { cursor: default; }
.revir-md-table tbody td { padding: 7px 11px; border-bottom: 1px solid var(--line); vertical-align: middle; }
.revir-md-table tbody tr:hover td { background: rgba(13,148,136,0.05); }
.revir-sticky-th { position: sticky !important; left: 0; z-index: 5 !important; background: var(--sidebar-bg) !important; box-shadow: 2px 0 6px rgba(15,23,42,0.06); }
.revir-sticky-td { position: sticky !important; left: 0; z-index: 2; background: var(--card-bg); box-shadow: 2px 0 6px rgba(15,23,42,0.05); }
.revir-table tbody tr:hover td.revir-sticky-td { background: var(--mint); }
.revir-row-selected td.revir-sticky-td { background: rgba(13,148,136,0.16) !important; }
.revir-sort-ind { opacity: 0.45; font-size: 11px; margin-left: 3px; }
.revir-sort-ind.active { opacity: 1; color: var(--teal); font-weight: 800; }
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
.revir-iframe-wrap { margin: 8px 0; border: 1px solid var(--line); border-radius: 10px; overflow: hidden; background: #cbd5e1; }
.revir-iframe { border: none; display: block; }

/* A4 Belge & Rapor Önizleme Sahnesi (210mm x 297mm Gerçek Baskı Oranı) */
.revir-a4-stage {
	background: #cbd5e1; border: 1px solid var(--line); border-radius: 12px;
	padding: 20px 16px; margin: 8px 0 12px 0;
	display: flex; flex-direction: column; align-items: center; gap: 18px;
	max-height: 76vh; overflow-y: auto;
}
body.revir-dark .revir-a4-stage, .dark .revir-a4-stage { background: #1e293b; }
.revir-a4-sheet {
	width: min(100%, 794px); min-height: 520px; aspect-ratio: 210 / 297;
	background: #ffffff !important; color: #0f172a !important;
	border: 1px solid #94a3b8; border-radius: 3px;
	box-shadow: 0 10px 30px rgba(15,23,42,0.22);
	padding: 26px 32px; position: relative; box-sizing: border-box;
	overflow: auto;
}
.revir-a4-sheet-img {
	width: min(100%, 794px);
	background: #ffffff !important; color: #0f172a !important;
	border: 1px solid #94a3b8; border-radius: 3px;
	box-shadow: 0 10px 30px rgba(15,23,42,0.22);
	padding: 12px; position: relative; box-sizing: border-box;
}
.revir-a4-sheet-img img { width: 100%; height: auto; display: block; border-radius: 2px; }
.revir-a4-sheet-landscape {
	width: min(100%, 960px); min-height: 460px; aspect-ratio: 297 / 210;
	background: #ffffff !important; color: #0f172a !important;
	border: 1px solid #94a3b8; border-radius: 3px;
	box-shadow: 0 10px 30px rgba(15,23,42,0.22);
	padding: 22px 26px; position: relative; box-sizing: border-box;
	overflow: auto;
}
.revir-a4-page-badge {
	display: inline-block; font-size: 11px; font-weight: 700; color: #475569;
	background: rgba(255,255,255,0.92); border: 1px solid #cbd5e1;
	padding: 2px 10px; border-radius: 999px; margin-bottom: 6px;
}
.revir-a4-sheet table, .revir-a4-sheet-landscape table {
	width: 100%; border-collapse: collapse; margin: 8px 0 14px 0; font-size: 11.5px; color: #0f172a !important;
}
.revir-a4-sheet th, .revir-a4-sheet td, .revir-a4-sheet-landscape th, .revir-a4-sheet-landscape td {
	border: 1px solid #cbd5e1; padding: 5px 7px; text-align: left; color: #0f172a !important;
}
.revir-a4-sheet th, .revir-a4-sheet-landscape th { background: #f1f5f9 !important; font-weight: 700; }

/* Sohbet Balonu İçi Onaylı Veri İşleme Kartı */
.revir-cmsg-write-card {
	margin-top: 8px; padding: 10px 11px; border-radius: 10px;
	background: rgba(13,148,136,0.08); border: 1px solid rgba(13,148,136,0.34);
	font-size: 12.2px;
}
.revir-cmsg-write-card.done {
	background: rgba(16,185,129,0.12); border-color: rgba(16,185,129,0.45);
}
.revir-cmsg-write-card select {
	width: 100%; margin: 5px 0 7px 0; padding: 5px 8px; border-radius: 7px;
	border: 1px solid var(--line); background: var(--input-bg); color: var(--input-text);
	font-size: 12px; font-family: inherit;
}
.revir-cmsg-write-list { margin: 4px 0 8px 16px; padding: 0; }
.revir-cmsg-write-list li { margin-bottom: 2px; }
.revir-cmsg-write-actions { display: flex; gap: 6px; margin-top: 6px; }
.revir-cmsg-write-actions button {
	flex: 1; padding: 6px 8px; border-radius: 7px; font-size: 12px; font-weight: 700; cursor: pointer;
	border: 1px solid var(--line); background: var(--card-bg); color: var(--text-main);
}
.revir-cmsg-write-actions button.confirm {
	background: var(--teal); color: #ffffff; border-color: var(--teal);
}
.revir-cmsg-write-actions button.confirm:hover { filter: brightness(1.06); }
.revir-cmsg-write-actions button:disabled { opacity: 0.55; cursor: not-allowed; }

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

/* Modal Diyalog & Sağdan Açılan Çekmece (Slide-Over Drawer) */
.revir-modal-backdrop {
	position: fixed; inset: 0; z-index: 2000;
	background: rgba(15, 23, 42, 0.55); backdrop-filter: blur(4px);
	display: none; align-items: center; justify-content: center; padding: 20px;
}
.revir-modal-backdrop.open { display: flex; animation: revirFadeIn 0.15s ease; }
@keyframes revirFadeIn {
	from { opacity: 0; }
	to { opacity: 1; }
}
@keyframes revirModalPop {
	from { opacity: 0; transform: scale(0.97) translateY(6px); }
	to { opacity: 1; transform: scale(1) translateY(0); }
}
.revir-modal-backdrop.drawer-mode { justify-content: flex-end; align-items: stretch; padding: 0; }
.revir-modal-card {
	position: relative;
	background: var(--card-bg); color: var(--text-main);
	border: 1px solid var(--line); border-radius: 16px;
	box-shadow: 0 24px 64px rgba(0,0,0,0.34);
	width: 93vw; max-width: 1220px; max-height: 92vh;
	display: flex; flex-direction: column; overflow: hidden;
	transition: width 0.2s ease, max-width 0.2s ease, height 0.2s ease, max-height 0.2s ease;
}
.revir-modal-backdrop.open .revir-modal-card:not(.drawer) { animation: revirModalPop 0.16s ease; }
.revir-modal-card.revir-modal-resizing {
	transition: none !important; user-select: none !important;
	border-color: var(--teal) !important;
	box-shadow: 0 24px 64px rgba(13,148,136,0.28) !important;
}
.revir-modal-resize-r, .revir-modal-resize-l {
	position: absolute; top: 10px; bottom: 16px; width: 7px;
	cursor: ew-resize; z-index: 25; touch-action: none;
}
.revir-modal-resize-r { right: 0; }
.revir-modal-resize-l { left: 0; }
.revir-modal-resize-b {
	position: absolute; left: 14px; right: 16px; bottom: 0; height: 7px;
	cursor: ns-resize; z-index: 25; touch-action: none;
}
.revir-modal-resize-se {
	position: absolute; right: 2px; bottom: 2px; width: 16px; height: 16px;
	cursor: nwse-resize; z-index: 26; touch-action: none; border-radius: 0 0 10px 0;
	background: linear-gradient(135deg, transparent 46%, rgba(13,148,136,0.45) 48%, transparent 56%, rgba(13,148,136,0.75) 76%);
	opacity: 0.55; transition: opacity 0.15s ease;
}
.revir-modal-resize-sw {
	position: absolute; left: 2px; bottom: 2px; width: 14px; height: 14px;
	cursor: nesw-resize; z-index: 26; touch-action: none;
}
.revir-modal-card:hover .revir-modal-resize-se { opacity: 0.95; }
.revir-modal-card.fullscreen .revir-modal-resize-r,
.revir-modal-card.fullscreen .revir-modal-resize-l,
.revir-modal-card.fullscreen .revir-modal-resize-b,
.revir-modal-card.fullscreen .revir-modal-resize-se,
.revir-modal-card.fullscreen .revir-modal-resize-sw,
.revir-modal-card.drawer .revir-modal-resize-r,
.revir-modal-card.drawer .revir-modal-resize-b,
.revir-modal-card.drawer .revir-modal-resize-se,
.revir-modal-card.drawer .revir-modal-resize-sw { display: none; }
#revir-resize-badge {
	position: fixed; display: none; z-index: 99999; pointer-events: none;
	background: #0f172a; color: #ffffff; font-size: 11px; font-weight: 800;
	padding: 3px 8px; border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.24);
}
.revir-modal-card.small { max-width: 640px; }
.revir-modal-card.drawer {
	width: 640px; max-width: 94vw; height: 100vh; max-height: 100vh;
	border-radius: 16px 0 0 16px; border-right: none;
	animation: revirSlideInRight 0.2s ease;
}
.revir-modal-card.fullscreen {
	width: 98vw !important; max-width: 98vw !important;
	height: 96vh !important; max-height: 96vh !important;
	border-radius: 12px !important;
}
@keyframes revirSlideInRight {
	from { transform: translateX(36px); opacity: 0.6; }
	to { transform: translateX(0); opacity: 1; }
}
.revir-modal-header {
	display: flex; align-items: center; justify-content: space-between; gap: 10px;
	padding: 12px 18px; border-bottom: 1px solid var(--line);
	background: var(--sidebar-bg); flex-shrink: 0;
}
.revir-modal-hdr-left { display: flex; align-items: center; gap: 8px; min-width: 0; }
.revir-modal-hdr-actions { display: flex; align-items: center; gap: 6px; flex-shrink: 0; }
.revir-modal-title { font-size: 1.05rem; font-weight: 800; color: var(--text-heading); margin: 0; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.revir-modal-back {
	display: none; align-items: center; gap: 4px;
	padding: 4px 10px; border-radius: 7px; border: 1px solid var(--teal);
	background: var(--mint); color: var(--teal); font-size: 12px; font-weight: 700; cursor: pointer;
}
.revir-modal-back.visible { display: inline-flex; }
.revir-modal-back:hover { background: var(--teal); color: #ffffff; }
.revir-modal-hbtn, .revir-modal-close {
	width: 32px; height: 32px; border-radius: 8px; border: 1px solid var(--line);
	background: var(--card-bg); color: var(--text-main); font-size: 16px;
	cursor: pointer; display: flex; align-items: center; justify-content: center;
}
.revir-modal-hbtn:hover { background: var(--mint); color: var(--teal); border-color: var(--teal); }
.revir-modal-hbtn.active { background: var(--teal); color: #ffffff; border-color: var(--teal); }
.revir-modal-dirty-pill {
	display: none; align-items: center; gap: 4px; padding: 2px 9px; border-radius: 999px;
	background: #fef3c7; color: #92400e; border: 1px solid #f59e0b;
	font-size: 11px; font-weight: 700; white-space: nowrap; flex-shrink: 0;
}
.revir-modal-dirty-pill.visible { display: inline-flex; }
.revir-modal-close:hover { background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }
.revir-modal-body { padding: 16px 22px 68px 22px; overflow-y: auto; flex: 1; }
body.modal-open #revir-fab-btn {
	padding: 0; width: 46px; height: 46px; min-width: 46px; opacity: 0.85;
}
body.modal-open #revir-fab-btn span:not(.revir-fab-icon) { display: none; }
body.modal-open #revir-fab-btn:hover { opacity: 1; }

/* Global Komut & Kişi Arama Paleti (Cmd/Ctrl+K) */
#revir-cmd-backdrop {
	position: fixed; inset: 0; z-index: 4200;
	background: rgba(15, 23, 42, 0.52); backdrop-filter: blur(5px);
	display: none; align-items: flex-start; justify-content: center; padding-top: 9vh;
}
#revir-cmd-backdrop.open { display: flex; animation: revirFadeIn 0.12s ease; }
#revir-cmd-card {
	width: 620px; max-width: 94vw; max-height: 74vh;
	background: var(--card-bg); color: var(--text-main);
	border: 1px solid var(--line); border-radius: 16px;
	box-shadow: 0 24px 64px rgba(15,23,42,0.38);
	display: flex; flex-direction: column; overflow: hidden;
	animation: revirModalPop 0.14s ease;
}
.revir-cmd-close {
	width: 26px; height: 26px; border-radius: 7px; border: 1px solid var(--line);
	background: var(--card-bg); color: var(--text-muted); font-size: 13px; font-weight: 700;
	cursor: pointer; display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0;
}
.revir-cmd-close:hover { background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }
.revir-cmd-top {
	display: flex; align-items: center; gap: 10px;
	padding: 13px 16px; border-bottom: 1px solid var(--line);
	background: var(--sidebar-bg);
}
#revir-cmd-input {
	flex: 1; border: none; outline: none; background: transparent;
	color: var(--text-heading); font-family: inherit; font-size: 15.5px; font-weight: 500;
}
#revir-cmd-results {
	flex: 1; overflow-y: auto; padding: 8px;
}
.revir-cmd-group {
	font-size: 11px; font-weight: 800; letter-spacing: 0.05em; text-transform: uppercase;
	color: var(--text-muted); padding: 8px 10px 4px 10px;
}
.revir-cmd-item {
	display: flex; align-items: center; justify-content: space-between; gap: 10px;
	padding: 9px 12px; border-radius: 10px; cursor: pointer;
	transition: background 0.12s;
}
.revir-cmd-item:hover, .revir-cmd-item.active { background: var(--mint); }
.revir-cmd-item-l { display: flex; align-items: center; gap: 10px; min-width: 0; }
.revir-cmd-item-ico { font-size: 18px; flex-shrink: 0; }
.revir-cmd-item-title { font-size: 13.8px; font-weight: 700; color: var(--text-heading); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.revir-cmd-item-sub { font-size: 12px; color: var(--text-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.revir-cmd-item-tag {
	font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 6px;
	background: var(--sidebar-bg); border: 1px solid var(--line); color: var(--teal); flex-shrink: 0;
}
.revir-cmd-footer {
	display: flex; align-items: center; justify-content: space-between;
	padding: 8px 14px; border-top: 1px solid var(--line);
	background: var(--sidebar-bg); font-size: 11.5px; color: var(--text-muted);
}

/* Tema Swatch Kartları, Klinik Rozetler & Sabit Şerit */
.revir-sticky-alert-bar {
	position: sticky; top: 0; z-index: 10;
	background: var(--card-bg); border: 1px solid var(--line);
	border-radius: 12px; padding: 11px 15px; margin-bottom: 12px;
	box-shadow: 0 4px 14px rgba(15,23,42,0.06);
}
.revir-op-row {
	background: var(--card-bg); border: 1px solid var(--line); color: var(--text-main);
	border-radius: 9px; padding: 10px 14px; margin-bottom: 6px;
	display: flex; flex-wrap: wrap; align-items: center; gap: 8px;
	box-shadow: 0 1px 3px rgba(15,23,42,0.04);
}
.revir-dose-hero {
	background: var(--mint); border: 2px solid var(--teal); border-radius: 13px;
	padding: 14px 16px; text-align: center; margin-bottom: 10px; color: var(--text-main);
}
.revir-dose-hero.danger {
	background: rgba(239,68,68,0.11); border-color: #ef4444;
}
.revir-dose-hero-val { font-size: 2rem; font-weight: 900; color: var(--teal); margin: 4px 0; }
.revir-dose-hero.danger .revir-dose-hero-val { color: #dc2626; }
.revir-hypo-timer {
	display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-wrap: wrap;
	background: rgba(239,68,68,0.10); border: 1.5px solid rgba(239,68,68,0.45);
	border-radius: 11px; padding: 10px 14px; margin: 8px 0 10px 0;
}
.revir-hypo-clock { font-size: 1.35rem; font-weight: 900; font-variant-numeric: tabular-nums; color: #dc2626; }

/* Bildirim Toast & Yükleme Çubuğu / Göstergesi */
#revir-busy-bar {
	position: fixed; top: 0; left: 0; height: 3px; width: 0%;
	background: var(--teal); z-index: 9999; transition: width 0.18s ease, opacity 0.2s;
	opacity: 0; pointer-events: none;
}
#revir-busy-bar.active { width: 78%; opacity: 1; }
#revir-loading-pill {
	position: fixed; top: 62px; left: 50%; transform: translateX(-50%); z-index: 9998;
	background: var(--card-bg); color: var(--text-heading);
	border: 1.5px solid var(--teal); border-radius: 999px;
	padding: 6px 15px; font-size: 12.8px; font-weight: 700;
	box-shadow: 0 8px 24px rgba(15,23,42,0.18);
	display: none; align-items: center; gap: 8px; pointer-events: none;
}
#revir-loading-pill.active { display: inline-flex; }
.revir-spin-dot {
	width: 13px; height: 13px; border-radius: 50%;
	border: 2.2px solid rgba(13,148,136,0.25); border-top-color: var(--teal);
	animation: revirSpin 0.7s linear infinite; display: inline-block;
}
@keyframes revirSpin { to { transform: rotate(360deg); } }
body.revir-busy .revir-btn { cursor: progress; }
#revir-toast {
	position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); z-index: 99999;
	background: #0f172a; color: #ffffff; padding: 10px 18px; border-radius: 10px;
	font-size: 13.5px; font-weight: 600; box-shadow: 0 10px 28px rgba(0,0,0,0.28);
	display: none; cursor: pointer; border-left: 4px solid var(--teal);
}
#revir-toast.toast-err { border-left-color: #ef4444; background: #7f1d1d; }
#revir-toast.toast-warn { border-left-color: #f59e0b; background: #78350f; }

/* Yüzen Gemini Sohbet Butonu & Konuşma Balonu */
#revir-fab-btn {
	position: fixed; bottom: 22px; right: 24px; z-index: 3400;
	width: 56px; height: 56px; min-width: 56px; padding: 0; border-radius: 50%;
	border: 1.5px solid rgba(255,255,255,0.45);
	background: linear-gradient(135deg, var(--teal, #0d9488) 0%, #0284c7 100%);
	color: #ffffff; font-family: inherit;
	display: none; align-items: center; justify-content: center;
	cursor: pointer; box-shadow: 0 10px 28px rgba(13,148,136,0.38);
	transition: transform 0.2s cubic-bezier(0.34,1.56,0.64,1), box-shadow 0.2s ease;
	user-select: none;
}
#revir-fab-btn.visible { display: inline-flex; }
#revir-fab-btn:hover { transform: translateY(-3px) scale(1.05); box-shadow: 0 14px 34px rgba(2,132,199,0.45); }
#revir-fab-btn svg { width: 28px; height: 28px; display: block; }

#revir-chat-bubble {
	position: fixed; bottom: 88px; right: 24px; z-index: 3450;
	width: 390px; height: 535px; max-width: calc(100vw - 32px); max-height: calc(100vh - 110px);
	background: var(--card-bg, #ffffff); color: var(--text-main, #0f172a);
	border: 1px solid var(--line, #e2e8f0); border-radius: 20px;
	box-shadow: 0 22px 60px rgba(15,23,42,0.28);
	display: none; flex-direction: column; overflow: hidden;
	transition: width 0.22s ease, height 0.22s ease;
}
#revir-chat-bubble.open { display: flex; animation: revirChatPop 0.16s ease; }
#revir-chat-bubble.revir-chat-resizing { transition: none !important; user-select: none !important; border-color: var(--teal) !important; }
@keyframes revirChatPop {
	from { opacity: 0; transform: translateY(10px) scale(0.97); }
	to { opacity: 1; transform: translateY(0) scale(1); }
}
.revir-chat-resize-nw {
	position: absolute; top: 2px; left: 2px; width: 15px; height: 15px;
	cursor: nwse-resize; z-index: 15; touch-action: none; border-radius: 12px 0 6px 0;
}
.revir-chat-resize-nw::after {
	content: ""; position: absolute; top: 4px; left: 4px; width: 7px; height: 7px;
	border-top: 2px solid rgba(255,255,255,0.65); border-left: 2px solid rgba(255,255,255,0.65);
	border-radius: 2px 0 0 0; opacity: 0.7;
}
.revir-chat-resize-nw:hover::after { opacity: 1; }
#revir-chat-bubble.expanded { width: 550px; height: 675px; }

/* Sayfa İçi 12-Sütunlu Sanal Izgara Kutucuk Yerleşimi & Boyutlandırma Modu (Takvim Hariç) */
.revir-grid-zone {
	position: relative;
	display: grid !important;
	grid-template-columns: repeat(12, minmax(0, 1fr));
	gap: 12px;
	align-items: stretch;
	grid-auto-flow: row;
	width: 100%;
}
.revir-tab-pane.revir-grid-zone:not(.active) {
	display: none !important;
}
.revir-grid-zone > * { grid-column: 1 / -1; min-width: 0; }
body.revir-layout-edit .revir-grid-zone {
	background-image: repeating-linear-gradient(
		to right,
		rgba(13,148,136,0.06) 0,
		rgba(13,148,136,0.06) calc((100% - 132px) / 12),
		transparent calc((100% - 132px) / 12),
		transparent calc((100% - 132px) / 12 + 12px)
	);
	border-radius: 12px;
	padding-bottom: 12px;
}
.revir-box-host {
	position: relative;
	box-sizing: border-box;
	container-type: inline-size;
	transition: transform 0.16s cubic-bezier(0.2, 0.8, 0.2, 1), border-color 0.15s ease, box-shadow 0.15s ease;
}
body.revir-layout-edit .revir-box-host {
	outline: 1.5px dashed rgba(13,148,136,0.55); outline-offset: 2px; border-radius: 12px;
}
body.revir-layout-edit .revir-box-host:hover {
	outline-color: var(--teal); box-shadow: 0 6px 20px rgba(13,148,136,0.12);
}
.revir-box-host.revir-box-dragging { opacity: 0.52; outline: 2px solid var(--teal) !important; transform: scale(0.99); }
.revir-box-host.revir-box-drop-before { box-shadow: 0 -4px 0 0 var(--teal), -4px 0 0 0 rgba(13,148,136,0.35), 0 0 0 2px rgba(13,148,136,0.22) !important; }
.revir-box-host.revir-box-drop-after { box-shadow: 0 4px 0 0 var(--teal), 4px 0 0 0 rgba(13,148,136,0.35), 0 0 0 2px rgba(13,148,136,0.22) !important; }
.revir-box-host.revir-box-swap-target {
	outline: 3px solid #f59e0b !important;
	box-shadow: 0 0 0 5px rgba(245,158,11,0.24), 0 12px 28px rgba(245,158,11,0.18) !important;
	transform: scale(1.008);
}
.revir-box-swap-pill {
	display: none; position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
	z-index: 30; background: #b45309; color: #ffffff; font: 800 12px sans-serif;
	padding: 6px 12px; border-radius: 999px; pointer-events: none;
	box-shadow: 0 6px 18px rgba(0,0,0,0.24);
}
.revir-box-host.revir-box-swap-target > .revir-box-swap-pill { display: block; }
.revir-box-host.revir-box-resizing {
	transition: none !important; user-select: none !important;
	outline: 2px dashed var(--teal) !important; box-shadow: 0 10px 26px rgba(13,148,136,0.18) !important;
}
.revir-box-toolbar {
	display: none; position: absolute; top: 6px; right: 8px; z-index: 24;
	align-items: center; gap: 3px; padding: 3px 5px; border-radius: 8px;
	background: var(--card-bg); border: 1px solid rgba(13,148,136,0.45);
	box-shadow: 0 4px 12px rgba(15,23,42,0.14);
}
body.revir-layout-edit .revir-box-host > .revir-box-toolbar { display: inline-flex; }
.revir-box-tbtn {
	border: none; background: transparent; color: var(--text-main);
	width: 23px; height: 22px; border-radius: 5px; cursor: pointer;
	font-size: 12px; font-weight: 800; display: inline-flex; align-items: center; justify-content: center;
	padding: 0; line-height: 1;
}
.revir-box-tbtn:hover { background: var(--mint); color: var(--teal); }
.revir-box-drag { cursor: grab; color: var(--teal); font-size: 13px; touch-action: none; }
.revir-box-drag:active { cursor: grabbing; }
.revir-box-span-tag { font-size: 10.5px; font-weight: 800; color: var(--teal); padding: 0 4px; white-space: nowrap; }
.revir-box-res-r {
	position: absolute; top: 10px; right: -4px; bottom: 14px; width: 8px;
	cursor: ew-resize; z-index: 21; touch-action: none; display: none;
}
.revir-box-res-b {
	position: absolute; left: 12px; right: 14px; bottom: -4px; height: 8px;
	cursor: ns-resize; z-index: 21; touch-action: none; display: none;
}
.revir-box-res-se {
	position: absolute; right: 0; bottom: 0; width: 15px; height: 15px;
	cursor: nwse-resize; z-index: 22; touch-action: none; border-radius: 0 0 10px 0;
	background: linear-gradient(135deg, transparent 46%, rgba(13,148,136,0.45) 48%, transparent 56%, rgba(13,148,136,0.75) 76%);
	opacity: 0; transition: opacity 0.14s ease; display: none;
}
body.revir-box-resize-on .revir-box-res-se { display: block; }
body.revir-box-resize-on .revir-box-host:hover > .revir-box-res-se { opacity: 0.72; }
body.revir-layout-edit .revir-box-res-r,
body.revir-layout-edit .revir-box-res-b,
body.revir-layout-edit .revir-box-res-se { display: block; opacity: 0.92; }
/* Daralan grid kutularında okunabilirliği korumak için içerik kırılım kuralları */
@container (max-width: 720px) {
	.revir-cols-6, .revir-cols-5 { grid-template-columns: repeat(3, minmax(0, 1fr)) !important; }
}
@container (max-width: 500px) {
	.revir-cols-6, .revir-cols-5, .revir-cols-4, .revir-cols-3 { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }
	.revir-cols-2 { grid-template-columns: repeat(1, minmax(0, 1fr)) !important; }
}
@media (max-width: 760px) {
	.revir-grid-zone > .revir-box-host { grid-column: 1 / -1 !important; width: auto !important; height: auto !important; }
}
/* Depo Kart & Önizleme Yatay Eylem Çubuğu — dar kutularda bile tek satırda yan yana kalır */
.revir-cols.revir-doc-action-row,
.revir-cols.revir-cols-5.revir-doc-action-row {
	display: grid !important;
	grid-template-columns: repeat(5, minmax(0, 1fr)) !important;
	gap: 4px !important;
	align-items: stretch !important;
	margin-top: 8px;
}
.revir-cols.revir-doc-action-row.is-4,
.revir-cols.revir-cols-4.revir-doc-action-row {
	grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
}
.revir-cols.revir-doc-action-row > .revir-col { min-width: 0 !important; }
.revir-cols.revir-doc-action-row .revir-btn,
.revir-cols.revir-doc-action-row .revir-popover-trigger,
.revir-cols.revir-doc-action-row .revir-pop-trigger {
	padding: 6px 2px !important;
	font-size: 13px !important;
	min-height: 32px !important;
	white-space: nowrap !important;
}
.revir-cols.revir-doc-prev-actions {
	display: flex !important;
	flex-wrap: nowrap !important;
	gap: 8px !important;
	align-items: stretch !important;
	overflow-x: auto !important;
	padding: 2px 0;
}
.revir-cols.revir-doc-prev-actions > .revir-col {
	flex: 1 1 0 !important;
	min-width: 96px !important;
}
.revir-doc-card-thumb {
	background: #f1f5f9;
	border: 1px solid #cbd5e1;
	border-radius: 8px;
	padding: 6px;
	overflow: hidden;
	box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);
}
.revir-doc-card-thumb img {
	width: 100%;
	max-height: 245px;
	object-fit: contain;
	border-radius: 4px;
	background: #ffffff;
	box-shadow: 0 2px 8px rgba(15,23,42,0.12);
	display: block;
}
#revir-layout-bar {
	position: fixed; bottom: 18px; left: 24px; z-index: 3390;
	display: none; align-items: center; gap: 8px; flex-wrap: wrap;
	background: var(--card-bg); color: var(--text-heading);
	border: 1.5px solid var(--teal); border-radius: 12px; padding: 7px 12px;
	box-shadow: 0 12px 32px rgba(15,23,42,0.22); font-size: 12.5px; font-weight: 700;
}
body.revir-layout-edit #revir-layout-bar { display: inline-flex; animation: revirPopIn 0.14s ease; }
.revir-layout-bar-btn {
	border: 1px solid var(--line); background: var(--sidebar-bg); color: var(--text-main);
	border-radius: 7px; padding: 4px 10px; font-size: 12px; font-weight: 700; cursor: pointer; font-family: inherit;
}
.revir-layout-bar-btn:hover { border-color: var(--teal); color: var(--teal); background: var(--mint); }
.revir-layout-bar-btn.primary { background: var(--teal); color: #ffffff; border-color: var(--teal); }
.revir-layout-bar-btn.primary:hover { background: var(--teal-hover); color: #ffffff; }
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
.revir-chat-hbtn.stop-tts { display: none; width: auto; padding: 0 9px; gap: 4px; font-size: 11.5px; font-weight: 800; background: #ef4444; color: #ffffff !important; border: 1px solid rgba(255,255,255,0.55); animation: revirPulse 1.15s infinite; }
.revir-chat-hbtn.stop-tts.speaking { display: inline-flex; }
.revir-chat-hbtn.stop-tts:hover { background: #dc2626; }

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
	display: flex; align-items: center; justify-content: flex-end; flex-wrap: wrap; gap: 4px;
	margin-top: 6px; padding-top: 4px; border-top: 1px dashed var(--line, #e2e8f0);
}
.revir-cmsg-tbtn {
	border: none; background: transparent; color: var(--text-muted, #64748b);
	font-size: 11px; font-weight: 600; cursor: pointer; padding: 2px 6px; border-radius: 5px;
}
.revir-cmsg-tbtn:hover { background: rgba(13,148,136,0.12); color: var(--teal, #0d9488); }
.revir-cmsg-tbtn.speaking { background: #fef2f2; color: #dc2626 !important; border: 1px solid #fecaca; font-weight: 700; }
.revir-cmsg-downloads {
	display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; padding-top: 7px;
	border-top: 1px solid var(--line, #e2e8f0);
}
.revir-cmsg-dlbtn {
	display: inline-flex; align-items: center; gap: 6px; padding: 6px 11px;
	border-radius: 9px; border: 1px solid rgba(13,148,136,0.35);
	background: linear-gradient(135deg, rgba(13,148,136,0.12) 0%, rgba(2,132,199,0.12) 100%);
	color: var(--teal, #0f766e); font-size: 12px; font-weight: 700; cursor: pointer;
	font-family: inherit; transition: all 0.15s;
}
.revir-cmsg-dlbtn:hover { background: var(--teal, #0d9488); color: #ffffff !important; border-color: var(--teal, #0d9488); transform: translateY(-1px); }

.revir-voice-bar {
	display: none; flex-direction: column; gap: 6px; padding: 8px 12px;
	background: linear-gradient(90deg, rgba(239,68,68,0.10) 0%, rgba(249,115,22,0.08) 100%);
	border-top: 1px solid rgba(239,68,68,0.28); flex-shrink: 0;
}
.revir-voice-bar.active { display: flex; }
.revir-voice-top { display: flex; align-items: center; justify-content: space-between; gap: 8px; font-size: 11.8px; font-weight: 700; color: #dc2626; }
.revir-voice-meter { width: 100%; height: 6px; border-radius: 999px; background: rgba(148,163,184,0.28); overflow: hidden; }
.revir-voice-meter-fill { height: 100%; width: 6%; border-radius: 999px; background: linear-gradient(90deg, #10b981, #f59e0b, #ef4444); transition: width 0.12s linear; }
.revir-voice-actions { display: flex; gap: 6px; justify-content: flex-end; }
.revir-voice-abtn {
	border: none; border-radius: 7px; padding: 4px 10px; font-size: 11.5px; font-weight: 700;
	cursor: pointer; font-family: inherit;
}
.revir-voice-abtn.send { background: #0d9488; color: #ffffff; }
.revir-voice-abtn.send:hover { background: #0f766e; }
.revir-voice-abtn.cancel { background: rgba(100,116,139,0.16); color: var(--text-main, #0f172a); }
.revir-voice-abtn.cancel:hover { background: rgba(239,68,68,0.18); color: #dc2626; }

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
<div id="revir-loading-pill"><span class="revir-spin-dot"></span><span>İşlem yürütülüyor...</span></div>
<button id="revir-sidebar-open" class="revir-sidebar-open-btn" type="button" title="Menüyü Aç" aria-label="Yan menüyü aç" onclick="window.__revirSetSidebar(false)">❯</button>
<div class="revir-app">
	<aside id="revir-sidebar" class="revir-sidebar">__REVIR_INIT_SIDEBAR__</aside>
	<div class="revir-workspace">
		<header id="revir-topbar" class="revir-topbar">__REVIR_INIT_TOPBAR__</header>
		<main id="revir-main" class="revir-main">
			<div id="revir-main-inner" class="revir-main-inner">__REVIR_INIT_MAIN__</div>
		</main>
	</div>
</div>
<div id="revir-modal-backdrop" class="revir-modal-backdrop" onclick="if(event.target===this) window.__revirTryBackdropClose()">
	<div id="revir-modal-card" class="revir-modal-card" role="dialog" aria-modal="true" aria-labelledby="revir-modal-title">
		<div class="revir-modal-header">
			<div class="revir-modal-hdr-left">
				<button id="revir-modal-back" type="button" class="revir-modal-back" title="Önceki Pencereye Dön" onclick="window.__revirCloseDialog()">← Geri</button>
				<h3 id="revir-modal-title" class="revir-modal-title"></h3>
				<span id="revir-modal-dirty-pill" class="revir-modal-dirty-pill" title="Kaydedilmemiş değişiklikler var">● Kaydedilmemiş</span>
			</div>
			<div class="revir-modal-hdr-actions">
				<button id="revir-modal-reset-sz" type="button" class="revir-modal-hbtn" title="Pencere Boyutunu Sıfırla" aria-label="Pencere boyutunu sıfırla" onclick="window.__revirResetModalSize()">↺</button>
				<button id="revir-modal-fs-btn" type="button" class="revir-modal-hbtn" title="Tam Ekran Büyüt / Küçült" aria-label="Tam ekranı büyüt veya küçült" onclick="window.__revirToggleModalFullscreen()">⤢</button>
				<button type="button" class="revir-modal-close" title="Kapat (ESC)" aria-label="Pencereyi kapat" onclick="window.__revirCloseDialog(true)">✕</button>
			</div>
		</div>
		<div id="revir-modal-body" class="revir-modal-body"></div>
		<div class="revir-modal-resize-l" data-mres="l" title="Pencere genişliğini ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın"></div>
		<div class="revir-modal-resize-r" data-mres="r" title="Pencere genişliğini ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın"></div>
		<div class="revir-modal-resize-b" data-mres="b" title="Pencere yüksekliğini ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın"></div>
		<div class="revir-modal-resize-sw" data-mres="sw" title="Boyutlandırmak için sürükleyin · Sıfırlamak için çift tıklayın"></div>
		<div class="revir-modal-resize-se" data-mres="se" title="Boyutlandırmak için sürükleyin · Sıfırlamak için çift tıklayın"></div>
	</div>
</div>
<div id="revir-cmd-backdrop" onclick="if(event.target===this) window.__revirCloseCmdPalette()">
	<div id="revir-cmd-card" role="dialog" aria-modal="true" aria-label="Kişi, sayfa ve ayar arama">
		<div class="revir-cmd-top">
			<span style="font-size:17px;">🔍</span>
			<input id="revir-cmd-input" type="text" aria-label="Öğrenci, sayfa veya ayar ara" placeholder="Öğrenci adı, okul no, sayfa veya ayar ara... (Örn: Ahmet, yedek, tema, aşı)" autocomplete="off" oninput="window.__revirCmdSearch(this.value)" onkeydown="window.__revirCmdKey(event)"/>
			<span class="revir-kbd">ESC</span>
			<button type="button" class="revir-cmd-close" title="Kapat (ESC)" aria-label="Komut paletini kapat" onclick="window.__revirCloseCmdPalette()">✕</button>
		</div>
		<div id="revir-cmd-results"></div>
		<div class="revir-cmd-footer">
			<span>↑↓ Gezin · ↵ Seç · ESC Kapat</span>
			<span>Kısayollar: Ctrl/⌘K Ara · Ctrl/⌘J Sohbet · Alt+V Yeni Vizit · Alt+B Menü</span>
		</div>
	</div>
</div>
<div id="revir-chart-tooltip"></div>
<div id="revir-resize-badge"></div>
<div id="revir-toast" onclick="this.style.display='none'" title="Kapatmak için tıklayın"></div>
<div id="revir-layout-bar">
	<span>🧩 Kutucuk Düzeni (Takvim Hariç) — Sürükle veya Kenardan Boyutlandır</span>
	<button type="button" class="revir-layout-bar-btn" onclick="window.__revirResetPageBoxLayout()" title="Bu sayfadaki tüm kutucuk konum ve boyutlarını varsayılana döndür">↺ Sayfayı Sıfırla</button>
	<button type="button" class="revir-layout-bar-btn primary" onclick="window.__revirToggleLayoutEdit(false)">✓ Bitir</button>
</div>

<!-- Yüzen Gemini Sohbet Butonu ve Konuşma Balonu -->
<button id="revir-fab-btn" type="button" title="Gemini Klinik Asistan ile Konuş (Ctrl/⌘J / Alt+A)" aria-label="Gemini Klinik Asistan" onclick="window.__revirToggleChatBubble()">
	<svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
		<path d="M12 2L14.6 9.4L22 12L14.6 14.6L12 22L9.4 14.6L2 12L9.4 9.4L12 2Z" fill="currentColor"/>
		<circle cx="19" cy="5" r="1.4" fill="currentColor" opacity="0.85"/>
	</svg>
</button>

<div id="revir-chat-bubble" ondragover="event.preventDefault()" ondrop="window.__revirChatDropFiles(event)">
	<div id="revir-chat-resize-nw" class="revir-chat-resize-nw" title="Sohbet penceresini sürükleyerek boyutlandırın · Sıfırlamak için çift tıklayın"></div>
	<div class="revir-chat-hdr">
		<div class="revir-chat-hdr-left">
			<div class="revir-chat-avatar">
				<svg viewBox="0 0 24 24" width="20" height="20" fill="none" aria-hidden="true">
					<path d="M12 2L14.6 9.4L22 12L14.6 14.6L12 22L9.4 14.6L2 12L9.4 9.4L12 2Z" fill="currentColor"/>
				</svg>
			</div>
			<div class="revir-chat-title-wrap">
				<div id="revir-chat-title" class="revir-chat-title">Gemini Klinik Asistan</div>
				<div id="revir-chat-sub" class="revir-chat-sub">🛡️ KVKK Korumalı · Akıllı Asistan &amp; Dosya Destekli</div>
			</div>
		</div>
		<div class="revir-chat-hdr-btns">
			<button id="revir-chat-pin-btn" type="button" class="revir-chat-hbtn" title="Sohbeti Sabitle (Dışarı Tıklayınca Kapanmaz)" onclick="window.__revirToggleChatPin()">📌</button>
			<button type="button" class="revir-chat-hbtn" title="Sohbeti Temizle" onclick="window.__revirClearChat()">🗑️</button>
			<button type="button" class="revir-chat-hbtn" title="Pencereyi Büyüt / Küçült" onclick="window.__revirExpandChat()">⤢</button>
			<button type="button" class="revir-chat-hbtn" title="Küçült" onclick="window.__revirToggleChatBubble()">✕</button>
		</div>
	</div>
	<div id="revir-chat-chips" class="revir-chat-chips">
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Seçili kişinin sağlık, büyüme, özel gereksinim ve revir profilini kısaca özetle.')">🩺 Kişiyi Özetle</button>
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Seçili kişinin kapsamlı sağlık raporunu Word (.docx) dosyası olarak hazırla.')">📄 Word Raporu</button>
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Bugünkü revir başvurularını Excel (.xlsx) tablosu olarak hazırla ve özetle.')">📊 Excel Liste</button>
		<button type="button" class="revir-chat-qchip" onclick="window.__revirQuickPrompt('Engelli veya özel gereksinimli öğrencilerin listesini Excel olarak çıkar.')">♿ Özel Gereksinim Excel</button>
		<button type="button" class="revir-chat-qchip" onclick="document.getElementById('revir-chat-file-inp').click()">📎 Dosya Ekle</button>
	</div>
	<div id="revir-chat-messages" class="revir-chat-msgs">
		<div class="revir-cmsg revir-cmsg-ai">
			<p>Merhaba! Ben <strong id="revir-chat-welcome-name">Gemini Klinik Asistan</strong>. 🩺</p>
			<p>Bana yazılı soru sorabilir, <strong>📎 dosya / PDF</strong> ekleyebilir veya <em>"Bu öğrencinin raporunu Word yap"</em> / <em>"Ekli PDF'den sınıfları ve öğrencileri oluştur"</em> diyerek işlem başlatabilirsiniz.</p>
		</div>
	</div>
	<div id="revir-chat-files" class="revir-chat-files"></div>
	<div class="revir-chat-composer">
		<label class="revir-chat-ibtn" title="Dosya / Tahlil / Görsel Ekle (.pdf, .png, .jpg, .xlsx, .docx, .csv)">
			📎
			<input id="revir-chat-file-inp" type="file" multiple style="display:none" accept=".pdf,.png,.jpg,.jpeg,.webp,.xlsx,.xls,.csv,.docx,.txt" onchange="window.__revirChatAddFiles(this.files);this.value='';"/>
		</label>
		<textarea id="revir-chat-inp" rows="1" placeholder="Mesaj yazın veya dosya sürükleyin..." onkeydown="if(event.key==='Enter' && !event.shiftKey){event.preventDefault();window.__revirSendChat();}"></textarea>
		<button id="revir-chat-send-btn" type="button" class="revir-chat-ibtn" title="Gönder" onclick="window.__revirSendChat()">➤</button>
	</div>
</div>

<script>
(function() {
	var CSRF = "__REVIR_CSRF__";
	var pendingInputs = {};
	var pendingEditorCells = [];
	var inflight = false;
	var queuedReqs = [];
	var activeClickKey = null;
	var lastFocusedEditorId = null;
	var modalDirty = false;
	var modalBackdropWarned = false;
	var lastDialogSig = "";
	var lastDialogTitle = "";
	var modalReturnFocus = null;
	var commandReturnFocus = null;
	var lastRenderedPage = "";
	var loadingPillTimer = null;
	var cmdItems = [];
	var cmdActiveIdx = 0;
	var cmdTimer = null;
	var hypoTimers = {};
	var chatPinned = false;
	var boxLayoutEnabled = true;

	function setModalDirty(dirty) {
		modalDirty = !!dirty;
		var pill = document.getElementById("revir-modal-dirty-pill");
		if (pill) pill.classList.toggle("visible", modalDirty);
	}

	function isCalendarPageName(pg) {
		return String(pg || "").toLocaleLowerCase("tr-TR").indexOf("takvim") !== -1;
	}

	function showToast(msg, kind) {
		var el = document.getElementById("revir-toast");
		if (!el) return;
		el.textContent = msg;
		el.classList.remove("toast-err", "toast-warn");
		var k = kind || (String(msg).indexOf("❌") !== -1 ? "err" : (String(msg).indexOf("⚠️") !== -1 ? "warn" : ""));
		if (k === "err") el.classList.add("toast-err");
		else if (k === "warn") el.classList.add("toast-warn");
		el.style.display = "block";
		clearTimeout(el.__t);
		el.__t = setTimeout(function() { el.style.display = "none"; }, 3600);
	}

	function setBusy(on, labelTxt) {
		var bar = document.getElementById("revir-busy-bar");
		var pill = document.getElementById("revir-loading-pill");
		if (bar) {
			if (on) bar.classList.add("active");
			else bar.classList.remove("active");
		}
		document.body.classList.toggle("revir-busy", !!on);
		clearTimeout(loadingPillTimer);
		if (on) {
			if (pill && labelTxt) {
				var sp = pill.querySelector("span:last-child");
				if (sp) sp.textContent = labelTxt;
			}
			loadingPillTimer = setTimeout(function() {
				if (pill && inflight) pill.classList.add("active");
			}, 260);
		} else {
			if (pill) {
				pill.classList.remove("active");
				var sp2 = pill.querySelector("span:last-child");
				if (sp2) sp2.textContent = "İşlem yürütülüyor...";
			}
		}
	}

	window.__revirStartHypoTimer = function(clockId, totalSec) {
		var sec = parseInt(totalSec || 900, 10);
		if (hypoTimers[clockId]) clearInterval(hypoTimers[clockId]);
		var updateDisp = function() {
			var el = document.getElementById(clockId);
			if (!el) {
				clearInterval(hypoTimers[clockId]);
				delete hypoTimers[clockId];
				return;
			}
			var mm = Math.floor(sec / 60);
			var ss = sec % 60;
			el.textContent = (mm < 10 ? "0" : "") + mm + ":" + (ss < 10 ? "0" : "") + ss;
		};
		updateDisp();
		showToast("⏱️ 15-15 Hipoglisemi sayacı başlatıldı — 15 dk sonra KŞ tekrar ölçülecek!", "warn");
		hypoTimers[clockId] = setInterval(function() {
			sec -= 1;
			if (sec <= 0) {
				sec = 0;
				updateDisp();
				clearInterval(hypoTimers[clockId]);
				delete hypoTimers[clockId];
				showToast("🚨 15 DAKİKA DOLDU! Öğrencinin kan şekerini şimdi tekrar ölçün!", "err");
				try {
					if (window.__revirSpeakText) window.__revirSpeakText("Dikkat, hipoglisemi on beş dakika sayacı doldu. Öğrencinin kan şekerini tekrar ölçün.", null);
				} catch (e) {}
				return;
			}
			updateDisp();
		}, 1000);
	};

	window.__revirResetHypoTimer = function(clockId) {
		if (hypoTimers[clockId]) {
			clearInterval(hypoTimers[clockId]);
			delete hypoTimers[clockId];
		}
		var el = document.getElementById(clockId);
		if (el) el.textContent = "15:00";
	};

	function sendAction(payload, opts) {
		opts = opts || {};
		payload = Object.assign({}, payload);
		var clickKey = payload.action === "click" && payload.key ? String(payload.key) : null;
		if (clickKey && (
			clickKey === activeClickKey ||
			queuedReqs.some(function(req) {
				return req.payload.action === "click" && String(req.payload.key || "") === clickKey;
			})
		)) {
			return;
		}

		payload.pending_inputs = Object.assign({}, payload.pending_inputs || {}, pendingInputs);
		payload.pending_editor_cells = (payload.pending_editor_cells || []).concat(pendingEditorCells);
		pendingInputs = {};
		pendingEditorCells = [];

		if (inflight) {
			queuedReqs.push({ payload: payload, opts: opts });
			return;
		}
		inflight = true;
		activeClickKey = clickKey;
		setBusy(true);

		var mainEl = document.getElementById("revir-main");
		var dlgBody = document.getElementById("revir-modal-body");
		var savedMainScroll = mainEl ? mainEl.scrollTop : 0;
		var savedDlgScroll = dlgBody ? dlgBody.scrollTop : 0;
		var savedSbScroll = (document.getElementById("revir-sidebar") || {}).scrollTop || 0;

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
			activeClickKey = null;
			setBusy(false);
			if (queuedReqs.length) {
				var nxt = queuedReqs.shift();
				sendAction(nxt.payload, nxt.opts);
			}
		});
	}

	function describeFocusTarget(el) {
		if (!el) return null;
		return { element: el, id: el.id || "", key: el.getAttribute("data-key") || "" };
	}

	function restoreFocusTarget(target) {
		if (!target) return;
		var el = target.element && target.element.isConnected ? target.element : null;
		if (!el && target.id) el = document.getElementById(target.id);
		if (!el && target.key) {
			el = document.querySelector('[data-key="' + CSS.escape(target.key) + '"]');
		}
		if (el && el.focus) {
			try { el.focus({ preventScroll: true }); } catch (_fe) { el.focus(); }
		}
	}

	function applyRender(data, mainScroll, dlgScroll, sbScroll, refocusId) {
		if (!data) return;
		cleanupPortaledMenus();
		if (data.title) document.title = data.title;

		var dynStyles = document.getElementById("revir-dynamic-styles");
		if (dynStyles && dynStyles.dataset.hash !== String((data.styles_html || "").length) + ":" + (data.styles_html || "").slice(0, 64)) {
			dynStyles.innerHTML = data.styles_html || "";
			dynStyles.dataset.hash = String((data.styles_html || "").length) + ":" + (data.styles_html || "").slice(0, 64);
		}

		var tb = document.getElementById("revir-topbar");
		if (tb && data.topbar_html !== undefined) {
			tb.innerHTML = data.topbar_html || "";
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
		var nextPage = String(data.active_page || "");
		var pageChanged = !!(lastRenderedPage && nextPage && nextPage !== lastRenderedPage);
		lastRenderedPage = nextPage || lastRenderedPage;
		if (mainEl) mainEl.scrollTop = pageChanged ? 0 : mainScroll;

		var backdrop = document.getElementById("revir-modal-backdrop");
		var card = document.getElementById("revir-modal-card");
		var mTitle = document.getElementById("revir-modal-title");
		var mBody = document.getElementById("revir-modal-body");
		var mBack = document.getElementById("revir-modal-back");
		var mFs = document.getElementById("revir-modal-fs-btn");
		if (data.dialog_html !== null && data.dialog_html !== undefined) {
			var modalWasOpen = !!(backdrop && backdrop.classList.contains("open"));
			var curSig = String(data.dialog_title || "") + "|" + String(data.dialog_depth || 1) + "|" + String(data.dialog_mode || "modal");
			var isNewSig = (curSig !== lastDialogSig);
			if (isNewSig) {
				if (!modalWasOpen) modalReturnFocus = describeFocusTarget(document.activeElement);
				setModalDirty(false);
				modalBackdropWarned = false;
				if (card) card.classList.remove("fullscreen");
				lastDialogSig = curSig;
			}
			if (mTitle) mTitle.textContent = data.dialog_title || "";
			if (mBody) {
				mBody.innerHTML = data.dialog_html;
				mBody.scrollTop = dlgScroll;
			}
			var isDrawer = (data.dialog_mode === "drawer");
			lastDialogTitle = String(data.dialog_title || "");
			if (backdrop) {
				backdrop.classList.add("open");
				backdrop.classList.toggle("drawer-mode", isDrawer);
			}
			if (card) {
				card.classList.toggle("small", !isDrawer && data.dialog_width === "small");
				card.classList.toggle("drawer", isDrawer);
				applySavedModalSize(card, lastDialogTitle, isDrawer);
			}
			if (mBack) mBack.classList.toggle("visible", Number(data.dialog_depth || 0) > 1);
			if (mFs) {
				mFs.style.display = isDrawer ? "none" : "flex";
				var isFsNow = !!(card && card.classList.contains("fullscreen"));
				mFs.classList.toggle("active", isFsNow);
				mFs.textContent = isFsNow ? "⤡" : "⤢";
			}
			document.body.classList.add("modal-open");
			if (isNewSig && mBody && !refocusId) {
				setTimeout(function() {
					var firstInp = mBody.querySelector("input:not([type=hidden]):not([type=file]):not([disabled]), textarea:not([disabled]), select:not([disabled])");
					if (firstInp && firstInp.focus) {
						try { firstInp.focus({ preventScroll: true }); } catch (_fe) {}
					} else if (card && card.focus) {
						card.setAttribute("tabindex", "-1");
						card.focus();
					}
				}, 40);
			}
		} else {
			var shouldRestoreModalFocus = !!(backdrop && backdrop.classList.contains("open"));
			lastDialogTitle = "";
			if (backdrop) {
				backdrop.classList.remove("open");
				backdrop.classList.remove("drawer-mode");
			}
			if (card) {
				card.classList.remove("drawer");
				card.classList.remove("fullscreen");
				card.style.width = "";
				card.style.maxWidth = "";
				card.style.height = "";
				card.style.maxHeight = "";
			}
			if (mBody) mBody.innerHTML = "";
			setModalDirty(false);
			modalBackdropWarned = false;
			lastDialogSig = "";
			if (mFs) {
				mFs.classList.remove("active");
				mFs.textContent = "⤢";
			}
			document.body.classList.remove("modal-open");
			if (shouldRestoreModalFocus && modalReturnFocus) {
				var returnTarget = modalReturnFocus;
				modalReturnFocus = null;
				setTimeout(function() { restoreFocusTarget(returnTarget); }, 0);
			}
		}

		// Expander durumlarını dinle
		document.querySelectorAll("details.revir-expander").forEach(function(det) {
			if (det.dataset.expBound === "1") return;
			det.dataset.expBound = "1";
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

		if (data.box_layout_enabled !== undefined) {
			boxLayoutEnabled = !!data.box_layout_enabled;
		}
		if (isCalendarPageName(lastRenderedPage) || !boxLayoutEnabled) {
			document.body.classList.remove("revir-layout-edit");
		}
		document.body.classList.toggle("revir-box-resize-on", !!(boxLayoutEnabled && !isCalendarPageName(lastRenderedPage)));
		if (window.__revirSyncLayoutBtn) window.__revirSyncLayoutBtn();

		if (window.__revirInitResizableTables) window.__revirInitResizableTables();
		if (window.__revirInitBoxLayout) window.__revirInitBoxLayout();

		// Konfeti animasyonu
		if (data.balloons) launchConfetti();

		// Yüzen Gemini sohbet butonu görünürlüğü ve alt başlığı
		var fabBtn = document.getElementById("revir-fab-btn");
		var chatBubble = document.getElementById("revir-chat-bubble");
		var chatSub = document.getElementById("revir-chat-sub");
		var chatTitleEl = document.getElementById("revir-chat-title");
		var chatWelcomeEl = document.getElementById("revir-chat-welcome-name");
		if (fabBtn) {
			if (data.floating_chat_enabled) {
				fabBtn.classList.add("visible");
			} else {
				fabBtn.classList.remove("visible");
				if (chatBubble) chatBubble.classList.remove("open");
			}
		}
		if (chatTitleEl && data.floating_chat_title) {
			chatTitleEl.textContent = data.floating_chat_title;
			if (chatWelcomeEl) chatWelcomeEl.textContent = data.floating_chat_title;
			if (fabBtn) {
				fabBtn.title = data.floating_chat_title + " ile Konuş (Ctrl/⌘J / Alt+A)";
				fabBtn.setAttribute("aria-label", data.floating_chat_title);
			}
		}
		if (chatSub && data.floating_chat_label) {
			chatSub.textContent = data.floating_chat_label;
		}
		var chatChipsEl = document.getElementById("revir-chat-chips");
		if (chatChipsEl && Array.isArray(data.floating_chat_chips) && data.floating_chat_chips.length) {
			var chipParts = data.floating_chat_chips.map(function(ch) {
				if (ch.action === "upload") {
					return '<button type="button" class="revir-chat-qchip" onclick="document.getElementById(\'revir-chat-file-inp\').click()">' + escHtml(ch.label || "📎 Dosya Ekle") + "</button>";
				}
				var pJson = JSON.stringify(String(ch.prompt || ""));
				return '<button type="button" class="revir-chat-qchip" onclick=\'window.__revirQuickPrompt(' + escHtml(pJson) + ')\'>' + escHtml(ch.label || "") + "</button>";
			});
			chatChipsEl.innerHTML = chipParts.join("");
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
		if (grp) {
			grp.dataset.value = String(idx);
			grp.querySelectorAll(".revir-nav-item").forEach(function(button, buttonIdx) {
				if (buttonIdx === idx) button.setAttribute("aria-current", "page");
				else button.removeAttribute("aria-current");
			});
		}
		if (grp && grp.dataset && grp.dataset.form) return;
		sendAction({ action: "change", key: key, value: idx });
	};

	window.__revirSubmitForm = function(formId) {
		setModalDirty(false);
		modalBackdropWarned = false;
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

	// Portal tabanlı Fixed-Viewport Açılır Menü Konumlandırıcı (Expander/Kart/Modal Kesilmesini Tamamen Önler)
	function portalAndPositionMenu(wrap, menu) {
		if (!wrap || !menu) return;
		var ctrl = wrap.querySelector(".revir-combo-control, .revir-multi-control") || wrap;
		if (menu.dataset.portaled !== "1") {
			menu.__origParent = menu.parentNode;
			menu.__origNext = menu.nextSibling;
			menu.__anchorWrap = wrap;
			menu.dataset.portaled = "1";
			menu.classList.add("revir-portaled");
			document.body.appendChild(menu);
		} else {
			menu.__anchorWrap = wrap;
		}
		var rect = ctrl.getBoundingClientRect();
		var vw = window.innerWidth || document.documentElement.clientWidth || 1200;
		var vh = window.innerHeight || document.documentElement.clientHeight || 800;
		var spaceBelow = vh - rect.bottom - 12;
		var spaceAbove = rect.top - 12;
		var menuW = Math.min(Math.max(rect.width, 220), vw - 20);
		var leftPos = Math.max(10, Math.min(rect.left, vw - menuW - 10));
		menu.style.left = leftPos + "px";
		menu.style.width = menuW + "px";
		if (spaceBelow < 190 && spaceAbove > spaceBelow) {
			var maxHUp = Math.max(120, Math.min(280, spaceAbove - 6));
			menu.style.top = "auto";
			menu.style.bottom = (vh - rect.top + 4) + "px";
			menu.style.maxHeight = maxHUp + "px";
		} else {
			var maxHDn = Math.max(120, Math.min(280, spaceBelow - 4));
			menu.style.bottom = "auto";
			menu.style.top = (rect.bottom + 4) + "px";
			menu.style.maxHeight = maxHDn + "px";
		}
	}

	function restorePortaledMenu(menu) {
		if (!menu) return;
		menu.classList.remove("open");
		if (menu.dataset.portaled === "1") {
			menu.classList.remove("revir-portaled");
			menu.dataset.portaled = "0";
			menu.style.left = "";
			menu.style.top = "";
			menu.style.bottom = "";
			menu.style.width = "";
			menu.style.maxHeight = "";
			if (menu.__origParent && menu.__origParent.isConnected) {
				if (menu.__origNext && menu.__origNext.parentNode === menu.__origParent) {
					menu.__origParent.insertBefore(menu, menu.__origNext);
				} else {
					menu.__origParent.appendChild(menu);
				}
			} else if (menu.parentNode === document.body) {
				menu.remove();
			}
		}
	}

	function cleanupPortaledMenus() {
		document.querySelectorAll('.revir-combo-menu[data-portaled="1"]').forEach(function(m) {
			m.remove();
		});
	}

	function repositionOpenMenus(ev) {
		var openMenus = document.querySelectorAll('.revir-combo-menu.open[data-portaled="1"]');
		if (!openMenus.length) return;
		openMenus.forEach(function(m) {
			if (ev && ev.target && (ev.target === m || (m.contains && m.contains(ev.target)))) return;
			var w = m.__anchorWrap;
			if (!w || !w.isConnected) {
				restorePortaledMenu(m);
				return;
			}
			var r = w.getBoundingClientRect();
			if (r.bottom < 0 || r.top > (window.innerHeight || 800)) {
				restorePortaledMenu(m);
				return;
			}
			portalAndPositionMenu(w, m);
		});
	}

	window.addEventListener("scroll", repositionOpenMenus, true);
	window.addEventListener("resize", repositionOpenMenus);

	// Aranabilir Selectbox (Klavye ↑↓↵ + Fare Destekli)
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
			var kbdCls = html.length === 0 ? " kbd-active" : "";
			html.push('<div class="revir-combo-opt' + sel + kbdCls + '" data-opt-idx="' + i + '" onmousedown="window.__revirSelectPick(\'' + key.replace(/'/g, "\\'") + '\', ' + i + ', event)">' + escHtml(txt) + '</div>');
			if (html.length >= 200) break;
		}
		if (!html.length) html.push('<div class="revir-combo-opt" style="color:var(--text-muted);cursor:default;">Sonuç bulunamadı</div>');
		menu.innerHTML = html.join("");
	}

	window.__revirComboKey = function(key, ev) {
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("combo_" + key);
		if (!menu) return;
		if (ev.key === "Escape") {
			restorePortaledMenu(menu);
			return;
		}
		if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
			ev.preventDefault();
			if (!menu.classList.contains("open")) {
				renderComboMenu(key, "");
				portalAndPositionMenu(wrap, menu);
				menu.classList.add("open");
			}
			var items = Array.prototype.slice.call(menu.querySelectorAll(".revir-combo-opt[data-opt-idx]"));
			if (!items.length) return;
			var curI = items.findIndex(function(x) { return x.classList.contains("kbd-active"); });
			items.forEach(function(x) { x.classList.remove("kbd-active"); });
			var nxt = ev.key === "ArrowDown" ? (curI + 1) % items.length : (curI <= 0 ? items.length - 1 : curI - 1);
			items[nxt].classList.add("kbd-active");
			if (items[nxt].scrollIntoView) { try { items[nxt].scrollIntoView({ block: "nearest" }); } catch (e) {} }
			return;
		}
		if (ev.key === "Enter") {
			var act = menu.querySelector(".revir-combo-opt.kbd-active[data-opt-idx]");
			if (act && menu.classList.contains("open")) {
				ev.preventDefault();
				window.__revirSelectPick(key, parseInt(act.dataset.optIdx, 10), ev);
			}
		}
	};

	window.__revirOpenCombo = function(key, ev) {
		if (ev) ev.stopPropagation();
		closeAllMenus("combo_" + key);
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("combo_" + key);
		if (!menu) return;
		renderComboMenu(key, "");
		portalAndPositionMenu(wrap, menu);
		menu.classList.add("open");
	};

	window.__revirFilterCombo = function(key, val) {
		var wrap = document.querySelector('.revir-combo[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("combo_" + key);
		if (menu) {
			portalAndPositionMenu(wrap, menu);
			if (!menu.classList.contains("open")) menu.classList.add("open");
		}
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
		if (menu) restorePortaledMenu(menu);
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

	// Multiselect (Klavye ↑↓↵ Destekli)
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
			var kbdCls = html.length === 0 ? " kbd-active" : "";
			html.push('<div class="revir-combo-opt' + (isSel ? " selected" : "") + kbdCls + '" data-opt-idx="' + i + '" onmousedown="window.__revirMultiToggle(\'' + key.replace(/'/g, "\\'") + '\', ' + i + ', event)">'
				+ '<span>' + (isSel ? "☑" : "☐") + '</span><span>' + escHtml(txt) + '</span></div>');
			if (html.length >= 200) break;
		}
		if (!html.length) html.push('<div class="revir-combo-opt" style="color:var(--text-muted);">Sonuç bulunamadı</div>');
		menu.innerHTML = html.join("");
	}

	window.__revirMultiKey = function(key, ev) {
		var wrap = document.querySelector('.revir-multi[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("multi_" + key);
		if (!menu) return;
		if (ev.key === "Escape") { restorePortaledMenu(menu); return; }
		if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
			ev.preventDefault();
			if (!menu.classList.contains("open")) { renderMultiMenu(key, ""); portalAndPositionMenu(wrap, menu); menu.classList.add("open"); }
			var items = Array.prototype.slice.call(menu.querySelectorAll(".revir-combo-opt[data-opt-idx]"));
			if (!items.length) return;
			var curI = items.findIndex(function(x) { return x.classList.contains("kbd-active"); });
			items.forEach(function(x) { x.classList.remove("kbd-active"); });
			var nxt = ev.key === "ArrowDown" ? (curI + 1) % items.length : (curI <= 0 ? items.length - 1 : curI - 1);
			items[nxt].classList.add("kbd-active");
			if (items[nxt].scrollIntoView) { try { items[nxt].scrollIntoView({ block: "nearest" }); } catch (e) {} }
			return;
		}
		if (ev.key === "Enter") {
			var act = menu.querySelector(".revir-combo-opt.kbd-active[data-opt-idx]");
			if (act && menu.classList.contains("open")) {
				ev.preventDefault();
				window.__revirMultiToggle(key, parseInt(act.dataset.optIdx, 10), ev);
			}
		}
	};

	window.__revirOpenMulti = function(key, ev) {
		if (ev) ev.stopPropagation();
		closeAllMenus("multi_" + key);
		var wrap = document.querySelector('.revir-multi[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("multi_" + key);
		if (!menu) return;
		renderMultiMenu(key, "");
		portalAndPositionMenu(wrap, menu);
		menu.classList.add("open");
	};

	window.__revirFilterMulti = function(key, val) {
		var wrap = document.querySelector('.revir-multi[data-key="' + CSS.escape(key) + '"]');
		var menu = document.getElementById("multi_" + key);
		if (menu) {
			portalAndPositionMenu(wrap, menu);
			if (!menu.classList.contains("open")) menu.classList.add("open");
		}
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
		var tablist = document.getElementById(groupId);
		var currentButton = tablist ? tablist.querySelector('.revir-tab-btn[aria-selected="true"]') : null;
		if (currentButton && parseInt(currentButton.dataset.tabIdx, 10) === idx) return;
		// Yalnızca ilgili sekme grubunun kendi lazy niteliğine bak (üst lazy sekme grubuna tırmanma)
		var ownTabsWrap = tablist ? tablist.closest(".revir-tabs") : null;
		var shouldRerender = Boolean(ownTabsWrap && ownTabsWrap.getAttribute("data-rerender-on-tab") === "true");
		document.querySelectorAll('.revir-tab-btn[data-tab-group="' + groupId + '"]').forEach(function(b) {
			var selected = parseInt(b.dataset.tabIdx, 10) === idx;
			b.classList.toggle("active", selected);
			b.setAttribute("aria-selected", String(selected));
			b.setAttribute("tabindex", selected ? "0" : "-1");
		});
		document.querySelectorAll('.revir-tab-pane[data-tab-group="' + groupId + '"]').forEach(function(p) {
			var selected = parseInt(p.dataset.tabIdx, 10) === idx;
			p.classList.toggle("active", selected);
			p.setAttribute("aria-hidden", String(!selected));
		});
		if (!shouldRerender) {
			setTimeout(function() {
				if (window.__revirInitResizableTables) window.__revirInitResizableTables();
				if (window.__revirInitBoxLayout) window.__revirInitBoxLayout();
			}, 0);
		}
		fetch("/api/ui_state", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({ kind: "tab", id: groupId, value: idx })
		}).then(function(response) {
			if (response.ok && shouldRerender) sendAction({ action: "render" });
		}).catch(function(){});
	};
	window.__revirTabKey = function(ev, groupId, idx, orientation) {
		var next = idx;
		var forward = orientation === "vertical" ? "ArrowDown" : "ArrowRight";
		var backward = orientation === "vertical" ? "ArrowUp" : "ArrowLeft";
		var buttons = Array.prototype.slice.call(document.querySelectorAll('.revir-tab-btn[data-tab-group="' + groupId + '"]'));
		if (ev.key === forward) next = (idx + 1) % buttons.length;
		else if (ev.key === backward) next = (idx - 1 + buttons.length) % buttons.length;
		else if (ev.key === "Home") next = 0;
		else if (ev.key === "End") next = buttons.length - 1;
		else return;
		ev.preventDefault();
		window.__revirSwitchTab(groupId, next);
		if (buttons[next]) buttons[next].focus();
	};

	// Popover (Ekran kenarlarında otomatik yukarı/sağa-sola hizalama destekli)
	window.__revirClosePopovers = function(ev) {
		if (ev) ev.stopPropagation();
		document.querySelectorAll(".revir-popover-body.open").forEach(function(m) {
			m.classList.remove("open", "flip-up", "align-left");
		});
	};

	window.__revirTogglePopover = function(popId, ev) {
		if (ev) ev.stopPropagation();
		var el = document.getElementById("pop_" + popId);
		if (!el) return;
		var wasOpen = el.classList.contains("open");
		closeAllMenus();
		if (!wasOpen) {
			el.classList.remove("flip-up", "align-left");
			el.style.maxHeight = "";
			el.classList.add("open");
			try {
				var vw = window.innerWidth || document.documentElement.clientWidth || 1200;
				var vh = window.innerHeight || document.documentElement.clientHeight || 800;
				var rect = el.getBoundingClientRect();
				if (rect.left < 12) {
					el.classList.add("align-left");
					rect = el.getBoundingClientRect();
					if (rect.right > vw - 12) el.classList.remove("align-left");
				}
				rect = el.getBoundingClientRect();
				if (rect.bottom > vh - 16) {
					var wrapRect = el.parentElement ? el.parentElement.getBoundingClientRect() : rect;
					var spaceAbove = wrapRect.top - 16;
					var spaceBelow = vh - wrapRect.bottom - 16;
					if (spaceAbove > spaceBelow) {
						el.classList.add("flip-up");
						el.style.maxHeight = Math.max(180, Math.min(Math.round(vh * 0.72), Math.round(spaceAbove))) + "px";
					} else {
						el.style.maxHeight = Math.max(180, Math.min(Math.round(vh * 0.72), Math.round(spaceBelow))) + "px";
					}
				}
			} catch (e) {}
		}
	};

	function closeAllMenus(exceptId) {
		document.querySelectorAll(".revir-combo-menu.open, .revir-popover-body.open").forEach(function(m) {
			if (m.id !== exceptId) {
				if (m.classList.contains("revir-combo-menu")) restorePortaledMenu(m);
				else m.classList.remove("open");
			}
		});
	}

	// Dataframe seçimi (metin seçilip kopyalanıyorsa tıklama tetikleme)
	window.__revirDfSelect = function(key, rowIdx, colName) {
		try {
			var sel = window.getSelection && window.getSelection();
			if (sel && String(sel.toString() || "").trim().length > 0) return;
		} catch (_e) {}
		sendAction({ action: "df_select", key: key, value: { rows: [rowIdx], cells: [[rowIdx, colName]] } });
	};

	// İstemci tarafı tablo sıralama (görsel yön oku destekli)
	window.__revirSortTable = function(th) {
		var tbl = th.closest("table");
		if (!tbl || tbl.classList.contains("revir-editor-table")) return;
		if (tbl.dataset.colResizeAt && (Date.now() - parseInt(tbl.dataset.colResizeAt, 10)) < 320) return;
		var tbody = tbl.querySelector("tbody");
		if (!tbody) return;
		var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
		var asc = th.dataset.sortDir !== "asc";
		Array.prototype.forEach.call(th.parentNode.children, function(sib) {
			if (sib !== th) sib.dataset.sortDir = "";
			var ind = sib.querySelector(".revir-sort-ind");
			if (ind) {
				ind.textContent = "⇅";
				ind.classList.remove("active");
			}
		});
		th.dataset.sortDir = asc ? "asc" : "desc";
		var curInd = th.querySelector(".revir-sort-ind");
		if (curInd) {
			curInd.textContent = asc ? "▲" : "▼";
			curInd.classList.add("active");
		}
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

	// Data Editor (Excel tarzı Enter ile alt satıra geçiş)
	var nextEditorFocusId = null;
	window.__revirEditorEnter = function(inp) {
		if (!inp) return;
		var tbl = inp.closest("table.revir-editor-table");
		var r = parseInt(inp.dataset.edRow || "-1", 10);
		var c = parseInt(inp.dataset.edCol || "-1", 10);
		if (tbl && r >= 0 && c >= 0) {
			var nextEl = tbl.querySelector('[data-ed-row="' + (r + 1) + '"][data-ed-col="' + c + '"]');
			if (nextEl && nextEl.id) {
				nextEditorFocusId = nextEl.id;
			}
		}
		inp.blur();
		if (nextEditorFocusId) {
			var targetNow = document.getElementById(nextEditorFocusId);
			if (targetNow && !inflight) {
				try { targetNow.focus(); if (targetNow.select) targetNow.select(); } catch (e) {}
				nextEditorFocusId = null;
			}
		}
	};

	window.__revirEditorDirty = function(key, rowKind, rowId, col, val) {
		pendingEditorCells.push({ key: key, op: "cell", row_kind: rowKind, row_id: rowId, col: col, val: val });
	};

	window.__revirEditorCell = function(key, rowKind, rowId, col, val, cellId) {
		pendingEditorCells = pendingEditorCells.filter(function(m) {
			return !(m.key === key && m.row_kind === rowKind && m.row_id === rowId && m.col === col);
		});
		var targetFocus = nextEditorFocusId || cellId;
		nextEditorFocusId = null;
		sendAction(
			{ action: "data_editor", key: key, value: { op: "cell", row_kind: rowKind, row_id: rowId, col: col, val: val } },
			{ refocusId: targetFocus }
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

	// Yerel Masaüstü Dosya İndirme (Native Save Dialog + Tembel Üretim destekli)
	window.__revirDownload = function(token, fileName) {
		showToast("⏳ Dosya hazırlanıyor: " + (fileName || "rapor"));
		if (window.pywebview && window.pywebview.api && window.pywebview.api.save_download) {
			window.pywebview.api.save_download(token, fileName).then(function(res) {
				if (res && res.saved) showToast("✅ Kaydedildi: " + res.path);
				else if (res && res.error) showToast("❌ İndirme hatası: " + res.error, "err");
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

	// Diyalog, Üst Bar, Komut Paleti & Yan Menü
	window.__revirCloseDialog = function(closeAll) {
		setModalDirty(false);
		modalBackdropWarned = false;
		sendAction({ action: "close_dialog", all: !!closeAll });
	};

	window.__revirTryBackdropClose = function() {
		if (window.__revirLastResizeEndAt && (Date.now() - window.__revirLastResizeEndAt) < 350) {
			return;
		}
		if (modalDirty && !modalBackdropWarned) {
			modalBackdropWarned = true;
			showToast("⚠️ Kaydedilmemiş değişiklikler var — kapatmak için ✕ tuşuna veya dışarı tekrar tıklayın.");
			return;
		}
		window.__revirCloseDialog(false);
	};

	window.__revirToggleModalFullscreen = function() {
		var card = document.getElementById("revir-modal-card");
		var fsBtn = document.getElementById("revir-modal-fs-btn");
		if (!card) return;
		card.classList.toggle("fullscreen");
		var isFs = card.classList.contains("fullscreen");
		if (fsBtn) {
			fsBtn.classList.toggle("active", isFs);
			fsBtn.textContent = isFs ? "⤡" : "⤢";
		}
	};

	window.__revirResetModalSize = function() {
		var card = document.getElementById("revir-modal-card");
		if (!card) return;
		var isDrawer = card.classList.contains("drawer");
		var store = loadResizeStore(MODAL_STORE_KEY);
		var mKey = (isDrawer ? "drawer:" : "modal:") + String(lastDialogTitle || "");
		delete store[mKey];
		saveResizeStore(MODAL_STORE_KEY, store);
		card.classList.remove("fullscreen");
		card.style.width = "";
		card.style.maxWidth = "";
		card.style.height = "";
		card.style.maxHeight = "";
		var fsBtn = document.getElementById("revir-modal-fs-btn");
		if (fsBtn) {
			fsBtn.classList.remove("active");
			fsBtn.textContent = "⤢";
		}
		showToast("↺ Pencere boyutu varsayılana döndürüldü");
	};

	window.__revirTopbarAction = function(actKey, actVal) {
		sendAction({ action: "topbar_action", key: actKey, value: actVal });
	};

	window.__revirToggleSidebar = function() {
		var sb = document.getElementById("revir-sidebar");
		var isCol = sb ? sb.classList.contains("collapsed") : false;
		window.__revirSetSidebar(!isCol);
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

	// Global Komut & Kişi Arama Paleti (Cmd/Ctrl+K)
	window.__revirOpenCmdPalette = function() {
		var bd = document.getElementById("revir-cmd-backdrop");
		var inp = document.getElementById("revir-cmd-input");
		if (!bd) return;
		if (!bd.classList.contains("open")) commandReturnFocus = document.activeElement;
		bd.classList.add("open");
		if (inp) {
			inp.value = "";
			setTimeout(function() { inp.focus(); }, 30);
		}
		window.__revirCmdSearch("");
	};

	window.__revirCloseCmdPalette = function() {
		var bd = document.getElementById("revir-cmd-backdrop");
		if (bd) bd.classList.remove("open");
		if (commandReturnFocus && commandReturnFocus.isConnected) {
			var returnTarget = commandReturnFocus;
			commandReturnFocus = null;
			setTimeout(function() {
				if (returnTarget.isConnected && returnTarget.focus) returnTarget.focus({ preventScroll: true });
			}, 0);
		}
	};

	function renderCmdList() {
		var box = document.getElementById("revir-cmd-results");
		if (!box) return;
		if (!cmdItems.length) {
			box.innerHTML = '<div style="padding:24px;text-align:center;color:var(--text-muted);font-size:13.5px;">Eşleşen öğrenci, sayfa veya ayar bulunamadı.</div>';
			return;
		}
		var html = [];
		var lastGrp = "";
		for (var i = 0; i < cmdItems.length; i++) {
			var it = cmdItems[i];
			if (it.group && it.group !== lastGrp) {
				lastGrp = it.group;
				html.push('<div class="revir-cmd-group">' + escHtml(lastGrp) + '</div>');
			}
			var act = (i === cmdActiveIdx) ? " active" : "";
			html.push(
				'<div class="revir-cmd-item' + act + '" data-cmd-idx="' + i + '" onmouseenter="window.__revirCmdHover(' + i + ')" onclick="window.__revirCmdPick(' + i + ')">'
				+ '<div class="revir-cmd-item-l">'
				+ '<span class="revir-cmd-item-ico">' + escHtml(it.icon || "•") + '</span>'
				+ '<div style="min-width:0;">'
				+ '<div class="revir-cmd-item-title">' + escHtml(it.title || "") + '</div>'
				+ (it.subtitle ? '<div class="revir-cmd-item-sub">' + escHtml(it.subtitle) + '</div>' : '')
				+ '</div></div>'
				+ (it.tag ? '<span class="revir-cmd-item-tag">' + escHtml(it.tag) + '</span>' : '')
				+ '</div>'
			);
		}
		box.innerHTML = html.join("");
		var activeEl = box.querySelector(".revir-cmd-item.active");
		if (activeEl && activeEl.scrollIntoView) {
			try { activeEl.scrollIntoView({ block: "nearest" }); } catch (e) {}
		}
	}

	window.__revirCmdHover = function(idx) {
		cmdActiveIdx = idx;
		document.querySelectorAll(".revir-cmd-item").forEach(function(el) {
			el.classList.toggle("active", parseInt(el.dataset.cmdIdx, 10) === idx);
		});
	};

	window.__revirCmdPick = function(idx) {
		var it = cmdItems[idx];
		if (!it) return;
		window.__revirCloseCmdPalette();
		if (it.kind === "open_chat") {
			if (window.__revirOpenChatBubble) window.__revirOpenChatBubble();
			else window.__revirToggleChatBubble();
			return;
		}
		sendAction({ action: "cmd_exec", item: it });
	};

	window.__revirCmdSearch = function(q) {
		clearTimeout(cmdTimer);
		cmdTimer = setTimeout(function() {
			fetch("/api/cmd_search", {
				method: "POST",
				headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
				body: JSON.stringify({ q: String(q || "") })
			})
			.then(function(r) { return r.json(); })
			.then(function(res) {
				cmdItems = (res && Array.isArray(res.items)) ? res.items : [];
				cmdActiveIdx = 0;
				renderCmdList();
			})
			.catch(function() {});
		}, 70);
	};

	window.__revirCmdKey = function(ev) {
		if (ev.key === "ArrowDown") {
			ev.preventDefault();
			if (cmdItems.length) {
				cmdActiveIdx = (cmdActiveIdx + 1) % cmdItems.length;
				renderCmdList();
			}
		} else if (ev.key === "ArrowUp") {
			ev.preventDefault();
			if (cmdItems.length) {
				cmdActiveIdx = (cmdActiveIdx - 1 + cmdItems.length) % cmdItems.length;
				renderCmdList();
			}
		} else if (ev.key === "Enter") {
			ev.preventDefault();
			window.__revirCmdPick(cmdActiveIdx);
		}
	};

	function escHtml(s) {
		return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
	}

	// Modal içi kirli form takibi
	document.addEventListener("input", function(ev) {
		if (ev.target && ev.target.closest && ev.target.closest("#revir-modal-body")) {
			setModalDirty(true);
			modalBackdropWarned = false;
		}
	});

	// Genel tıklama & klavye delegasyonu
	document.addEventListener("click", function(ev) {
		var closeSb = ev.target && ev.target.closest ? ev.target.closest(".revir-sidebar-close-btn") : null;
		if (closeSb) { window.__revirToggleSidebar(); return; }
		if (!ev.target.closest(".revir-combo") && !ev.target.closest(".revir-multi") && !ev.target.closest(".revir-combo-menu") && !ev.target.closest(".revir-popover-wrap")) {
			closeAllMenus();
		}
		var chatBubble = document.getElementById("revir-chat-bubble");
		if (chatBubble && chatBubble.classList.contains("open") && ev.target && document.contains(ev.target)) {
			var insideChat = ev.target.closest && ev.target.closest("#revir-chat-bubble, #revir-fab-btn");
			if (!insideChat && !chatPinned) {
				chatBubble.classList.remove("open");
			}
		}
	});

	document.addEventListener("keydown", function(ev) {
		if (ev.key === "Tab") {
			var modalBd = document.getElementById("revir-modal-backdrop");
			var cmdBd = document.getElementById("revir-cmd-backdrop");
			var activeBackdrop = modalBd && modalBd.classList.contains("open")
				? modalBd
				: (cmdBd && cmdBd.classList.contains("open") ? cmdBd : null);
			if (activeBackdrop) {
				var focusables = Array.prototype.slice.call(activeBackdrop.querySelectorAll(
					'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])'
				)).filter(function(el) { return el.offsetWidth || el.offsetHeight || el.getClientRects().length; });
				if (!focusables.length) {
					ev.preventDefault();
					var dialog = activeBackdrop.querySelector('[role="dialog"]');
					if (dialog) {
						dialog.setAttribute("tabindex", "-1");
						dialog.focus();
					}
					return;
				}
				var first = focusables[0];
				var last = focusables[focusables.length - 1];
				if (ev.shiftKey && (document.activeElement === first || !activeBackdrop.contains(document.activeElement))) {
					ev.preventDefault();
					last.focus();
				} else if (!ev.shiftKey && (document.activeElement === last || !activeBackdrop.contains(document.activeElement))) {
					ev.preventDefault();
					first.focus();
				}
			}
		}
		if ((ev.metaKey || ev.ctrlKey) && (ev.key === "k" || ev.key === "K")) {
			ev.preventDefault();
			var cmdBd = document.getElementById("revir-cmd-backdrop");
			if (cmdBd && cmdBd.classList.contains("open")) window.__revirCloseCmdPalette();
			else window.__revirOpenCmdPalette();
			return;
		}
		if (((ev.metaKey || ev.ctrlKey) && (ev.key === "j" || ev.key === "J")) || (ev.altKey && (ev.key === "a" || ev.key === "A"))) {
			ev.preventDefault();
			var fabEl = document.getElementById("revir-fab-btn");
			if (fabEl && fabEl.classList.contains("visible")) {
				window.__revirToggleChatBubble();
			}
			return;
		}
		if (ev.altKey && (ev.key === "b" || ev.key === "B")) {
			ev.preventDefault();
			window.__revirToggleSidebar();
			return;
		}
		if (ev.altKey && (ev.key === "v" || ev.key === "V")) {
			ev.preventDefault();
			window.__revirTopbarAction("goto_page", "🩺 Revir Kaydı");
			return;
		}
		if (ev.key === "Escape") {
			closeAllMenus();
			var cmdBd2 = document.getElementById("revir-cmd-backdrop");
			if (cmdBd2 && cmdBd2.classList.contains("open")) { window.__revirCloseCmdPalette(); return; }
			var bd = document.getElementById("revir-modal-backdrop");
			if (bd && bd.classList.contains("open")) { window.__revirCloseDialog(); return; }
			var cb = document.getElementById("revir-chat-bubble");
			if (cb && cb.classList.contains("open")) { window.__revirToggleChatBubble(); return; }
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
	var chatBusy = false;

	window.__revirToggleChatBubble = function() {
		var b = document.getElementById("revir-chat-bubble");
		if (!b) return;
		b.classList.toggle("open");
		if (b.classList.contains("open")) {
			var inp = document.getElementById("revir-chat-inp");
			if (inp) setTimeout(function() { inp.focus(); }, 60);
		}
	};

	window.__revirOpenChatBubble = function() {
		var b = document.getElementById("revir-chat-bubble");
		if (!b) return;
		b.classList.add("open");
		var inp = document.getElementById("revir-chat-inp");
		if (inp) setTimeout(function() { inp.focus(); }, 60);
	};

	window.__revirExpandChat = function() {
		var b = document.getElementById("revir-chat-bubble");
		if (!b) return;
		b.classList.toggle("expanded");
		if (b.classList.contains("expanded")) {
			b.style.width = "";
			b.style.height = "";
		} else {
			applySavedChatSize();
		}
	};

	window.__revirToggleChatPin = function(forceVal) {
		chatPinned = (forceVal !== undefined) ? !!forceVal : !chatPinned;
		var pinBtn = document.getElementById("revir-chat-pin-btn");
		if (pinBtn) {
			pinBtn.classList.toggle("active", chatPinned);
			pinBtn.title = chatPinned ? "Sabitlemeyi Kaldır (Dışarı tıklayınca kapanır)" : "Sohbeti Sabitle (Dışarı Tıklayınca Kapanmaz)";
		}
		showToast(chatPinned ? "📌 Sohbet penceresi sabitlendi" : "📌 Sohbet sabitlemesi kaldırıldı");
	};

	window.__revirExportChatMsg = function(plainText, fmt) {
		if (!plainText) return;
		showToast(fmt === "xlsx" ? "📊 Excel dosyası hazırlanıyor..." : "📄 Word dosyası hazırlanıyor...");
		fetch("/api/gemini_export_msg", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({ text: plainText, format: fmt || "docx" })
		})
		.then(function(r) { return r.json(); })
		.then(function(res) {
			if (res && res.ok && res.token) {
				window.__revirDownload(res.token, res.file_name || ("gemini_rapor." + (fmt || "docx")));
			} else {
				showToast("❌ Dosya oluşturulamadı: " + ((res && res.error) || "Bilinmeyen hata"));
			}
		})
		.catch(function(err) {
			showToast("❌ Dışa aktarma hatası: " + err);
		});
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

	function appendChatMsg(role, htmlContent, plainText, downloads, writeProposal) {
		var box = document.getElementById("revir-chat-messages");
		if (!box) return null;
		var div = document.createElement("div");
		div.className = "revir-cmsg " + (role === "user" ? "revir-cmsg-user" : "revir-cmsg-ai");
		div.innerHTML = htmlContent;
		if (role === "ai" && writeProposal && writeProposal.token && !writeProposal.text_confirm) {
			var isBatch = (writeProposal.mode === "batch_action");
			var isDestructive = !!writeProposal.is_destructive;
			var wc = document.createElement("div");
			wc.className = "revir-cmsg-write-card";
			if (isDestructive) {
				wc.style.borderColor = "#ef4444";
				wc.style.background = "rgba(254,242,242,0.85)";
			}
			var hdrIcon = isDestructive ? "⚠️ " : "📝 ";
			var hdr = '<div style="font-weight:800;margin-bottom:4px;' + (isDestructive ? 'color:#b91c1c;' : '') + '">' + hdrIcon + escHtml(writeProposal.title || "Veritabanı İşlem Onayı") + "</div>";
			var selHtml = "";
			if (!isBatch && Array.isArray(writeProposal.candidates) && writeProposal.candidates.length) {
				selHtml += '<label style="font-size:11.5px;font-weight:700;color:var(--text-muted);">🎯 Hedef Öğrenci / Kişi:</label>';
				selHtml += '<select class="revir-cw-target">';
				writeProposal.candidates.forEach(function(c) {
					var selAttr = (String(c.id) === String(writeProposal.default_student_id)) ? " selected" : "";
					selHtml += '<option value="' + escHtml(String(c.id)) + '"' + selAttr + ">" + escHtml(c.label) + "</option>";
				});
				selHtml += "</select>";
			}
			var itemsHtml = "";
			if (Array.isArray(writeProposal.items) && writeProposal.items.length) {
				itemsHtml = '<ul class="revir-cmsg-write-list">' + writeProposal.items.map(function(it) { return "<li>" + escHtml(it) + "</li>"; }).join("") + "</ul>";
			}
			var archHtml = (!isBatch && writeProposal.has_attachment)
				? '<label style="display:flex;align-items:center;gap:6px;font-size:11.5px;margin:4px 0;"><input type="checkbox" class="revir-cw-arch" checked/> 📁 Ekli belgeyi Evrak Deposuna da arşivle</label>'
				: "";
			var confirmLabel = writeProposal.confirm_label || (isBatch ? "✅ Onayla ve İşlemi Uygula" : "✅ Onayla ve Kişiye İşle");
			wc.innerHTML = hdr + selHtml + itemsHtml + archHtml
				+ '<div class="revir-cmsg-write-actions">'
				+ '<button type="button" class="confirm"' + (isDestructive ? ' style="background:#dc2626;border-color:#dc2626;"' : '') + '>' + escHtml(confirmLabel) + '</button>'
				+ '<button type="button" class="cancel">✖ İptal</button>'
				+ "</div>";
			var okBtn = wc.querySelector("button.confirm");
			var noBtn = wc.querySelector("button.cancel");
			if (okBtn) {
				okBtn.onclick = function() {
					var selEl = wc.querySelector(".revir-cw-target");
					var archEl = wc.querySelector(".revir-cw-arch");
					var targetId = selEl ? parseInt(selEl.value, 10) : (writeProposal.default_student_id || null);
					okBtn.disabled = true;
					if (noBtn) noBtn.disabled = true;
					okBtn.textContent = "⏳ Uygulanıyor...";
					fetch("/api/gemini_chat_apply_write", {
						method: "POST",
						headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
						body: JSON.stringify({
							token: writeProposal.token,
							student_id: targetId,
							archive_doc: archEl ? !!archEl.checked : false
						})
					})
					.then(function(r) { return r.json(); })
					.then(function(ares) {
						if (ares && ares.ok) {
							wc.classList.add("done");
							wc.innerHTML = '<div style="font-weight:800;color:#059669;">✅ ' + escHtml(ares.summary || "İşlem başarıyla uygulandı ve ekranlar eşitlendi!") + "</div>";
							showToast("✅ " + (ares.summary || "İşlem uygulandı!"));
							sendAction({ action: "refresh" });
						} else {
							okBtn.disabled = false;
							if (noBtn) noBtn.disabled = false;
							okBtn.textContent = confirmLabel;
							showToast("❌ " + ((ares && ares.error) || "İşlem başarısız."), "err");
						}
					})
					.catch(function(err) {
						okBtn.disabled = false;
						if (noBtn) noBtn.disabled = false;
						okBtn.textContent = confirmLabel;
						showToast("❌ Bağlantı hatası: " + String(err), "err");
					});
				};
			}
			if (noBtn) {
				noBtn.onclick = function() {
					wc.innerHTML = '<div style="color:var(--text-muted);font-size:11.5px;">✖ İşlem iptal edildi (veritabanında değişiklik yapılmadı).</div>';
				};
			}
			div.appendChild(wc);
		}
		if (role === "ai" && Array.isArray(downloads) && downloads.length) {
			var dlWrap = document.createElement("div");
			dlWrap.className = "revir-cmsg-downloads";
			downloads.forEach(function(d) {
				if (!d || !d.token) return;
				var dbtn = document.createElement("button");
				dbtn.type = "button";
				dbtn.className = "revir-cmsg-dlbtn";
				dbtn.textContent = d.label || ("📥 " + (d.file_name || "Raporu İndir"));
				dbtn.onclick = function() { window.__revirDownload(d.token, d.file_name || "rapor"); };
				dlWrap.appendChild(dbtn);
			});
			div.appendChild(dlWrap);
		}
		if (role === "ai" && plainText) {
			var tools = document.createElement("div");
			tools.className = "revir-cmsg-tools";
			var docBtn = document.createElement("button");
			docBtn.type = "button";
			docBtn.className = "revir-cmsg-tbtn";
			docBtn.title = "Bu yanıtı resmi Word (.docx) dosyasına aktar";
			docBtn.textContent = "📄 Word";
			docBtn.onclick = function() { window.__revirExportChatMsg(plainText, "docx"); };

			var xlsBtn = document.createElement("button");
			xlsBtn.type = "button";
			xlsBtn.className = "revir-cmsg-tbtn";
			xlsBtn.title = "Bu yanıtı Excel (.xlsx) tablosuna aktar";
			xlsBtn.textContent = "📊 Excel";
			xlsBtn.onclick = function() { window.__revirExportChatMsg(plainText, "xlsx"); };

			var cpBtn = document.createElement("button");
			cpBtn.type = "button";
			cpBtn.className = "revir-cmsg-tbtn";
			cpBtn.textContent = "📋 Kopyala";
			cpBtn.onclick = function() {
				if (navigator.clipboard) navigator.clipboard.writeText(plainText);
				showToast("📋 Yanıt panoya kopyalandı");
			};
			tools.appendChild(docBtn);
			tools.appendChild(xlsBtn);
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
		fetch("/api/gemini_chat_clear", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: "{}"
		}).catch(function(){});
	};

	window.__revirSendChat = function() {
		if (chatBusy) return;
		var inp = document.getElementById("revir-chat-inp");
		var msgTxt = inp ? inp.value.trim() : "";
		var filesToSend = chatFiles.slice();

		if (!msgTxt && !filesToSend.length) return;

		if (inp) inp.value = "";
		chatFiles = [];
		renderChatFileChips();

		var badges = [];
		filesToSend.forEach(function(f) { badges.push("📎 " + escHtml(f.name)); });
		var userHtml = "";
		if (badges.length) {
			userHtml += '<div style="font-size:11.5px;opacity:0.9;margin-bottom:4px;">' + badges.join(" · ") + "</div>";
		}
		if (msgTxt) {
			userHtml += "<p>" + escHtml(msgTxt) + "</p>";
		}
		var chatTitleNow = (document.getElementById("revir-chat-title") || {}).textContent || "Klinik Asistan";
		appendChatMsg("user", userHtml, "");
		var waitBubble = appendChatMsg("ai", '<p style="color:var(--text-muted);">⏳ <em>' + escHtml(chatTitleNow) + ' yanıt hazırlıyor...</em></p>', "");
		chatBusy = true;

		fetch("/api/gemini_chat", {
			method: "POST",
			headers: { "Content-Type": "application/json", "X-Revir-Token": CSRF },
			body: JSON.stringify({
				message: msgTxt,
				attachments: filesToSend
			})
		})
		.then(function(r) { return r.json(); })
		.then(function(res) {
			if (waitBubble) waitBubble.remove();
			if (res.ok) {
				appendChatMsg(
					"ai",
					res.reply_html || ("<p>" + escHtml(res.reply_text || "") + "</p>"),
					res.reply_text || "",
					res.downloads || [],
					res.write_proposal || null
				);
				if (res.refresh_ui) {
					sendAction({ action: "refresh" });
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

	// Tablo & Modal Pencere Kenarından/Köşesinden ve Sütun Çizgisinden Canlı Boyutlandırma
	var TBL_STORE_KEY = "revir_table_sizes_v1";
	var MODAL_STORE_KEY = "revir_modal_sizes_v1";

	function loadResizeStore(k) {
		try { return JSON.parse(localStorage.getItem(k) || "{}") || {}; } catch (e) { return {}; }
	}
	function saveResizeStore(k, obj) {
		try { localStorage.setItem(k, JSON.stringify(obj)); } catch (e) {}
	}
	function showResizeBadge(txt, x, y) {
		var b = document.getElementById("revir-resize-badge");
		if (!b) return;
		b.textContent = txt;
		b.style.display = "block";
		b.style.left = Math.min(window.innerWidth - 130, x + 14) + "px";
		b.style.top = Math.max(8, Math.min(window.innerHeight - 34, y + 14)) + "px";
	}
	function hideResizeBadge() {
		var b = document.getElementById("revir-resize-badge");
		if (b) b.style.display = "none";
	}

	function applySavedModalSize(card, titleStr, isDrawer) {
		if (!card || card.classList.contains("fullscreen")) return;
		var store = loadResizeStore(MODAL_STORE_KEY);
		var mKey = (isDrawer ? "drawer:" : "modal:") + String(titleStr || "");
		var sv = store[mKey];
		if (sv && sv.w) {
			var maxW = Math.round((window.innerWidth || 1200) * (isDrawer ? 0.96 : 0.98));
			card.style.width = Math.max(360, Math.min(maxW, sv.w)) + "px";
			card.style.maxWidth = isDrawer ? "96vw" : "98vw";
		} else {
			card.style.width = "";
			card.style.maxWidth = "";
		}
		if (!isDrawer && sv && sv.h) {
			var maxH = Math.round((window.innerHeight || 800) * 0.96);
			card.style.height = Math.max(240, Math.min(maxH, sv.h)) + "px";
			card.style.maxHeight = "96vh";
		} else {
			card.style.height = "";
			card.style.maxHeight = "";
		}
	}

	function initResizableModal() {
		var card = document.getElementById("revir-modal-card");
		if (!card || card.dataset.modalResizeInit === "1") return;
		card.dataset.modalResizeInit = "1";
		card.querySelectorAll("[data-mres]").forEach(function(handle) {
			var mode = handle.dataset.mres;
			handle.addEventListener("pointerdown", function(e) {
				if (e.button !== 0 || card.classList.contains("fullscreen")) return;
				e.preventDefault();
				e.stopPropagation();
				var isDrawer = card.classList.contains("drawer");
				var rect = card.getBoundingClientRect();
				var startX = e.clientX;
				var startY = e.clientY;
				var startW = rect.width;
				var startH = rect.height;
				var vw = window.innerWidth || 1200;
				var vh = window.innerHeight || 800;
				var store = loadResizeStore(MODAL_STORE_KEY);
				var mKey = (isDrawer ? "drawer:" : "modal:") + String(lastDialogTitle || "");
				card.classList.add("revir-modal-resizing");

				function onMove(mv) {
					var dx = mv.clientX - startX;
					var dy = mv.clientY - startY;
					store[mKey] = store[mKey] || {};
					if (mode === "r" || mode === "l" || mode === "se" || mode === "sw") {
						var newW;
						if (isDrawer) {
							newW = Math.round(Math.max(360, Math.min(vw * 0.96, startW - dx)));
							card.style.maxWidth = "96vw";
						} else {
							var deltaW = (mode === "l" || mode === "sw") ? (-dx * 2) : (dx * 2);
							newW = Math.round(Math.max(360, Math.min(vw * 0.98, startW + deltaW)));
							card.style.maxWidth = "98vw";
						}
						card.style.width = newW + "px";
						store[mKey].w = newW;
					}
					if (!isDrawer && (mode === "b" || mode === "se" || mode === "sw")) {
						var newH = Math.round(Math.max(240, Math.min(vh * 0.96, startH + dy * 2)));
						card.style.height = newH + "px";
						card.style.maxHeight = "96vh";
						store[mKey].h = newH;
					}
					var curR = card.getBoundingClientRect();
					showResizeBadge(Math.round(curR.width) + " × " + Math.round(curR.height) + " px", mv.clientX, mv.clientY);
				}

				function onUp() {
					window.__revirLastResizeEndAt = Date.now();
					card.classList.remove("revir-modal-resizing");
					hideResizeBadge();
					saveResizeStore(MODAL_STORE_KEY, store);
					window.removeEventListener("pointermove", onMove);
					window.removeEventListener("pointerup", onUp);
					window.removeEventListener("pointercancel", onUp);
				}

				window.addEventListener("pointermove", onMove);
				window.addEventListener("pointerup", onUp);
				window.addEventListener("pointercancel", onUp);
			});

			handle.addEventListener("dblclick", function(e) {
				e.preventDefault();
				e.stopPropagation();
				var isDrawer = card.classList.contains("drawer");
				var store = loadResizeStore(MODAL_STORE_KEY);
				var mKey = (isDrawer ? "drawer:" : "modal:") + String(lastDialogTitle || "");
				delete store[mKey];
				saveResizeStore(MODAL_STORE_KEY, store);
				card.style.width = "";
				card.style.maxWidth = "";
				card.style.height = "";
				card.style.maxHeight = "";
			});
		});
	}

	function computeTableStoreId(outer, tbl, seenMap) {
		var scope = outer.closest("#revir-modal-body") ? ("modal:" + (lastDialogTitle || "dlg")) : ("page:" + (lastRenderedPage || "main"));
		var expKey = outer.dataset.tblExplicit === "1" ? (outer.dataset.tblKey || "") : "";
		var part;
		if (expKey) {
			part = "k:" + expKey;
		} else {
			var hdrs = [];
			if (tbl) {
				tbl.querySelectorAll("thead th").forEach(function(th) {
					if (hdrs.length < 8) {
						hdrs.push((th.textContent || "").replace(/[⇅▲▼]/g, "").trim().slice(0, 24));
					}
				});
			}
			part = "h:" + (hdrs.join("|") || (outer.dataset.tblKey || "tbl"));
		}
		var base = scope + "::" + part;
		seenMap[base] = (seenMap[base] || 0) + 1;
		return base + "#" + seenMap[base];
	}

	function initResizableTables() {
		var wraps = document.querySelectorAll("#revir-main .revir-table-wrap, #revir-main .revir-md-table-wrap, #revir-modal-body .revir-table-wrap, #revir-modal-body .revir-md-table-wrap");
		wraps.forEach(function(wrap) {
			if (wrap.closest(".revir-a4-sheet, .revir-a4-sheet-landscape, #revir-chat-messages")) return;
			if (!wrap.parentElement || !wrap.parentElement.classList.contains("revir-table-outer")) {
				var outerNew = document.createElement("div");
				outerNew.className = "revir-table-outer";
				outerNew.dataset.tblExplicit = "0";
				wrap.parentNode.insertBefore(outerNew, wrap);
				outerNew.appendChild(wrap);
			}
		});

		var store = loadResizeStore(TBL_STORE_KEY);
		var seenMap = {};
		document.querySelectorAll(".revir-table-outer").forEach(function(outer) {
			if (outer.closest("#revir-chat-messages, .revir-a4-sheet, .revir-a4-sheet-landscape")) return;
			var wrap = outer.querySelector(":scope > .revir-table-wrap, :scope > .revir-md-table-wrap");
			if (!wrap) return;
			var tbl = wrap.querySelector("table");
			if (!tbl) return;

			if (wrap.dataset.origMaxHeight === undefined) {
				wrap.dataset.origMaxHeight = wrap.style.maxHeight || "";
			}
			var storeId = computeTableStoreId(outer, tbl, seenMap);
			outer.dataset.tblStoreId = storeId;

			var sv = store[storeId];
			if (sv) {
				if (sv.h) {
					wrap.style.height = sv.h + "px";
					wrap.style.maxHeight = sv.h + "px";
				}
				if (sv.ml) {
					outer.style.marginLeft = sv.ml + "px";
				}
				if (sv.w) {
					outer.style.width = sv.w + "px";
					var pw0 = outer.parentElement ? outer.parentElement.clientWidth : 0;
					if (pw0 > 0 && sv.w > pw0 + 2) outer.style.maxWidth = "none";
				}
				var ths0 = tbl.querySelectorAll("thead th");
				if (Array.isArray(sv.cols) && sv.cols.length === ths0.length && sv.cols.every(function(x) { return typeof x === "number" && x >= 36; })) {
					var totalColW = 0;
					tbl.classList.add("revir-table-fixed-cols");
					tbl.style.tableLayout = "fixed";
					ths0.forEach(function(thEl, idx) {
						var cw = Math.round(sv.cols[idx]);
						thEl.style.width = cw + "px";
						thEl.style.minWidth = cw + "px";
						totalColW += cw;
					});
					tbl.style.width = "max(100%, " + totalColW + "px)";
				}
			}

			// Sütun kenar çizgisi boyutlandırıcıları
			tbl.querySelectorAll("thead th").forEach(function(th, cIdx, allThs) {
				if (th.classList.contains("revir-ed-act-th")) return;
				if (th.querySelector(":scope > .revir-col-resizer")) return;
				var colHandle = document.createElement("span");
				colHandle.className = "revir-col-resizer";
				colHandle.title = "Sütun genişliğini ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın";
				colHandle.addEventListener("click", function(e) {
					e.preventDefault();
					e.stopPropagation();
				});
				colHandle.addEventListener("pointerdown", function(e) {
					if (e.button !== 0) return;
					e.preventDefault();
					e.stopPropagation();
					var colWidths = [];
					allThs.forEach(function(h) {
						colWidths.push(Math.max(48, Math.round(h.getBoundingClientRect().width)));
					});
					tbl.classList.add("revir-table-fixed-cols");
					tbl.style.tableLayout = "fixed";
					allThs.forEach(function(h, idx) {
						h.style.width = colWidths[idx] + "px";
						h.style.minWidth = colWidths[idx] + "px";
					});
					var startX = e.clientX;
					var startColW = colWidths[cIdx];
					colHandle.classList.add("active");
					outer.classList.add("revir-tbl-resizing");

					function onColMove(mv) {
						var newCW = Math.round(Math.max(48, Math.min(960, startColW + (mv.clientX - startX))));
						colWidths[cIdx] = newCW;
						th.style.width = newCW + "px";
						th.style.minWidth = newCW + "px";
						var sumW = colWidths.reduce(function(acc, v) { return acc + v; }, 0);
						tbl.style.width = "max(100%, " + sumW + "px)";
						showResizeBadge("Sütun: " + newCW + " px", mv.clientX, mv.clientY);
					}

					function onColUp() {
						tbl.dataset.colResizeAt = String(Date.now());
						window.__revirLastResizeEndAt = Date.now();
						colHandle.classList.remove("active");
						outer.classList.remove("revir-tbl-resizing");
						hideResizeBadge();
						var stNow = loadResizeStore(TBL_STORE_KEY);
						stNow[storeId] = stNow[storeId] || {};
						stNow[storeId].cols = colWidths.slice();
						saveResizeStore(TBL_STORE_KEY, stNow);
						window.removeEventListener("pointermove", onColMove);
						window.removeEventListener("pointerup", onColUp);
						window.removeEventListener("pointercancel", onColUp);
					}

					window.addEventListener("pointermove", onColMove);
					window.addEventListener("pointerup", onColUp);
					window.addEventListener("pointercancel", onColUp);
				});
				colHandle.addEventListener("dblclick", function(e) {
					e.preventDefault();
					e.stopPropagation();
					tbl.dataset.colResizeAt = String(Date.now());
					tbl.classList.remove("revir-table-fixed-cols");
					tbl.style.tableLayout = "";
					tbl.style.width = "";
					allThs.forEach(function(h) {
						h.style.width = "";
						h.style.minWidth = "";
					});
					var stNow = loadResizeStore(TBL_STORE_KEY);
					if (stNow[storeId]) {
						delete stNow[storeId].cols;
						if (!stNow[storeId].w && !stNow[storeId].h && !stNow[storeId].ml) delete stNow[storeId];
						saveResizeStore(TBL_STORE_KEY, stNow);
					}
				});
				th.appendChild(colHandle);
			});

			if (outer.dataset.tblResizeInit === "1") return;
			outer.dataset.tblResizeInit = "1";

			function startOuterResize(e, mode) {
				if (e.button !== 0) return;
				e.preventDefault();
				e.stopPropagation();
				var startX = e.clientX;
				var startY = e.clientY;
				var rect = outer.getBoundingClientRect();
				var wrapRect = wrap.getBoundingClientRect();
				var startW = rect.width;
				var startH = wrapRect.height;
				var startML = parseFloat(getComputedStyle(outer).marginLeft) || 0;
				var parentW = outer.parentElement ? outer.parentElement.getBoundingClientRect().width : startW;
				var maxOuterW = Math.max(parentW, Math.round((window.innerWidth || 1200) * 1.4), 1600);
				var liveState = Object.assign({}, (loadResizeStore(TBL_STORE_KEY)[storeId] || {}));
				outer.classList.add("revir-tbl-resizing");

				function onMove(mv) {
					var dx = mv.clientX - startX;
					var dy = mv.clientY - startY;

					if (mode === "r" || mode === "se") {
						var rawW = Math.round(Math.max(220, Math.min(maxOuterW - startML, startW + dx)));
						if (startML === 0 && Math.abs(rawW - parentW) <= 6) {
							outer.style.width = "100%";
							outer.style.maxWidth = "100%";
							delete liveState.w;
						} else {
							outer.style.width = rawW + "px";
							outer.style.maxWidth = rawW > parentW + 2 ? "none" : "100%";
							liveState.w = rawW;
						}
					} else if (mode === "l" || mode === "sw") {
						var newML = Math.round(Math.max(0, Math.min(startML + startW - 220, startML + dx)));
						var newWL = Math.round(Math.max(220, startW - (newML - startML)));
						if (newML === 0 && Math.abs(newWL - parentW) <= 6) {
							outer.style.marginLeft = "";
							outer.style.width = "100%";
							outer.style.maxWidth = "100%";
							delete liveState.ml;
							delete liveState.w;
						} else {
							outer.style.marginLeft = newML ? (newML + "px") : "";
							outer.style.width = newWL + "px";
							outer.style.maxWidth = (newML + newWL) > parentW + 2 ? "none" : "100%";
							if (newML) liveState.ml = newML;
							else delete liveState.ml;
							liveState.w = newWL;
						}
					}

					if (mode === "b" || mode === "se" || mode === "sw") {
						var newH = Math.round(Math.max(96, Math.min(2200, startH + dy)));
						wrap.style.height = newH + "px";
						wrap.style.maxHeight = newH + "px";
						liveState.h = newH;
					}

					var curW = Math.round(outer.getBoundingClientRect().width);
					var curH = Math.round(wrap.getBoundingClientRect().height);
					showResizeBadge(curW + " × " + curH + " px", mv.clientX, mv.clientY);
				}

				function onUp() {
					window.__revirLastResizeEndAt = Date.now();
					outer.classList.remove("revir-tbl-resizing");
					hideResizeBadge();
					var stFinal = loadResizeStore(TBL_STORE_KEY);
					if (liveState.w || liveState.h || liveState.ml || liveState.cols) {
						stFinal[storeId] = liveState;
					} else {
						delete stFinal[storeId];
					}
					store = stFinal;
					saveResizeStore(TBL_STORE_KEY, stFinal);
					window.removeEventListener("pointermove", onMove);
					window.removeEventListener("pointerup", onUp);
					window.removeEventListener("pointercancel", onUp);
				}

				window.addEventListener("pointermove", onMove);
				window.addEventListener("pointerup", onUp);
				window.addEventListener("pointercancel", onUp);
			}

			function resetOuter(resetW, resetH, resetCols) {
				var stNow = loadResizeStore(TBL_STORE_KEY);
				stNow[storeId] = stNow[storeId] || {};
				if (resetW) {
					outer.style.width = "";
					outer.style.maxWidth = "";
					outer.style.marginLeft = "";
					delete stNow[storeId].w;
					delete stNow[storeId].ml;
				}
				if (resetH) {
					wrap.style.height = "";
					wrap.style.maxHeight = wrap.dataset.origMaxHeight || "";
					delete stNow[storeId].h;
				}
				if (resetCols) {
					tbl.classList.remove("revir-table-fixed-cols");
					tbl.style.tableLayout = "";
					tbl.style.width = "";
					tbl.querySelectorAll("thead th").forEach(function(h) {
						h.style.width = "";
						h.style.minWidth = "";
					});
					delete stNow[storeId].cols;
				}
				if (!stNow[storeId].w && !stNow[storeId].h && !stNow[storeId].ml && !stNow[storeId].cols) {
					delete stNow[storeId];
				}
				store = stNow;
				saveResizeStore(TBL_STORE_KEY, stNow);
			}

			[
				["l", "revir-tbl-handle-l", "Tablo sol kenarını sürükleyerek genişliği ayarlayın · Sıfırlamak için çift tıklayın", true, false, false],
				["r", "revir-tbl-handle-r", "Tablo sağ kenarını sürükleyerek genişliği ayarlayın · Sıfırlamak için çift tıklayın", true, false, false],
				["b", "revir-tbl-handle-b", "Tablo alt kenarını sürükleyerek yüksekliği ayarlayın · Sıfırlamak için çift tıklayın", false, true, false],
				["sw", "revir-tbl-handle-sw", "Köşeden sürükleyerek boyutlandırın · Sıfırlamak için çift tıklayın", true, true, true],
				["se", "revir-tbl-handle-se", "Köşeden sürükleyerek tablo boyutunu ayarlayın · Sıfırlamak için çift tıklayın", true, true, true]
			].forEach(function(cfg) {
				var hEl = document.createElement("div");
				hEl.className = cfg[1];
				hEl.title = cfg[2];
				hEl.addEventListener("pointerdown", function(e) { startOuterResize(e, cfg[0]); });
				hEl.addEventListener("dblclick", function(e) {
					e.preventDefault();
					e.stopPropagation();
					resetOuter(cfg[3], cfg[4], cfg[5]);
				});
				outer.appendChild(hEl);
			});
		});
	}

	window.__revirInitResizableTables = initResizableTables;
	initResizableModal();

	// Yüzen AI Sohbet Penceresi Sol Üst Köşeden Boyutlandırma
	var CHAT_STORE_KEY = "revir_chat_size_v1";
	function applySavedChatSize() {
		var b = document.getElementById("revir-chat-bubble");
		if (!b || b.classList.contains("expanded")) return;
		var sv = loadResizeStore(CHAT_STORE_KEY);
		if (sv && sv.w) b.style.width = Math.max(320, Math.min((window.innerWidth || 1200) - 36, sv.w)) + "px";
		if (sv && sv.h) b.style.height = Math.max(360, Math.min((window.innerHeight || 800) - 110, sv.h)) + "px";
	}
	(function initChatResize() {
		var b = document.getElementById("revir-chat-bubble");
		var h = document.getElementById("revir-chat-resize-nw");
		if (!b || !h || h.dataset.init === "1") return;
		h.dataset.init = "1";
		applySavedChatSize();
		h.addEventListener("pointerdown", function(e) {
			if (e.button !== 0) return;
			e.preventDefault();
			e.stopPropagation();
			b.classList.remove("expanded");
			b.classList.add("revir-chat-resizing");
			var rect = b.getBoundingClientRect();
			var startX = e.clientX, startY = e.clientY, startW = rect.width, startH = rect.height;
			var vw = window.innerWidth || 1200, vh = window.innerHeight || 800;
			function onMove(mv) {
				var newW = Math.round(Math.max(320, Math.min(vw - 36, startW - (mv.clientX - startX))));
				var newH = Math.round(Math.max(360, Math.min(vh - 110, startH - (mv.clientY - startY))));
				b.style.width = newW + "px";
				b.style.height = newH + "px";
				showResizeBadge(newW + " × " + newH + " px", mv.clientX, mv.clientY);
			}
			function onUp() {
				b.classList.remove("revir-chat-resizing");
				hideResizeBadge();
				var rNow = b.getBoundingClientRect();
				saveResizeStore(CHAT_STORE_KEY, { w: Math.round(rNow.width), h: Math.round(rNow.height) });
				window.removeEventListener("pointermove", onMove);
				window.removeEventListener("pointerup", onUp);
				window.removeEventListener("pointercancel", onUp);
			}
			window.addEventListener("pointermove", onMove);
			window.addEventListener("pointerup", onUp);
			window.addEventListener("pointercancel", onUp);
		});
		h.addEventListener("dblclick", function(e) {
			e.preventDefault();
			e.stopPropagation();
			try { localStorage.removeItem(CHAT_STORE_KEY); } catch (_e) {}
			b.style.width = "";
			b.style.height = "";
		});
	})();

	// Sayfa İçi 12-Sütunlu Sanal Izgara Yerleşim, İtme, Takas & Boyutlandırma Motoru (TAKVİM HARİÇ)
	var BOX_STORE_KEY = "revir_box_layout_v1";
	var GRID_COLS = 12;
	function loadBoxStore() {
		var obj = loadResizeStore(BOX_STORE_KEY);
		if (!obj || typeof obj !== "object") obj = {};
		if (!obj.pages || typeof obj.pages !== "object") obj.pages = {};
		return obj;
	}
	function saveBoxStore(obj) {
		saveResizeStore(BOX_STORE_KEY, obj);
	}
	function getPageBoxData(store, pg) {
		store.pages[pg] = store.pages[pg] || {};
		store.pages[pg].orders = store.pages[pg].orders || {};
		store.pages[pg].sizes = store.pages[pg].sizes || {};
		return store.pages[pg];
	}
	function boxTextFingerprint(el) {
		var hdr = el.querySelector("h1, h2, h3, h4, h5, h6, summary, .revir-metric-label, .revir-caption, strong");
		var raw = (hdr ? hdr.textContent : el.textContent) || "kutu";
		return raw.replace(/\s+/g, " ").trim().slice(0, 40) || "kutu";
	}
	function inferMinGridCols(el) {
		if (!el) return 4;
		if (el.querySelector(".revir-table-outer, table, .revir-chart-card, .revir-cols-5, .revir-cols-6")) return 6;
		if (el.classList.contains("revir-cols")) return 6;
		return 4;
	}
	function isEligibleBoxNode(el) {
		if (!el || el.nodeType !== 1) return false;
		if (el.closest(".cal-leaf, .cal-grid, [data-no-box-layout='1'], [data-layout-canvas], .revir-a4-sheet, .revir-a4-sheet-landscape, #revir-modal-body, #revir-sidebar, #revir-chat-bubble")) return false;
		if (el.hasAttribute("data-layout-widget") || el.hasAttribute("data-layout-canvas")) return false;
		if (el.classList.contains("revir-card") || el.classList.contains("revir-chart-card") || el.classList.contains("revir-metric")) return true;
		if (el.tagName === "DETAILS" && el.classList.contains("revir-expander")) return true;
		if (el.classList.contains("revir-cols") && !el.querySelector(":scope > .revir-sticky-col")) {
			if (el.querySelector(".revir-card, .revir-metric, .revir-chart-card, details.revir-expander")) return true;
		}
		return false;
	}
	function isEligibleGridCol(colEl) {
		if (!colEl || !colEl.classList || !colEl.classList.contains("revir-col")) return false;
		if (colEl.classList.contains("revir-sticky-col")) return false;
		if (colEl.closest("[data-layout-canvas], [data-layout-widget]")) return false;
		return !!colEl.querySelector(".revir-card, .revir-metric, .revir-chart-card, details.revir-expander, .revir-btn");
	}
	function clearBoxDrops() {
		document.querySelectorAll(".revir-box-drop-before, .revir-box-drop-after, .revir-box-swap-target").forEach(function(n) {
			n.classList.remove("revir-box-drop-before", "revir-box-drop-after", "revir-box-swap-target");
		});
	}
	function syncLayoutBtnState() {
		var btn = document.getElementById("revir-tb-layout-btn");
		if (!btn) return;
		var isCal = isCalendarPageName(lastRenderedPage);
		btn.disabled = (isCal || !boxLayoutEnabled);
		btn.classList.toggle("active", document.body.classList.contains("revir-layout-edit"));
		if (isCal) btn.title = "Takvim sayfası kutucuk düzenlemesinden hariçtir";
		else if (!boxLayoutEnabled) btn.title = "Kutucuk düzeni Ayarlar bölümünden kapalı";
		else btn.title = "Sayfa içindeki kutucukların 12 sütunlu ızgarada yerini ve boyutunu değiştir (Takvim hariç)";
	}
	window.__revirSyncLayoutBtn = syncLayoutBtnState;

	window.__revirToggleLayoutEdit = function(forceVal) {
		if (isCalendarPageName(lastRenderedPage)) {
			document.body.classList.remove("revir-layout-edit");
			syncLayoutBtnState();
			showToast("ℹ️ Takvim sayfası kutucuk düzenlemesinden hariçtir.", "warn");
			return;
		}
		if (!boxLayoutEnabled) {
			document.body.classList.remove("revir-layout-edit");
			syncLayoutBtnState();
			showToast("ℹ️ Kutucuk düzenleme Ayarlar → Görünüm & Düzen bölümünden etkinleştirilebilir.", "warn");
			return;
		}
		var nxt = (forceVal !== undefined) ? !!forceVal : !document.body.classList.contains("revir-layout-edit");
		document.body.classList.toggle("revir-layout-edit", nxt);
		if (String(lastRenderedPage || "").includes("Genel Bakış")) {
			var overviewEditState = document.querySelector('input[data-key="layout_overview_edit"]');
			if (overviewEditState) {
				var targetValue = nxt ? "true" : "false";
				if (overviewEditState.value !== targetValue) {
					overviewEditState.value = targetValue;
					overviewEditState.dispatchEvent(new Event("change", { bubbles: true }));
				}
			}
		}
		syncLayoutBtnState();
		if (nxt) {
			showToast("🧩 12 sütunlu akıllı düzen açıldı — merkeze sürükleyince takas, araya bırakınca itme yapar");
		}
	};
	document.addEventListener("change", function(e) {
		var t = e && e.target;
		if (!t || t.dataset.key !== "layout_overview_edit") return;
		document.body.classList.toggle("revir-layout-edit", String(t.value).toLowerCase() === "true");
		syncLayoutBtnState();
	});

	window.__revirResetPageBoxLayout = function(pageName) {
		var pg = String(pageName || lastRenderedPage || "");
		if (!pg || isCalendarPageName(pg)) return;
		var store = loadBoxStore();
		if (store.pages && store.pages[pg]) {
			delete store.pages[pg];
			saveBoxStore(store);
		}
		if (typeof window.__revirResetSmartCanvas === "function") {
			window.__revirResetSmartCanvas();
		}
		showToast("↺ Bu sayfanın kutucuk yerleşimi ve boyutları sıfırlandı");
		sendAction({ action: "refresh" });
	};

	window.__revirResetAllBoxLayouts = function() {
		try {
			localStorage.removeItem(BOX_STORE_KEY);
			Object.keys(localStorage).forEach(function(k) {
				if (k.indexOf("revir_smart_grid_v2:") === 0) localStorage.removeItem(k);
			});
		} catch (_e) {}
		document.body.classList.remove("revir-layout-edit");
		syncLayoutBtnState();
		showToast("↺ Tüm sayfa kutucuk yerleşimleri ve boyutları sıfırlandı");
		sendAction({ action: "refresh" });
	};

	window.__revirResetTableModalSizes = function() {
		try {
			localStorage.removeItem(TBL_STORE_KEY);
			localStorage.removeItem(MODAL_STORE_KEY);
			localStorage.removeItem(CHAT_STORE_KEY);
		} catch (_e) {}
		showToast("↺ Tablo, pencere ve sohbet balonu boyutları sıfırlandı");
		sendAction({ action: "refresh" });
	};

	function persistZoneOrder(pg, zoneId, hostEls) {
		var store = loadBoxStore();
		var pData = getPageBoxData(store, pg);
		pData.orders[zoneId] = hostEls.map(function(h) { return h.dataset.boxId; }).filter(Boolean);
		saveBoxStore(store);
	}

	function attachBoxDecorators(hostEl, zoneEl, zoneId, pg, isHoriz) {
		var boxId = hostEl.dataset.boxId;
		if (!boxId) return;
		hostEl.classList.add("revir-box-host");
		var minCol = inferMinGridCols(hostEl);
		var isGridZone = !isHoriz && zoneEl.classList.contains("revir-grid-zone");
		var store = loadBoxStore();
		var pData = getPageBoxData(store, pg);
		var sv = pData.sizes[boxId];

		function applyBoxSpanAndSize(sizeObj) {
			if (isGridZone) {
				var cs = (sizeObj && sizeObj.colSpan) ? Math.max(minCol, Math.min(GRID_COLS, Number(sizeObj.colSpan))) : GRID_COLS;
				hostEl.style.gridColumn = "span " + cs;
				hostEl.style.width = "";
				hostEl.style.maxWidth = "";
				if (spanBadge) spanBadge.textContent = cs + "/" + GRID_COLS;
			} else if (sizeObj && sizeObj.w) {
				hostEl.style.width = sizeObj.w + "px";
				hostEl.style.maxWidth = "none";
			}
			if (sizeObj && sizeObj.h) {
				hostEl.style.height = sizeObj.h + "px";
				hostEl.style.overflow = "auto";
			}
		}

		if (hostEl.dataset.boxDecorated === "1") {
			applyBoxSpanAndSize(sv);
			return;
		}
		hostEl.dataset.boxDecorated = "1";

		var swapPill = document.createElement("div");
		swapPill.className = "revir-box-swap-pill";
		swapPill.textContent = "🔁 Yer Değiştir";
		hostEl.appendChild(swapPill);

		// Araç çubuğu (Düzenleme modunda görünür)
		var tb = document.createElement("div");
		tb.className = "revir-box-toolbar";
		tb.addEventListener("click", function(e) { e.stopPropagation(); });

		var spanBadge = document.createElement("span");
		spanBadge.className = "revir-box-span-tag";
		if (isGridZone) {
			var curCs = (sv && sv.colSpan) ? Math.max(minCol, Math.min(GRID_COLS, Number(sv.colSpan))) : GRID_COLS;
			spanBadge.textContent = curCs + "/" + GRID_COLS;
		}

		var dragBtn = document.createElement("button");
		dragBtn.type = "button";
		dragBtn.className = "revir-box-tbtn revir-box-drag";
		dragBtn.title = "Sürükle: araya bırakınca iter, kutu ortasına bırakınca yer değiştirir";
		dragBtn.textContent = "⋮⋮";
		dragBtn.addEventListener("pointerdown", function(e) {
			if (e.button !== 0 || !document.body.classList.contains("revir-layout-edit")) return;
			e.preventDefault();
			e.stopPropagation();
			hostEl.classList.add("revir-box-dragging");
			var activeTarget = null;
			var dropMode = "after";

			function zoneSiblings() {
				return Array.prototype.slice.call(zoneEl.children).filter(function(ch) {
					return ch !== hostEl && ch.dataset && ch.dataset.boxZone === zoneId;
				});
			}

			function onDragMove(mv) {
				clearBoxDrops();
				activeTarget = null;
				var sibs = zoneSiblings();
				for (var i = 0; i < sibs.length; i++) {
					var cand = sibs[i];
					var r = cand.getBoundingClientRect();
					if (mv.clientX >= r.left && mv.clientX <= r.right && mv.clientY >= r.top && mv.clientY <= r.bottom) {
						activeTarget = cand;
						var inCenterX = mv.clientX >= r.left + r.width * 0.25 && mv.clientX <= r.right - r.width * 0.25;
						var inCenterY = mv.clientY >= r.top + r.height * 0.22 && mv.clientY <= r.bottom - r.height * 0.22;
						if (inCenterX && inCenterY) {
							dropMode = "swap";
							cand.classList.add("revir-box-swap-target");
						} else {
							var before = isGridZone ? (mv.clientX < r.left + r.width / 2 && mv.clientY < r.bottom - r.height * 0.25 || mv.clientY < r.top + r.height * 0.28) : (isHoriz ? (mv.clientX < r.left + r.width / 2) : (mv.clientY < r.top + r.height / 2));
							dropMode = before ? "before" : "after";
							cand.classList.add(before ? "revir-box-drop-before" : "revir-box-drop-after");
						}
						break;
					}
				}
			}

			function onDragUp() {
				window.removeEventListener("pointermove", onDragMove);
				window.removeEventListener("pointerup", onDragUp);
				window.removeEventListener("pointercancel", onDragUp);
				hostEl.classList.remove("revir-box-dragging");
				clearBoxDrops();
				if (activeTarget && activeTarget.parentNode === zoneEl) {
					if (dropMode === "swap") {
						var placeholder = document.createElement("span");
						zoneEl.insertBefore(placeholder, hostEl);
						zoneEl.insertBefore(hostEl, activeTarget);
						zoneEl.insertBefore(activeTarget, placeholder);
						placeholder.parentNode.removeChild(placeholder);
					} else if (dropMode === "before") {
						zoneEl.insertBefore(hostEl, activeTarget);
					} else {
						zoneEl.insertBefore(hostEl, activeTarget.nextSibling);
					}
					var newOrder = Array.prototype.slice.call(zoneEl.children).filter(function(ch) { return ch.dataset && ch.dataset.boxZone === zoneId; });
					persistZoneOrder(pg, zoneId, newOrder);
				}
			}

			window.addEventListener("pointermove", onDragMove);
			window.addEventListener("pointerup", onDragUp);
			window.addEventListener("pointercancel", onDragUp);
		});

		function moveSibling(dir) {
			var sibs = Array.prototype.slice.call(zoneEl.children).filter(function(ch) { return ch.dataset && ch.dataset.boxZone === zoneId; });
			var idx = sibs.indexOf(hostEl);
			if (idx === -1) return;
			var nxtIdx = idx + dir;
			if (nxtIdx < 0 || nxtIdx >= sibs.length) return;
			var targetSib = sibs[nxtIdx];
			if (dir < 0) zoneEl.insertBefore(hostEl, targetSib);
			else zoneEl.insertBefore(hostEl, targetSib.nextSibling);
			var newOrder = Array.prototype.slice.call(zoneEl.children).filter(function(ch) { return ch.dataset && ch.dataset.boxZone === zoneId; });
			persistZoneOrder(pg, zoneId, newOrder);
		}

		var prevBtn = document.createElement("button");
		prevBtn.type = "button";
		prevBtn.className = "revir-box-tbtn revir-box-mv-prev";
		prevBtn.title = isHoriz ? "Sola Taşı" : "Önceye / Yukarı Taşı";
		prevBtn.textContent = isHoriz ? "◀" : "▲";
		prevBtn.addEventListener("click", function(e) { e.preventDefault(); e.stopPropagation(); moveSibling(-1); });

		var nextBtn = document.createElement("button");
		nextBtn.type = "button";
		nextBtn.className = "revir-box-tbtn revir-box-mv-next";
		nextBtn.title = isHoriz ? "Sağa Taşı" : "Sonraya / Aşağı Taşı";
		nextBtn.textContent = isHoriz ? "▶" : "▼";
		nextBtn.addEventListener("click", function(e) { e.preventDefault(); e.stopPropagation(); moveSibling(1); });

		var rstBtn = document.createElement("button");
		rstBtn.type = "button";
		rstBtn.className = "revir-box-tbtn revir-box-reset";
		rstBtn.title = "Bu kutucuğun boyutunu sıfırla";
		rstBtn.textContent = "↺";
		rstBtn.addEventListener("click", function(e) {
			e.preventDefault();
			e.stopPropagation();
			hostEl.style.width = "";
			hostEl.style.maxWidth = "";
			hostEl.style.height = "";
			hostEl.style.overflow = "";
			if (isGridZone) {
				hostEl.style.gridColumn = "span " + GRID_COLS;
				spanBadge.textContent = GRID_COLS + "/" + GRID_COLS;
			}
			var stNow = loadBoxStore();
			var pd = getPageBoxData(stNow, pg);
			delete pd.sizes[boxId];
			saveBoxStore(stNow);
			showToast("↺ Kutucuk boyutu sıfırlandı");
		});

		tb.appendChild(dragBtn);
		if (isGridZone) tb.appendChild(spanBadge);
		tb.appendChild(prevBtn);
		tb.appendChild(nextBtn);
		tb.appendChild(rstBtn);
		hostEl.appendChild(tb);
		applyBoxSpanAndSize(sv);

		// Kenar & Köşe 12-Sütuna Kilitli Boyutlandırma Tutamaçları
		[
			["r", "revir-box-res-r", "Genişliği 12 sütunlu ızgarada ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın"],
			["b", "revir-box-res-b", "Yüksekliği ayarlamak için sürükleyin · Sıfırlamak için çift tıklayın"],
			["se", "revir-box-res-se", "Kutucuğu köşeden sürükleyerek boyutlandırın · Sıfırlamak için çift tıklayın"]
		].forEach(function(hcfg) {
			var hEl = document.createElement("div");
			hEl.className = hcfg[1];
			hEl.title = hcfg[2];
			hEl.addEventListener("click", function(e) { e.stopPropagation(); });
			hEl.addEventListener("pointerdown", function(e) {
				if (e.button !== 0 || !boxLayoutEnabled || isCalendarPageName(lastRenderedPage)) return;
				e.preventDefault();
				e.stopPropagation();
				var rect = hostEl.getBoundingClientRect();
				var zRect = zoneEl.getBoundingClientRect();
				var startX = e.clientX, startY = e.clientY, startW = rect.width, startH = rect.height;
				var live = Object.assign({}, (getPageBoxData(loadBoxStore(), pg).sizes[boxId] || {}));
				hostEl.classList.add("revir-box-resizing");
				function onMove(mv) {
					var badgeTxt = "";
					if (hcfg[0] === "r" || hcfg[0] === "se") {
						if (isGridZone && zRect.width > 0) {
							var tgtW = Math.max(160, startW + (mv.clientX - startX));
							var colUnit = (zRect.width + 12) / GRID_COLS;
							var nSpan = Math.max(minCol, Math.min(GRID_COLS, Math.round((tgtW + 12) / colUnit)));
							hostEl.style.gridColumn = "span " + nSpan;
							hostEl.style.width = "";
							hostEl.style.maxWidth = "";
							live.colSpan = nSpan;
							delete live.w;
							if (spanBadge) spanBadge.textContent = nSpan + "/" + GRID_COLS;
							badgeTxt = nSpan + "/" + GRID_COLS + " sütun (%" + Math.round((nSpan / GRID_COLS) * 100) + ")";
						} else {
							var nw = Math.round(Math.max(180, Math.min(1800, startW + (mv.clientX - startX))));
							hostEl.style.width = nw + "px";
							hostEl.style.maxWidth = "none";
							live.w = nw;
							badgeTxt = nw + " px";
						}
					}
					if (hcfg[0] === "b" || hcfg[0] === "se") {
						var nh = Math.round(Math.max(72, Math.min(1600, startH + (mv.clientY - startY))));
						hostEl.style.height = nh + "px";
						hostEl.style.overflow = "auto";
						live.h = nh;
					}
					var cr = hostEl.getBoundingClientRect();
					var txt = badgeTxt ? (badgeTxt + " · " + Math.round(cr.height) + " px") : (Math.round(cr.width) + " × " + Math.round(cr.height) + " px");
					showResizeBadge(txt, mv.clientX, mv.clientY);
				}
				function onUp() {
					hostEl.classList.remove("revir-box-resizing");
					hideResizeBadge();
					var stNow = loadBoxStore();
					var pd = getPageBoxData(stNow, pg);
					if (live.colSpan || live.w || live.h) pd.sizes[boxId] = live;
					else delete pd.sizes[boxId];
					saveBoxStore(stNow);
					window.removeEventListener("pointermove", onMove);
					window.removeEventListener("pointerup", onUp);
					window.removeEventListener("pointercancel", onUp);
				}
				window.addEventListener("pointermove", onMove);
				window.addEventListener("pointerup", onUp);
				window.addEventListener("pointercancel", onUp);
			});
			hEl.addEventListener("dblclick", function(e) {
				e.preventDefault();
				e.stopPropagation();
				var stNow = loadBoxStore();
				var pd = getPageBoxData(stNow, pg);
				pd.sizes[boxId] = pd.sizes[boxId] || {};
				if (hcfg[0] === "r" || hcfg[0] === "se") {
					hostEl.style.width = "";
					hostEl.style.maxWidth = "";
					if (isGridZone) {
						hostEl.style.gridColumn = "span " + GRID_COLS;
						if (spanBadge) spanBadge.textContent = GRID_COLS + "/" + GRID_COLS;
					}
					delete pd.sizes[boxId].colSpan;
					delete pd.sizes[boxId].w;
				}
				if (hcfg[0] === "b" || hcfg[0] === "se") {
					hostEl.style.height = "";
					hostEl.style.overflow = "";
					delete pd.sizes[boxId].h;
				}
				if (!pd.sizes[boxId].colSpan && !pd.sizes[boxId].w && !pd.sizes[boxId].h) delete pd.sizes[boxId];
				saveBoxStore(stNow);
			});
			hostEl.appendChild(hEl);
		});
	}

	function initBoxLayout() {
		var mainInner = document.getElementById("revir-main-inner");
		if (!mainInner) return;
		var pg = String(lastRenderedPage || "main");
		syncLayoutBtnState();
		if (isCalendarPageName(pg) || !boxLayoutEnabled) return;

		var store = loadBoxStore();
		var pData = getPageBoxData(store, pg);

		function setupZone(zoneEl, zoneId, items, isHoriz) {
			if (!zoneEl || items.length === 0) return;
			if (!isHoriz) zoneEl.classList.add("revir-grid-zone");
			var seen = {};
			items.forEach(function(item) {
				var fp = boxTextFingerprint(item);
				seen[fp] = (seen[fp] || 0) + 1;
				item.dataset.boxId = zoneId + "::" + fp + "#" + seen[fp];
				item.dataset.boxZone = zoneId;
			});
			var savedOrd = pData.orders[zoneId];
			if (Array.isArray(savedOrd) && savedOrd.length > 1 && items.length > 1) {
				var byId = {};
				items.forEach(function(it) { byId[it.dataset.boxId] = it; });
				var ordered = [];
				savedOrd.forEach(function(bid) { if (byId[bid]) { ordered.push(byId[bid]); delete byId[bid]; } });
				items.forEach(function(it) { if (byId[it.dataset.boxId]) ordered.push(it); });
				if (isHoriz) {
					ordered.forEach(function(it) { zoneEl.appendChild(it); });
				} else {
					var anchor = items[items.length - 1].nextSibling;
					ordered.forEach(function(it) { zoneEl.insertBefore(it, anchor); });
				}
			}
			items.forEach(function(item) {
				attachBoxDecorators(item, zoneEl, zoneId, pg, isHoriz);
			});
		}

		// 1. 12-Sütunlu Dikey/Izgara Kutucuk Bölgeleri (özel tuval dışındaki tüm bölümler)
		var vZones = [mainInner].concat(
			Array.prototype.slice.call(mainInner.querySelectorAll(".revir-tab-pane, .revir-col"))
		);
		vZones.forEach(function(zEl, zIdx) {
			if (zEl.closest(".cal-leaf, .cal-grid, [data-no-box-layout='1'], [data-layout-canvas], #revir-modal-body")) return;
			var ch = Array.prototype.slice.call(zEl.children).filter(isEligibleBoxNode);
			if (ch.length >= 1) setupZone(zEl, "vz:" + zIdx, ch, false);
		});

		// 2. Yatay sütun kutucuk bölgeleri (yan yana duran 2-4 kart/metrik sütunları)
		var hZones = Array.prototype.slice.call(mainInner.querySelectorAll(".revir-cols"));
		hZones.forEach(function(colsEl, hIdx) {
			if (colsEl.closest(".cal-leaf, .cal-grid, [data-no-box-layout='1'], [data-layout-canvas], #revir-modal-body")) return;
			if (colsEl.querySelector(":scope > .revir-sticky-col")) return;
			var cols = Array.prototype.slice.call(colsEl.children).filter(isEligibleGridCol);
			if (cols.length >= 2) setupZone(colsEl, "hz:" + hIdx, cols, true);
		});
	}
	window.__revirInitBoxLayout = initBoxLayout;

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
# 10. YEREL MİKROFON KAYDEDİCİSİ & SESLİ OKUMA (TTS) VE HTTP SUNUCUSU
# ==============================================================================

class _NativeAudioController:
	"""Otomatik sessizlik algılamalı (VAD) yerel Python mikrofon kaydedicisi ve anında durdurulabilir Türkçe sesli okuma."""
	def __init__(self) -> None:
		self._ctrl_lock = threading.Lock()
		self._buf_lock = threading.Lock()
		self._tts_lock = threading.Lock()
		self._stream: Any = None
		self._active: bool = False
		self._chunks: list[Any] = []
		self._samplerate: int = 44100
		self._started_at: float = 0.0
		self._noise_samples: list[float] = []
		self._noise_floor: float = 200.0
		self._speech_detected: bool = False
		self._speech_duration: float = 0.0
		self._silence_after_speech: float = 0.0
		self._current_level: int = 0
		self._auto_stopped: bool = False
		self._stream_error: str = ""
		self._completed_wav: bytes | None = None
		self._completed_err: str = ""
		self._tts_procs: list[Any] = []

	def start_recording(self) -> tuple[bool, str]:
		self.stop_speaking()
		with self._ctrl_lock:
			self._close_stream_unlocked()
			try:
				import sounddevice as sd
				import numpy as np

				dev_info = sd.query_devices(kind="input")
				dev_name = str(dev_info.get("name") or "Mikrofon") if isinstance(dev_info, dict) else "Mikrofon"
				native_sr = int(round(float(dev_info.get("default_samplerate") or 44100))) if isinstance(dev_info, dict) else 44100
				if native_sr < 8000 or native_sr > 96000:
					native_sr = 44100

				with self._buf_lock:
					self._chunks = []
					self._samplerate = native_sr
					self._started_at = time.monotonic()
					self._noise_samples = []
					self._noise_floor = 180.0
					self._speech_detected = False
					self._speech_duration = 0.0
					self._silence_after_speech = 0.0
					self._current_level = 0
					self._auto_stopped = False
					self._stream_error = ""
					self._completed_wav = None
					self._completed_err = ""
					self._active = True

				def _cb(indata, frames, _time_info, _status):
					if not self._active:
						return
					try:
						chunk_copy = indata.copy()
						rms = float(np.sqrt(np.mean(chunk_copy.astype(np.float32) ** 2))) if frames > 0 else 0.0
						dt = float(frames) / float(self._samplerate or 44100)
						lvl = int(min(100, max(0, (rms / 2400.0) * 100.0)))
						with self._buf_lock:
							if not self._active:
								return
							self._chunks.append(chunk_copy)
							self._current_level = lvl
							elapsed = time.monotonic() - self._started_at
							if elapsed <= 0.40:
								self._noise_samples.append(rms)
								if self._noise_samples:
									self._noise_floor = max(120.0, float(np.median(self._noise_samples)))
							else:
								thresh = max(self._noise_floor * 2.35, 320.0)
								if rms >= thresh:
									self._speech_duration += dt
									self._silence_after_speech = 0.0
									if self._speech_duration >= 0.18:
										self._speech_detected = True
								else:
									if self._speech_detected:
										self._silence_after_speech += dt
										if self._silence_after_speech >= 1.85 and self._speech_duration >= 0.35:
											self._auto_stopped = True
											self._active = False
							if elapsed >= 45.0:
								self._auto_stopped = True
								self._active = False
					except Exception as cb_exc:
						with self._buf_lock:
							self._stream_error = str(cb_exc)
							self._active = False

				try:
					self._stream = sd.InputStream(
						samplerate=native_sr,
						channels=1,
						dtype="int16",
						blocksize=1024,
						callback=_cb,
					)
					self._stream.start()
				except Exception:
					# Fallback to 16kHz default blocksize
					self._samplerate = 16000
					self._stream = sd.InputStream(samplerate=16000, channels=1, dtype="int16", callback=_cb)
					self._stream.start()
				return True, dev_name
			except Exception as exc:
				self._active = False
				self._stream = None
				return False, str(exc)

	def _close_stream_unlocked(self) -> None:
		self._active = False
		stream = self._stream
		self._stream = None
		if stream is not None:
			try:
				stream.stop()
			except Exception:
				pass
			try:
				stream.close()
			except Exception:
				pass

	def _encode_chunks_to_wav(self, chunks: list[Any], src_sr: int) -> tuple[bool, bytes, str]:
		if not chunks:
			return False, b"", "Ses verisi yakalanamadı. Mikrofon bağlantınızı ve izinlerinizi kontrol edin."
		try:
			import numpy as np
			import wave

			pcm = np.concatenate(chunks, axis=0).reshape(-1)
			if len(pcm) < int(src_sr * 0.25):
				return False, b"", "Ses kaydı çok kısa (en az yarım saniye konuşun)."

			max_abs = int(np.max(np.abs(pcm)))
			rms_all = float(np.sqrt(np.mean(pcm.astype(np.float32) ** 2)))
			if max_abs < 60 and rms_all < 18.0:
				return False, b"", "Mikrofondan ses alınamadı (tamamen sessiz). Sistem Ayarları > Gizlilik > Mikrofon iznini ve mikrofon sesini kontrol edin."

			# 16 kHz mono int16'ya temiz örnekleme indirgeme (küçük boyut + yüksek tanıma doğruluğu)
			target_sr = 16000
			if src_sr != target_sr and len(pcm) > 4:
				new_len = max(1, int(round(len(pcm) * target_sr / float(src_sr))))
				x_old = np.linspace(0.0, 1.0, len(pcm), endpoint=False)
				x_new = np.linspace(0.0, 1.0, new_len, endpoint=False)
				pcm_out = np.interp(x_new, x_old, pcm.astype(np.float32)).astype(np.int16)
			else:
				pcm_out = pcm.astype(np.int16)

			buf = io.BytesIO()
			with wave.open(buf, "wb") as wf:
				wf.setnchannels(1)
				wf.setsampwidth(2)
				wf.setframerate(target_sr)
				wf.writeframes(pcm_out.tobytes())
			return True, buf.getvalue(), ""
		except Exception as exc:
			return False, b"", f"Ses kodlama hatası: {exc}"

	def get_voice_status(self) -> dict[str, Any]:
		with self._buf_lock:
			elapsed = max(0.0, time.monotonic() - self._started_at) if self._started_at else 0.0
			auto_st = bool(self._auto_stopped)
			st_dict = {
				"recording": bool(self._active),
				"level": int(self._current_level),
				"elapsed": round(elapsed, 2),
				"speech_detected": bool(self._speech_detected),
				"silence_sec": round(self._silence_after_speech, 2),
				"auto_stopped": auto_st,
				"error": self._stream_error,
			}
		if auto_st and self._stream is not None:
			with self._ctrl_lock:
				self._close_stream_unlocked()
				with self._buf_lock:
					if self._completed_wav is None and not self._completed_err:
						chunks_snap = list(self._chunks)
						self._chunks = []
						sr_snap = self._samplerate
						ok_e, wav_b, err_e = self._encode_chunks_to_wav(chunks_snap, sr_snap)
						if ok_e:
							self._completed_wav = wav_b
						else:
							self._completed_err = err_e
		return st_dict

	def cancel_recording(self) -> None:
		with self._ctrl_lock:
			self._close_stream_unlocked()
			with self._buf_lock:
				self._chunks = []
				self._completed_wav = None
				self._completed_err = ""
				self._auto_stopped = False
				self._speech_detected = False
				self._current_level = 0

	def stop_recording(self) -> tuple[bool, bytes, str]:
		with self._ctrl_lock:
			self._close_stream_unlocked()
			with self._buf_lock:
				if self._completed_wav is not None:
					wav_ready = self._completed_wav
					self._completed_wav = None
					self._completed_err = ""
					self._chunks = []
					return True, wav_ready, ""
				if self._completed_err:
					err_ready = self._completed_err
					self._completed_err = ""
					self._chunks = []
					return False, b"", err_ready
				chunks_snap = list(self._chunks)
				self._chunks = []
				sr_snap = self._samplerate
			return self._encode_chunks_to_wav(chunks_snap, sr_snap)

	def speak_text(self, text: str) -> bool:
		self.stop_speaking()
		clean = (text or "").strip()
		if not clean:
			return False
		import sys, subprocess, shutil
		if sys.platform == "darwin" and shutil.which("say"):
			cmds = [
				["/usr/bin/say", "-v", "Yelda", "-r", "195", clean[:3500]],
				["/usr/bin/say", "-r", "195", clean[:3500]],
			]
			for cmd in cmds:
				try:
					proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
					with self._tts_lock:
						self._tts_procs.append(proc)
					return True
				except Exception:
					continue
		return False

	def is_speaking(self) -> bool:
		with self._tts_lock:
			alive: list[Any] = []
			for p in self._tts_procs:
				try:
					if p.poll() is None:
						alive.append(p)
				except Exception:
					pass
			self._tts_procs = alive
			return bool(alive)

	def stop_speaking(self) -> None:
		with self._tts_lock:
			procs = list(self._tts_procs)
			self._tts_procs = []
		for p in procs:
			try:
				if p.poll() is None:
					p.terminate()
					try:
						p.wait(timeout=0.18)
					except Exception:
						p.kill()
						try:
							p.wait(timeout=0.2)
						except Exception:
							pass
			except Exception:
				pass


_AUDIO_CTRL = _NativeAudioController()


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
						init_tb = init_res.get("topbar_html") or ""
						init_main = init_res.get("main_html") or ""
					except Exception:
						init_json = "null"
						init_styles = ""
						init_sb = ""
						init_tb = ""
						init_main = ""
					page_html = (
						_SHELL_HTML
						.replace("__REVIR_CSRF__", server_ref.csrf_token)
						.replace("__REVIR_INIT_STYLES__", init_styles)
						.replace("__REVIR_INIT_SIDEBAR__", init_sb)
						.replace("__REVIR_INIT_TOPBAR__", init_tb)
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
					item = server_ref.session.resolve_download_item(token)
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

				if path == "/api/cmd_search":
					try:
						payload = self._read_json()
						q_str = str(payload.get("q") or "")
						handler = server_ref.session.cmd_search_handler
						items = handler(q_str) if callable(handler) else []
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", json.dumps({"ok": True, "items": items}, ensure_ascii=False).encode("utf-8"))
					except Exception as exc:
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", json.dumps({"ok": False, "items": [], "error": str(exc)}, ensure_ascii=False).encode("utf-8"))
					return

				if path == "/api/gemini_chat_clear":
					with server_ref.session.lock:
						server_ref.session.chat_history.clear()
					_AUDIO_CTRL.stop_speaking()
					self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					return

				if path == "/api/voice_start":
					ok_v, info_v = _AUDIO_CTRL.start_recording()
					if ok_v:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": True, "device": info_v}, ensure_ascii=False).encode("utf-8"))
					else:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": info_v}, ensure_ascii=False).encode("utf-8"))
					return

				if path == "/api/voice_status":
					st_v = _AUDIO_CTRL.get_voice_status()
					self._send_bytes(HTTPStatus.OK, "application/json", json.dumps(st_v, ensure_ascii=False).encode("utf-8"))
					return

				if path == "/api/voice_cancel":
					_AUDIO_CTRL.cancel_recording()
					self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					return

				if path == "/api/voice_stop":
					ok_s, wav_bytes, err_s = _AUDIO_CTRL.stop_recording()
					if ok_s and wav_bytes:
						b64_wav = base64.b64encode(wav_bytes).decode("ascii")
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": True, "data_b64": b64_wav, "mime": "audio/wav"}).encode("utf-8"))
					else:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": err_s or "Ses kaydı alınamadı."}, ensure_ascii=False).encode("utf-8"))
					return

				if path == "/api/tts_speak":
					payload = self._read_json()
					handled = _AUDIO_CTRL.speak_text(str(payload.get("text") or ""))
					self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": True, "handled": handled}).encode("utf-8"))
					return

				if path == "/api/tts_status":
					self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"speaking": _AUDIO_CTRL.is_speaking()}).encode("utf-8"))
					return

				if path == "/api/tts_stop":
					_AUDIO_CTRL.stop_speaking()
					self._send_bytes(HTTPStatus.OK, "application/json", b'{"ok":true}')
					return

				if path == "/api/gemini_chat_apply_write":
					try:
						payload = self._read_json()
						w_handler = server_ref.session.floating_chat_write_handler
						if not callable(w_handler):
							self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": "Veri işleme yöneticisi hazır değil."}, ensure_ascii=False).encode("utf-8"))
							return
						w_res = w_handler(payload)
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", json.dumps(w_res, ensure_ascii=False).encode("utf-8"))
					except Exception as exc:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode("utf-8"))
					return

				if path == "/api/gemini_export_msg":
					try:
						payload = self._read_json()
						exp_handler = server_ref.session.floating_chat_export_handler
						if not callable(exp_handler):
							self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": "Dışa aktarma işleyicisi hazır değil."}, ensure_ascii=False).encode("utf-8"))
							return
						exp_res = exp_handler(payload)
						self._send_bytes(HTTPStatus.OK, "application/json; charset=utf-8", json.dumps(exp_res, ensure_ascii=False).encode("utf-8"))
					except Exception as exc:
						self._send_bytes(HTTPStatus.OK, "application/json", json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode("utf-8"))
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
			_AUDIO_CTRL.stop_speaking()
			_AUDIO_CTRL.cancel_recording()
		except Exception:
			pass
		try:
			self._httpd.shutdown()
			self._httpd.server_close()
		except Exception:
			pass

	def get_download_item(self, token: str) -> tuple[bytes, str, str] | None:
		return self.session.resolve_download_item(token)


def reset_session() -> None:
	"""Test veya oturum sıfırlama için aktif oturumu temizler."""
	global _ACTIVE_SESSION
	_ACTIVE_SESSION = _AppSession()


def run_action_headless(app_fn: Callable[[], None], payload: dict[str, Any] | None = None) -> dict[str, Any]:
	"""Otomatik doğrulama ve testler için başsız render döngüsü çalıştırır."""
	return _execute_render_cycle(app_fn, _ACTIVE_SESSION, payload or {"action": "init"})
