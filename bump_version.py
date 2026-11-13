"""Okul Revir Defteri - Tek Merkezden Sürüm Yükseltme (SemVer) Aracı.

Kullanım:
    python bump_version.py             -> Varsayılan olarak 'patch' artırır (örn. 1.1.0 -> 1.1.1)
    python bump_version.py patch       -> Düzeltme sürümü artırır (X.Y.Z+1)
    python bump_version.py minor       -> Özellik/optimizasyon sürümü artırır (X.Y+1.0)
    python bump_version.py major       -> Ana sürümü artırır (X+1.0.0)
    python bump_version.py 1.2.3       -> Sürümü doğrudan belirtilen değere sabitler
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent
VERSION_FILE = ROOT_DIR / "VERSION"
SEMVER_RE = re.compile(r"^(\d+)\.(\d+)\.(\d+)$")


def read_current_version() -> str:
	if not VERSION_FILE.exists():
		return "1.1.0"
	txt = VERSION_FILE.read_text(encoding="utf-8").strip()
	if not SEMVER_RE.match(txt):
		raise ValueError(f"Geçersiz VERSION biçimi (X.Y.Z bekleniyor): {txt!r}")
	return txt


def compute_next_version(current: str, mode: str) -> str:
	m = SEMVER_RE.match(current.strip())
	if not m:
		raise ValueError(f"Geçersiz mevcut sürüm: {current!r}")
	major, minor, patch = int(m.group(1)), int(m.group(2)), int(m.group(3))

	norm = mode.strip().lower()
	if norm == "patch":
		return f"{major}.{minor}.{patch + 1}"
	if norm == "minor":
		return f"{major}.{minor + 1}.0"
	if norm == "major":
		return f"{major + 1}.0.0"
	if SEMVER_RE.match(norm):
		return norm
	raise ValueError(f"Desteklenmeyen sürüm modu veya değeri: {mode!r}. Kullanım: patch | minor | major | X.Y.Z")


def sync_version_references(new_version: str) -> list[str]:
	changed: list[str] = []

	old_v_txt = VERSION_FILE.read_text(encoding="utf-8").strip() if VERSION_FILE.exists() else None
	if old_v_txt != new_version:
		VERSION_FILE.write_text(new_version + "\n", encoding="utf-8")
		changed.append("VERSION")

	replacements: list[tuple[Path, str, str]] = [
		(
			ROOT_DIR / "persantil.py",
			r'(_DEFAULT_APP_VERSION\s*=\s*")\d+\.\d+\.\d+(")',
			rf'\g<1>{new_version}\g<2>',
		),
		(
			ROOT_DIR / "desktop_app.py",
			r'(_DEFAULT_APP_VERSION\s*=\s*")\d+\.\d+\.\d+(")',
			rf'\g<1>{new_version}\g<2>',
		),
		(
			ROOT_DIR / "RevirDefteri_Setup.iss",
			r'(#define\s+MyAppVersion\s+")[^"]+(")',
			rf'\g<1>{new_version}\g<2>',
		),
		(
			ROOT_DIR / "RevirDefteri_Setup.iss",
			r'(; Çıktı:\s*installer\\RevirDefteri_Kurulum_v)[^\s]+(\.exe)',
			rf'\g<1>{new_version}\g<2>',
		),
		(
			ROOT_DIR / "kurulum_olustur.iss",
			r'(#define\s+MyAppVersion\s+")[^"]+(")',
			rf'\g<1>{new_version}\g<2>',
		),
	]

	for path, pattern, repl in replacements:
		if not path.exists():
			continue
		content = path.read_text(encoding="utf-8")
		updated, count = re.subn(pattern, repl, content)
		if count > 0 and updated != content:
			path.write_text(updated, encoding="utf-8")
			changed.append(path.name)

	return changed


def main(argv: list[str] | None = None) -> int:
	args = argv if argv is not None else sys.argv[1:]
	mode = args[0] if args else "patch"
	current = read_current_version()
	if mode == "--current":
		print(current)
		return 0
	if mode == "--sync":
		changed = sync_version_references(current)
		print(f"Sürüm senkronize edildi: v{current} (Güncellenenler: {changed or 'yok'})")
		return 0

	nxt = compute_next_version(current, mode)
	changed = sync_version_references(nxt)
	print(f"Sürüm yükseltildi: v{current} -> v{nxt} (Güncellenen dosyalar: {', '.join(changed)})")
	return 0


if __name__ == "__main__":
	raise SystemExit(main())
