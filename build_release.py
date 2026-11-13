"""Okul Revir Defteri - Tek Tıkla Sürüm Yükseltme, PyInstaller ve Akıllı Kurulum (.EXE) Üreticisi.

Kullanım:
    python build_release.py                       -> Etkileşimli sürüm sorar, PyInstaller + ISCC çalıştırır
    python build_release.py --version minor --yes -> Sürümü minor artırır (örn. 1.1.1 -> 1.2.0) ve derler
    python build_release.py --version 1.2.0 --yes -> Sürümü 1.2.0 yapıp tüm paketleri derler
"""
from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

from bump_version import (
	compute_next_version,
	read_current_version,
	sync_version_references,
)

PROJECT_ROOT = Path(__file__).resolve().parent
SPEC_FILE = PROJECT_ROOT / "RevirDefteri.spec"
ISS_FILE = PROJECT_ROOT / "RevirDefteri_Setup.iss"
DIST_EXE = PROJECT_ROOT / "dist" / "RevirDefteri" / "RevirDefteri.exe"
INSTALLER_DIR = PROJECT_ROOT / "installer"
PACKAGES_DIR = PROJECT_ROOT / "Kurulum_Paketleri"


def find_iscc_executable() -> Path | None:
	"""Bilgisayardaki Inno Setup Compiler (ISCC.exe) yolunu otomatik bulur."""
	which_hit = shutil.which("ISCC.exe")
	if which_hit:
		return Path(which_hit)

	local_app = Path(os.environ.get("LocalAppData") or (Path.home() / "AppData" / "Local"))
	candidates = [
		Path(r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"),
		Path(r"C:\Program Files\Inno Setup 6\ISCC.exe"),
		local_app / "Programs" / "Inno Setup 6" / "ISCC.exe",
		Path(r"C:\Program Files (x86)\Inno Setup 7\ISCC.exe"),
		Path(r"C:\Program Files\Inno Setup 7\ISCC.exe"),
		local_app / "Programs" / "Inno Setup 7" / "ISCC.exe",
		Path(r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe"),
	]
	for candidate in candidates:
		if candidate.exists():
			return candidate
	return None


def _prompt_target_version(current: str, default_mode: str = "patch") -> str:
	"""Etkileşimli kip için kullanıcıdan yeni sürüm numarasını veya SemVer modunu alır."""
	auto_next = compute_next_version(current, default_mode)
	prompt_msg = (
		f"Mevcut Sürüm: {current} -> Yeni Sürüm Numarasını Girin "
		f"(Enter'a basarsanız otomatik {auto_next} olur): "
	)
	raw = input(prompt_msg).strip()
	if not raw:
		return auto_next
	return compute_next_version(current, raw)


def _clean_old_installers(folder: Path, keep_name: str) -> None:
	if not folder.exists():
		return
	for old_file in folder.glob("RevirDefteri_Kurulum*.exe"):
		if old_file.name.lower() != keep_name.lower():
			try:
				old_file.unlink()
			except OSError:
				pass


def build_release(target_version: str) -> tuple[Path, list[Path]]:
	"""Sürümü senkronize eder, PyInstaller ve Inno Setup ile .exe kurulumunu üretir."""
	changed = sync_version_references(target_version)
	print(f"[1/4] Sürüm v{target_version} olarak eşitlendi (Güncellenen: {', '.join(changed) or 'zaten güncel'}).")

	if not SPEC_FILE.exists():
		raise FileNotFoundError(f"PyInstaller spec dosyası bulunamadı: {SPEC_FILE}")

	print("[2/4] PyInstaller ile RevirDefteri.exe derleniyor...")
	subprocess.run(
		[sys.executable, "-m", "PyInstaller", str(SPEC_FILE), "--clean", "-y"],
		cwd=str(PROJECT_ROOT),
		check=True,
	)
	if not DIST_EXE.exists():
		raise FileNotFoundError(f"Beklenen masaüstü .exe üretilemedi: {DIST_EXE}")
	print(f"[OK] Masaüstü uygulaması hazır: {DIST_EXE}")

	iscc_exe = find_iscc_executable()
	if iscc_exe is None:
		raise FileNotFoundError(
			"Inno Setup Compiler (ISCC.exe) bulunamadı. Lütfen Inno Setup 6 kurun."
		)
	if not ISS_FILE.exists():
		raise FileNotFoundError(f"Inno Setup scripti bulunamadı: {ISS_FILE}")

	INSTALLER_DIR.mkdir(parents=True, exist_ok=True)
	print(f"[3/4] Inno Setup ({iscc_exe}) ile kurulum sihirbazı derleniyor...")
	subprocess.run(
		[str(iscc_exe), str(ISS_FILE)],
		cwd=str(PROJECT_ROOT),
		check=True,
	)

	setup_name = f"RevirDefteri_Kurulum_v{target_version}.exe"
	built_setup = INSTALLER_DIR / setup_name
	if not built_setup.exists():
		raise FileNotFoundError(f"Beklenen kurulum dosyası üretilemedi: {built_setup}")

	print("[4/4] Çıktılar düzenleniyor ve eski kök kopyalar temizleniyor...")
	_clean_old_installers(PROJECT_ROOT, setup_name)
	_clean_old_installers(INSTALLER_DIR, setup_name)
	PACKAGES_DIR.mkdir(parents=True, exist_ok=True)
	_clean_old_installers(PACKAGES_DIR, setup_name)

	root_copy = PROJECT_ROOT / setup_name
	pkg_copy = PACKAGES_DIR / setup_name
	shutil.copy2(built_setup, root_copy)
	shutil.copy2(built_setup, pkg_copy)

	outputs = [built_setup, root_copy, pkg_copy]
	print(f"[BASARILI] Kurulum ve Güncelleme Paketi hazır (v{target_version}):")
	for out_p in outputs:
		print(f"    - {out_p} ({out_p.stat().st_size:,} bayt)")
	return DIST_EXE, outputs


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description="Okul Revir Defteri sürüm yükseltme ve tek tıkla .EXE oluşturma aracı."
	)
	parser.add_argument(
		"--version",
		"-v",
		dest="version_arg",
		default="",
		help="Yeni sürüm (patch | minor | major | X.Y.Z)",
	)
	parser.add_argument(
		"--yes",
		"-y",
		action="store_true",
		help="Soru sormadan doğrudan derlemeyi başlat",
	)
	return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
	args = parse_args(argv)
	current = read_current_version()

	if args.version_arg:
		target_version = compute_next_version(current, args.version_arg)
	elif args.yes:
		target_version = compute_next_version(current, "patch")
	else:
		target_version = _prompt_target_version(current, "patch")

	build_release(target_version)
	return 0


if __name__ == "__main__":
	raise SystemExit(main())
