"""
Okul Revir Defteri — Microsoft Store MSIX Paketleyici ve Görsel Varlık Üreticisi (build_msix.py)

Özellikler:
- `VERSION` (X.Y.Z) dosyasını Microsoft Store zorunlu 4 haneli sürümüne (`X.Y.Z.0`) dönüştürür.
- `app_icon.png` dosyasından tüm zorunlu MSIX `Assets/*.png` logolarını ve Partner Center vitrin görsellerini üretir.
- Microsoft Store kurallarına uyumlu (`runFullTrust`, `microphone`, `Windows.Desktop`) `AppxManifest.xml` oluşturur.
- `dist/RevirDefteri` çıktısını veritabanı, log ve kök sertifika kalıntılarından arındırarak sahneler,
  `MS_STORE_BUILD` işaretçisini ekler ve Windows SDK `makeappx.exe` ile `.msix` paketini derler.

Kullanım:
    python build_msix.py                   -> PyInstaller + Assets + AppxManifest + .msix derler
    python build_msix.py --skip-pyinstaller -> Mevcut dist/RevirDefteri üzerinden hızlı .msix üretir
    python build_msix.py --assets-only     -> Yalnızca ikon setini ve AppxManifest.xml önizlemesini üretir
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape as xml_escape

from PIL import Image, ImageDraw

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "msix_config.json"

EXCLUDED_FILE_NAMES = {
	"revir_defteri.db",
	"revir_defteri.db-shm",
	"revir_defteri.db-wal",
	"revir_defteri.db-journal",
	".revir_server.log",
	".window_state.json",
	".revir_desktop.pid",
	"RevirDefteri_RootCA.cer",
	"RevirDefteri_CodeSign.pfx",
	"[Content_Types].xml",
}

EXCLUDED_DIR_NAMES = {
	"kisi_verileri",
	"egitim_yillari",
	"revir_evraklari",
	"oto_yedekler",
	"guncellemeler",
	".webview2_cache",
	"__pycache__",
	"default-docx-template",
}

# MSIX paketi içindeki zorunlu ve önerilen ölçekli görsel varlıklar: (dosya_adı, genişlik, yükseklik, dolgu_oranı)
MSIX_ASSET_SPECS: tuple[tuple[str, int, int, float], ...] = (
	("Square44x44Logo.png", 44, 44, 0.88),
	("Square44x44Logo.scale-100.png", 44, 44, 0.88),
	("Square44x44Logo.scale-200.png", 88, 88, 0.88),
	("Square44x44Logo.targetsize-24_altform-unplated.png", 24, 24, 0.92),
	("Square44x44Logo.targetsize-48_altform-unplated.png", 48, 48, 0.90),
	("Square44x44Logo.targetsize-256_altform-unplated.png", 256, 256, 0.90),
	("Square71x71Logo.png", 71, 71, 0.84),
	("Square71x71Logo.scale-200.png", 142, 142, 0.84),
	("Square150x150Logo.png", 150, 150, 0.82),
	("Square150x150Logo.scale-100.png", 150, 150, 0.82),
	("Square150x150Logo.scale-200.png", 300, 300, 0.82),
	("Wide310x150Logo.png", 310, 150, 0.78),
	("Wide310x150Logo.scale-200.png", 620, 300, 0.78),
	("Square310x310Logo.png", 310, 310, 0.80),
	("Square310x310Logo.scale-200.png", 620, 620, 0.80),
	("StoreLogo.png", 50, 50, 0.86),
	("StoreLogo.scale-100.png", 50, 50, 0.86),
	("StoreLogo.scale-200.png", 100, 100, 0.86),
)


def to_msix_quad_version(semver: str) -> str:
	"""
	SemVer (`X.Y.Z`) sürümünü Microsoft Store Partner Center zorunlu 4 haneli
	`Major.Minor.Build.0` biçimine dönüştürür (4. hane Store yüklemelerinde daima 0 olmalıdır).
	"""
	clean = str(semver or "").strip().lstrip("vV")
	m = re.match(r"^(\d+)\.(\d+)\.(\d+)(?:\.\d+)?$", clean)
	if not m:
		raise ValueError(f"Geçersiz sürüm numarası (X.Y.Z bekleniyor): {semver!r}")
	maj, min_v, pat = int(m.group(1)), int(m.group(2)), int(m.group(3))
	if any(v > 65535 for v in (maj, min_v, pat)):
		raise ValueError(f"MSIX sürüm bileşenleri 0..65535 aralığında olmalıdır: {semver!r}")
	return f"{maj}.{min_v}.{pat}.0"


def load_msix_config(config_path: Path = DEFAULT_CONFIG_PATH) -> dict[str, Any]:
	defaults: dict[str, Any] = {
		"identity_name": "blueddreamer.RevirDefteri",
		"publisher": "CN=1096CE83-946D-45FB-AD8B-94E6437B5EFD",
		"publisher_display_name": "blueddreamer",
		"display_name": "Revir Defteri",
		"package_family_name": "blueddreamer.RevirDefteri_93bhc84zrfnny",
		"package_sid": "S-1-15-2-821227329-492232997-165842215-2630038997-1000387688-1387758346-1805854164",
		"store_id": "9N8629W34880",
		"description": (
			"Okul revirleri için yerel veritabanlı sağlık izlem, WHO büyüme persantili, "
			"ilaç ve diyabet karar destek, okul taraması ve resmi raporlama sistemi."
		),
		"background_color": "#0D9488",
		"min_windows_version": "10.0.17763.0",
		"max_windows_version_tested": "10.0.22621.0",
		"language": "tr-tr",
		"privacy_policy_url": "https://github.com/blue-doggy/revir_defteri/blob/main/PRIVACY_POLICY.md",
	}
	if config_path.exists():
		loaded = json.loads(config_path.read_text(encoding="utf-8"))
		if isinstance(loaded, dict):
			for k, v in loaded.items():
				if not str(k).startswith("_") and v is not None and str(v).strip() != "":
					defaults[k] = v
	return defaults


def _place_centered_icon(
	src_img: Image.Image,
	canvas_w: int,
	canvas_h: int,
	scale_ratio: float = 0.84,
	bg_rgba: tuple[int, int, int, int] = (0, 0, 0, 0),
) -> Image.Image:
	canvas = Image.new("RGBA", (canvas_w, canvas_h), bg_rgba)
	max_box = max(1, int(round(min(canvas_w, canvas_h) * scale_ratio)))
	src_w, src_h = src_img.size
	ratio = min(max_box / max(1, src_w), max_box / max(1, src_h))
	new_w = max(1, int(round(src_w * ratio)))
	new_h = max(1, int(round(src_h * ratio)))
	resized = src_img.resize((new_w, new_h), Image.Resampling.LANCZOS)
	off_x = (canvas_w - new_w) // 2
	off_y = (canvas_h - new_h) // 2
	canvas.alpha_composite(resized, (off_x, off_y))
	return canvas


def _build_gradient_banner(src_img: Image.Image, width: int, height: int) -> Image.Image:
	"""Partner Center mağaza vitrini (Poster 720x1080 / Hero 1920x1080) için temiz kurumsal kapak üretir."""
	banner = Image.new("RGBA", (width, height), (15, 23, 42, 255))
	draw = ImageDraw.Draw(banner)
	for y in range(height):
		t = y / max(1, height - 1)
		r = int(13 + (15 - 13) * t)
		g = int(148 + (38 - 148) * t)
		b = int(136 + (66 - 136) * t)
		draw.line([(0, y), (width, y)], fill=(r, g, b, 255))

	icon_box = int(min(width, height) * 0.44)
	resized = src_img.resize((icon_box, icon_box), Image.Resampling.LANCZOS)
	ox = (width - icon_box) // 2
	oy = (height - icon_box) // 2
	banner.alpha_composite(resized, (ox, oy))
	return banner


def generate_store_assets(
	icon_path: Path,
	msix_assets_dir: Path,
	submission_assets_dir: Path | None = None,
) -> list[Path]:
	"""`app_icon.png` dosyasından tüm MSIX kutucuk/mağaza ikonlarını ve Partner Center görsellerini üretir."""
	if not icon_path.exists():
		raise FileNotFoundError(f"Kaynak uygulama ikonu bulunamadı: {icon_path}")

	msix_assets_dir.mkdir(parents=True, exist_ok=True)
	created: list[Path] = []
	with Image.open(icon_path) as raw_im:
		src_rgba = raw_im.convert("RGBA")
		for fname, w, h, ratio in MSIX_ASSET_SPECS:
			out_file = msix_assets_dir / fname
			tile = _place_centered_icon(src_rgba, w, h, scale_ratio=ratio)
			tile.save(out_file, format="PNG", optimize=True)
			created.append(out_file)

		if submission_assets_dir is not None:
			submission_assets_dir.mkdir(parents=True, exist_ok=True)
			for s_name, sw, sh in (
				("StoreListing_Logo_300x300.png", 300, 300),
				("StoreListing_Logo_150x150.png", 150, 150),
			):
				p_logo = submission_assets_dir / s_name
				_place_centered_icon(src_rgba, sw, sh, scale_ratio=0.88, bg_rgba=(240, 253, 250, 255)).convert("RGB").save(
					p_logo, format="PNG", optimize=True
				)
				created.append(p_logo)

			for b_name, bw, bh in (
				("StoreListing_Poster_720x1080.png", 720, 1080),
				("StoreListing_Hero_1920x1080.png", 1920, 1080),
			):
				p_ban = submission_assets_dir / b_name
				_build_gradient_banner(src_rgba, bw, bh).convert("RGB").save(p_ban, format="PNG", optimize=True)
				created.append(p_ban)

	return created


def render_appx_manifest(cfg: dict[str, Any], quad_version: str, arch: str = "x64") -> str:
	"""Microsoft Store kurallarına uyumlu `AppxManifest.xml` içeriğini üretir ve XML sözdizimini doğrular."""
	id_name = str(cfg["identity_name"]).strip()
	publisher = str(cfg["publisher"]).strip()
	if not publisher.upper().startswith("CN="):
		raise ValueError(f"Publisher değeri 'CN=...' ile başlamalıdır: {publisher!r}")
	if not re.match(r"^[A-Za-z0-9.-]{3,50}$", id_name):
		raise ValueError(f"Identity Name yalnızca harf, rakam, nokta ve tire içermelidir (3-50 krk): {id_name!r}")

	disp_name = xml_escape(str(cfg["display_name"]).strip())
	pub_disp = xml_escape(str(cfg["publisher_display_name"]).strip())
	desc = xml_escape(str(cfg["description"]).strip())
	bg_col = xml_escape(str(cfg.get("background_color") or "#0D9488").strip())
	min_ver = xml_escape(str(cfg.get("min_windows_version") or "10.0.17763.0").strip())
	max_ver = xml_escape(str(cfg.get("max_windows_version_tested") or "10.0.22621.0").strip())
	lang = xml_escape(str(cfg.get("language") or "tr-tr").strip())

	xml_text = f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap rescap">

  <Identity
    Name="{xml_escape(id_name)}"
    Publisher="{xml_escape(publisher)}"
    Version="{xml_escape(quad_version)}"
    ProcessorArchitecture="{xml_escape(arch)}" />

  <Properties>
    <DisplayName>{disp_name}</DisplayName>
    <PublisherDisplayName>{pub_disp}</PublisherDisplayName>
    <Logo>Assets\\StoreLogo.png</Logo>
  </Properties>

  <Dependencies>
    <TargetDeviceFamily
      Name="Windows.Desktop"
      MinVersion="{min_ver}"
      MaxVersionTested="{max_ver}" />
  </Dependencies>

  <Resources>
    <Resource Language="{lang}" />
  </Resources>

  <Applications>
    <Application
      Id="RevirDefteri"
      Executable="RevirDefteri.exe"
      EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements
        DisplayName="{disp_name}"
        Description="{desc}"
        BackgroundColor="{bg_col}"
        Square150x150Logo="Assets\\Square150x150Logo.png"
        Square44x44Logo="Assets\\Square44x44Logo.png">
        <uap:DefaultTile
          Wide310x150Logo="Assets\\Wide310x150Logo.png"
          Square71x71Logo="Assets\\Square71x71Logo.png"
          Square310x310Logo="Assets\\Square310x310Logo.png">
          <uap:ShowNameOnTiles>
            <uap:ShowOn Tile="square150x150Logo" />
            <uap:ShowOn Tile="wide310x150Logo" />
            <uap:ShowOn Tile="square310x310Logo" />
          </uap:ShowNameOnTiles>
        </uap:DefaultTile>
      </uap:VisualElements>
    </Application>
  </Applications>

  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
    <DeviceCapability Name="microphone" />
  </Capabilities>
</Package>
"""
	# XML bütünlüğünü doğrula
	ET.fromstring(xml_text.encode("utf-8"))
	return xml_text


def stage_msix_directory(
	dist_dir: Path,
	staging_dir: Path,
	icon_path: Path,
	cfg: dict[str, Any],
	quad_version: str,
	submission_assets_dir: Path | None = None,
) -> Path:
	"""PyInstaller `dist/RevirDefteri` çıktısını temiz bir MSIX sahneleme klasörüne hazırlar."""
	exe_file = dist_dir / "RevirDefteri.exe"
	if not exe_file.exists():
		raise FileNotFoundError(f"Paketlenmiş uygulama bulunamadı: {exe_file}")

	if staging_dir.exists():
		shutil.rmtree(staging_dir)
	staging_dir.mkdir(parents=True, exist_ok=True)

	def _ignore_fn(cur_dir: str, names: list[str]) -> set[str]:
		ignored: set[str] = set()
		for n in names:
			low = n.lower()
			if n in EXCLUDED_FILE_NAMES or n in EXCLUDED_DIR_NAMES:
				ignored.add(n)
			elif (
				low.endswith((".db", ".db-wal", ".db-shm", ".cer", ".pfx", ".pdb"))
				or "[" in n
				or "]" in n
				or ("persantil veriler" in cur_dir.lower() and low.startswith("ogrenci_"))
			):
				ignored.add(n)
		return ignored

	for item in dist_dir.iterdir():
		if item.name in EXCLUDED_FILE_NAMES or item.name in EXCLUDED_DIR_NAMES:
			continue
		if item.name.lower().endswith((".db", ".db-wal", ".db-shm", ".cer", ".pfx", ".pdb")):
			continue
		target = staging_dir / item.name
		if item.is_dir():
			shutil.copytree(item, target, ignore=_ignore_fn)
		else:
			shutil.copy2(item, target)

	# Microsoft Store çalışma zamanı işaretçisi (uygulama içi .exe güncelleyiciyi otomatik kapatır)
	(staging_dir / "MS_STORE_BUILD").write_text("microsoft-store\n", encoding="utf-8")
	internal_dir = staging_dir / "_internal"
	if internal_dir.is_dir():
		(internal_dir / "MS_STORE_BUILD").write_text("microsoft-store\n", encoding="utf-8")

	generate_store_assets(icon_path, staging_dir / "Assets", submission_assets_dir=submission_assets_dir)
	manifest_xml = render_appx_manifest(cfg, quad_version=quad_version)
	manifest_path = staging_dir / "AppxManifest.xml"
	manifest_path.write_text(manifest_xml, encoding="utf-8")
	return manifest_path


def find_makeappx_exe() -> Path | None:
	"""Sistemdeki veya önbellekteki Windows SDK `makeappx.exe` aracını otomatik bulur."""
	from_path = shutil.which("makeappx.exe") or shutil.which("makeappx")
	if from_path:
		return Path(from_path)

	local_cache_base = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local")) / "RevirDefteri_BuildTools" / "bin"
	roots = [
		PROJECT_ROOT / "build" / "sdk_tools" / "bin",
		local_cache_base,
		Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Windows Kits" / "10" / "bin",
		Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Windows Kits" / "10" / "bin",
	]
	candidates: list[Path] = []
	for root in roots:
		if not root.exists():
			continue
		for sub in root.glob("10.0.*/x64/makeappx.exe"):
			candidates.append(sub)
		direct = root / "x64" / "makeappx.exe"
		if direct.exists():
			candidates.append(direct)
	if not candidates:
		return None
	candidates.sort(key=lambda p: str(p), reverse=True)
	return candidates[0]


def ensure_makeappx_exe(auto_download: bool = True) -> Path | None:
	"""
	`makeappx.exe` kurulu değilse (ör. MEB ağında tam Windows SDK kurulumu 15605 hatası verdiğinde),
	resmi Microsoft NuGet `Microsoft.Windows.SDK.BuildTools` paketinden (~20 MB) taşınabilir
	ve Microsoft imzalı `makeappx.exe` + `signtool.exe` araçlarını otomatik indirip hazırlar.
	"""
	existing = find_makeappx_exe()
	if existing is not None:
		return existing
	if not auto_download:
		return None

	import io
	import urllib.request
	import zipfile

	try:
		import truststore
		truststore.inject_into_ssl()
	except Exception:
		pass

	index_url = "https://api.nuget.org/v3-flatcontainer/microsoft.windows.sdk.buildtools/index.json"
	print("[*] Sistemde makeappx.exe bulunamadı; resmi Microsoft.Windows.SDK.BuildTools (~20 MB) indiriliyor...")
	req_idx = urllib.request.Request(index_url, headers={"User-Agent": "RevirDefteri-MSIXBuilder/1.0"})
	with urllib.request.urlopen(req_idx, timeout=25) as resp:
		idx_data = json.loads(resp.read().decode("utf-8"))

	versions = [str(v) for v in (idx_data.get("versions") or []) if "-" not in str(v)]
	if not versions:
		versions = [str(v) for v in (idx_data.get("versions") or [])]
	if not versions:
		raise RuntimeError("NuGet üzerinde Microsoft.Windows.SDK.BuildTools sürümü bulunamadı.")

	ver = versions[-1]
	nupkg_url = f"https://api.nuget.org/v3-flatcontainer/microsoft.windows.sdk.buildtools/{ver}/microsoft.windows.sdk.buildtools.{ver}.nupkg"
	req_pkg = urllib.request.Request(nupkg_url, headers={"User-Agent": "RevirDefteri-MSIXBuilder/1.0"})
	with urllib.request.urlopen(req_pkg, timeout=120) as resp_pkg:
		raw_nupkg = resp_pkg.read()

	dest_root = PROJECT_ROOT / "build" / "sdk_tools"
	dest_root.mkdir(parents=True, exist_ok=True)
	with zipfile.ZipFile(io.BytesIO(raw_nupkg)) as zf:
		for member in zf.namelist():
			if member.startswith("bin/") and not member.endswith("/"):
				target_path = dest_root / Path(member)
				target_path.parent.mkdir(parents=True, exist_ok=True)
				target_path.write_bytes(zf.read(member))

	found = find_makeappx_exe()
	if found is not None:
		print(f"[OK] Resmi Microsoft SDK araçları hazırlandı (v{ver}): {found}")
	return found


def pack_msix(makeappx_exe: Path, staging_dir: Path, output_msix: Path) -> None:
	output_msix.parent.mkdir(parents=True, exist_ok=True)
	for old_msix in output_msix.parent.glob("RevirDefteri_Store_v*.msix"):
		if old_msix.resolve() != output_msix.resolve():
			old_msix.unlink(missing_ok=True)
	if output_msix.exists():
		output_msix.unlink()
	cmd = [str(makeappx_exe), "pack", "/d", str(staging_dir), "/p", str(output_msix), "/o"]
	res = subprocess.run(cmd, capture_output=True)
	if res.returncode != 0 or not output_msix.exists():
		out_txt = (res.stdout or b"").decode("cp1254", errors="replace") + "\n" + (res.stderr or b"").decode("cp1254", errors="replace")
		raise RuntimeError(f"makeappx.exe paketleme hatası (kod {res.returncode}):\n{out_txt.strip()}")


def main(argv: list[str] | None = None) -> int:
	parser = argparse.ArgumentParser(description="Okul Revir Defteri — Microsoft Store MSIX Paketleyici")
	parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH, help="msix_config.json dosya yolu")
	parser.add_argument("--skip-pyinstaller", action="store_true", help="Mevcut dist/RevirDefteri klasörünü kullan")
	parser.add_argument("--assets-only", action="store_true", help="Yalnızca MSIX ikonlarını, vitrin görsellerini ve AppxManifest.xml önizlemesini üret")
	args = parser.parse_args(argv)

	ver_raw = (PROJECT_ROOT / "VERSION").read_text(encoding="utf-8").strip()
	quad_ver = to_msix_quad_version(ver_raw)
	cfg = load_msix_config(args.config)
	icon_path = PROJECT_ROOT / "app_icon.png"
	sub_assets_dir = PROJECT_ROOT / "store_submission_assets"

	if args.assets_only:
		preview_dir = PROJECT_ROOT / "build" / "msix_preview"
		created = generate_store_assets(icon_path, preview_dir / "Assets", submission_assets_dir=sub_assets_dir)
		manifest_file = preview_dir / "AppxManifest.xml"
		manifest_file.write_text(render_appx_manifest(cfg, quad_ver), encoding="utf-8")
		print(f"[OK] {len(created)} görsel varlık ve AppxManifest.xml hazırlandı (v{quad_ver}):")
		print(f"     - Manifest : {manifest_file}")
		print(f"     - Vitrin   : {sub_assets_dir}")
		return 0

	dist_dir = PROJECT_ROOT / "dist" / "RevirDefteri"
	if not args.skip_pyinstaller:
		print(f"[1/3] PyInstaller ile RevirDefteri.spec derleniyor (v{ver_raw})...")
		py_cmd = [sys.executable, "-m", "PyInstaller", str(PROJECT_ROOT / "RevirDefteri.spec"), "--clean", "-y"]
		rc = subprocess.call(py_cmd, cwd=PROJECT_ROOT)
		if rc != 0:
			print("[HATA] PyInstaller derlemesi başarısız oldu!", file=sys.stderr)
			return rc

	staging_dir = PROJECT_ROOT / "build" / "msix_staging"
	print(f"[2/3] MSIX sahneleme klasörü ve Store görselleri hazırlanıyor ({staging_dir})...")
	stage_msix_directory(
		dist_dir=dist_dir,
		staging_dir=staging_dir,
		icon_path=icon_path,
		cfg=cfg,
		quad_version=quad_ver,
		submission_assets_dir=sub_assets_dir,
	)

	if cfg["publisher"] == "CN=00000000-0000-0000-0000-000000000000":
		print(
			"[UYARI] msix_config.json içinde örnek Publisher ('CN=00000000-...') kullanılıyor.\n"
			"        Partner Center'a yüklemeden önce Product Identity değerlerinizi msix_config.json dosyasına girin."
		)

	try:
		makeappx = ensure_makeappx_exe(auto_download=True)
	except Exception as exc:
		print(f"[UYARI] Taşınabilir SDK araçları otomatik indirilemedi: {exc}", file=sys.stderr)
		makeappx = None
	out_msix = PROJECT_ROOT / "installer" / f"RevirDefteri_Store_v{ver_raw}.msix"
	if makeappx is None:
		print(
			"\n[BİLGİ] Sahneleme klasörü ve AppxManifest.xml eksiksiz hazırlandı!\n"
			"        Ancak sistemde Windows SDK 'makeappx.exe' bulunamadı.\n"
			"        Windows SDK kurmak için PowerShell'de şu komutu çalıştırın:\n"
			"            winget install Microsoft.WindowsSDK.10.0.22621\n"
			"        Ardından 'python build_msix.py --skip-pyinstaller' komutuyla .msix dosyasını anında üretebilirsiniz."
		)
		return 0

	print(f"[3/3] makeappx.exe ile MSIX paketi oluşturuluyor: {out_msix}")
	pack_msix(makeappx, staging_dir, out_msix)
	print(f"\n[BAŞARILI] Microsoft Store MSIX paketi hazır: {out_msix}")
	print(f"           Partner Center vitrin görselleri : {sub_assets_dir}")
	return 0


if __name__ == "__main__":
	raise SystemExit(main())
