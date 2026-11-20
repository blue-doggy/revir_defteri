"""
Okul Revir Defteri ve Sağlık Takip Sistemi - Masaüstü Başlatıcı Köprüsü (desktop_launcher.py)

Tüm masaüstü başlatma mantığı saf Python masaüstü yöneticisi olan `desktop_app.py`
üzerinden yürütülür.
"""
from __future__ import annotations

from desktop_app import launch_desktop_app, main


if __name__ == "__main__":
	raise SystemExit(main())
