"""
Revir Defteri - Konsolsuz Python Masaüstü Başlatıcısı (.pyw)
Doğrudan yerel Python masaüstü penceresini başlatır.
"""
from __future__ import annotations

import socket
import subprocess
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _quick_git_sync() -> None:
	try:
		s = socket.create_connection(("8.8.8.8", 53), timeout=1.0)
		s.close()
		r = subprocess.run(
			["git", "remote", "get-url", "origin"],
			cwd=APP_DIR,
			capture_output=True,
			text=True,
			timeout=2,
			creationflags=CREATE_NO_WINDOW,
		)
		if r.returncode != 0:
			subprocess.run(
				["git", "remote", "add", "origin", "https://github.com/blue-doggy/revir_defteri.git"],
				cwd=APP_DIR,
				capture_output=True,
				timeout=2,
				creationflags=CREATE_NO_WINDOW,
			)
		subprocess.run(
			["git", "pull", "--ff-only", "origin", "main"],
			cwd=APP_DIR,
			capture_output=True,
			timeout=5,
			creationflags=CREATE_NO_WINDOW,
		)
	except Exception:
		pass


def main() -> int:
	_quick_git_sync()
	if str(APP_DIR) not in sys.path:
		sys.path.insert(0, str(APP_DIR))
	from desktop_app import launch_desktop_app
	return launch_desktop_app()


if __name__ == "__main__":
	raise SystemExit(main())
