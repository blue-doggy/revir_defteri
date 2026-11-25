"""Application navigation metadata shared by the main UI and tests."""

PAGE_NAV_GROUPS: tuple[tuple[str, tuple[tuple[str, str, str, str | None], ...]], ...] = (
	("GÜNLÜK KLİNİK", (
		("📊 Genel Bakış", "📊", "Genel Bakış", "open_followups"),
		("🩺 Revir Kaydı", "🩺", "Revir Kaydı", "today_visits"),
		("💊 İlaç Takip Sistemi", "💊", "İlaç Takip", "active_meds"),
		("💉 İnsülin & Diyabet Yönetimi", "💉", "İnsülin & Diyabet", "diabetes_profiles"),
		("💓 Vital Takip", "💓", "Vital Takip", None),
	)),
	("KİŞİ & TARAMA", (
		("👥 Öğrenciler", "👥", "Öğrenciler", None),
		("👫 Yetişkinler", "👫", "Yetişkinler", None),
		("🏫 Okul Taraması", "🏫", "Okul Taraması", None),
	)),
	("GELİŞİM & PLAN", (
		("📋 Sağlık Profili ve Bakım Planı", "📋", "Sağlık Profili & Plan", None),
		("📏 Ölçüm ve Gelişim", "📏", "Ölçüm & Gelişim", None),
		("📅 Takvim & Planlama", "📅", "Takvim & Planlama", None),
		("🍽️ Yemekhane Yönetimi", "🍽️", "Yemekhane Yönetimi", None),
	)),
	("BELGE & SİSTEM", (
		("🖨️ Raporlar", "🖨️", "Raporlar", None),
		("📁 Evraklar", "📁", "Evraklar", None),
		("⚙️ Ayarlar", "⚙️", "Ayarlar", None),
	)),
)

PAGES = [item[0] for group in PAGE_NAV_GROUPS for item in group[1]]
