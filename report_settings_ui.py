"""Reusable report output settings panel."""

from typing import Any, Callable

import revir_ui as ui


def render_report_settings_panel(
	*,
	get_report_settings: Callable[[], dict[str, Any]],
	save_report_settings: Callable[[dict[str, Any]], None],
	default_report_settings: dict[str, Any],
	key_prefix: str = "rep",
	in_dialog: bool = False,
) -> None:
	"""Render shared report section, detail, and preset controls."""
	kp = str(key_prefix or "rep").strip() or "rep"
	ui.markdown("### 🖨️ Rapor Çıktı & Yazdırma Ayarları")
	ui.caption("Kişi ve sınıf raporu çıktılarında (Word `.docx` ve yazdırma) yer alacak bölümleri ve detay tercihlerini özelleştirin.")
	cur_cfg = get_report_settings()

	tab_sections, tab_fields, tab_presets = ui.tabs([
		"📑 Rapor Bölümleri",
		"🔍 Detay Alanları",
		"⚡ Hızlı Şablonlar"
	], orientation="segmented")

	def _after_save(msg: str) -> None:
		ui.success(msg)
		if in_dialog:
			ui.rerun(scope="fragment")
		else:
			ui.rerun()

	with tab_presets:
		ui.markdown("##### ⚡ Hazır Rapor Şablonları")
		ui.caption("Tek tıkla ihtiyacınıza uygun bir rapor şablonu seçebilirsiniz:")
		pr_c1, pr_c2 = ui.columns(2)
		with pr_c1:
			if ui.button("✅ Tam Kapsamlı Resmi Sağlık Raporu", use_container_width=True, help="Tüm bölümler ve geçmiş kayıtlar eksiksiz dahil edilir.", key=f"{kp}_preset_full"):
				save_report_settings(dict(default_report_settings))
				_after_save("Tam kapsamlı resmi rapor şablonu uygulandı!")
			if ui.button("🩺 Yalnızca Revir & Muayene Raporu", use_container_width=True, help="Sadece kimlik, uyarılar ve revir vizit geçmişi dahil edilir.", key=f"{kp}_preset_visits"):
				custom_s = dict(default_report_settings)
				custom_s.update({
					"include_health_profile": False,
					"include_growth_summary": False,
					"include_lab_results": False,
					"include_measurements_history": False,
					"include_care_plans": False,
					"include_vaccinations": False,
				})
				save_report_settings(custom_s)
				_after_save("Yalnızca revir muayene raporu şablonu uygulandı!")
		with pr_c2:
			if ui.button("📏 Yalnızca Büyüme & Persantil Raporu", use_container_width=True, help="DSÖ büyüme eğrileri, boy-kilo geçmişi ve persantil değerlendirmesi.", key=f"{kp}_preset_growth"):
				custom_s = dict(default_report_settings)
				custom_s.update({
					"include_visits": False,
					"include_care_plans": False,
					"include_lab_results": False,
					"include_vaccinations": False,
				})
				save_report_settings(custom_s)
				_after_save("Büyüme ve persantil raporu şablonu uygulandı!")
			if ui.button("📋 Özet Sağlık Künyesi (Vizitsiz Form)", use_container_width=True, help="Alerji, kronik durum, acil notu ve vital özetinden oluşan tek sayfalık form.", key=f"{kp}_preset_summary"):
				custom_s = dict(default_report_settings)
				custom_s.update({
					"include_visits": False,
					"include_care_plans": False,
					"include_measurements_history": False,
					"include_lab_results": False,
				})
				save_report_settings(custom_s)
				_after_save("Özet sağlık künyesi şablonu uygulandı!")

	with tab_sections:
		ui.markdown("##### 📑 Rapora Dahil Edilecek Ana Bölümler")
		ui.caption("Rapor çıktısında yer almasını istemediğiniz bölümlerin işaretini kaldırabilirsiniz:")
		sec_c1, sec_c2 = ui.columns(2)
		with sec_c1:
			chk_id = ui.checkbox("👤 Öğrenci Kimlik ve Temel Bilgileri", value=cur_cfg.get("include_id_box", True), key=f"{kp}_chk_id")
			chk_alert = ui.checkbox("⚠️ Kritik Sağlık & Alerji Uyarı Kutusu", value=cur_cfg.get("include_alert_box", True), key=f"{kp}_chk_alert")
			chk_hp = ui.checkbox("📋 Sağlık Profili & Tarama Bilgileri", value=cur_cfg.get("include_health_profile", True), key=f"{kp}_chk_hp")
			chk_growth = ui.checkbox("📏 DSÖ Klinik Büyüme & Persantil Taraması", value=cur_cfg.get("include_growth_summary", True), key=f"{kp}_chk_growth")
			chk_vitals = ui.checkbox("💓 Vital Bulgular Ölçüm Geçmişi", value=cur_cfg.get("include_vitals", True), key=f"{kp}_chk_vitals")
			chk_labs = ui.checkbox("🧪 Laboratuvar Tahlil Sonuçları", value=cur_cfg.get("include_lab_results", True), key=f"{kp}_chk_labs")
		with sec_c2:
			chk_meas = ui.checkbox("⚖️ Boy, Kilo ve Baş Çevresi Ölçüm Geçmişi", value=cur_cfg.get("include_measurements_history", True), key=f"{kp}_chk_meas")
			chk_visits = ui.checkbox("🩺 Revir Başvuru ve Muayene Kayıtları", value=cur_cfg.get("include_visits", True), key=f"{kp}_chk_visits")
			chk_cp = ui.checkbox("📝 Hemşirelik Bakım Planları (NANDA-I)", value=cur_cfg.get("include_care_plans", True), key=f"{kp}_chk_cp")
			chk_vac = ui.checkbox("💉 Aşı ve Bağışıklama Kayıtları", value=cur_cfg.get("include_vaccinations", True), key=f"{kp}_chk_vac")
			chk_dis = ui.checkbox("♿ Engellilik / Özel Gereksinim Profili (Varsa)", value=cur_cfg.get("include_disability_profile", True), key=f"{kp}_chk_dis")
			chk_ai = ui.checkbox("🤖 AI Klinik Analiz & Eksik Veri Bölümü (Varsayılan)", value=cur_cfg.get("include_ai_analysis", False), key=f"{kp}_chk_ai")
			chk_sig = ui.checkbox("✍️ Yetkili ve Veli İmza / Kaşe Bölümü", value=cur_cfg.get("include_signatures", True), key=f"{kp}_chk_sig")
			chk_foot = ui.checkbox("ℹ️ Yasal Sorumluluk Reddi / Dipnot", value=cur_cfg.get("include_footer_disclaimer", True), key=f"{kp}_chk_foot")

	with tab_fields:
		ui.markdown("##### 🔍 Detay Bilgi ve Filtre Tercihleri")
		fld_c1, fld_c2 = ui.columns(2)
		with fld_c1:
			chk_guard = ui.checkbox("📞 Veli Adı ve Telefon Bilgisini Göster", value=cur_cfg.get("show_guardian_contact", True), key=f"{kp}_chk_guard")
			chk_bg = ui.checkbox("🩸 Kan Grubu Bilgisini Göster", value=cur_cfg.get("show_blood_group", True), key=f"{kp}_chk_bg")
			chk_sym = ui.checkbox("🩺 Vizit Kayıtlarında Şikayet / Belirtileri Göster", value=cur_cfg.get("show_visit_complaint", True), key=f"{kp}_chk_sym")
			chk_act = ui.checkbox("🛠️ Vizit Kayıtlarında Yapılan İşlemi Göster", value=cur_cfg.get("show_visit_action", True), key=f"{kp}_chk_act")
		with fld_c2:
			chk_obs = ui.checkbox("🔍 Vizit Kayıtlarında Fizik Muayene / Gözlemi Göster", value=cur_cfg.get("show_visit_observation", True), key=f"{kp}_chk_obs")
			chk_ref = ui.checkbox("🚑 Sevk ve Yönlendirme Durumunu Göster", value=cur_cfg.get("show_visit_referral", True), key=f"{kp}_chk_ref")
			chk_z = ui.checkbox("📊 Ölçümlerde DSÖ Z-Skorlarını (SDS) Göster", value=cur_cfg.get("show_meas_zscores", True), key=f"{kp}_chk_z")
			chk_p = ui.checkbox("📈 Ölçümlerde Persantil Dilimlerini (%p) Göster", value=cur_cfg.get("show_meas_percentiles", True), key=f"{kp}_chk_p")

		ui.markdown("---")
		ui.markdown("###### 🕒 Vizit Kayıt Limiti")
		limit_opts = [0, 5, 10, 15, 20, 30]
		cur_lim = cur_cfg.get("visits_limit", 0)
		lim_idx = limit_opts.index(cur_lim) if cur_lim in limit_opts else 0
		sel_limit = ui.selectbox(
			"Raporda listelenecek maksimum vizit sayısı",
			limit_opts,
			index=lim_idx,
			format_func=lambda x: "Tüm Geçmiş Vizitler (Limitsiz)" if x == 0 else f"Son {x} Vizit Kaydı",
			key=f"{kp}_sel_visits_limit"
		)

	ui.markdown("---")
	if ui.button("💾 Rapor Çıktı Ayarlarını Kaydet", type="primary", use_container_width=True, key=f"btn_save_report_settings_{kp}"):
		new_saved = {
			"include_id_box": chk_id,
			"include_alert_box": chk_alert,
			"include_health_profile": chk_hp,
			"include_growth_summary": chk_growth,
			"include_vitals": chk_vitals,
			"include_lab_results": chk_labs,
			"include_measurements_history": chk_meas,
			"include_visits": chk_visits,
			"include_care_plans": chk_cp,
			"include_vaccinations": chk_vac,
			"include_disability_profile": chk_dis,
			"include_ai_analysis": chk_ai,
			"include_signatures": chk_sig,
			"include_footer_disclaimer": chk_foot,
			"show_guardian_contact": chk_guard,
			"show_blood_group": chk_bg,
			"show_visit_complaint": chk_sym,
			"show_visit_action": chk_act,
			"show_visit_observation": chk_obs,
			"show_visit_referral": chk_ref,
			"show_meas_zscores": chk_z,
			"show_meas_percentiles": chk_p,
			"visits_limit": sel_limit,
		}
		save_report_settings(new_saved)
		_after_save("Rapor çıktı ayarları başarıyla kaydedildi!")
