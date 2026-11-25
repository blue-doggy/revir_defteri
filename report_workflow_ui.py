"""Shared presentation helpers for report creation workflows."""

from collections.abc import Sequence

import revir_ui as ui


def render_step_indicator(step_index: int, step_labels: Sequence[str]) -> None:
	if not step_labels or not 0 <= step_index < len(step_labels):
		raise ValueError("step_index must refer to an existing report workflow step")
	ui.progress((step_index + 1) / len(step_labels))
	ui.markdown(f"##### Adım {step_index + 1}/{len(step_labels)} · {step_labels[step_index]}")


def render_confirmation_summary(report_title: str, details: Sequence[tuple[str, str]]) -> None:
	with ui.container(border=True):
		ui.markdown("#### Son Kontrol")
		ui.caption("İndirmeden önce rapor kapsamını ve içerik ayarlarını doğrulayın.")
		ui.caption("Rapor türü")
		ui.text(report_title)
		columns = ui.columns(2)
		for index, (label, value) in enumerate(details):
			with columns[index % len(columns)]:
				ui.caption(label)
				ui.text(str(value))
