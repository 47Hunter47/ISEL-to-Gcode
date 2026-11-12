//! Industrial-SCADA palette: deep navy panels, cyan accent, mint-white text.
//!
//! Sampled from the reference keyvisual (Siemens SCADA): the dark blue family
//! (#001030..#004060), the cyan glow (#00C0E0) and the mint-white text
//! (#D0E8E0). Sharp corners, hairline edges, no shadows — panels read as
//! instruments and the toolpath stays the only loud thing on screen.
//!
//! `visuals()` is pure data: no window, no GPU, so it is testable headless.

use egui::style::{WidgetVisuals, Widgets};
use egui::{Color32, CornerRadius, Shadow, Stroke, Visuals};

/// Window and panel background (measured #001830..#002040).
pub const PANEL: Color32 = Color32::from_rgb(0x00, 0x1C, 0x30);
/// Raised boxes: popup/dialog frames, faint stripes.
pub const PANEL_ALT: Color32 = Color32::from_rgb(0x0B, 0x2C, 0x44);
/// Canvas: preview background, log well, text inputs (measured #001030, deepened).
pub const DEEP: Color32 = Color32::from_rgb(0x00, 0x0E, 0x1A);
/// Inactive widget fill (measured #003050).
pub const BTN: Color32 = Color32::from_rgb(0x0B, 0x35, 0x52);
/// Hovered widget fill.
pub const BTN_HOVER: Color32 = Color32::from_rgb(0x14, 0x48, 0x6B);
/// Open/pressed widget fill.
pub const BTN_ACTIVE: Color32 = Color32::from_rgb(0x1B, 0x5A, 0x80);
/// The accent: cyan glow (measured #00C0E0). Cut path, primary button, progress,
/// sliders, checkboxes, selection.
pub const ACCENT: Color32 = Color32::from_rgb(0x00, 0xC0, 0xE0);
/// Dark glyphs placed on cyan fills.
pub const ON_ACCENT: Color32 = Color32::from_rgb(0x01, 0x1A, 0x24);
/// Secondary cyan: hover text, quiet highlights.
pub const ACCENT_SOFT: Color32 = Color32::from_rgb(0x6F, 0xD0, 0xE8);
/// Body text (measured #D0E8E0).
pub const TEXT: Color32 = Color32::from_rgb(0xD0, 0xE8, 0xE0);
/// Secondary text: stats, hints, placeholders.
pub const TEXT_WEAK: Color32 = Color32::from_rgb(0x7F, 0xA8, 0xB8);
/// Grid numbers and ticks — must not compete with the toolpath.
pub const TEXT_FAINT: Color32 = Color32::from_rgb(0x2F, 0x5F, 0x73);
/// Hairlines: panel borders, widget edges.
pub const LINE: Color32 = Color32::from_rgb(0x2A, 0x6A, 0x86);
/// Log lines (phosphor mint on the deep well).
pub const LOG: Color32 = Color32::from_rgb(0xA8, 0xEC, 0xC8);
/// G2/G3 arcs.
pub const ARC: Color32 = Color32::from_rgb(0x3F, 0xD0, 0xB0);
/// Rapid moves: SCADA warning amber (the first themes used red).
pub const RAPID: Color32 = Color32::from_rgb(0xF0, 0xA0, 0x20);
/// Success ("Conversion complete!").
pub const OK: Color32 = Color32::from_rgb(0x5F, 0xE0, 0xA8);
/// Errors.
pub const ERR: Color32 = Color32::from_rgb(0xFF, 0x6B, 0x5A);
/// Preview grid lines.
pub const GRID: Color32 = Color32::from_rgb(0x10, 0x37, 0x4F);
/// Preview X/Y zero axes.
pub const AXIS: Color32 = Color32::from_rgb(0x1C, 0x5A, 0x78);
/// Selected text background.
pub const SELECTION: Color32 = Color32::from_rgb(0x1A, 0x5A, 0x70);

fn widget(bg: Color32, fg: Color32, edge: Color32, width: f32) -> WidgetVisuals {
    WidgetVisuals {
        bg_fill: bg,
        weak_bg_fill: bg,
        bg_stroke: Stroke::new(width, edge),
        corner_radius: CornerRadius::same(2),
        fg_stroke: Stroke::new(width, fg),
        expansion: 2.0,
    }
}

/// Build the whole egui visual set. Pure: no context, no window, testable.
pub fn visuals() -> Visuals {
    let mut v = Visuals::dark();

    v.panel_fill = PANEL;
    v.window_fill = PANEL;
    v.extreme_bg_color = DEEP;
    v.faint_bg_color = PANEL_ALT;
    v.code_bg_color = DEEP;
    // The path field sits in a deep well with a hairline edge.
    v.text_edit_bg_color = Some(DEEP);
    v.hyperlink_color = ACCENT;
    v.weak_text_color = Some(TEXT_WEAK);
    v.warn_fg_color = RAPID;
    v.error_fg_color = ERR;

    // SCADA panels are rectangles: no rounded window, no drop shadows.
    v.window_corner_radius = CornerRadius::ZERO;
    v.menu_corner_radius = CornerRadius::same(2);
    v.window_stroke = Stroke::new(1.0, LINE);
    v.window_shadow = Shadow::NONE;
    v.popup_shadow = Shadow::NONE;

    // Plain labels read their colour from the noninteractive foreground, so this
    // one field sets the body-text colour everywhere.
    let mut w = Widgets::default();
    w.noninteractive = widget(PANEL, TEXT, LINE, 1.0);
    w.inactive = widget(BTN, TEXT, LINE, 1.0);
    w.hovered = widget(BTN_HOVER, ACCENT_SOFT, ACCENT_SOFT, 1.5);
    w.active = widget(ACCENT, ON_ACCENT, ACCENT, 1.5);
    w.open = widget(BTN_ACTIVE, ACCENT_SOFT, ACCENT_SOFT, 1.5);
    v.widgets = w;

    // Sliders/checkboxes light up cyan; the filled part of a slider reads as a gauge.
    v.slider_trailing_fill = true;

    v.selection.bg_fill = SELECTION;
    v.selection.stroke = Stroke::new(1.0, ACCENT);
    v
}

/// Apply the palette to a live egui context.
pub fn apply(ctx: &egui::Context) {
    let v = visuals();
    // egui keeps one style per theme bucket and picks the bucket from the OS
    // theme. Setting only the current bucket (what `set_visuals` does) leaks
    // egui's own light defaults on a light-mode Windows: white progress track,
    // white checkbox. Fill both buckets and pin the preference.
    ctx.set_theme(egui::ThemePreference::Dark);
    ctx.set_visuals_of(egui::Theme::Dark, v.clone());
    ctx.set_visuals_of(egui::Theme::Light, v);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn apply_survives_the_light_system_theme() {
        // Regression: a light-mode Windows used to show egui's stock light
        // defaults (white progress track and checkbox) because only the active
        // theme bucket had been filled.
        let ctx = egui::Context::default();
        apply(&ctx);
        for theme in [egui::Theme::Dark, egui::Theme::Light] {
            let v = &ctx.style_of(theme).visuals;
            assert_eq!(v.panel_fill, PANEL, "{theme:?} panel_fill");
            assert_eq!(
                v.extreme_bg_color, DEEP,
                "{theme:?} extreme_bg_color is the progress bar track"
            );
            assert_eq!(v.widgets.inactive.bg_fill, BTN, "{theme:?} widget fill");
            assert_eq!(v.widgets.noninteractive.fg_stroke.color, TEXT, "{theme:?} label text");
        }
    }
}
