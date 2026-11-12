//! The theme is pure data: testable with no window and no GPU.

use egui::Color32;
use isel_to_gcode::theme;

fn lum(c: Color32) -> f32 {
    let [r, g, b, _a] = c.to_array();
    (0.2126 * r as f32 + 0.7152 * g as f32 + 0.0722 * b as f32) / 255.0
}

#[test]
fn test_dark_surfaces_and_bright_accent() {
    assert!(lum(theme::DEEP) < lum(theme::PANEL), "canvas must sit behind panels");
    assert!(lum(theme::PANEL) < lum(theme::BTN), "widgets must lift off the panel");
    assert!(lum(theme::BTN) < lum(theme::BTN_HOVER));
    assert!(lum(theme::ACCENT) > 0.45, "cyan has to glow on navy");
    assert!(lum(theme::TEXT) > 0.75, "body text must be near-white");
    assert!(lum(theme::TEXT_FAINT) < 0.35, "grid ticks must not fight the toolpath");
    for bg in [theme::DEEP, theme::PANEL, theme::BTN] {
        assert_ne!(bg, theme::ACCENT);
        assert!(lum(bg) < lum(theme::ACCENT));
    }
}

#[test]
fn test_toolpath_colors_are_distinguishable() {
    let trio = [theme::ACCENT, theme::ARC, theme::RAPID];
    for (i, a) in trio.iter().enumerate() {
        for b in &trio[i + 1..] {
            assert_ne!(a, b, "cut / arc / rapid must never look alike");
        }
        assert!(lum(*a) > lum(theme::GRID) + 0.15, "toolpath must beat the grid");
    }
}

#[test]
fn test_visuals_actually_carry_the_palette() {
    let v = theme::visuals();
    assert_eq!(v.panel_fill, theme::PANEL);
    assert_eq!(v.window_fill, theme::PANEL);
    assert_eq!(v.extreme_bg_color, theme::DEEP);
    assert_eq!(v.hyperlink_color, theme::ACCENT);
    // plain labels read their colour here, so this is the body-text colour
    assert_eq!(v.widgets.noninteractive.fg_stroke.color, theme::TEXT);
    assert_eq!(v.weak_text_color, Some(theme::TEXT_WEAK));
    // sliders / checkboxes / pressed buttons turn cyan with dark glyphs
    assert_eq!(v.widgets.active.bg_fill, theme::ACCENT);
    assert_eq!(v.widgets.active.fg_stroke.color, theme::ON_ACCENT);
    assert_eq!(v.selection.stroke.color, theme::ACCENT);
    // flat and sharp: no floating shadows, corners stay near-square
    assert_eq!(v.window_shadow, egui::Shadow::NONE);
    assert!(v.widgets.inactive.corner_radius.average() <= 2.0);
    assert_eq!(v.window_corner_radius, egui::CornerRadius::ZERO);
}
