"""Cross-platform Dear PyGui desktop GUI for New QC Sheet."""

from __future__ import annotations

import argparse
import sys
from importlib.metadata import PackageNotFoundError, version as package_version

import dearpygui.dearpygui as dpg

from xlsx_clean.core import (
    create_datasheet,
    list_ink_colors,
    list_sets,
    load_config,
)
from xlsx_clean.paths import default_backend, package_file

try:
    _VERSION = package_version('xlsx-clean')
except PackageNotFoundError:
    _VERSION = 'dev'

# Module-level variables
rows = []
addin_paths = []
config = {}


# Tags
SET_COMBO = "set_combo"
INK_COMBO = "ink_combo"
SERIAL_INPUT = "serial_input"
STATUS_TEXT = "status_text"
HEADING_FONT = "heading_font"

def on_set_change(sender, app_data):
    """Callback when the set combo changes."""
    new_inks = list_ink_colors(rows, app_data)
    dpg.configure_item(INK_COMBO, items=new_inks)
    dpg.set_value(INK_COMBO, new_inks[0] if new_inks else '')

def on_create(sender, app_data):
    """Callback for the Create datasheet button."""
    selected_set = dpg.get_value(SET_COMBO)
    selected_dir = dpg.get_value(INK_COMBO)
    batch_serial = dpg.get_value(SERIAL_INPUT)
    
    try:
        result = create_datasheet(
            selected_set=selected_set,
            selected_dir=selected_dir,
            batch_serial=batch_serial,
            backend=None,
            rows=rows,
            addin_paths=addin_paths,
            config=config
        )
        
        msg = result.message
        if result.template:
            msg += f"\nTemplate: {result.template}"
        if result.destination:
            msg += f"\nDestination: {result.destination}"
        if result.backend:
            msg += f"\nBackend: {result.backend}"
            
        dpg.set_value(STATUS_TEXT, msg)
        
        if result.skipped:
            dpg.configure_item(STATUS_TEXT, color=(230, 180, 30))  # Yellow
        elif result.ok:
            dpg.configure_item(STATUS_TEXT, color=(0, 200, 80))  # Green
        else:
            dpg.configure_item(STATUS_TEXT, color=(220, 60, 60))  # Red
    except Exception as e:
        dpg.set_value(STATUS_TEXT, f"Error: {e}")
        dpg.configure_item(STATUS_TEXT, color=(220, 60, 60))

def _create_modern_theme() -> int:
    """Build a global theme that gives DPG widgets a modern, web-like feel, matching the HTML demo."""
    with dpg.theme() as theme:
        with dpg.theme_component(dpg.mvAll):
            # Rounded corners (like CSS border-radius)
            dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 4)
            dpg.add_theme_style(dpg.mvStyleVar_PopupRounding, 4)
            dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 8)
            dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 4)
            dpg.add_theme_style(dpg.mvStyleVar_TabRounding, 4)
            dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 4)
            dpg.add_theme_style(dpg.mvStyleVar_ScrollbarRounding, 4)

            # Generous padding (like CSS padding)
            dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 12, 10)
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 10)
            dpg.add_theme_style(dpg.mvStyleVar_ItemInnerSpacing, 8, 6)
            dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 48, 48)

            # Scrollbar
            dpg.add_theme_style(dpg.mvStyleVar_ScrollbarSize, 12)

            # Colors from demo.html
            bg = (248, 250, 252, 255)           # Slate 50 (#f8fafc)
            surface = (255, 255, 255, 255)      # White (#ffffff)
            frame = (255, 255, 255, 255)        # White (inputs)
            frame_hover = (240, 253, 250, 255)  # Teal 50 (#f0fdfa)
            frame_active = (204, 251, 241, 255) # Teal 100 (#ccfbf1)
            accent = (15, 118, 110, 255)        # Teal 700 (#0f766e) - Primary
            accent_hover = (17, 94, 89, 255)    # Teal 800 (#115e59) - Secondary
            accent_active = (19, 78, 74, 255)   # Teal 900 (#134e4a)
            text = (15, 23, 42, 255)            # Slate 900 (#0f172a)
            text_dim = (100, 116, 139, 255)     # Slate 500 (#64748b)
            border = (203, 213, 225, 255)       # Slate 300 (#cbd5e1)
            popup_bg = (255, 255, 255, 255)     # White dropdowns

            # Window
            dpg.add_theme_color(dpg.mvThemeCol_WindowBg, bg)
            dpg.add_theme_color(dpg.mvThemeCol_PopupBg, popup_bg)
            dpg.add_theme_color(dpg.mvThemeCol_TitleBg, bg)
            dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, bg)

            # Frame (combos, inputs)
            dpg.add_theme_color(dpg.mvThemeCol_FrameBg, frame)
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, frame_hover)
            dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, frame_active)

            # Border
            dpg.add_theme_color(dpg.mvThemeCol_Border, border)
            dpg.add_theme_color(dpg.mvThemeCol_BorderShadow, (0, 0, 0, 0))
            dpg.add_theme_style(dpg.mvStyleVar_FrameBorderSize, 1)

            # Button (Combo box drop arrows) - blend into frame
            dpg.add_theme_color(dpg.mvThemeCol_Button, frame)
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, frame_hover)
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, frame_active)

            # Header (combo dropdown items on hover)
            dpg.add_theme_color(dpg.mvThemeCol_Header, frame_hover)
            dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, frame_active)
            dpg.add_theme_color(dpg.mvThemeCol_HeaderActive, accent)

            # Text
            dpg.add_theme_color(dpg.mvThemeCol_Text, text)
            dpg.add_theme_color(dpg.mvThemeCol_TextDisabled, text_dim)

            # Separator
            dpg.add_theme_color(dpg.mvThemeCol_Separator, border)

            # Scrollbar
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg, bg)
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrab, border)
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrabHovered, text_dim)
            dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrabActive, text)

            # Check mark / selection
            dpg.add_theme_color(dpg.mvThemeCol_CheckMark, accent)

    return theme



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="New QC Sheet — Dear PyGui desktop application.",
    )
    parser.add_argument('--width', type=int, default=650, help='Window width')
    parser.add_argument('--height', type=int, default=700, help='Window height')
    return parser.parse_args()


def main() -> None:
    global rows, addin_paths, config
    args = parse_args()

    dpg.create_context()

    # --- Modern theme ---------------------------------------------------
    modern_theme = _create_modern_theme()
    dpg.bind_theme(modern_theme)

    # --- Load Inter font ------------------------------------------------
    font_regular = package_file("fonts/Inter-Regular.ttf")
    font_bold = package_file("fonts/Inter-Bold.ttf")
    default_font = None
    heading_font = None
    label_font = None
    if font_regular.is_file():
        with dpg.font_registry():
            default_font = dpg.add_font(str(font_regular), 18)
            label_font = dpg.add_font(str(font_regular), 14)
            if font_bold.is_file():
                heading_font = dpg.add_font(str(font_bold), 28, tag=HEADING_FONT)

    # Load config and initialize variables
    sets = []
    initial_inks = []
    error_msg = ""

    label_set = 'Select Set'
    label_ink = 'Select Ink Color'
    label_serial = 'Serial:'

    try:
        content, loaded_rows, loaded_config = load_config()
        rows = loaded_rows
        config = loaded_config
        if len(content) >= 5:
            label_set = content[0]
            label_ink = content[1]
            label_serial = content[2]
            addin_paths = content[3:5]

        sets = list_sets(rows)
        if sets:
            initial_inks = list_ink_colors(rows, sets[0])
    except Exception as e:
        error_msg = f"Failed to load config: {e}"

    with dpg.theme() as tight_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, -8)

    with dpg.theme() as normal_spacing_theme:
        with dpg.theme_component(dpg.mvAll):
            dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 10, 10)

    with dpg.theme() as btn_theme:
        with dpg.theme_component(dpg.mvButton):
            dpg.add_theme_color(dpg.mvThemeCol_Text, (255, 255, 255, 255))
            dpg.add_theme_color(dpg.mvThemeCol_Button, (15, 118, 110, 255))
            dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (17, 94, 89, 255))
            dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (19, 78, 74, 255))

    with dpg.window(tag="primary_window"):
        if default_font:
            dpg.bind_font(default_font)

        title = dpg.add_text('New QC Sheet', color=(17, 94, 89, 255))
        if heading_font:
            dpg.bind_item_font(title, heading_font)
        dpg.add_text(f"v{_VERSION}", color=(100, 106, 120))
        dpg.add_spacer(height=4)
        dpg.add_separator()
        dpg.add_spacer(height=10)

        def add_floating_label(text):
            lbl = dpg.add_text(text, color=(15, 118, 110, 255))
            if label_font:
                dpg.bind_item_font(lbl, label_font)

        with dpg.group() as g1:
            add_floating_label(label_set)
            c1 = dpg.add_combo(items=sets, default_value=sets[0] if sets else '',
                          tag=SET_COMBO, callback=on_set_change, width=-1)
            dpg.bind_item_theme(c1, normal_spacing_theme)
        dpg.bind_item_theme(g1, tight_theme)
        dpg.add_spacer(height=8)

        with dpg.group() as g2:
            add_floating_label(label_ink)
            c2 = dpg.add_combo(items=initial_inks,
                          default_value=initial_inks[0] if initial_inks else '',
                          tag=INK_COMBO, width=-1)
            dpg.bind_item_theme(c2, normal_spacing_theme)
        dpg.bind_item_theme(g2, tight_theme)
        dpg.add_spacer(height=8)

        with dpg.group() as g3:
            add_floating_label(label_serial)
            c3 = dpg.add_input_text(tag=SERIAL_INPUT, width=-1, hint="e.g. 1234/A")
            dpg.bind_item_theme(c3, normal_spacing_theme)
        dpg.bind_item_theme(g3, tight_theme)
        dpg.add_spacer(height=10)

        dpg.add_separator()
        dpg.add_spacer(height=10)

        btn = dpg.add_button(label='Create datasheet', width=-1, height=38,
                             callback=on_create)
        dpg.bind_item_theme(btn, btn_theme)
        dpg.add_spacer(height=14)

        dpg.add_text(error_msg, tag=STATUS_TEXT, wrap=args.width - 40)
        if error_msg:
            dpg.configure_item(STATUS_TEXT, color=(220, 60, 60))

    dpg.create_viewport(title=f'New QC Sheet · v{_VERSION}',
                        width=args.width, height=args.height)
    dpg.setup_dearpygui()
    dpg.show_viewport()
    dpg.set_primary_window("primary_window", True)
    dpg.start_dearpygui()
    dpg.destroy_context()


if __name__ == '__main__':
    main()

