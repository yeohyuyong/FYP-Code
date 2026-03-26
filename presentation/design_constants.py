"""Design constants for the DSO DIIM Presentation."""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

# Slide dimensions (16:9 widescreen)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Colors
PRIMARY_DARK = RGBColor(0x1B, 0x2A, 0x4A)
PRIMARY_BLUE = RGBColor(0x2E, 0x50, 0x90)
ACCENT_TEAL = RGBColor(0x00, 0x97, 0xA7)
ACCENT_GOLD = RGBColor(0xD4, 0xA8, 0x43)
LIGHT_GRAY = RGBColor(0xF4, 0xF5, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED_ALERT = RGBColor(0xC0, 0x39, 0x2B)
GREEN_SUCCESS = RGBColor(0x27, 0xAE, 0x60)
TEXT_DARK = RGBColor(0x2C, 0x3E, 0x50)
TEXT_MID = RGBColor(0x6C, 0x7A, 0x89)

# Layout
MARGIN_LEFT = Inches(0.6)
MARGIN_RIGHT = Inches(0.6)
TITLE_BAR_HEIGHT = Inches(0.85)
BREADCRUMB_HEIGHT = Inches(0.30)
BREADCRUMB_TOP = TITLE_BAR_HEIGHT
CONTENT_TOP = Inches(1.25)
CONTENT_LEFT = Inches(0.7)
CONTENT_WIDTH = Inches(11.9)
CONTENT_HEIGHT = Inches(5.8)

# Font
FONT_FAMILY = 'Calibri'

# Breadcrumb sections
SECTIONS = [
    "Introduction",
    "DIIM Model",
    "DIIM Results",
    "Simplified Methods",
    "Monte Carlo",
    "Results",
    "Recommendations",
]
