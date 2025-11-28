import math

# GAS Config Port
# Base dimensions from GAS: W: 1123, H: 794 (Pixels at ~96 DPI)
# VBA uses Points (1/72 inch).
# Conversion: Points = Pixels * 0.75

class PPTConfig:
    BASE_WIDTH_PX = 960
    BASE_HEIGHT_PX = 720 # 4:3 aspect ratio typically, but we only care about width scaling for now
    
    # A4 Size in Points (Standard PowerPoint)
    SLIDE_WIDTH_PT = 841.68  # 11.69 inches * 72
    SLIDE_HEIGHT_PT = 595.44 # 8.27 inches * 72

    @staticmethod
    def px_to_pt(px):
        # Scale based on width ratio to ensure full width usage
        return px * (PPTConfig.SLIDE_WIDTH_PT / PPTConfig.BASE_WIDTH_PX)

    # Layout Positions (Ported from GAS POS_PX)
    # Using the same structure for easy mapping
    POS_PX = {
        "titleSlide": {
            "logo": {"left": 55, "top": 60, "width": 135},
            "title": {"left": 50, "top": 200, "width": 830, "height": 90},
            "date": {"left": 50, "top": 450, "width": 250, "height": 40}
        },
        "contentSlide": {
            "headerLogo": {"right": 20, "top": 20, "width": 75},
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "body": {"left": 25, "top": 132, "width": 910, "height": 330},
            "twoColLeft": {"left": 25, "top": 132, "width": 440, "height": 330},
            "twoColRight": {"left": 495, "top": 132, "width": 440, "height": 330}
        },
        "sectionSlide": {
            "title": {"left": 55, "top": 230, "width": 840, "height": 80},
            "ghostNum": {"left": 35, "top": 120, "width": 400, "height": 200}
        },
        "processSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "timelineSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "cycleSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "body": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "cardsSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "gridArea": {"left": 25, "top": 120, "width": 910, "height": 340}
        },
        "pyramidSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "pyramidArea": {"left": 25, "top": 120, "width": 910, "height": 360}
        },
        "triangleSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 110, "width": 910, "height": 350}
        },
        "compareSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "leftBox": {"left": 25, "top": 112, "width": 445, "height": 350},
            "rightBox": {"left": 490, "top": 112, "width": 445, "height": 350}
        },
        "diagramSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "flowChartSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "stepUpSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "imageTextSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "imageArea": {"left": 25, "top": 132, "width": 440, "height": 330},
            "textArea": {"left": 495, "top": 132, "width": 440, "height": 330}
        },
        "tableSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "tableArea": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "progressSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "quoteSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "quoteArea": {"left": 100, "top": 150, "width": 760, "height": 200},
            "authorArea": {"left": 100, "top": 360, "width": 760, "height": 50}
        },
        "kpiSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "bulletCardsSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "faqSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "statsCompareSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "leftBox": {"left": 25, "top": 132, "width": 440, "height": 330},
            "rightBox": {"left": 495, "top": 132, "width": 440, "height": 330}
        },
        "barCompareSlide": {
            "title": {"left": 25, "top": 20, "width": 830, "height": 65},
            "titleUnderline": {"left": 25, "top": 80, "width": 260, "height": 4},
            "subhead": {"left": 25, "top": 90, "width": 910, "height": 40},
            "area": {"left": 25, "top": 132, "width": 910, "height": 330}
        },
        "footer": {
            "leftText": {"left": 15, "top": 511, "width": 250, "height": 20},
            "rightPage": {"right": 15, "top": 511, "width": 50, "height": 20}
        },
        "bottomBar": {
            "left": 0, "top": 534, "width": 960, "height": 6 # Note: GAS width might be inconsistent with BASE, but we'll scale
        }
    }

    FONTS = {
        "family": "Meiryo", # Unified to Meiryo
        "sizes": {
            "title": 48,       # Increased from 41
            "date": 24,        # Increased from 16
            "sectionTitle": 44,# Increased from 38
            "contentTitle": 32,# Increased from 24
            "subhead": 24,     # Increased from 16
            "body": 24,        # Minimum requested size
            "footer": 12       # Footer can be smaller, but kept readable
        }
    }

class ColorUtils:
    @staticmethod
    def hex_to_rgb(hex_str):
        hex_str = hex_str.lstrip('#')
        return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

    @staticmethod
    def rgb_to_long(r, g, b):
        # VBA uses BGR format for colors in some contexts, but RGB function is usually R + G*256 + B*65536
        # However, standard RGB function in VBA: RGB(r, g, b)
        # We will generate the string "RGB(r, g, b)" for the macro.
        return f"RGB({r}, {g}, {b})"

    @staticmethod
    def lighten_color(hex_color, amount):
        # Simple lightening logic
        r, g, b = ColorUtils.hex_to_rgb(hex_color)
        r = int(min(255, r + (255 - r) * amount))
        g = int(min(255, g + (255 - g) * amount))
        b = int(min(255, b + (255 - b) * amount))
        return f"#{r:02x}{g:02x}{b:02x}"

    @staticmethod
    def generate_tinted_gray(base_hex, saturation, lightness):
        # Simplified version: mix base with gray
        # For VBA, we might just return a standard gray if this is too complex to replicate exactly without numpy/colorsys
        # Or just return a hex string.
        # Let's use a fixed light gray for now to match the "background_gray" feel
        return "#F8F9FA" 

    @staticmethod
    def generate_process_colors(base_hex, steps):
        colors = []
        for i in range(steps):
            denom = max(1, steps - 1)
            lighten_amount = 0.5 * (1 - (i / denom))
            colors.append(ColorUtils.lighten_color(base_hex, lighten_amount))
        return colors

    @staticmethod
    def generate_timeline_colors(base_hex, count):
        colors = []
        for i in range(count):
            denom = max(1, count - 1)
            lighten_amount = 0.6 * (1 - (i / denom))
            colors.append(ColorUtils.lighten_color(base_hex, lighten_amount))
        return colors

    @staticmethod
    def generate_cycle_colors(base_hex, count):
        colors = []
        for i in range(count):
            colors.append(base_hex) # Cycle usually uses same color or slight variation
        return colors

    @staticmethod
    def generate_pyramid_colors(base_hex, count):
        colors = []
        for i in range(count):
            denom = max(1, count - 1)
            lighten_amount = 0.6 * (i / denom) # Top is lighter or darker? GAS: top is base, bottom is lighter? Or reverse.
            # Let's assume top (i=0) is base, bottom is lighter.
            colors.append(ColorUtils.lighten_color(base_hex, lighten_amount))
        return colors
