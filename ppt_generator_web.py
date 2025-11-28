import json
import os
import google.generativeai as genai
from config import PPTConfig, ColorUtils

# Re-use the system prompt from the original file, or import it if it was in a separate module.
# Assuming prompts.py is in the parent or same directory. 
# We will copy the essential parts or import if possible.
try:
    from prompts import SYSTEM_PROMPT
except ImportError:
    # Fallback if running from a different context
    SYSTEM_PROMPT = """
    (Paste the full system prompt here if import fails, but for now we assume it exists or we pass it in)
    """

def generate_json_from_text(text_input, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.0-flash')
    
    try:
        response = model.generate_content(
            contents=[SYSTEM_PROMPT, f"Input Text:\n{text_input}"],
            generation_config={"response_mime_type": "application/json"}
        )
        text = response.text
        if "```json" in text:
            text = text.split("```json")[1].split("```")[0]
        elif "```" in text:
            text = text.split("```")[1].split("```")[0]
        return json.loads(text)
    except Exception as e:
        print(f"Error generating JSON: {e}")
        return None

def escape_vba(text):
    if not text:
        return ""
    return str(text).replace('"', '""').replace('\n', '" & vbCr & "')

def get_rgb_string(hex_color):
    try:
        r, g, b = ColorUtils.hex_to_rgb(hex_color)
        return f"RGB({r}, {g}, {b})"
    except:
        return "RGB(0, 0, 0)"

def json_to_vba(data, settings):
    """
    Converts slideData JSON to VBA with custom styling and A4 size.
    settings: dict with keys 'primary_color', 'font_family', 'logo_path' (optional)
    """
    if not data:
        return ""

    primary_color_rgb = get_rgb_string(settings.get('primary_color', '#4285F4'))
    title_color_rgb = get_rgb_string(settings.get('title_color', '#333333'))
    body_color_rgb = get_rgb_string(settings.get('body_color', '#333333'))
    font_family = settings.get('font_family', 'Meiryo')
    
    vba = []
    vba.append("Sub CreateCustomPresentation()")
    vba.append("    Dim pptApp As Object")
    vba.append("    Dim pptPres As Object")
    vba.append("    Dim pptSlide As Object")
    vba.append("    Dim pptShape As Object")
    vba.append("    Dim slideIndex As Integer")
    vba.append("")
    vba.append("    Set pptApp = CreateObject(\"PowerPoint.Application\")")
    vba.append("    pptApp.Visible = True")
    vba.append("    Set pptPres = pptApp.Presentations.Add")
    vba.append("")
    
    # 1. Set A4 Size
    vba.append(f"    ' Set to A4 Size")
    vba.append(f"    pptPres.PageSetup.SlideWidth = {PPTConfig.SLIDE_WIDTH_PT}")
    vba.append(f"    pptPres.PageSetup.SlideHeight = {PPTConfig.SLIDE_HEIGHT_PT}")
    vba.append("")

    for i, slide in enumerate(data):
        slide_type = slide.get("type", "content")
        title = escape_vba(slide.get("title", ""))
        subhead = escape_vba(slide.get("subhead", ""))
        notes = escape_vba(slide.get("notes", ""))
        
        vba.append(f"    ' === Slide {i+1}: {slide_type} ===")
        vba.append(f"    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12) ' 12 = ppLayoutBlank")
        
        # --- Common Elements ---
        vba.append("    pptSlide.FollowMasterBackground = msoFalse")
        vba.append("    pptSlide.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)")
        
        if slide_type == "title":
            pos = PPTConfig.POS_PX["titleSlide"]
            t_rect = pos["title"]
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(t_rect['left'])}, {PPTConfig.px_to_pt(t_rect['top'])}, {PPTConfig.px_to_pt(t_rect['width'])}, {PPTConfig.px_to_pt(t_rect['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{title}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['title']}")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {title_color_rgb}")
            vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 ' Center")
            vba.append(f"    pptShape.TextFrame2.AutoSize = 2 ' msoAutoSizeTextToFitShape")
            
            d_rect = pos["date"]
            date_str = escape_vba(slide.get("date", ""))
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(d_rect['left'])}, {PPTConfig.px_to_pt(d_rect['top'])}, {PPTConfig.px_to_pt(d_rect['width'])}, {PPTConfig.px_to_pt(d_rect['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{date_str}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['date']}")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
            
        elif slide_type == "section":
            pos = PPTConfig.POS_PX["sectionSlide"]
            g_rect = pos["ghostNum"]
            section_no = slide.get("sectionNo", i)
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(g_rect['left'])}, {PPTConfig.px_to_pt(g_rect['top'])}, {PPTConfig.px_to_pt(g_rect['width'])}, {PPTConfig.px_to_pt(g_rect['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{section_no}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 180")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(240, 240, 240)")
            
            t_rect = pos["title"]
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(t_rect['left'])}, {PPTConfig.px_to_pt(t_rect['top'])}, {PPTConfig.px_to_pt(t_rect['width'])}, {PPTConfig.px_to_pt(t_rect['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{title}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['sectionTitle']}")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {title_color_rgb}")
            vba.append(f"    pptShape.TextFrame2.AutoSize = 2 ' msoAutoSizeTextToFitShape")
            
        elif slide_type == "process":
            # --- Process Slide ---
            pos = PPTConfig.POS_PX["processSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            area = pos["area"]
            steps = slide.get("steps", [])[:4] # Max 4 steps
            if steps:
                n = len(steps)
                colors = ColorUtils.generate_process_colors(settings['primary_color'], n)
                
                # Dimensions
                box_h_px = 65 if n > 3 else (80 if n == 3 else 100)
                arrow_h_px = 15 if n > 3 else (20 if n == 3 else 25)
                font_size = 24 # Minimum 24pt
                
                box_h_pt = PPTConfig.px_to_pt(box_h_px)
                arrow_h_pt = PPTConfig.px_to_pt(arrow_h_px)
                header_w_pt = PPTConfig.px_to_pt(120)
                
                start_y = PPTConfig.px_to_pt(area['top'] + 10)
                current_y = start_y
                
                body_left = PPTConfig.px_to_pt(area['left']) + header_w_pt
                body_w_pt = PPTConfig.px_to_pt(area['width']) - header_w_pt
                
                for i, step in enumerate(steps):
                    # Header (Step N)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(area['left'])}, {current_y}, {header_w_pt}, {box_h_pt})") # msoShapeRectangle
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {colors[i]}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"STEP {i+1}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {font_size}")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2") # Center
                    
                    # Body
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {body_left}, {current_y}, {body_w_pt}, {box_h_pt})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {ColorUtils.generate_tinted_gray(settings['primary_color'], 10, 95)}") # Light gray
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    
                    # Text
                    text_shape_left = body_left + PPTConfig.px_to_pt(20)
                    text_shape_w = body_w_pt - PPTConfig.px_to_pt(40)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {text_shape_left}, {current_y}, {text_shape_w}, {box_h_pt})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(step)}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {font_size}")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
                    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")
                    
                    current_y += box_h_pt
                    
                    # Arrow
                    if i < n - 1:
                        arrow_left = PPTConfig.px_to_pt(area['left']) + header_w_pt / 2 - PPTConfig.px_to_pt(8)
                        vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(66, {arrow_left}, {current_y}, {PPTConfig.px_to_pt(16)}, {arrow_h_pt})") # msoShapeDownArrow
                        vba.append(f"    pptShape.Fill.ForeColor.RGB = {ColorUtils.generate_tinted_gray(settings['primary_color'], 38, 88)}") # Ghost gray
                        vba.append(f"    pptShape.Line.Visible = msoFalse")
                        current_y += arrow_h_pt

        elif slide_type == "timeline":
            # --- Timeline Slide ---
            pos = PPTConfig.POS_PX["timelineSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            area = pos["area"]
            milestones = slide.get("milestones", [])
            if milestones:
                n = len(milestones)
                colors = ColorUtils.generate_timeline_colors(settings['primary_color'], n)
                
                base_y = PPTConfig.px_to_pt(area['top'] + area['height'] * 0.5)
                inner_margin = PPTConfig.px_to_pt(80)
                left_x = PPTConfig.px_to_pt(area['left']) + inner_margin
                right_x = PPTConfig.px_to_pt(area['left'] + area['width']) - inner_margin
                
                # Main Line
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddLine({left_x}, {base_y}, {right_x}, {base_y})")
                vba.append(f"    pptShape.Line.ForeColor.RGB = RGB(200, 200, 200)")
                vba.append(f"    pptShape.Line.Weight = 2")
                
                gap = (right_x - left_x) / (n - 1) if n > 1 else 0
                card_w = PPTConfig.px_to_pt(180)
                v_offset = PPTConfig.px_to_pt(40)
                header_h = PPTConfig.px_to_pt(28)
                body_h = PPTConfig.px_to_pt(80)
                
                for i, m in enumerate(milestones):
                    x = left_x + gap * i
                    is_above = (i % 2 == 0)
                    
                    card_left = x - (card_w / 2)
                    card_top = (base_y - v_offset - header_h - body_h) if is_above else (base_y + v_offset)
                    
                    # Connector
                    conn_y1 = (card_top + header_h + body_h) if is_above else base_y
                    conn_y2 = base_y if is_above else card_top
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddLine({x}, {conn_y1}, {x}, {conn_y2})")
                    vba.append(f"    pptShape.Line.ForeColor.RGB = RGB(150, 150, 150)")
                    
                    # Dot
                    dot_r = PPTConfig.px_to_pt(10)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(9, {x - dot_r/2}, {base_y - dot_r/2}, {dot_r}, {dot_r})") # msoShapeOval
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(colors[i])}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    
                    # Card Header
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {card_left}, {card_top}, {card_w}, {header_h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(colors[i])}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(m.get('date', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    
                    # Card Body
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {card_left}, {card_top + header_h}, {card_w}, {body_h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 10, 95))}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(m.get('label', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 24")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")

        elif slide_type == "cycle":
            # --- Cycle Slide ---
            pos = PPTConfig.POS_PX["cycleSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            area = pos["body"]
            items = slide.get("items", [])[:4]
            if items:
                center_x = PPTConfig.px_to_pt(area['left'] + area['width'] / 2)
                center_y = PPTConfig.px_to_pt(area['top'] + area['height'] / 2)
                radius_x = PPTConfig.px_to_pt(area['width'] / 3.2)
                radius_y = PPTConfig.px_to_pt(area['height'] / 2.6)
                
                card_w = PPTConfig.px_to_pt(200)
                card_h = PPTConfig.px_to_pt(90)
                
                positions = [
                    (center_x + radius_x, center_y),
                    (center_x, center_y + radius_y),
                    (center_x - radius_x, center_y),
                    (center_x, center_y - radius_y)
                ]
                
                for i, item in enumerate(items):
                    if i >= 4: break
                    pos_x, pos_y = positions[i]
                    card_left = pos_x - card_w / 2
                    card_top = pos_y - card_h / 2
                    
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {card_left}, {card_top}, {card_w}, {card_h})") # msoShapeRoundedRectangle
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {primary_color_rgb}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    
                    label = item.get("label", "")
                    sub = item.get("subLabel", f"Phase {i+1}")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{sub}\" & vbCrLf & \"{label}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")

        elif slide_type == "cards":
            # --- Cards Slide ---
            pos = PPTConfig.POS_PX["cardsSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            area = pos["gridArea"]
            items = slide.get("items", [])
            if items:
                cols = 3 if len(items) > 4 else 2
                rows = (len(items) + cols - 1) // cols
                gap = PPTConfig.px_to_pt(16)
                
                area_w = PPTConfig.px_to_pt(area['width'])
                area_h = PPTConfig.px_to_pt(area['height'])
                card_w = (area_w - gap * (cols - 1)) / cols
                card_h = (area_h - gap * (rows - 1)) / rows
                
                for i, item in enumerate(items):
                    r = i // cols
                    c = i % cols
                    left = PPTConfig.px_to_pt(area['left']) + c * (card_w + gap)
                    top = PPTConfig.px_to_pt(area['top']) + r * (card_h + gap)
                    
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {left}, {top}, {card_w}, {card_h})") # msoShapeRoundedRectangle
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 10, 95))}")
                    vba.append(f"    pptShape.Line.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 15, 88))}")
                    
                    title = item.get("title", "")
                    desc = item.get("desc", "")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{title}\" & vbCrLf & vbCrLf & \"{desc}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")

        elif slide_type == "pyramid":
            # --- Pyramid Slide ---
            pos = PPTConfig.POS_PX["pyramidSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            area = pos["pyramidArea"]
            levels = slide.get("levels", [])[:4]
            if levels:
                n = len(levels)
                colors = ColorUtils.generate_pyramid_colors(settings['primary_color'], n)
                
                level_h = PPTConfig.px_to_pt(70)
                gap = PPTConfig.px_to_pt(2)
                total_h = (level_h * n) + (gap * (n - 1))
                
                start_y = PPTConfig.px_to_pt(area['top']) + (PPTConfig.px_to_pt(area['height']) - total_h) / 2
                pyramid_w = PPTConfig.px_to_pt(480)
                center_x = PPTConfig.px_to_pt(area['left']) + pyramid_w / 2
                
                text_col_left = PPTConfig.px_to_pt(area['left']) + pyramid_w + PPTConfig.px_to_pt(30)
                text_col_w = PPTConfig.px_to_pt(400)
                
                base_w = pyramid_w
                w_decrement = base_w / n
                
                for i, level in enumerate(levels):
                    level_w = base_w - (w_decrement * (n - 1 - i))
                    level_x = center_x - level_w / 2
                    level_y = start_y + i * (level_h + gap)
                    
                    # Pyramid Level
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {level_x}, {level_y}, {level_w}, {level_h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(colors[i])}")
                    vba.append(f"    pptShape.Line.Visible = msoFalse")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{level.get('title', '')}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    
                    # Description
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {text_col_left}, {level_y}, {text_col_w}, {level_h})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{level.get('description', '')}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 24")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
                    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")

        elif slide_type == "compare":
            # --- Compare Slide ---
            pos = PPTConfig.POS_PX["compareSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            l_rect = pos["leftBox"]
            r_rect = pos["rightBox"]
            
            # Left Box
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(l_rect['left'])}, {PPTConfig.px_to_pt(l_rect['top'])}, {PPTConfig.px_to_pt(l_rect['width'])}, {PPTConfig.px_to_pt(l_rect['height'])})")
            vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 10, 95))}")
            vba.append(f"    pptShape.Line.Visible = msoFalse")
            
            # Left Title
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(l_rect['left'])}, {PPTConfig.px_to_pt(l_rect['top'])}, {PPTConfig.px_to_pt(l_rect['width'])}, 40)")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{slide.get('leftTitle', 'Option A')}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
            vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
            
            # Left Items
            l_items = "\\r".join(slide.get("leftItems", []))
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(l_rect['left']) + 10}, {PPTConfig.px_to_pt(l_rect['top']) + 40}, {PPTConfig.px_to_pt(l_rect['width']) - 20}, {PPTConfig.px_to_pt(l_rect['height']) - 50})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(l_items.replace('\\r', '\n'))}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 24")
            
            # Right Box
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(r_rect['left'])}, {PPTConfig.px_to_pt(r_rect['top'])}, {PPTConfig.px_to_pt(r_rect['width'])}, {PPTConfig.px_to_pt(r_rect['height'])})")
            vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 5, 98))}")
            vba.append(f"    pptShape.Line.Visible = msoFalse")
            
            # Right Title
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(r_rect['left'])}, {PPTConfig.px_to_pt(r_rect['top'])}, {PPTConfig.px_to_pt(r_rect['width'])}, 40)")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{slide.get('rightTitle', 'Option B')}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
            vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
            
            # Right Items
            r_items = "\\r".join(slide.get("rightItems", []))
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(r_rect['left']) + 10}, {PPTConfig.px_to_pt(r_rect['top']) + 40}, {PPTConfig.px_to_pt(r_rect['width']) - 20}, {PPTConfig.px_to_pt(r_rect['height']) - 50})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(r_items.replace('\\r', '\n'))}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 24")

        elif slide_type == "diagram":
            pos = PPTConfig.POS_PX["diagramSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            shapes = slide.get("shapes", [])
            for shp in shapes:
                # Basic shape mapping
                st = shp.get("shapeType", "rect")
                mso_shape = 1 # msoShapeRectangle
                if st == "oval": mso_shape = 9 # msoShapeOval
                elif st == "rounded_rect": mso_shape = 5 # msoShapeRoundedRectangle
                
                x = PPTConfig.px_to_pt(shp.get("x", 100))
                y = PPTConfig.px_to_pt(shp.get("y", 100))
                w = PPTConfig.px_to_pt(shp.get("w", 100))
                h = PPTConfig.px_to_pt(shp.get("h", 50))
                
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape({mso_shape}, {x}, {y}, {w}, {h})")
                vba.append(f"    pptShape.Fill.ForeColor.RGB = {primary_color_rgb}")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(shp.get('label', ''))}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")

        elif slide_type == "flowChart":
            pos = PPTConfig.POS_PX["flowChartSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            flows = slide.get("flows", [])
            if flows:
                steps = flows[0].get("steps", [])
                n = len(steps)
                if n > 0:
                    area = pos["area"]
                    box_w = PPTConfig.px_to_pt(150)
                    box_h = PPTConfig.px_to_pt(60)
                    gap = PPTConfig.px_to_pt(30)
                    
                    # Center the flow chart
                    total_w = n * box_w + (n - 1) * gap
                    center_x = PPTConfig.px_to_pt(area['left'] + area['width'] / 2)
                    start_x = center_x - total_w / 2
                    
                    start_y = PPTConfig.px_to_pt(area['top']) + PPTConfig.px_to_pt(50)
                    
                    for i, step in enumerate(steps):
                        x = start_x + i * (box_w + gap)
                        vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {x}, {start_y}, {box_w}, {box_h})")
                        vba.append(f"    pptShape.Fill.ForeColor.RGB = {primary_color_rgb}")
                        vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(step)}\"")
                        vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                        vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                        
                        if i < n - 1:
                            arrow_x = x + box_w
                            arrow_y = start_y + box_h / 2 - PPTConfig.px_to_pt(5)
                            vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(33, {arrow_x}, {arrow_y}, {gap}, {PPTConfig.px_to_pt(10)})") # msoShapeRightArrow
                            vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.generate_tinted_gray(settings['primary_color'], 20, 80))}")

        elif slide_type == "stepUp":
            pos = PPTConfig.POS_PX["stepUpSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            steps = slide.get("steps", [])
            n = len(steps)
            if n > 0:
                area = pos["area"]
                step_w = PPTConfig.px_to_pt(area['width']) / n
                step_h = PPTConfig.px_to_pt(50)
                base_y = PPTConfig.px_to_pt(area['top']) + PPTConfig.px_to_pt(area['height'])
                
                for i, step in enumerate(steps):
                    h = (i + 1) * step_h
                    x = PPTConfig.px_to_pt(area['left']) + i * step_w
                    y = base_y - h
                    
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {x}, {y}, {step_w}, {h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {get_rgb_string(ColorUtils.lighten_color(settings['primary_color'], 0.1 * i))}")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(step.get('label', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")

        elif slide_type == "imageText":
            pos = PPTConfig.POS_PX["imageTextSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            # Image Placeholder (Left)
            ia = pos["imageArea"]
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(ia['left'])}, {PPTConfig.px_to_pt(ia['top'])}, {PPTConfig.px_to_pt(ia['width'])}, {PPTConfig.px_to_pt(ia['height'])})")
            vba.append(f"    pptShape.Fill.ForeColor.RGB = RGB(230, 230, 230)")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"[IMAGE: {escape_vba(slide.get('imageDesc', ''))}]\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            
            # Text (Right)
            ta = pos["textArea"]
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(ta['left'])}, {PPTConfig.px_to_pt(ta['top'])}, {PPTConfig.px_to_pt(ta['width'])}, {PPTConfig.px_to_pt(ta['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(slide.get('text', ''))}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['body']}")

        elif slide_type == "table":
            pos = PPTConfig.POS_PX["tableSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            headers = slide.get("headers", [])
            rows = slide.get("rows", [])
            if headers:
                num_rows = len(rows) + 1
                num_cols = len(headers)
                area = pos["tableArea"]
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTable({num_rows}, {num_cols}, {PPTConfig.px_to_pt(area['left'])}, {PPTConfig.px_to_pt(area['top'])}, {PPTConfig.px_to_pt(area['width'])}, {PPTConfig.px_to_pt(200)})")
                
                # Headers
                for c, h in enumerate(headers):
                    vba.append(f"    pptShape.Table.Cell(1, {c+1}).Shape.TextFrame.TextRange.Text = \"{escape_vba(h)}\"")
                    vba.append(f"    pptShape.Table.Cell(1, {c+1}).Shape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.Table.Cell(1, {c+1}).Shape.Fill.ForeColor.RGB = {primary_color_rgb}")
                    vba.append(f"    pptShape.Table.Cell(1, {c+1}).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)")
                
                # Rows
                for r, row in enumerate(rows):
                    for c, cell in enumerate(row):
                        if c < num_cols:
                            vba.append(f"    pptShape.Table.Cell({r+2}, {c+1}).Shape.TextFrame.TextRange.Text = \"{escape_vba(cell)}\"")
                            vba.append(f"    pptShape.Table.Cell({r+2}, {c+1}).Shape.TextFrame.TextRange.Font.Name = \"{font_family}\"")

        elif slide_type == "progress":
            pos = PPTConfig.POS_PX["progressSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            items = slide.get("items", [])
            area = pos["area"]
            bar_h = PPTConfig.px_to_pt(30)
            gap = PPTConfig.px_to_pt(20)
            start_y = PPTConfig.px_to_pt(area['top'])
            
            for i, item in enumerate(items):
                y = start_y + i * (bar_h + gap + 30) # +30 for label
                
                # Label
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(area['left'])}, {y}, {PPTConfig.px_to_pt(300)}, 20)")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(item.get('label', ''))}\"")
                
                # Track
                y_bar = y + 25
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {PPTConfig.px_to_pt(area['left'])}, {y_bar}, {PPTConfig.px_to_pt(area['width'])}, {bar_h})")
                vba.append(f"    pptShape.Fill.ForeColor.RGB = RGB(230, 230, 230)")
                
                # Fill
                pct = item.get("percent", 0)
                fill_w = PPTConfig.px_to_pt(area['width']) * (pct / 100.0)
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {PPTConfig.px_to_pt(area['left'])}, {y_bar}, {fill_w}, {bar_h})")
                vba.append(f"    pptShape.Fill.ForeColor.RGB = {primary_color_rgb}")

        elif slide_type == "quote":
            pos = PPTConfig.POS_PX["quoteSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            qa = pos["quoteArea"]
            aa = pos["authorArea"]
            
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(qa['left'])}, {PPTConfig.px_to_pt(qa['top'])}, {PPTConfig.px_to_pt(qa['width'])}, {PPTConfig.px_to_pt(qa['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"“{escape_vba(slide.get('quote', ''))}”\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 32")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Italic = msoTrue")
            vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2") # Center
            
            vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(aa['left'])}, {PPTConfig.px_to_pt(aa['top'])}, {PPTConfig.px_to_pt(aa['width'])}, {PPTConfig.px_to_pt(aa['height'])})")
            vba.append(f"    pptShape.TextFrame.TextRange.Text = \"— {escape_vba(slide.get('author', ''))}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
            vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 3") # Right

        elif slide_type == "kpi":
            pos = PPTConfig.POS_PX["kpiSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            kpis = slide.get("kpis", [])
            if kpis:
                area = pos["area"]
                cols = 3
                rows = (len(kpis) + cols - 1) // cols
                gap = PPTConfig.px_to_pt(20)
                w = (PPTConfig.px_to_pt(area['width']) - gap * (cols - 1)) / cols
                h = PPTConfig.px_to_pt(150)
                
                for i, kpi in enumerate(kpis):
                    r = i // cols
                    c = i % cols
                    x = PPTConfig.px_to_pt(area['left']) + c * (w + gap)
                    y = PPTConfig.px_to_pt(area['top']) + r * (h + gap)
                    
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(5, {x}, {y}, {w}, {h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = RGB(245, 245, 245)")
                    
                    # Value
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {x}, {y + 10}, {w}, {h/2})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(kpi.get('value', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = 36")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {primary_color_rgb}")
                    
                    # Label
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {x}, {y + h/2}, {w}, {h/2})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(kpi.get('label', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")

        elif slide_type == "bulletCards":
            pos = PPTConfig.POS_PX["bulletCardsSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            cards = slide.get("cards", [])
            if cards:
                area = pos["area"]
                cols = 2
                gap = PPTConfig.px_to_pt(20)
                w = (PPTConfig.px_to_pt(area['width']) - gap * (cols - 1)) / cols
                h = PPTConfig.px_to_pt(300)
                
                for i, card in enumerate(cards):
                    if i >= 2: break # Limit to 2 for simplicity
                    x = PPTConfig.px_to_pt(area['left']) + i * (w + gap)
                    y = PPTConfig.px_to_pt(area['top'])
                    
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {x}, {y}, {w}, {h})")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = RGB(250, 250, 250)")
                    vba.append(f"    pptShape.Line.ForeColor.RGB = {primary_color_rgb}")
                    
                    # Title
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {x + 10}, {y + 10}, {w - 20}, 40)")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(card.get('title', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    
                    # Points
                    points = "\\r".join(["・" + p for p in card.get("points", [])])
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {x + 10}, {y + 50}, {w - 20}, {h - 60})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(points.replace('\\r', '\n'))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")

        elif slide_type == "faq":
            pos = PPTConfig.POS_PX["faqSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            items = slide.get("items", [])
            area = pos["area"]
            y = PPTConfig.px_to_pt(area['top'])
            w = PPTConfig.px_to_pt(area['width'])
            
            for item in items:
                # Q
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(area['left'])}, {y}, {w}, 30)")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"Q. {escape_vba(item.get('q', ''))}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {primary_color_rgb}")
                y += 30
                
                # A
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(area['left'])}, {y}, {w}, 40)")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"A. {escape_vba(item.get('a', ''))}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                y += 50

        elif slide_type == "statsCompare":
            pos = PPTConfig.POS_PX["statsCompareSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            stats = slide.get("stats", [])
            if stats:
                lb = pos["leftBox"]
                rb = pos["rightBox"]
                
                # Titles
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(lb['left'])}, {PPTConfig.px_to_pt(lb['top']) - 30}, {PPTConfig.px_to_pt(lb['width'])}, 30)")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(slide.get('leftTitle', ''))}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(rb['left'])}, {PPTConfig.px_to_pt(rb['top']) - 30}, {PPTConfig.px_to_pt(rb['width'])}, 30)")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(slide.get('rightTitle', ''))}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                
                y = PPTConfig.px_to_pt(lb['top'])
                h = PPTConfig.px_to_pt(50)
                
                for stat in stats:
                    # Label (Center)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(460)}, {y}, {PPTConfig.px_to_pt(200)}, {h})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(stat.get('label', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2")
                    
                    # Left Value
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(lb['left'])}, {y}, {PPTConfig.px_to_pt(lb['width'])}, {h})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(stat.get('leftValue', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 3") # Right
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    
                    # Right Value
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(rb['left'])}, {y}, {PPTConfig.px_to_pt(rb['width'])}, {h})")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(stat.get('rightValue', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 1") # Left
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
                    
                    y += h + 10

        elif slide_type == "barCompare":
            pos = PPTConfig.POS_PX["barCompareSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            items = slide.get("items", [])
            if items:
                area = pos["area"]
                y = PPTConfig.px_to_pt(area['top'])
                max_val = 100 # Assumed max
                w_base = PPTConfig.px_to_pt(300)
                
                for item in items:
                    valA = item.get("valueA", 0)
                    valB = item.get("valueB", 0)
                    
                    # Label
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(area['left'])}, {y}, {PPTConfig.px_to_pt(area['width'])}, 20)")
                    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{escape_vba(item.get('label', ''))}\"")
                    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                    y += 25
                    
                    # Bar A
                    wa = w_base * (valA / max_val)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(area['left'])}, {y}, {wa}, 20)")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = {primary_color_rgb}")
                    
                    # Bar B
                    wb = w_base * (valB / max_val)
                    vba.append(f"    Set pptShape = pptSlide.Shapes.AddShape(1, {PPTConfig.px_to_pt(area['left'])}, {y + 25}, {wb}, 20)")
                    vba.append(f"    pptShape.Fill.ForeColor.RGB = RGB(150, 150, 150)")
                    
                    y += 60

        else:
            # --- Standard Content Slide (Fallback) ---
            pos = PPTConfig.POS_PX["contentSlide"]
            draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type)
            
            b_rect = pos["body"]
            content_text = ""
            if "points" in slide:
                # Manual bullets
                content_text = "・" + "\\r・".join(slide["points"])
            elif "items" in slide:
                items = slide["items"]
                if items and isinstance(items[0], str):
                    content_text = "・" + "\\r・".join(items)
                elif items and isinstance(items[0], dict):
                     content_text = "・" + "\\r・".join([f"{item.get('title','')}: {item.get('desc','')}" for item in items])
            elif "steps" in slide:
                content_text = "・" + "\\r・".join(slide["steps"])
            
            if content_text:
                content_text = escape_vba(content_text.replace("\\r", "\n"))
                vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(b_rect['left'])}, {PPTConfig.px_to_pt(b_rect['top'])}, {PPTConfig.px_to_pt(b_rect['width'])}, {PPTConfig.px_to_pt(b_rect['height'])})")
                vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{content_text}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['body']}")
                vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {body_color_rgb}")
                # Removed automatic bullet type assignment

    vba.append("    MsgBox \"Presentation Created!\", vbInformation")
    vba.append("End Sub")
    
    return "\n".join(vba)

def draw_common_header(vba, slide, pos, font_family, title_color_rgb, primary_color_rgb, slide_type):
    title = escape_vba(slide.get("title", ""))
    subhead = escape_vba(slide.get("subhead", ""))
    
    t_rect = pos["title"]
    vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(t_rect['left'])}, {PPTConfig.px_to_pt(t_rect['top'])}, {PPTConfig.px_to_pt(t_rect['width'])}, {PPTConfig.px_to_pt(t_rect['height'])})")
    vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{title}\"")
    vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
    vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['contentTitle']}")
    vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
    vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {primary_color_rgb}")
    vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2") # Center
    vba.append(f"    pptShape.TextFrame2.AutoSize = 2")
    
    u_rect = pos["titleUnderline"]
    line_start_x = PPTConfig.px_to_pt(u_rect['left'])
    line_start_y = PPTConfig.px_to_pt(u_rect['top'])
    line_end_x = line_start_x + PPTConfig.px_to_pt(u_rect['width'])
    vba.append(f"    Set pptShape = pptSlide.Shapes.AddLine({line_start_x}, {line_start_y}, {line_end_x}, {line_start_y})")
    vba.append(f"    pptShape.Line.ForeColor.RGB = {primary_color_rgb}")
    vba.append(f"    pptShape.Line.Weight = 2")

    if subhead:
        s_rect = pos["subhead"]
        vba.append(f"    Set pptShape = pptSlide.Shapes.AddTextbox(1, {PPTConfig.px_to_pt(s_rect['left'])}, {PPTConfig.px_to_pt(s_rect['top'])}, {PPTConfig.px_to_pt(s_rect['width'])}, {PPTConfig.px_to_pt(s_rect['height'])})")
        vba.append(f"    pptShape.TextFrame.TextRange.Text = \"{subhead}\"")
        vba.append(f"    pptShape.TextFrame.TextRange.Font.Name = \"{font_family}\"")
        vba.append(f"    pptShape.TextFrame.TextRange.Font.Size = {PPTConfig.FONTS['sizes']['subhead']}")
        vba.append(f"    pptShape.TextFrame.TextRange.Font.Bold = msoTrue")
        vba.append(f"    pptShape.TextFrame.TextRange.Font.Color.RGB = {title_color_rgb}")
        vba.append(f"    pptShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2") # Center

    
    # Notes
    notes = escape_vba(slide.get("notes", ""))
    if notes:
         vba.append(f"    pptSlide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = \"{notes}\"")

    vba.append("")


