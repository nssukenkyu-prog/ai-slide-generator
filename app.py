from flask import Flask, render_template, request, send_file, Response
import os
import sys

# Add parent directory to path to import prompts if needed, 
# but we are self-contained in ppt_web_app for now except for prompts.py
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'ppt_gen'))
# Also add current dir
sys.path.append(os.path.dirname(__file__))

from ppt_generator_web import generate_json_from_text, json_to_vba
try:
    from prompts import SYSTEM_PROMPT
except ImportError:
    # If prompts.py is in ../ppt_gen/prompts.py, we need to make sure we can import it
    # We added the path above, so it should work if the file exists.
    pass

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    text_input = request.form.get('text_input')
    api_key = request.form.get('api_key') or os.environ.get("GOOGLE_API_KEY")
    
    # Capture settings to pass through
    settings = {
        'primary_color': request.form.get('primary_color', '#4285F4'),
        'title_color': request.form.get('title_color', '#333333'),
        'body_color': request.form.get('body_color', '#333333'),
        'font_family': request.form.get('font_family', 'Meiryo')
    }

    if not text_input:
        return "Error: No text input provided.", 400
    if not api_key:
        return "Error: API Key is required.", 400

    # Generate JSON
    slide_data = generate_json_from_text(text_input, api_key)
    
    if not slide_data:
        return "Error: Failed to generate slide data from AI.", 500

    # Pre-process slides for the editor (flatten lists to strings)
    for slide in slide_data:
        content_parts = []
        if "points" in slide:
            content_parts = slide["points"]
        elif "items" in slide:
            items = slide["items"]
            if items and isinstance(items[0], str):
                content_parts = items
            elif items and isinstance(items[0], dict):
                content_parts = [f"{item.get('title','')}: {item.get('desc','')}" if 'title' in item else f"{item.get('label','')}: {item.get('subLabel','')}" for item in items]
        elif "milestones" in slide:
            content_parts = [f"{m.get('date','')}: {m.get('label','')}" for m in slide["milestones"]]
        elif "levels" in slide:
            content_parts = [f"{l.get('title','')}: {l.get('description','')}" for l in slide["levels"]]
        elif "leftItems" in slide or "rightItems" in slide:
            content_parts.append("--- Left ---")
            content_parts.extend(slide.get("leftItems", []))
            content_parts.append("--- Right ---")
            content_parts.extend(slide.get("rightItems", []))
        elif "shapes" in slide:
            content_parts = [f"{s.get('label','')}" for s in slide["shapes"]]
        elif "flows" in slide:
            for flow in slide["flows"]:
                content_parts.extend(flow.get("steps", []))
        elif "imageDesc" in slide or "text" in slide:
            content_parts.append(f"Image: {slide.get('imageDesc','')}")
            content_parts.append(f"Text: {slide.get('text','')}")
        elif "headers" in slide:
            content_parts.append(" | ".join(slide.get("headers", [])))
            for row in slide.get("rows", []):
                content_parts.append(" | ".join(row))
        elif "quote" in slide:
            content_parts.append(f"Quote: {slide.get('quote','')}")
            content_parts.append(f"Author: {slide.get('author','')}")
        elif "kpis" in slide:
            content_parts = [f"{k.get('label','')}: {k.get('value','')} ({k.get('change','')})" for k in slide["kpis"]]
        elif "cards" in slide: # bulletCards
            for c in slide["cards"]:
                content_parts.append(f"Title: {c.get('title','')}")
                for p in c.get("points", []):
                    content_parts.append(f"- {p}")
                content_parts.append("---")
        elif "stats" in slide: # statsCompare
            content_parts = [f"{s.get('label','')}: {s.get('leftValue','')} / {s.get('rightValue','')}" for s in slide["stats"]]
        elif "items" in slide: # Generic items fallback (faq, barCompare, progress, etc)
            items = slide["items"]
            if items and isinstance(items[0], dict):
                if "q" in items[0]: # FAQ
                    content_parts = [f"Q: {i.get('q','')}\nA: {i.get('a','')}" for i in items]
                elif "valueA" in items[0]: # barCompare
                    content_parts = [f"{i.get('label','')}: {i.get('valueA','')} / {i.get('valueB','')}" for i in items]
                elif "percent" in items[0]: # progress
                    content_parts = [f"{i.get('label','')}: {i.get('percent','')}%" for i in items]
                else:
                    # Fallback for other dict items
                    content_parts = [f"{item.get('title','')}: {item.get('desc','')}" if 'title' in item else f"{item.get('label','')}: {item.get('subLabel','')}" for item in items]
        
        # Join with newlines for textarea
        slide['content_text'] = "\n".join(content_parts)

    return render_template('edit.html', slides=slide_data, settings=settings)

@app.route('/download', methods=['POST'])
def download():
    # Reconstruct slide_data from form
    print("DEBUG: Form Data:", request.form)
    slide_count = int(request.form.get('slide_count', 0))
    print(f"DEBUG: slide_count = {slide_count}")
    slide_data = []
    
    for i in range(slide_count):
        slide = {}
        slide_type = request.form.get(f'slide_{i}_type')
        slide['type'] = slide_type
        slide['title'] = request.form.get(f'slide_{i}_title')
        slide['subhead'] = request.form.get(f'slide_{i}_subhead')
        
        # Reconstruct content list
        content_text = request.form.get(f'slide_{i}_content', '')
        content_list = [line.strip() for line in content_text.split('\n') if line.strip()]
        
        # Assign back to the appropriate key based on type
        if slide_type == 'process':
            slide['steps'] = content_list
        elif slide_type == 'timeline':
            milestones = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    milestones.append({"date": parts[0].strip(), "label": parts[1].strip()})
                else:
                    milestones.append({"date": "", "label": line})
            slide['milestones'] = milestones
        elif slide_type == 'cycle':
            items = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    items.append({"label": parts[1].strip(), "subLabel": parts[0].strip()})
                else:
                    items.append({"label": line, "subLabel": ""})
            slide['items'] = items
        elif slide_type == 'cards':
            items = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    items.append({"title": parts[0].strip(), "desc": parts[1].strip()})
                else:
                    items.append({"title": line, "desc": ""})
            slide['items'] = items
        elif slide_type == 'pyramid':
            levels = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    levels.append({"title": parts[0].strip(), "description": parts[1].strip()})
                else:
                    levels.append({"title": line, "description": ""})
            slide['levels'] = levels
        elif slide_type == 'compare':
            left_items = []
            right_items = []
            current_list = left_items
            for line in content_list:
                if "--- Left ---" in line:
                    current_list = left_items
                    continue
                elif "--- Right ---" in line:
                    current_list = right_items
                    continue
                current_list.append(line)
            slide['leftItems'] = left_items
            slide['rightItems'] = right_items
        elif slide_type == 'diagram':
            slide['shapes'] = [{"label": line} for line in content_list]
        elif slide_type == 'flowChart':
            slide['flows'] = [{"steps": content_list}]
        elif slide_type == 'stepUp':
            slide['steps'] = [{"label": line} for line in content_list]
        elif slide_type == 'imageText':
            slide['imageDesc'] = ""
            slide['text'] = ""
            for line in content_list:
                if line.startswith("Image:"): slide['imageDesc'] = line.replace("Image:", "").strip()
                elif line.startswith("Text:"): slide['text'] = line.replace("Text:", "").strip()
                else: slide['text'] += "\n" + line
        elif slide_type == 'table':
            if content_list:
                slide['headers'] = [c.strip() for c in content_list[0].split('|')]
                slide['rows'] = [[c.strip() for c in row.split('|')] for row in content_list[1:]]
        elif slide_type == 'quote':
            for line in content_list:
                if line.startswith("Quote:"): slide['quote'] = line.replace("Quote:", "").strip()
                elif line.startswith("Author:"): slide['author'] = line.replace("Author:", "").strip()
        elif slide_type == 'kpi':
            kpis = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    val_change = parts[1].strip().split('(')
                    val = val_change[0].strip()
                    change = val_change[1].replace(')', '').strip() if len(val_change) > 1 else ""
                    kpis.append({"label": parts[0].strip(), "value": val, "change": change})
                else:
                    kpis.append({"label": line, "value": "", "change": ""})
            slide['kpis'] = kpis
        elif slide_type == 'bulletCards':
            cards = []
            current_card = None
            for line in content_list:
                if line.startswith("Title:"):
                    if current_card: cards.append(current_card)
                    current_card = {"title": line.replace("Title:", "").strip(), "points": []}
                elif line.startswith("-"):
                    if current_card: current_card["points"].append(line.replace("-", "").strip())
                elif line == "---":
                    if current_card: 
                        cards.append(current_card)
                        current_card = None
            if current_card: cards.append(current_card)
            slide['cards'] = cards
        elif slide_type == 'faq':
            items = []
            current_q = None
            for line in content_list:
                if line.startswith("Q:"):
                    if current_q: items.append(current_q)
                    current_q = {"q": line.replace("Q:", "").strip(), "a": ""}
                elif line.startswith("A:"):
                    if current_q: current_q["a"] = line.replace("A:", "").strip()
            if current_q: items.append(current_q)
            slide['items'] = items
        elif slide_type == 'statsCompare':
            stats = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    vals = parts[1].strip().split('/')
                    l = vals[0].strip()
                    r = vals[1].strip() if len(vals) > 1 else ""
                    stats.append({"label": parts[0].strip(), "leftValue": l, "rightValue": r})
            slide['stats'] = stats
        elif slide_type == 'barCompare':
            items = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    vals = parts[1].strip().split('/')
                    vA = float(vals[0].strip()) if vals[0].strip().replace('.','').isdigit() else 0
                    vB = float(vals[1].strip()) if len(vals) > 1 and vals[1].strip().replace('.','').isdigit() else 0
                    items.append({"label": parts[0].strip(), "valueA": vA, "valueB": vB})
            slide['items'] = items
        elif slide_type == 'progress':
            items = []
            for line in content_list:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    pct = float(parts[1].replace('%','').strip()) if parts[1].replace('%','').strip().replace('.','').isdigit() else 0
                    items.append({"label": parts[0].strip(), "percent": pct})
            slide['items'] = items
        else:
            slide['points'] = content_list
            slide['items'] = content_list # Fallback
        
        # Pass through other fields
        if request.form.get(f'slide_{i}_sectionNo'):
             slide['sectionNo'] = request.form.get(f'slide_{i}_sectionNo')
             
        slide_data.append(slide)

    settings = {
        'primary_color': request.form.get('primary_color'),
        'title_color': request.form.get('title_color'),
        'body_color': request.form.get('body_color'),
        'font_family': request.form.get('font_family')
    }
    
    print(f"DEBUG: slide_data length: {len(slide_data)}")
    # print(f"DEBUG: slide_data[0]: {slide_data[0] if slide_data else 'Empty'}")
    
    vba_code = json_to_vba(slide_data, settings)
    print(f"DEBUG: vba_code length: {len(vba_code)}")
    
    return Response(
        vba_code,
        mimetype="text/plain",
        headers={"Content-disposition": "attachment; filename=presentation_macro.vba"}
    )

if __name__ == '__main__':
    app.run(debug=True, port=5000)
