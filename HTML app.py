import streamlit as st
import json
import re
from datetime import datetime

# ---- Optional document imports ----
try:
    from docx import Document
except ImportError:
    Document = None
try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

# ===================== STYLE DEFINITIONS =====================
CARD_STYLES = {
    "Elevated Shadow": "bg-white rounded-2xl shadow-lg p-6 hover:shadow-xl transition-shadow",
    "Bordered Accent": "bg-white rounded-xl border-l-4 border-{color}-500 p-6 shadow-sm",
    "Flat Minimal": "bg-gray-50 rounded-lg p-6 border border-gray-200",
    "Dark Card": "bg-gray-800 text-white rounded-2xl p-6 shadow-md"
}

TABLE_STYLES = {
    "Striped Rows": "striped",
    "Dark Header": "dark-header",
    "Bordered": "bordered",
    "Minimal": "minimal"
}

COLOR_THEMES = {
    "Blue (Default)": "indigo",
    "Green": "emerald",
    "Purple": "violet",
    "Rose": "rose",
    "Amber": "amber",
    "Slate": "slate"
}

FONTS = ["Inter", "System UI", "Serif"]

# ===================== UTILS =====================
def read_text_file(uploaded_file):
    return uploaded_file.getvalue().decode("utf-8")

def read_docx_file(uploaded_file):
    if Document is None:
        raise ImportError("python-docx not installed")
    doc = Document(uploaded_file)
    return "\n".join([p.text for p in doc.paragraphs])

def read_pdf_file(uploaded_file):
    if PdfReader is None:
        raise ImportError("PyPDF2 not installed")
    reader = PdfReader(uploaded_file)
    text = ""
    for page in reader.pages:
        t = page.extract_text()
        if t:
            text += t + "\n"
    return text.strip()

def get_color(color_name):
    return COLOR_THEMES.get(color_name, "indigo")

def build_card_html(cards_list, card_style, color_theme):
    """Builds a grid of cards from a list of dicts {title, items, image}."""
    color = get_color(color_theme)
    style_template = CARD_STYLES[card_style]
    if "{color}" in style_template:
        style_class = style_template.replace("{color}", color)
    else:
        style_class = style_template

    cards_html = []
    for card in cards_list:
        items = "".join([f"<li class='text-sm'>{item}</li>" for item in card.get("items", [])])
        image_tag = ""
        if card.get("image"):
            image_tag = f'<img src="{card["image"]}" class="w-full h-32 object-cover rounded-lg mb-3" />'
        cards_html.append(f"""
        <div class="{style_class} flex flex-col">
            {image_tag}
            <h4 class="font-bold text-lg mb-2">{card['title']}</h4>
            <ul class="list-disc list-inside space-y-1">{items}</ul>
        </div>
        """)
    return f'<div class="grid md:grid-cols-2 gap-6 my-6">{"".join(cards_html)}</div>'

def build_table_html(rows, table_style, color_theme):
    """rows: list of lists, first row is header."""
    color = get_color(color_theme)
    header_html = "".join([f"<th class='px-4 py-3'>{cell}</th>" for cell in rows[0]])
    body_rows = ""
    for i, row in enumerate(rows[1:]):
        row_class = ""
        if table_style == "Striped Rows" and i % 2 == 1:
            row_class = "bg-gray-50"
        cells = "".join([f"<td class='px-4 py-3 border-b'>{cell}</td>" for cell in row])
        body_rows += f"<tr class='{row_class}'>{cells}</tr>"

    # Header styling
    if table_style == "Dark Header":
        thead_style = f"bg-{color}-600 text-white"
    elif table_style == "Bordered":
        thead_style = f"bg-{color}-100 border-b-2 border-{color}-500"
    else:  # minimal / striped
        thead_style = f"bg-{color}-50 text-{color}-800"

    table_html = f"""<div class="overflow-x-auto my-6 rounded-xl shadow">
    <table class="w-full bg-white">
        <thead><tr class="{thead_style}">{header_html}</tr></thead>
        <tbody>{body_rows}</tbody>
    </table>
</div>"""
    return table_html

def parse_advanced_plain_text(raw_text, card_style, table_style, color_theme):
    """
    Parses plain text with markers:
      ## Cards: title
      **Card Title** + - items
      | table |
      # heading, > quote, etc.
    """
    lines = raw_text.splitlines()
    output = []
    i = 0
    in_card_section = False
    card_buffer = []  # list of dicts
    current_card = None
    color = get_color(color_theme)

    def flush_cards():
        nonlocal card_buffer
        if card_buffer:
            output.append(build_card_html(card_buffer, card_style, color_theme))
            card_buffer = []

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            flush_cards()
            in_card_section = False
            output.append("<br>")
            i += 1
            continue

        # Headings
        if line.startswith("# "):
            flush_cards()
            output.append(f'<h1 class="text-3xl font-extrabold mt-8 mb-4 text-gray-900">{line[2:]}</h1>')
            i += 1
            continue
        if line.startswith("## "):
            flush_cards()
            title = line[3:]
            if title.lower().startswith("cards:"):
                in_card_section = True
                output.append(f'<h3 class="text-2xl font-bold mt-8 mb-4 text-{color}-600">{title[6:].strip()}</h3>')
            else:
                output.append(f'<h2 class="text-2xl font-bold mt-8 mb-3 text-gray-900">{title}</h2>')
            i += 1
            continue
        if line.startswith("### "):
            flush_cards()
            output.append(f'<h3 class="text-xl font-semibold mt-6 mb-2 text-gray-800">{line[4:]}</h3>')
            i += 1
            continue

        # Card section logic
        if in_card_section and line.startswith("**") and line.endswith("**"):
            if current_card:
                card_buffer.append(current_card)
            current_card = {"title": line[2:-2], "items": []}
            i += 1
            continue
        if in_card_section and (line.startswith("- ") or line.startswith("* ")):
            if current_card is not None:
                current_card["items"].append(line[2:])
            i += 1
            continue
        # If we are in card section but line doesn't match card syntax, close
        if in_card_section and current_card:
            card_buffer.append(current_card)
            current_card = None
            in_card_section = False  # exit card mode

        # Table detection (pipe)
        if "|" in line and line.count("|") >= 2:
            flush_cards()
            table_rows = [line]
            i += 1
            while i < len(lines) and "|" in lines[i] and lines[i].strip().count("|") >= 2:
                table_rows.append(lines[i].strip())
                i += 1
            # parse
            parsed_rows = []
            for r in table_rows:
                cells = [c.strip() for c in r.split("|") if c.strip()]
                parsed_rows.append(cells)
            output.append(build_table_html(parsed_rows, table_style, color_theme))
            continue

        # CSV table (comma separated)
        if line.count(",") >= 2 and i+1 < len(lines) and lines[i+1].strip().count(",") >= 2:
            flush_cards()
            table_rows = [line]
            i += 1
            while i < len(lines) and lines[i].strip().count(",") >= 2:
                table_rows.append(lines[i].strip())
                i += 1
            parsed_rows = [r.split(",") for r in table_rows]
            output.append(build_table_html(parsed_rows, table_style, color_theme))
            continue

        # Blockquote
        if line.startswith("> "):
            flush_cards()
            output.append(f'<blockquote class="border-l-4 border-{color}-300 pl-4 italic text-gray-600 my-4">{line[2:]}</blockquote>')
            i += 1
            continue

        # Bullet list
        if re.match(r"^[-*]\s", line):
            flush_cards()
            output.append(f'<ul class="list-disc list-inside ml-4 space-y-1 mb-4"><li>{line[2:]}</li></ul>')
            i += 1
            continue

        # Numbered list
        if re.match(r"^\d+\.\s", line):
            flush_cards()
            content = re.sub(r"^\d+\.\s", "", line)
            output.append(f'<ol class="list-decimal list-inside ml-4 space-y-1 mb-4"><li>{content}</li></ol>')
            i += 1
            continue

        # Normal paragraph
        flush_cards()
        output.append(f'<p class="mb-4 text-gray-700 leading-relaxed">{line}</p>')
        i += 1

    # flush remaining
    if current_card:
        card_buffer.append(current_card)
    flush_cards()
    return "\n".join(output)

# ===================== FINAL HTML GENERATION =====================
def generate_html(content, is_json, meta, style_settings):
    card_style = style_settings["card_style"]
    table_style = style_settings["table_style"]
    color_theme = style_settings["color_theme"]
    font = style_settings["font"]

    color = get_color(color_theme)
    hero_bg = f"linear-gradient(135deg, #1e293b, #{color})"  # simple gradient
    if color_theme == "Slate":
        hero_bg = "#f8fafc"
    font_family = {
        "Inter": "'Inter', sans-serif",
        "System UI": "system-ui, -apple-system, sans-serif",
        "Serif": "'Georgia', serif"
    }[font]

    title = meta.get("title", "Report")
    subtitle = meta.get("subtitle", "")
    badge = meta.get("badge", "Report")

    hero_text = "text-white" if color_theme != "Slate" else "text-gray-900"

    if is_json:
        # Use JSON structure to build content (we'll reuse the previous ev_report builder if needed)
        body = "<p class='text-gray-500'>JSON structured mode – insert your structured report builder here.</p>"
    else:
        body = parse_advanced_plain_text(content, card_style, table_style, color_theme)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        body {{ font-family: {font_family}; background: #f8fafc; margin:0; }}
        .hero {{ background: {hero_bg}; }}
    </style>
</head>
<body>
<header class="hero {hero_text} py-16 md:py-24 text-center">
    <div class="max-w-3xl mx-auto px-4">
        <span class="inline-block bg-white/20 px-4 py-1 rounded-full text-sm mb-4">{badge}</span>
        <h1 class="text-4xl md:text-5xl font-extrabold">{title}</h1>
        <p class="mt-4 text-xl opacity-80">{subtitle}</p>
    </div>
</header>
<main class="max-w-4xl mx-auto px-4 py-8">
    <div class="bg-white rounded-2xl shadow-lg p-6 md:p-10">
        {body}
    </div>
    <footer class="text-center text-sm text-gray-400 mt-10">© 2026 – Generated Report</footer>
</main>
</body>
</html>"""
    return html

# ===================== STREAMLIT UI =====================
st.set_page_config(page_title="Report Generator", layout="wide")
st.title("📄 HTML Report Generator – Custom Styles")

# ---------- Sidebar Options ----------
with st.sidebar:
    st.header("🎨 Style Settings")
    card_style = st.selectbox("Card Style", list(CARD_STYLES.keys()))
    table_style = st.selectbox("Table Style", list(TABLE_STYLES.keys()))
    color_theme = st.selectbox("Colour Theme", list(COLOR_THEMES.keys()) + ["Dark Mode"])
    font = st.selectbox("Font", FONTS)

    st.markdown("---")
    st.markdown("✏️ Use `## Cards:` and `**Title**` to create card grids.")
    st.markdown("✏️ Use `| col1 | col2 |` for tables.")

# ---------- Main Content ----------
col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("Upload file", type=["json","txt","docx","pdf"])
with col2:
    input_mode = st.radio("Input type", ["Plain Text (markers)", "JSON"])

meta_title = st.text_input("Report Title", "EV Market Outlook")
meta_subtitle = st.text_input("Subtitle", "Your Electric Vehicle Report")

if input_mode == "Plain Text (markers)":
    raw_text = st.text_area("Paste your text with markers", height=300, value="""
# Budget-Friendly EVs
## Cards: Affordable City Cars
**VinFast VF 3**
- Price: ₹7.5 – ₹10 Lakh
- Range: 215 km

**MG Bingo EV**
- Price: ₹9 – ₹12 Lakh
- Range: Up to 410 km

## Quick Comparison
| Model | Price | Range |
|-------|-------|-------|
| Maruti e Vitara | ₹16–20 Lakh | 543 km |
| Toyota Ebella | ₹23–28 Lakh | 543 km |

> 2026 promises major advances in battery tech.
""")
else:
    raw_text = st.text_area("Paste JSON configuration", height=300, value='{"meta":{"title":"EV Report"},"sections":{}}')

# File override
if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    try:
        if file_type in ("json", "txt"):
            raw_text = read_text_file(uploaded_file)
        elif file_type == "docx":
            raw_text = read_docx_file(uploaded_file)
        elif file_type == "pdf":
            raw_text = read_pdf_file(uploaded_file)
    except Exception as e:
        st.error(f"File error: {e}")

is_json = False
if input_mode == "JSON":
    try:
        json.loads(raw_text)
        is_json = True
    except:
        st.warning("Invalid JSON – falling back to plain text.")
        is_json = False

if st.button("✨ Generate HTML Report"):
    if not raw_text.strip():
        st.warning("Please provide content.")
    else:
        meta = {"title": meta_title, "subtitle": meta_subtitle, "badge": "Report"}
        style_settings = {
            "card_style": card_style,
            "table_style": table_style,
            "color_theme": color_theme,
            "font": font
        }
        with st.spinner("Generating..."):
            final_html = generate_html(raw_text, is_json, meta, style_settings)
        st.success("Ready!")
        st.download_button("📥 Download HTML", data=final_html, file_name="report.html", mime="text/html")
        with st.expander("👁 Preview HTML", expanded=True):
            st.components.v1.html(final_html, height=700, scrolling=True)
