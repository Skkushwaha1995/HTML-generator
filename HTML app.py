import streamlit as st
import json
import re

# Document reading libraries (for DOCX, PDF)
try:
    from docx import Document
except ImportError:
    st.error("python-docx missing. Add it to requirements.txt")
try:
    from PyPDF2 import PdfReader
except ImportError:
    st.error("PyPDF2 missing. Add it to requirements.txt")

# ---------- file text extraction ----------
def read_text_file(uploaded_file):
    return uploaded_file.getvalue().decode("utf-8")

def read_docx_file(uploaded_file):
    doc = Document(uploaded_file)
    return "\n".join([para.text for para in doc.paragraphs])

def read_pdf_file(uploaded_file):
    pdf_reader = PdfReader(uploaded_file)
    text = ""
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text.strip()

# ---------- plain-text HTML wrapper ----------
def build_plain_html(raw_text, meta=None):
    """Wrap any plain text inside the hero-styled report layout."""
    if meta is None:
        meta = {}
    title = meta.get("title", "EV Market Outlook")
    subtitle = meta.get("subtitle", "Your Electric Vehicle Report")
    badge = meta.get("badge", "Report")
    read_time = meta.get("read_time", "")
    date = meta.get("date", "")

    # Convert plain text to paragraphs
    paragraphs = raw_text.strip().split("\n")
    content_html = ""
    for p in paragraphs:
        if p.strip():
            content_html += f'<p class="mb-4 text-gray-700 leading-relaxed">{p.strip()}</p>'

    return f"""<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        body {{ margin:0; padding:0; font-family:'Inter',sans-serif; background:#f8fafc; }}
        .hero-gradient {{ background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 40%,#0ea5e9 100%); }}
        .pulse-dot {{ width:10px; height:10px; border-radius:50%; background:#10b981; display:inline-block; animation:pulse 2s infinite; }}
        @keyframes pulse {{ 0%,100% {{ box-shadow:0 0 0 0 rgba(16,185,129,0.6); }} 50% {{ box-shadow:0 0 0 14px rgba(16,185,129,0); }} }}
    </style>
</head>
<body>
<header class="hero-gradient text-white relative overflow-hidden">
    <div class="absolute top-0 left-0 w-full h-full opacity-10">
        <div class="absolute top-10 left-10 w-72 h-72 bg-white rounded-full blur-3xl"></div>
        <div class="absolute bottom-10 right-10 w-96 h-96 bg-cyan-400 rounded-full blur-3xl"></div>
    </div>
    <div class="max-w-[900px] mx-auto px-4 md:px-8 py-12 md:py-20 relative z-10 text-center">
        <span class="inline-flex items-center gap-2 bg-white/15 backdrop-blur-sm px-4 py-1.5 rounded-full text-sm font-medium mb-6 border border-white/20">
            <span class="pulse-dot"></span> {badge}
        </span>
        <h1 class="text-4xl md:text-5xl lg:text-6xl font-extrabold mb-5 leading-tight">{title}</h1>
        <p class="text-xl md:text-2xl text-white/80 mb-3">{subtitle}</p>
        <p class="text-white/60 text-lg mb-8">{read_time} {date}</p>
    </div>
    <div class="absolute bottom-0 left-0 w-full"><svg viewBox="0 0 1440 80" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none"><path d="M0 40C240 0 480 80 720 40C960 0 1200 80 1440 40V80H0V40Z" fill="#f8fafc"/></svg></div>
</header>
<main class="max-w-[900px] mx-auto px-4 md:px-8 py-8 font-sans relative z-10">
    <div class="bg-white p-6 md:p-10 rounded-2xl shadow-lg">
        {content_html}
    </div>
    <footer class="text-center text-sm text-gray-400 mt-10 pt-6 border-t border-gray-200">
        <p>© 2026 EV Market Outlook – Generated from your text</p>
    </footer>
</main>
</body>
</html>"""

# ---------- existing full JSON builder (unchanged) ----------
def build_html_report_from_dict(data):
    # ... (same as in the previous answer, omitted for brevity but must be included) ...
    # I'll paste the full function here to keep everything self-contained.
    meta = data["meta"]
    sections = data["sections"]
    vehicles = data.get("vehicles", {})
    comparison = data.get("comparison_table", [])
    faqs = data.get("faqs", [])

    def card_html(car, color):
        badge = f'<span class="inline-block mt-3 text-xs bg-{color}-100 text-{color}-700 px-2 py-1 rounded-full font-medium">{car["badge"]}</span>' if car.get("badge") else ""
        note = f'<p class="text-xs text-gray-400 mt-1 text-center">{car["note"]}</p>' if car.get("note") else ""
        return f"""
        <article class="bg-white rounded-2xl shadow-md card-hover border border-gray-100 overflow-hidden flex flex-col">
            <div class="card-image-wrapper">
                <img src="{car['image']}" alt="{car['name']}" loading="lazy">
            </div>
            <div class="p-5 flex-1 flex flex-col">
                <h3 class="text-xl font-bold text-gray-900 mb-2">{car['name']}</h3>
                <p class="text-base font-semibold text-{color}-600 mb-2">{car['price']}</p>
                <p class="text-gray-600 text-sm flex-1">{car['features']}</p>
                {badge}
                {note}
            </div>
        </article>"""

    budget_cards   = "\n".join([card_html(c, "emerald") for c in vehicles.get("budget", [])])
    midsize_cards  = "\n".join([card_html(c, "amber") for c in vehicles.get("midsize", [])])
    premium_cards  = "\n".join([card_html(c, "violet") for c in vehicles.get("premium", [])])

    table_rows = ""
    for row in comparison:
        table_rows += f"""
        <tr class="border-b border-gray-100 hover:bg-gray-50 transition-colors">
            <td class="p-4 font-medium text-gray-800">{row['name']}</td>
            <td class="p-4 text-gray-700">{row['price']}</td>
            <td class="p-4 text-gray-700">{row['range']}</td>
            <td class="p-4 text-gray-600 hidden md:table-cell">{row.get('highlight','')}</td>
        </tr>"""

    faq_html = ""
    for i, (q, a) in enumerate(faqs, 1):
        faq_html += f"""
        <div class="border border-gray-200 rounded-xl bg-white shadow-sm overflow-hidden">
            <input type="checkbox" id="faq{i}" class="hidden peer">
            <label for="faq{i}" class="flex justify-between items-center p-5 cursor-pointer text-lg font-semibold text-gray-800 hover:bg-gray-50 transition-colors">
                <span>{q}</span>
                <svg class="w-5 h-5 transition-transform duration-300 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M19 9l-7 7-7-7"></path></svg>
            </label>
            <div class="max-h-0 overflow-hidden transition-all duration-400 peer-checked:max-h-96">
                <div class="p-5 pt-0 text-gray-700 border-t border-gray-100 leading-relaxed">{a}</div>
            </div>
        </div>"""

    def build_section(section_id, title, body_html, color="indigo"):
        return f"""
        <section id="{section_id}" class="mb-12">
            <div class="flex items-center gap-3 mb-2"><span class="text-sm font-semibold text-{color}-500 uppercase tracking-wider">Section</span></div>
            <h2 class="text-3xl md:text-4xl font-extrabold text-gray-900 mb-5 leading-tight">{title}</h2>
            <div class="section-divider"></div>
            {body_html}
        </section>"""

    html = f"""<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Devanagari:wght@400;500;700&family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        body {{ margin:0; padding:0; font-family:'Inter',system-ui,sans-serif; background:#f8fafc; scroll-behavior:smooth; }}
        .hero-gradient {{ background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 40%,#0ea5e9 100%); }}
        .card-image-wrapper {{ position:relative; overflow:hidden; border-radius:0.75rem 0.75rem 0 0; aspect-ratio:16/10; background:#e2e8f0; }}
        .card-image-wrapper img {{ width:100%; height:100%; object-fit:cover; transition:transform 0.5s; }}
        .card-hover:hover .card-image-wrapper img {{ transform:scale(1.06); }}
        .card-hover {{ transition:all 0.35s cubic-bezier(0.4,0,0.2,1); }}
        .card-hover:hover {{ transform:translateY(-6px); box-shadow:0 25px 50px -12px rgba(0,0,0,0.18); }}
        .section-divider {{ height:4px; background:linear-gradient(90deg,transparent,#6366f1,#0ea5e9,#6366f1,transparent); border-radius:2px; margin:1.5rem 0 2.5rem; }}
        .faq-accordion input[type="checkbox"]:checked+label svg {{ transform:rotate(180deg); }}
        .pulse-dot {{ width:10px; height:10px; border-radius:50%; background:#10b981; display:inline-block; animation:pulse 2s infinite; }}
        @keyframes pulse {{ 0%,100% {{ box-shadow:0 0 0 0 rgba(16,185,129,0.6); }} 50% {{ box-shadow:0 0 0 14px rgba(16,185,129,0); }} }}
    </style>
</head>
<body>
<header class="hero-gradient text-white relative overflow-hidden">
    <div class="absolute top-0 left-0 w-full h-full opacity-10">
        <div class="absolute top-10 left-10 w-72 h-72 bg-white rounded-full blur-3xl"></div>
        <div class="absolute bottom-10 right-10 w-96 h-96 bg-cyan-400 rounded-full blur-3xl"></div>
    </div>
    <div class="max-w-[900px] mx-auto px-4 md:px-8 py-12 md:py-20 relative z-10 text-center">
        <span class="inline-flex items-center gap-2 bg-white/15 backdrop-blur-sm px-4 py-1.5 rounded-full text-sm font-medium mb-6 border border-white/20">
            <span class="pulse-dot"></span> {meta.get('badge','')}
        </span>
        <h1 class="text-4xl md:text-5xl lg:text-6xl font-extrabold mb-5 leading-tight">{meta['title']}</h1>
        <p class="text-xl md:text-2xl text-white/80 mb-3">{meta.get('subtitle','')}</p>
        <p class="text-white/60 text-lg mb-8">{meta.get('description','')}</p>
        <div class="flex flex-wrap justify-center gap-3 text-sm text-white/70">
            <span class="bg-white/10 px-3 py-1 rounded-full">⚡ {meta.get('read_time','')}</span>
            <span class="bg-white/10 px-3 py-1 rounded-full">📅 {meta.get('date','')}</span>
        </div>
    </div>
    <div class="absolute bottom-0 left-0 w-full"><svg viewBox="0 0 1440 80" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none"><path d="M0 40C240 0 480 80 720 40C960 0 1200 80 1440 40V80H0V40Z" fill="#f8fafc"/></svg></div>
</header>
<main class="max-w-[900px] mx-auto px-4 md:px-8 py-8 font-sans relative z-10">
    <figure class="mb-10 -mt-4 relative">
        <img src="{meta.get('hero_image','')}" alt="{meta.get('hero_image_alt','')}" class="w-full h-auto rounded-2xl shadow-2xl object-cover aspect-[2/1]" loading="eager">
        <figcaption class="text-center text-sm text-gray-500 mt-2 italic">{meta.get('hero_image_caption','')}</figcaption>
    </figure>
    <nav class="bg-white p-6 md:p-8 rounded-2xl shadow-lg mb-10 border border-gray-100">
        <h2 class="text-2xl font-bold text-gray-900 mb-5">📋 Table of Contents</h2>
        <ul class="grid sm:grid-cols-2 gap-2 text-indigo-700">{sections['toc']}</ul>
    </nav>
    {build_section("introduction", sections['intro']['heading'], sections['intro']['body'], "indigo")}
    {build_section("budget-friendly-evs", sections['budget']['heading'], f'<p class="mb-7 text-gray-700 text-lg">{sections["budget"]["intro_text"]}</p><div class="grid md:grid-cols-2 gap-6">{budget_cards}</div>', "emerald")}
    {build_section("mid-size-electric-suvs", sections['midsize']['heading'], f'<p class="mb-7 text-gray-700 text-lg">{sections["midsize"]["intro_text"]}</p><div class="grid md:grid-cols-2 gap-6">{midsize_cards}</div>', "amber")}
    {build_section("quick-comparison", sections['comparison']['heading'], f'<div class="overflow-x-auto rounded-2xl shadow-lg border border-gray-200"><table class="w-full border-collapse bg-white"><thead><tr class="bg-gradient-to-r from-indigo-500 to-purple-600 text-white"><th class="p-4 text-left font-semibold">Model</th><th class="p-4 text-left font-semibold">Price</th><th class="p-4 text-left font-semibold">Range</th><th class="p-4 text-left font-semibold hidden md:table-cell">Highlight</th></tr></thead><tbody>{table_rows}</tbody></table></div><p class="text-xs text-gray-400 mt-2 text-center">* Estimates based on industry reports.</p>', "rose")}
    {build_section("premium-evs", sections['premium']['heading'], f'<p class="mb-7 text-gray-700 text-lg">{sections["premium"]["intro_text"]}</p><div class="grid md:grid-cols-2 gap-6">{premium_cards}</div>', "violet")}
    {build_section("buy-now-or-wait", sections['buy_now']['heading'], sections['buy_now']['body'], "teal")}
    {build_section("conclusion", sections['conclusion']['heading'], sections['conclusion']['body'], "indigo")}
    <section id="faq" class="mb-12">
        <h2 class="text-3xl md:text-4xl font-extrabold text-gray-900 mb-5">❓ Frequently Asked Questions</h2>
        <div class="section-divider"></div>
        <div class="space-y-4 faq-accordion">{faq_html}</div>
    </section>
    <div class="bg-gradient-to-r from-indigo-600 via-purple-600 to-indigo-700 text-white p-8 rounded-2xl text-center shadow-2xl relative overflow-hidden mt-8">
        <div class="absolute top-0 left-0 w-full h-full opacity-10"><div class="absolute top-5 right-10 w-40 h-40 bg-white rounded-full blur-2xl"></div><div class="absolute bottom-5 left-10 w-52 h-52 bg-cyan-400 rounded-full blur-2xl"></div></div>
        <h3 class="text-2xl font-extrabold mb-3 relative z-10">Ready to Join the EV Revolution?</h3>
        <p class="mb-6 text-white/80 relative z-10">Find your perfect electric vehicle today.</p>
        <a href="#" class="inline-block bg-white text-indigo-700 font-bold py-3 px-8 rounded-full hover:bg-indigo-50 transition shadow-lg relative z-10">⚡ Discover EVs</a>
    </div>
    <footer class="text-center text-sm text-gray-400 mt-10 pt-6 border-t border-gray-200">
        <p>📝 <strong>Note:</strong> Replace placeholder images with actual model photos.</p>
        <p class="mt-1">© 2026 EV Market Outlook – Prices are industry estimates.</p>
    </footer>
</main>
</body>
</html>"""
    return html

# ---------- Streamlit UI ----------
st.set_page_config(page_title="EV Report Generator", layout="wide")
st.title("⚡ EV Market Outlook HTML Generator")
st.markdown(
    "Upload a **JSON, text, Word, or PDF** file, or paste content below. "
    "If the content is valid JSON, a full interactive report is created. "
    "Otherwise, your text will be turned into a clean, styled HTML page."
)

uploaded_file = st.file_uploader(
    "Upload file (JSON, TXT, DOCX, PDF)",
    type=["json", "txt", "docx", "pdf"]
)

json_text = st.text_area(
    "Or paste JSON / text content here (overrides file upload if both provided)",
    height=200
)

raw_text = None
report_data = None
is_json = False

# Process uploaded file
if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    try:
        if file_type in ("json", "txt"):
            raw_text = read_text_file(uploaded_file)
        elif file_type == "docx":
            raw_text = read_docx_file(uploaded_file)
        elif file_type == "pdf":
            raw_text = read_pdf_file(uploaded_file)
        else:
            st.error("Unsupported file type.")
    except Exception as e:
        st.error(f"Error reading file: {e}")

# Paste overrides file
if json_text.strip():
    raw_text = json_text.strip()

# Attempt to parse as JSON
if raw_text:
    try:
        report_data = json.loads(raw_text)
        is_json = True
        st.success("✅ Valid JSON detected – full report mode!")
    except json.JSONDecodeError:
        is_json = False
        st.info("ℹ️ Plain text detected – will generate a simple styled page.")

# Generate button
if st.button("✨ Generate HTML Report"):
    if raw_text is None:
        st.warning("Please upload a file or paste some content first.")
    else:
        with st.spinner("Generating report..."):
            if is_json and report_data is not None:
                html_output = build_html_report_from_dict(report_data)
            else:
                # plain text fallback – use a default meta for the hero
                default_meta = {
                    "title": "EV Market Outlook",
                    "subtitle": "Your Electric Vehicle Report",
                    "badge": "Report",
                    "read_time": "",
                    "date": ""
                }
                html_output = build_plain_html(raw_text, default_meta)

        st.success("Report generated!")
        st.download_button(
            label="📥 Download HTML Report",
            data=html_output,
            file_name="ev_report.html",
            mime="text/html"
        )
        with st.expander("Preview HTML"):
            st.components.v1.html(html_output, height=600, scrolling=True)
