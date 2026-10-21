import os
import io
import json
import time
import base64
from datetime import date
import streamlit as st
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()

st.set_page_config(page_title="Naqvi QTY Pro", page_icon="🏗️", layout="wide")


def get_secret(name):
    try:
        return st.secrets[name]
    except Exception:
        return os.getenv(name)


client = genai.Client(api_key=get_secret("GEMINI_API_KEY"))

CATEGORIES = {
    "Building": [
        "Excavation", "PCC", "RCC (Concrete)", "Steel / Reinforcement",
        "Brickwork", "Plaster", "Flooring / Tiles", "Paint",
        "Doors & Windows", "Roofing", "Plumbing", "Electrical",
    ],
    "Road": [
        "Earthwork", "Sub-base", "Base Course", "Asphalt / Carpeting",
        "Kerb & Channel", "Drains / Culverts", "Road Marking",
    ],
}

CSS = """
<style>
#MainMenu, footer, header {visibility: hidden;}
.block-container {padding-top: 1.5rem; max-width: 1100px;}
.hero {background: linear-gradient(135deg, #0F2547 0%, #1F3864 60%, #2E5AA8 100%);
  border-radius: 18px; padding: 28px 32px; color: #fff; display: flex;
  align-items: center; justify-content: space-between; gap: 28px; margin-bottom: 22px;}
.hero h1 {color: #fff; margin: 0; font-size: 2.2rem; line-height: 1.2;}
.hero .brand {color: #E8C26A; margin: 8px 0 0 0; font-size: 1.1rem; font-weight: 600;}
.hero p {color: #D6E0F5; margin: 6px 0 0 0; font-size: 1rem;}
.hero .pic {width: 300px; height: 210px; border-radius: 18px; object-fit: cover;
  object-position: center 20%; border: 3px solid #E8C26A; flex-shrink: 0;}
.badge {display: inline-block; background: #E8C26A; color: #0F2547; font-weight: 700;
  border-radius: 999px; padding: 3px 14px; font-size: 0.75rem; margin-bottom: 10px;}
.step {display: flex; align-items: center; gap: 12px; margin: 22px 0 8px 0;}
.step .num {background: #1F3864; color: #fff; width: 34px; height: 34px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center; font-weight: 700;}
.step .ttl {font-size: 1.25rem; font-weight: 700; color: #0F2547;}
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  background: #1F3864; color: #fff; border: none; border-radius: 10px;
  padding: 0.6rem 1.6rem; font-weight: 600;}
.stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
  background: #2E5AA8; color: #fff;}
.footer {text-align: center; color: #6B7686; font-size: 0.85rem; margin-top: 34px;
  padding: 16px; border-top: 1px solid #DDE3EE;}
.login-title {text-align: center; color: #0F2547; font-size: 1.8rem; font-weight: 800; margin: 10px 0 0 0;}
.login-sub {text-align: center; color: #6B7686; margin-bottom: 14px;}
.frame {display: block; width: 100%; height: auto; border-radius: 16px;
  border: 3px solid #E8C26A; margin-bottom: 6px;}
@media (max-width: 700px) {
  .hero {flex-direction: column; text-align: center; padding: 22px;}
  .hero h1 {font-size: 1.6rem;}
  .hero .pic {width: 100%; height: auto; max-height: 260px;}
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def photo_src():
    base = os.path.dirname(os.path.abspath(__file__))
    for name, mime in [("friends.jpg", "jpeg"), ("friends.jpeg", "jpeg"), ("friends.png", "png")]:
        path = os.path.join(base, name)
        if os.path.exists(path):
            with open(path, "rb") as fh:
                return f"data:image/{mime};base64," + base64.b64encode(fh.read()).decode()
    return None


def step(num, title):
    st.markdown(
        f'<div class="step"><div class="num">{num}</div><div class="ttl">{title}</div></div>',
        unsafe_allow_html=True,
    )


def footer():
    st.markdown(
        '<div class="footer">A gift from Shoaib Taji &nbsp;|&nbsp; Naqvi QTY Pro &middot; '
        "Master Five Architecture &amp; Survey<br>AI-generated quantities must be "
        "verified by a qualified engineer before use.</div>",
        unsafe_allow_html=True,
    )


def ask_ai(contents):
    models = ["gemini-2.5-flash", "gemini-2.5-flash-lite"]
    last_error = None
    for model in models:
        for attempt in range(3):
            try:
                return client.models.generate_content(
                    model=model,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    ),
                )
            except Exception as e:
                last_error = e
                time.sleep(3 * (attempt + 1))
    raise last_error


def make_excel(info, df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quantities"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "NAQVI QTY PRO - Bill of Quantities"
    ws["A1"].font = Font(size=16, bold=True, color="1F3864")
    ws.merge_cells("A1:H1")
    ws["A2"] = "Quantities only - Rates to be filled by user"
    ws["A2"].font = Font(size=11, italic=True)
    ws.merge_cells("A2:H2")

    row = 4
    for label, value in info:
        ws.cell(row=row, column=2, value=label).font = Font(bold=True)
        ws.cell(row=row, column=3, value=value)
        row += 1

    head_row = row + 1
    headers = ["S.No", "Item", "Description", "Unit", "Qty", "Rate", "Amount", "Remarks"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=head_row, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    r = head_row + 1
    for i, rec in enumerate(df.to_dict("records"), 1):
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=rec.get("Item", ""))
        ws.cell(row=r, column=3, value=rec.get("Description", ""))
        ws.cell(row=r, column=4, value=rec.get("Unit", ""))
        ws.cell(row=r, column=5, value=rec.get("Qty", 0))
        ws.cell(row=r, column=6, value=None)
        ws.cell(row=r, column=7, value=f"=E{r}*F{r}")
        ws.cell(row=r, column=8, value=rec.get("Remarks", ""))
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border
        r += 1

    ws.cell(row=r, column=6, value="TOTAL").font = Font(bold=True)
    ws.cell(row=r, column=7, value=f"=SUM(G{head_row + 1}:G{r - 1})").font = Font(bold=True)

    for col, w in zip("ABCDEFGH", [7, 24, 40, 10, 12, 12, 14, 35]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- LOGIN ----------------
if not st.session_state.get("ok"):
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
        with st.container(border=True):
            src = photo_src()
            if src:
                st.markdown(f'<img class="frame" src="{src}">', unsafe_allow_html=True)
            st.markdown('<p class="login-title">A gift from Shoaib Taji</p>', unsafe_allow_html=True)
            st.markdown(
                '<p class="login-sub">Naqvi QTY Pro &middot; Master Five Architecture &amp; Survey</p>',
                unsafe_allow_html=True,
            )
            with st.form("login"):
                pw = st.text_input("Access code", type="password", placeholder="Enter your access code")
                go = st.form_submit_button("Unlock", use_container_width=True)
            if go:
                if pw and pw == get_secret("APP_PASSWORD"):
                    st.session_state["ok"] = True
                    st.rerun()
                else:
                    st.error("Incorrect access code.")
    footer()
    st.stop()

# ---------------- HERO ----------------
src = photo_src()
img_tag = f'<img class="pic" src="{src}">' if src else ""
st.markdown(
    f"""
<div class="hero">
  <div>
    <span class="badge">QUANTITY TAKE-OFF</span>
    <h1>A gift from Shoaib Taji</h1>
    <div class="brand">Naqvi QTY Pro &middot; Master Five Architecture &amp; Survey</div>
    <p>Upload a drawing, choose what you need, and get a ready-to-price BOQ in Excel.</p>
  </div>
  {img_tag}
</div>
""",
    unsafe_allow_html=True,
)

# ---------------- STEP 1 ----------------
step(1, "Client &amp; Project Details")
with st.container(border=True):
    c1, c2, c3 = st.columns(3)
    client_name = c1.text_input("Client Name *")
    client_phone = c2.text_input("Phone / WhatsApp")
    client_city = c3.text_input("Client City")
    p1, p2, p3 = st.columns(3)
    project_name = p1.text_input("Project Name *")
    location = p2.text_input("Location / Address *")
    project_date = p3.date_input("Date", date.today())
    q1, q2 = st.columns([1, 2])
    project_type = q1.radio("Project Type", list(CATEGORIES.keys()), horizontal=True)
    plot_size = q2.text_input("Plot Size / Road Length", placeholder="e.g. 10 Marla or 2.5 km")
    notes = st.text_area("Additional Notes (optional)", height=70)

# ---------------- STEP 2 ----------------
step(2, "Upload Drawing")
with st.container(border=True):
    u1, u2 = st.columns([1, 1])
    f = u1.file_uploader("PDF or Image", type=["pdf", "png", "jpg", "jpeg"])
    if f:
        if f.type.startswith("image"):
            u2.image(f, caption=f.name, use_container_width=True)
        else:
            u2.info(f"📄 {f.name} uploaded")

# ---------------- STEP 3 ----------------
if f:
    step(3, "Select Quantities Required")
    with st.container(border=True):
        all_items = CATEGORIES[project_type]
        pick_all = st.checkbox("Select all items")
        items = st.multiselect(
            "Choose items",
            all_items,
            default=all_items if pick_all else all_items[:3],
        )
        extra = st.text_input("Other items (optional)", placeholder="e.g. boundary wall, septic tank")
        run = st.button("⚡ Extract Quantities", type="primary")

    if run:
        if not client_name or not project_name or not location:
            st.warning("Please fill in Client Name, Project Name and Location.")
        elif not items and not extra:
            st.warning("Please select at least one item.")
        else:
            wanted = ", ".join(items + ([extra] if extra else []))
            prompt = f"""You are an expert quantity surveyor.
Project: {project_name}, Location: {location}, Type: {project_type}, Size: {plot_size}.
Notes: {notes}
From this drawing/text, extract quantities ONLY for: {wanted}.
Do not include any other items. No rates.
Return a JSON list:
[{{"Item":"","Description":"","Unit":"","Qty":0,"Remarks":""}}]
If a dimension is missing, put Qty 0 and explain in Remarks."""
            with st.spinner("AI is reading the drawing... (may take up to 20 seconds)"):
                try:
                    r = ask_ai([
                        types.Part.from_bytes(data=f.getvalue(), mime_type=f.type),
                        prompt,
                    ])
                    st.session_state["df"] = pd.DataFrame(json.loads(r.text))
                    st.session_state["info"] = [
                        ("Client", client_name),
                        ("Phone", client_phone),
                        ("City", client_city),
                        ("Project", project_name),
                        ("Location", location),
                        ("Type", project_type),
                        ("Size", plot_size),
                        ("Date", project_date.strftime("%d-%m-%Y")),
                    ]
                except Exception as e:
                    st.error(f"Error: {e}")

# ---------------- RESULTS ----------------
if "df" in st.session_state:
    step(4, "Review &amp; Download")
    with st.container(border=True):
        m1, m2, m3 = st.columns(3)
        m1.metric("Items extracted", len(st.session_state["df"]))
        m2.metric("Project", project_name or "-")
        m3.metric("Location", location or "-")
        edited = st.data_editor(
            st.session_state["df"], num_rows="dynamic", use_container_width=True
        )
        st.download_button(
            "📥 Download Excel (BOQ)",
            make_excel(st.session_state["info"], edited),
            f"BOQ_{project_name or 'project'}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

footer()
