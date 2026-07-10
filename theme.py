"""
Visual theme for the Gen AI Champions Dashboard.

Injected once via st.markdown(..., unsafe_allow_html=True) in app.py.
Restyles Streamlit's default chrome (sidebar, buttons, metrics, dividers)
to match the navy / Space Grotesk "executive dashboard" reference design,
and defines the CSS classes used by components.py for stat cards, badges,
progress bars, and member cards.
"""

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');

:root{
  --bg:#F5F7FA;
  --panel:#FFFFFF;
  --ink:#1A202C;
  --muted:#64748B;
  --line:#E2E8F0;
  --navy:#0F172A;
  --navy-soft:#1E293B;
  --accent:#3B82F6;
  --accent-light:#DBEAFE;

  --completed:#059669;   --completed-bg:#D1FAE5;
  --inprogress:#3B82F6;  --inprogress-bg:#DBEAFE;
  --started:#F59E0B;     --started-bg:#FEF3C7;
  --notstarted:#94A3B8;  --notstarted-bg:#F1F5F9;
}

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
h1, h2, h3, h4 { font-family: 'Space Grotesk', sans-serif !important; }

/* App background */
.stApp { 
  background: linear-gradient(135deg, #F5F7FA 0%, #E5E9F2 100%);
}

/* Hide default Streamlit chrome we don't want */
#MainMenu, footer, header[data-testid="stHeader"] { visibility: hidden; height: 0; }

/* ---------------- Sidebar ---------------- */
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, var(--navy) 0%, #1E293B 100%);
  box-shadow: 4px 0 24px rgba(0,0,0,0.1);
}
section[data-testid="stSidebar"] * { color: #CBD5E1; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #FFFFFF; }

.brand-title{
  font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:18px;
  color:#fff;line-height:1.4;margin-bottom:4px;
  background: linear-gradient(135deg, #60A5FA 0%, #3B82F6 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}
.brand-sub{ font-size:12px;color:#94A3B8;margin-bottom:22px; }

.team-label{
  font-size:11px;text-transform:uppercase;letter-spacing:0.6px;
  color:#64748B;margin:22px 0 10px 0;font-weight:700;
  padding-bottom:6px;
  border-bottom:1px solid rgba(255,255,255,0.08);
}
.team-name{ font-size:13px;color:#CBD5E1;margin-bottom:8px;line-height:1.5; }
.team-name b{ color:#fff;font-weight:600; }
.team-name .id{ color:#64748B;font-size:11px; }

/* Sidebar nav -> st.sidebar.radio styled as pill nav items */
section[data-testid="stSidebar"] div[role="radiogroup"] label{
  display:flex;align-items:center;width:100%;
  padding:11px 14px;border-radius:10px;margin-bottom:4px;cursor:pointer;
  color:#CBD5E1;font-size:14px;font-weight:500;
  transition:all .2s cubic-bezier(0.4, 0, 0.2, 1);
  border:1px solid transparent;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover{
  background:rgba(59,130,246,0.1);
  border-color:rgba(59,130,246,0.2);
  color:#fff;
  transform:translateX(4px);
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked){
  background:linear-gradient(135deg, rgba(59,130,246,0.15) 0%, rgba(37,99,235,0.15) 100%);
  border-color:rgba(59,130,246,0.4);
  color:#fff;
  box-shadow: 0 4px 12px rgba(59,130,246,0.2);
}
section[data-testid="stSidebar"] div[role="radiogroup"] input[type="radio"]{
  /* keep in DOM for :has() and accessibility, just visually hidden */
  position:absolute; opacity:0; width:1px; height:1px;
}
section[data-testid="stSidebar"] div[data-testid="stWidgetLabel"]{ display:none; }

/* ---------------- Top bar ---------------- */
.topbar-title{ 
  font-size:28px;font-weight:700;margin:0 0 6px 0;color:var(--ink);
  background: linear-gradient(135deg, var(--ink) 0%, var(--accent) 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}
.topbar-sub{ margin:0;color:var(--muted);font-size:14px; }

/* ---------------- Stat cards ---------------- */
.stat-row{ display:flex; gap:16px; margin-bottom:28px; flex-wrap:wrap; }
.stat-card{
  background:var(--panel);
  border:1px solid var(--line);
  border-radius:16px;
  padding:20px 22px;
  flex:1;
  min-width:160px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.04);
  transition:all .3s cubic-bezier(0.4, 0, 0.2, 1);
  position:relative;
  overflow:hidden;
}
.stat-card:hover{
  transform:translateY(-4px);
  box-shadow: 0 12px 28px rgba(0,0,0,0.1);
  border-color:var(--accent);
}
.stat-card::before{
  content:'';
  position:absolute;
  top:0;left:0;right:0;
  height:4px;
  background:linear-gradient(90deg, var(--accent) 0%, var(--inprogress) 100%);
  opacity:0;
  transition:opacity .3s ease;
}
.stat-card:hover::before{ opacity:1; }
.stat-card .label{ font-size:12px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:0.5px; }
.stat-card .value{ 
  font-family:'Space Grotesk',sans-serif;font-size:32px;font-weight:700;margin-top:8px;color:var(--ink);
  background: linear-gradient(135deg, var(--ink) 0%, var(--accent) 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
}
.stat-card .sub{ font-size:12px;color:var(--muted);margin-top:6px; }

/* ---------------- Panels ---------------- */
.panel{
  background:var(--panel);
  border:1px solid var(--line);
  border-radius:16px;
  padding:24px 28px;
  margin-bottom:24px;
  box-shadow: 0 4px 20px rgba(0,0,0,0.04);
  transition:all .3s ease;
}
.panel:hover{
  box-shadow: 0 8px 32px rgba(0,0,0,0.08);
}
.panel h3{ 
  font-size:18px;margin:0 0 18px 0;color:var(--ink);font-weight:700;
  padding-bottom:12px;
  border-bottom:2px solid var(--line);
}

/* ---------------- Badges ---------------- */
.badge{
  display:inline-flex;align-items:center;gap:6px;padding:6px 12px;
  border-radius:24px;font-size:11px;font-weight:700;white-space:nowrap;
  transition:all .2s ease;
  border:1px solid transparent;
}
.badge:hover{ transform:scale(1.05); }
.badge .dot{ width:7px;height:7px;border-radius:50%;animation:pulse 2s ease-in-out infinite; }
@keyframes pulse {
  0%, 100% { opacity:1; }
  50% { opacity:0.6; }
}
.badge.completed{ background:var(--completed-bg);color:var(--completed);border-color:var(--completed); }
.badge.completed .dot{ background:var(--completed); }
.badge.inprogress{ background:var(--inprogress-bg);color:var(--inprogress);border-color:var(--inprogress); }
.badge.inprogress .dot{ background:var(--inprogress); }
.badge.started{ background:var(--started-bg);color:var(--started);border-color:var(--started); }
.badge.started .dot{ background:var(--started); }
.badge.notstarted{ background:var(--notstarted-bg);color:var(--notstarted);border-color:var(--notstarted); }
.badge.notstarted .dot{ background:var(--notstarted); }

/* ---------------- Progress bars ---------------- */
.progress-wrap{ display:flex;align-items:center;gap:10px;min-width:140px; }
.progress-track{ 
  flex:1;height:8px;border-radius:24px;background:var(--notstarted-bg);overflow:hidden;
  box-shadow: inset 0 2px 4px rgba(0,0,0,0.06);
}
.progress-fill{ 
  height:100%;border-radius:24px;
  background:linear-gradient(90deg, currentColor 0%, currentColor 100%);
  transition:width .6s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
.progress-pct{ font-size:13px;font-weight:700;font-family:'Space Grotesk',sans-serif;width:40px;text-align:right; }

/* ---------------- Topics table ---------------- */
.topics-table{ width:100%;border-collapse:collapse;font-size:14px; }
.topics-table th{
  text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:0.5px;
  color:var(--muted);padding:12px 14px;border-bottom:2px solid var(--line);
  background:var(--bg);font-weight:700;
}
.topics-table td{ padding:16px 14px;border-bottom:1px solid var(--line);vertical-align:top; }
.topics-table tr{ transition:all .2s ease; }
.topics-table tr:hover{ background:rgba(59,130,246,0.02); }
.topics-table tr:last-child td{ border-bottom:none; }
.topic-name{ font-weight:600;color:var(--ink);font-size:14px; }
.muted-italic{ font-style:italic;color:#AEB4C0;font-size:13px; }
.cell-muted{ color:var(--muted);font-size:13px;line-height:1.5; }
.note-tag{
  display:inline-block;font-size:10px;background:var(--started-bg);color:var(--started);
  padding:3px 8px;border-radius:8px;margin-left:8px;font-weight:700;text-transform:uppercase;
  border:1px solid var(--started);
}

/* ---------------- Team cards ---------------- */
.team-grid{ display:grid;grid-template-columns:repeat(auto-fill, minmax(200px, 1fr));gap:18px; }
.member-card{ 
  background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:20px;
  transition:all .3s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 2px 12px rgba(0,0,0,0.04);
}
.member-card:hover{
  transform:translateY(-6px);
  box-shadow: 0 12px 32px rgba(0,0,0,0.12);
  border-color:var(--accent);
}
.member-card .avatar{
  width:48px;height:48px;border-radius:50%;
  background:linear-gradient(135deg, var(--accent) 0%, var(--inprogress) 100%);
  color:#fff;
  display:flex;align-items:center;justify-content:center;
  font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:16px;margin-bottom:14px;
  box-shadow: 0 4px 16px rgba(59,130,246,0.3);
}
.member-card .name{ font-weight:700;font-size:15px;color:var(--ink); }
.member-card .role{ font-size:12px;color:var(--inprogress);font-weight:600;margin-top:4px; }
.member-card .dept{ font-size:11px;color:var(--muted);margin-top:6px; }

/* ---------------- Notes ---------------- */
.note-item{ 
  margin-bottom:16px;padding:16px;
  border-left:4px solid var(--accent);
  background:rgba(59,130,246,0.03);
  border-radius:8px;
}
.note-item:last-child{ margin-bottom:0; }
.note-item .n-body{ font-size:14px;color:var(--ink);line-height:1.6; }

/* ---------------- Report ---------------- */
.report-header{ 
  text-align:center;
  padding:24px 0 28px 0;
  border-bottom:3px solid var(--accent);
  margin-bottom:28px;
  background:linear-gradient(135deg, rgba(59,130,246,0.05) 0%, transparent 100%);
  border-radius:12px;
}
.report-header h1{ font-size:26px;margin:0 0 8px 0;color:var(--ink);font-weight:700; }
.report-header p{ margin:0;color:var(--muted);font-size:14px; }

/* ---------------- Streamlit Components Styling ---------------- */
/* Buttons */
button[kind="primary"], button[kind="secondary"], .stButton > button {
  border-radius:10px !important;
  font-weight:600 !important;
  font-size:14px !important;
  padding:10px 20px !important;
  border:1px solid transparent !important;
  transition:all .3s cubic-bezier(0.4, 0, 0.2, 1) !important;
  box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
}
button[kind="primary"]:hover, .stButton > button:hover {
  transform:translateY(-2px) !important;
  box-shadow: 0 6px 20px rgba(59,130,246,0.3) !important;
  border-color:var(--accent) !important;
}

/* Chat Messages */
div[data-testid="stChatMessage"] {
  border-radius:12px !important;
  padding:16px !important;
  margin-bottom:12px !important;
  box-shadow: 0 2px 12px rgba(0,0,0,0.04) !important;
  border:1px solid var(--line) !important;
}
div[data-testid="stChatMessage"]:hover {
  box-shadow: 0 4px 20px rgba(0,0,0,0.08) !important;
}

/* Chat Input */
div[data-testid="stChatInput"] {
  border-radius:12px !important;
  border:2px solid var(--line) !important;
  transition:all .3s ease !important;
}
div[data-testid="stChatInput"]:focus-within {
  border-color:var(--accent) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,0.1) !important;
}

/* Text Area */
textarea {
  border-radius:10px !important;
  border:2px solid var(--line) !important;
  padding:14px !important;
  font-size:14px !important;
  transition:all .3s ease !important;
}
textarea:focus {
  border-color:var(--accent) !important;
  box-shadow: 0 0 0 3px rgba(59,130,246,0.1) !important;
  outline:none !important;
}

/* File Uploader */
div[data-testid="stFileUploader"] {
  border-radius:12px !important;
  border:2px dashed var(--line) !important;
  padding:24px !important;
  transition:all .3s ease !important;
}
div[data-testid="stFileUploader"]:hover {
  border-color:var(--accent) !important;
  background:rgba(59,130,246,0.02) !important;
}

/* Spinner */
div[data-testid="stSpinner"] > div {
  border-color:var(--accent) transparent transparent transparent !important;
}

/* Success/Warning/Error Messages */
div[data-testid="stAlert"] {
  border-radius:12px !important;
  padding:14px 18px !important;
  border-left-width:4px !important;
}

/* Streamlit metric/progress native overrides (used sparingly as fallback) */
div[data-testid="stMetric"]{
  background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:18px 20px;
  box-shadow: 0 4px 16px rgba(0,0,0,0.04);
  transition:all .3s ease;
}
div[data-testid="stMetric"]:hover{
  transform:translateY(-2px);
  box-shadow: 0 8px 24px rgba(0,0,0,0.1);
}
</style>
"""
