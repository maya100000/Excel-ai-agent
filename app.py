import streamlit as st
import pandas as pd
import anthropic
import json
import os
import re
import tempfile
import traceback
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Contact Center AI Agent", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
section[data-testid="stSidebar"] { background: #0f1117; border-right: 1px solid #1e2130; }
section[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
.agent-header { background: #0f1117; border: 1px solid #1e2130; border-radius: 12px;
    padding: 1.2rem 1.8rem; margin-bottom: 1.5rem; display: flex; align-items: center; gap: 16px; }
.agent-header h1 { font-size: 1.4rem; font-weight: 600; color: #ffffff; margin: 0; letter-spacing: -0.02em; }
.agent-header p { font-size: 0.8rem; color: #6b7280; margin: 0; font-family: 'DM Mono', monospace; }
.status-dot { width:10px;height:10px;background:#22c55e;border-radius:50%;box-shadow:0 0 8px #22c55e88;flex-shrink:0; }
.kpi-card { background:#0f1117;border:1px solid #1e2130;border-radius:10px;padding:1rem 1.2rem; }
.kpi-label { font-size:0.72rem;color:#6b7280;text-transform:uppercase;letter-spacing:0.08em;font-family:'DM Mono',monospace;margin-bottom:6px; }
.kpi-value { font-size:1.6rem;font-weight:600;color:#ffffff;line-height:1; }
.kpi-sub { font-size:0.72rem;color:#6b7280;margin-top:4px;font-family:'DM Mono',monospace; }
.chat-wrap { background:#0f1117;border:1px solid #1e2130;border-radius:12px;padding:1.2rem;margin-bottom:1rem;max-height:500px;overflow-y:auto; }
.msg-user { background:#1a1f2e;border:1px solid #252b3b;border-radius:10px 10px 2px 10px;padding:0.7rem 1rem;margin-bottom:0.8rem;color:#e0e0e0;font-size:0.88rem;margin-left:20%; }
.msg-agent { background:#0d1f12;border:1px solid #1a3322;border-radius:10px 10px 10px 2px;padding:0.7rem 1rem;margin-bottom:0.8rem;color:#d1fae5;font-size:0.88rem;margin-right:10%;white-space:pre-wrap; }
.msg-label { font-size:0.68rem;font-family:'DM Mono',monospace;color:#6b7280;margin-bottom:4px; }
.msg-error { background:#1f0f0f;border:1px solid #3f1010;border-radius:10px 10px 10px 2px;padding:0.7rem 1rem;margin-bottom:0.8rem;color:#fca5a5;font-size:0.88rem;margin-right:10%; }
.section-title { font-size:0.72rem;font-family:'DM Mono',monospace;color:#6b7280;text-transform:uppercase;letter-spacing:0.1em;margin-bottom:0.8rem;padding-bottom:0.4rem;border-bottom:1px solid #1e2130; }
.badge-spike{background:#3f1010;color:#fca5a5;padding:2px 8px;border-radius:4px;font-size:0.72rem;font-family:'DM Mono',monospace;}
.badge-dip{background:#0d1f3a;color:#93c5fd;padding:2px 8px;border-radius:4px;font-size:0.72rem;font-family:'DM Mono',monospace;}
.badge-exp{background:#2d2000;color:#fde68a;padding:2px 8px;border-radius:4px;font-size:0.72rem;font-family:'DM Mono',monospace;}
.progress-step { padding: 6px 0; font-size: 0.82rem; font-family: 'DM Mono', monospace; }
.step-done  { color: #22c55e; }
.step-active{ color: #ffffff; }
.step-wait  { color: #374151; }
.error-box { background:#1f0f0f;border:1px solid #3f1010;border-radius:10px;padding:1rem 1.2rem;margin:1rem 0; }
.error-title { color:#fca5a5;font-weight:600;font-size:0.9rem;margin-bottom:4px; }
.error-detail { color:#9ca3af;font-size:0.8rem;font-family:'DM Mono',monospace; }
</style>
""", unsafe_allow_html=True)

# ── Schema ──────────────────────────────────────────────────────────
SHEET_TRANSLATION = {
    "אפריל 24":"April_2024","מאי 24":"May_2024","יוני 24":"June_2024",
    "אוגוסט 24":"August_2024","ספט 24":"September_2024","ינואר 25":"January_2025",
    "פברואר 25":"February_2025","מרץ 25":"March_2025","אפריל 25":"April_2025",
    "מאי 25":"May_2025","יוני 25":"June_2025","יולי 25":"July_2025",
    "אוגוסט 25":"August_2025","ספטמבר 25":"September_2025",
    "נציגים":"Agents","חיזוי אוגוסט+ספט":"Forecast_Aug_Sep",
    "חיזוי-תוצאות וסטיות":"Forecast_vs_Actuals",
}
COLUMN_SCHEMA = {
    "תאריך":"date","יום":"day_of_week","פניות כתובות נפתחו":"written_inquiries",
    "שיחות נכנסות":"inbound_calls","הערות":"notes","מספר נציגים":"num_agents",
    "שיחות ללא כפילויות":"unique_calls","פיק":"is_peak_day","עזר":"aux_calls","אחוז":"aux_pct",
    "נציג":"agent_name","זמן שיחה+החזק":"avg_handle_time","עמידה בשעות":"hours_compliance",
    "שיחות לשעה גולמי":"calls_per_hour_gross","תחזית פניות":"forecast_inquiries",
    "חיזוי שיחות סופי":"final_forecast_calls","פניות נתון אמת":"actual_inquiries",
    "שיחות נתון אמת":"actual_calls","סטייה פניות":"deviation_inquiries","סטייה שיחות":"deviation_calls",
}
DAY_TRANSLATION = {
    "יום ראשון":"Sunday","יום שני":"Monday","יום שלישי":"Tuesday",
    "יום רביעי":"Wednesday","יום חמישי":"Thursday","יום שישי":"Friday","שבת":"Saturday",
}
NOTES_TRANSLATION = {
    "חופש":"holiday","נוכחות":"payroll_day","שכר":"payroll_day","שכר,תלוש?":"payroll_day",
    "תלוש":"payslip_day","תקשור היעדרות מלחמה":"war_absence_notice","ט באב":"Tisha_BAv",
    "ערב חג (ללא שיחות)":"holiday_eve_no_calls","חזרה ללימודים":"back_to_school",
    "תחילת לימודים":"start_of_school_year","חזרה מחופש ראש השנה":"return_from_rosh_hashana",
    "52":"week_52_annotation","ממוצע שיחות ללא משה ותמר":"avg_excl_two_agents",
}
HIGH_VOLUME_EVENTS = {"payroll_day","payslip_day","war_absence_notice","back_to_school","start_of_school_year","return_from_rosh_hashana","Tisha_BAv"}
HOLIDAY_EVENTS     = {"holiday","holiday_eve_no_calls"}
WEEKEND_DAYS       = {"Friday","Saturday"}
MONTH_ORDER        = ["April_2024","May_2024","June_2024","August_2024","September_2024",
                      "January_2025","February_2025","March_2025","April_2025",
                      "May_2025","June_2025","July_2025","August_2025","September_2025"]

# ── Progress bar helper ─────────────────────────────────────────────
def show_progress(container, steps, current):
    """Render a step-by-step progress indicator."""
    html = ""
    for i, step in enumerate(steps):
        if i < current:
            html += f'<div class="progress-step step-done">✅ {step}</div>'
        elif i == current:
            html += f'<div class="progress-step step-active">⏳ {step}...</div>'
        else:
            html += f'<div class="progress-step step-wait">○ {step}</div>'
    container.markdown(html, unsafe_allow_html=True)

# ── Error display helper ────────────────────────────────────────────
def show_error(title, detail=None, suggestion=None):
    detail_html   = f'<div class="error-detail">{detail}</div>' if detail else ""
    suggest_html  = f'<div class="error-detail" style="margin-top:6px;color:#6b7280;">{suggestion}</div>' if suggestion else ""
    st.markdown(f'<div class="error-box"><div class="error-title">❌ {title}</div>{detail_html}{suggest_html}</div>', unsafe_allow_html=True)

# ── Validate uploaded file ──────────────────────────────────────────
def validate_file(uploaded_file):
    """Returns (ok, error_message)"""
    if uploaded_file is None:
        return False, "No file uploaded."
    if uploaded_file.size == 0:
        return False, "The file is empty."
    if uploaded_file.size > 50 * 1024 * 1024:
        return False, f"File is too large ({uploaded_file.size/1024/1024:.1f} MB). Maximum is 50 MB."
    name = uploaded_file.name.lower()
    if not (name.endswith(".xlsx") or name.endswith(".xls")):
        return False, f"File type not supported: {uploaded_file.name}. Please upload an .xlsx or .xls file."
    return True, None

# ── Validate API key ────────────────────────────────────────────────
def validate_api_key(api_key):
    if not api_key:
        return False, "No API key entered."
    if not api_key.startswith("sk-ant-"):
        return False, "Invalid API key format. Anthropic keys start with 'sk-ant-'."
    if len(api_key) < 20:
        return False, "API key is too short."
    return True, None

# ── Core processing functions ───────────────────────────────────────
def load_and_translate(path):
    xl = pd.ExcelFile(path)
    if len(xl.sheet_names) == 0:
        raise ValueError("The Excel file has no sheets.")
    sheets = {}
    for heb in xl.sheet_names:
        eng = SHEET_TRANSLATION.get(heb, heb)
        df = xl.parse(heb)
        df = df.rename(columns={h:e for h,e in COLUMN_SCHEMA.items() if h in df.columns})
        if "day_of_week" in df.columns:
            df["day_of_week"] = df["day_of_week"].map(lambda x: DAY_TRANSLATION.get(str(x),x) if pd.notna(x) else x)
        if "notes" in df.columns:
            df["notes"] = df["notes"].map(lambda x: NOTES_TRANSLATION.get(str(x).strip(),str(x)) if pd.notna(x) else x)
        if "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
        if "day_of_week" in df.columns:
            df["is_weekend"] = df["day_of_week"].isin(WEEKEND_DAYS)
        sheets[eng] = df
    return sheets

def clean_sheets(sheets):
    cleaned = {}
    for name, df in sheets.items():
        df = df.copy().drop_duplicates()
        for col in ["inbound_calls","written_inquiries","unique_calls","num_agents"]:
            if col in df.columns: df[col] = df[col].fillna(0)
        cleaned[name] = df
    return cleaned

def build_monthly_summary(sheets):
    rows = []
    for mn in MONTH_ORDER:
        df = sheets.get(mn)
        if df is None: continue
        is_weekend = df.get("is_weekend", pd.Series([False]*len(df), index=df.index))
        is_holiday = df["notes"].isin(HOLIDAY_EVENTS) if "notes" in df.columns else pd.Series([False]*len(df), index=df.index)
        is_workday = ~is_weekend & ~is_holiday
        w = df[is_workday]
        row = {"Month":mn,"Workdays":int(is_workday.sum())}
        if "notes" in df.columns:
            row["Payroll_Days"]    = int((df["notes"]=="payroll_day").sum())
            row["War_Notice_Days"] = int((df["notes"]=="war_absence_notice").sum())
        for col,lbl in [("inbound_calls","Total_Inbound_Calls"),("written_inquiries","Total_Written_Inquiries"),
                        ("unique_calls","Total_Unique_Calls"),("num_agents","Avg_Agents")]:
            if col in w.columns:
                vals = pd.to_numeric(w[col],errors="coerce").dropna()
                row[lbl] = round(float(vals.sum()),0) if lbl.startswith("Total") else round(float(vals.mean()),1)
        rows.append(row)
    if not rows:
        raise ValueError("No monthly data sheets found. Check that your file contains monthly sheets.")
    return pd.DataFrame(rows)

def compute_kpis(summary):
    kpis = {}
    if "Total_Inbound_Calls" not in summary.columns: return kpis
    kpis["total_calls"]     = int(summary["Total_Inbound_Calls"].sum())
    kpis["total_inquiries"] = int(summary["Total_Written_Inquiries"].sum()) if "Total_Written_Inquiries" in summary.columns else 0
    peak = summary.loc[summary["Total_Inbound_Calls"].idxmax()]
    low  = summary.loc[summary["Total_Inbound_Calls"].idxmin()]
    kpis["peak_month"]   = peak["Month"]
    kpis["peak_calls"]   = int(peak["Total_Inbound_Calls"])
    kpis["lowest_month"] = low["Month"]
    kpis["lowest_calls"] = int(low["Total_Inbound_Calls"])
    if len(summary)>=2:
        last = float(summary["Total_Inbound_Calls"].iloc[-1])
        prev = float(summary["Total_Inbound_Calls"].iloc[-2])
        kpis["mom_change"] = round((last-prev)/prev*100,1) if prev else 0
        kpis["last_month"] = summary["Month"].iloc[-1]
        kpis["prev_month"] = summary["Month"].iloc[-2]
    if "Avg_Agents" in summary.columns:
        kpis["avg_agents"] = round(float(summary["Avg_Agents"].mean()),1)
    return kpis

def detect_anomalies(sheets):
    anomalies = []
    for mn in MONTH_ORDER:
        df = sheets.get(mn)
        if df is None or "inbound_calls" not in df.columns: continue
        is_weekend = df.get("is_weekend", pd.Series([False]*len(df), index=df.index))
        is_holiday = df["notes"].isin(HOLIDAY_EVENTS) if "notes" in df.columns else pd.Series([False]*len(df), index=df.index)
        work = df[~is_weekend & ~is_holiday].copy()
        vals = pd.to_numeric(work["inbound_calls"],errors="coerce").dropna()
        if len(vals)<5: continue
        mean,std = float(vals.mean()),float(vals.std())
        if std==0: continue
        for idx in vals.index:
            v = float(vals.loc[idx])
            z = (v-mean)/std
            if abs(z)>2.0:
                row = work.loc[idx]
                note = str(row.get("notes",""))
                date_val = row.get("date","")
                date_str = date_val.strftime("%Y-%m-%d") if pd.notna(date_val) and hasattr(date_val,"strftime") else str(date_val)[:10]
                explanation = f"Expected — {note}" if note in HIGH_VOLUME_EVENTS else ("Unexplained spike" if z>0 else "Unexplained dip")
                anomalies.append({"Month":mn,"Date":date_str,"Day":str(row.get("day_of_week","")),"Inbound_Calls":v,
                                  "Month_Avg":round(mean,0),"Z_Score":round(z,2),
                                  "Flag":"SPIKE" if z>0 else "DIP","Event":note,"Explanation":explanation})
    return pd.DataFrame(anomalies)

def analyze_agents(sheets):
    ag = sheets.get("Agents")
    if ag is None or "agent_name" not in ag.columns: return pd.DataFrame()
    cols = [c for c in ["agent_name","avg_handle_time","hours_compliance","calls_per_hour_gross"] if c in ag.columns]
    df = ag[cols].dropna(subset=["agent_name"]).copy()
    if "calls_per_hour_gross" in df.columns:
        df = df.sort_values("calls_per_hour_gross",ascending=False).reset_index(drop=True)
        df.insert(0,"Rank",range(1,len(df)+1))
    return df

def build_context(sheets,summary,kpis,anomalies,agents):
    ctx = [
        "=== ISRAELI CONTACT CENTER — April 2024 to September 2025 ===",
        "Work week: Sunday to Thursday. Friday and Saturday are weekend — excluded from all KPIs.",
        "Business events: payroll_day (+30-60%), payslip_day (+74-190%), war_absence_notice (+93%), back_to_school (+210%), holiday=off.",
        f"\nMONTHLY SUMMARY:\n{summary.to_string(index=False)}",
        f"\nKPIs: {json.dumps(kpis)}",
    ]
    if not anomalies.empty:
        ctx.append(f"\nANOMALIES:\n{anomalies.to_string(index=False)}")
    if not agents.empty:
        ctx.append(f"\nAGENTS:\n{agents.to_string(index=False)}")
    return "\n".join(ctx)

# ── Chart detection ─────────────────────────────────────────────────
CHART_KEYWORDS = ["chart","graph","plot","show me","visualize","draw","display","bar","line","trend","compare visually","over time"]

def wants_chart(question):
    return any(kw in question.lower() for kw in CHART_KEYWORDS)

def build_chart_from_question(question, summary, sheets, agents):
    q = question.lower()
    if any(w in q for w in ["agent","rep","performer"]):
        if not agents.empty and "calls_per_hour_gross" in agents.columns:
            df = agents[["agent_name","calls_per_hour_gross"]].rename(columns={"agent_name":"Agent","calls_per_hour_gross":"Calls per Hour"})
            return "bar", df.set_index("Agent"), "Agent Performance — Calls per Hour"
    if any(w in q for w in ["day","weekday","sunday","monday","busiest day"]):
        day_data = []
        for mn in MONTH_ORDER:
            df = sheets.get(mn)
            if df is None or "inbound_calls" not in df.columns or "day_of_week" not in df.columns: continue
            is_weekend = df.get("is_weekend", pd.Series([False]*len(df), index=df.index))
            is_holiday = df["notes"].isin(HOLIDAY_EVENTS) if "notes" in df.columns else pd.Series([False]*len(df), index=df.index)
            work = df[~is_weekend & ~is_holiday]
            for day in ["Sunday","Monday","Tuesday","Wednesday","Thursday"]:
                vals = pd.to_numeric(work[work["day_of_week"]==day]["inbound_calls"],errors="coerce").dropna()
                if not vals.empty: day_data.append({"Day":day,"Avg Calls":round(float(vals.mean()),0)})
        if day_data:
            df = pd.DataFrame(day_data).groupby("Day")["Avg Calls"].mean().reset_index()
            df["Day"] = pd.Categorical(df["Day"],categories=["Sunday","Monday","Tuesday","Wednesday","Thursday"],ordered=True)
            return "bar", df.sort_values("Day").set_index("Day"), "Average Inbound Calls by Day of Week"
    if any(w in q for w in ["inquir","written","vs","both"]):
        if "Total_Inbound_Calls" in summary.columns and "Total_Written_Inquiries" in summary.columns:
            df = summary[["Month","Total_Inbound_Calls","Total_Written_Inquiries"]].copy()
            df["Month"] = df["Month"].str.replace("_"," ")
            df = df.rename(columns={"Total_Inbound_Calls":"Inbound Calls","Total_Written_Inquiries":"Written Inquiries"})
            return "bar", df.set_index("Month"), "Inbound Calls vs Written Inquiries"
    if any(w in q for w in ["staff","staffing","num agent","headcount"]):
        if "Avg_Agents" in summary.columns:
            df = summary[["Month","Avg_Agents"]].copy()
            df["Month"] = df["Month"].str.replace("_"," ")
            return "line", df.rename(columns={"Avg_Agents":"Avg Agents"}).set_index("Month"), "Avg Agents on Shift per Month"
    if "Total_Inbound_Calls" in summary.columns:
        df = summary[["Month","Total_Inbound_Calls"]].copy()
        df["Month"] = df["Month"].str.replace("_"," ")
        return "line", df.rename(columns={"Total_Inbound_Calls":"Inbound Calls"}).set_index("Month"), "Monthly Inbound Call Volume"
    return None

# ── Ask agent ───────────────────────────────────────────────────────
def ask_agent_smart(context, question, api_key, summary, sheets, agents):
    ok, err = validate_api_key(api_key)
    if not ok:
        return None, err, None

    chart = build_chart_from_question(question, summary, sheets, agents) if wants_chart(question) else None

    try:
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-opus-4-5", max_tokens=1500,
            system="You are an expert contact center operations analyst for an Israeli contact center. Work week is Sunday to Thursday. Friday and Saturday are always excluded from averages. Business events like payroll days and war notices significantly affect call volumes. Be specific, quantitative, and professional. When the user asks for a chart, briefly describe what it shows in 1-2 sentences then give key insights.",
            messages=[{"role":"user","content":f"{context}\n\nQuestion: {question}"}]
        )
        return response.content[0].text, None, chart
    except anthropic.AuthenticationError:
        return None, "Invalid API key. Please check your Anthropic API key in the sidebar.", None
    except anthropic.RateLimitError:
        return None, "Rate limit reached. Please wait a moment and try again.", None
    except anthropic.APIConnectionError:
        return None, "Could not connect to Anthropic. Please check your internet connection.", None
    except Exception as e:
        return None, f"Unexpected error: {str(e)}", None

def generate_excel(sheets,summary,kpis,anomalies,agents):
    output_path = os.path.join(tempfile.gettempdir(),"contact_center_analyzed.xlsx")
    with pd.ExcelWriter(output_path,engine="openpyxl") as writer:
        for mn in MONTH_ORDER:
            df = sheets.get(mn)
            if df is None: continue
            d = df.copy()
            if "date" in d.columns:
                d["date"] = d["date"].apply(lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) and hasattr(x,"strftime") else "")
            d.to_excel(writer,sheet_name=mn[:31],index=False)
        summary.to_excel(writer,sheet_name="Monthly_Summary",index=False)
        pd.DataFrame([{"KPI":k,"Value":str(v)} for k,v in kpis.items()]).to_excel(writer,sheet_name="KPIs",index=False)
        if not anomalies.empty: anomalies.to_excel(writer,sheet_name="Anomalies",index=False)
        if not agents.empty:    agents.to_excel(writer,sheet_name="Agent_Performance",index=False)
    wb = load_workbook(output_path)
    HDR=PatternFill("solid",fgColor="1F4E79"); HFONT=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    SUM=PatternFill("solid",fgColor="E2EFDA"); KPI=PatternFill("solid",fgColor="FFF2CC")
    SPK=PatternFill("solid",fgColor="FFAAAA"); DIP=PatternFill("solid",fgColor="AAC8FF")
    EXP=PatternFill("solid",fgColor="FFE5A0"); WKD=PatternFill("solid",fgColor="E0E0E0")
    HOL=PatternFill("solid",fgColor="D0E8FF"); PAY=PatternFill("solid",fgColor="FFF2CC")
    WAR=PatternFill("solid",fgColor="FFD0D0"); ALT=PatternFill("solid",fgColor="F7F7F7")
    thin=Side(style="thin",color="CCCCCC"); BDR=Border(left=thin,right=thin,top=thin,bottom=thin)
    CTR=Alignment(horizontal="center",vertical="center",wrap_text=True)
    def style(ws):
        for c in ws[1]: c.fill=HDR;c.font=HFONT;c.alignment=CTR;c.border=BDR
        ws.freeze_panes="A2"
        for i,row in enumerate(ws.iter_rows(min_row=2),2):
            for c in row: c.border=BDR
            if i%2==0:
                for c in row: c.fill=ALT
        for col in ws.columns:
            w=max((len(str(c.value or "")) for c in col),default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(w+4,42)
    for nm in wb.sheetnames: style(wb[nm])
    if "Monthly_Summary" in wb.sheetnames:
        for row in wb["Monthly_Summary"].iter_rows(min_row=2):
            for c in row: c.fill=SUM
    if "KPIs" in wb.sheetnames:
        for row in wb["KPIs"].iter_rows(min_row=2):
            for c in row: c.fill=KPI
    if "Anomalies" in wb.sheetnames:
        ws=wb["Anomalies"]; fc,ec=None,None
        for c in ws[1]:
            if c.value=="Flag": fc=c.column
            if c.value=="Explanation": ec=c.column
        if fc:
            for row in ws.iter_rows(min_row=2):
                flag=row[fc-1].value; expl=row[ec-1].value if ec else ""
                fill=EXP if "Expected" in str(expl) else (SPK if flag=="SPIKE" else DIP if flag=="DIP" else ALT)
                for c in row: c.fill=fill
    for mn in MONTH_ORDER:
        if mn not in wb.sheetnames: continue
        ws=wb[mn]; dc,nc=None,None
        for c in ws[1]:
            if c.value=="day_of_week": dc=c.column
            if c.value=="notes": nc=c.column
        for row in ws.iter_rows(min_row=2):
            day=str(row[dc-1].value or "") if dc else ""
            note=str(row[nc-1].value or "") if nc else ""
            if day in WEEKEND_DAYS: fill=WKD
            elif note in HOLIDAY_EVENTS: fill=HOL
            elif note in ("payroll_day","payslip_day"): fill=PAY
            elif note=="war_absence_notice": fill=WAR
            else: fill=None
            if fill:
                for c in row: c.fill=fill
    wb.save(output_path)
    return output_path

# ── Session state ───────────────────────────────────────────────────
for key in ["messages","sheets","summary","kpis","anomalies","agents","context","file_loaded","excel_path","load_error","last_chart"]:
    if key not in st.session_state:
        st.session_state[key] = [] if key=="messages" else None if key not in ("file_loaded",) else False

# ── Sidebar ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    api_key = st.text_input("Anthropic API Key", type="password", value=os.environ.get("ANTHROPIC_API_KEY",""),
                            help="Get your key at console.anthropic.com — starts with sk-ant-")
    if api_key:
        ok, err = validate_api_key(api_key)
        if ok:
            st.markdown('<span style="color:#22c55e;font-size:0.78rem;">✅ Key format valid</span>', unsafe_allow_html=True)
        else:
            st.markdown(f'<span style="color:#ef4444;font-size:0.78rem;">❌ {err}</span>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 📁 Upload File")
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx","xls"],
                                help="Max 50MB. Hebrew or English column names supported.")
    st.markdown("---")
    st.markdown("### 📋 Work Week")
    st.markdown("🟢 **Sun–Thu** = Workdays\n\n🔴 **Fri–Sat** = Weekend")
    st.markdown("---")
    st.markdown("### 💬 Chart Commands")
    st.markdown("Try:\n- *Show me a chart of monthly calls*\n- *Chart agent performance*\n- *Show calls vs inquiries*\n- *Chart by day of week*")

# ── Header ──────────────────────────────────────────────────────────
st.markdown("""
<div class="agent-header">
    <div class="status-dot"></div>
    <div>
        <h1>Contact Center AI Agent</h1>
        <p>Hebrew → English · April 2024 – September 2025 · Israeli work week (Sun–Thu)</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Process uploaded file with step-by-step progress ────────────────
if uploaded and not st.session_state.file_loaded:
    # Validate file first
    file_ok, file_err = validate_file(uploaded)
    if not file_ok:
        show_error("File Error", file_err, "Please upload a valid .xlsx or .xls file under 50MB.")
        st.session_state.load_error = file_err
    else:
        st.session_state.load_error = None
        STEPS = [
            "Reading Excel file",
            "Translating Hebrew to English",
            "Cleaning data",
            "Building monthly summary",
            "Computing KPIs",
            "Detecting anomalies",
            "Analyzing agent performance",
            "Generating Excel output",
        ]
        progress_container = st.empty()
        error_occurred = False

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            show_progress(progress_container, STEPS, 0)
            sheets = load_and_translate(tmp_path)

            show_progress(progress_container, STEPS, 1)
            # small pause so user can see each step
            import time; time.sleep(0.3)

            show_progress(progress_container, STEPS, 2)
            sheets = clean_sheets(sheets)
            time.sleep(0.2)

            show_progress(progress_container, STEPS, 3)
            summary = build_monthly_summary(sheets)
            time.sleep(0.2)

            show_progress(progress_container, STEPS, 4)
            kpis = compute_kpis(summary)
            time.sleep(0.2)

            show_progress(progress_container, STEPS, 5)
            anomalies = detect_anomalies(sheets)
            time.sleep(0.2)

            show_progress(progress_container, STEPS, 6)
            agents = analyze_agents(sheets)
            context = build_context(sheets, summary, kpis, anomalies, agents)
            time.sleep(0.2)

            show_progress(progress_container, STEPS, 7)
            excel_path = generate_excel(sheets, summary, kpis, anomalies, agents)
            time.sleep(0.3)

            # All done
            progress_container.markdown(
                "".join([f'<div class="progress-step step-done">✅ {s}</div>' for s in STEPS]),
                unsafe_allow_html=True
            )
            time.sleep(0.5)
            progress_container.empty()

            st.session_state.sheets      = sheets
            st.session_state.summary     = summary
            st.session_state.kpis        = kpis
            st.session_state.anomalies   = anomalies
            st.session_state.agents      = agents
            st.session_state.context     = context
            st.session_state.excel_path  = excel_path
            st.session_state.file_loaded = True
            st.session_state.messages    = []
            st.session_state.last_chart  = None
            st.success(f"✅ {uploaded.name} loaded — {len(sheets)} sheets, {len(anomalies)} anomalies detected")

        except ValueError as e:
            progress_container.empty()
            show_error("Data Error", str(e), "Make sure your file contains the expected monthly sheets.")
            st.session_state.load_error = str(e)
        except Exception as e:
            progress_container.empty()
            show_error("Processing Error", str(e), "Try re-uploading the file. If the problem persists, check that the file isn't password-protected or corrupted.")
            st.session_state.load_error = str(e)

elif uploaded is None:
    st.session_state.file_loaded = False
    st.session_state.load_error  = None

# ── Main dashboard ───────────────────────────────────────────────────
if st.session_state.file_loaded:
    kpis    = st.session_state.kpis
    summary = st.session_state.summary

    # KPI cards
    st.markdown('<div class="section-title">Key Performance Indicators</div>', unsafe_allow_html=True)
    c1,c2,c3,c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Total Inbound Calls</div><div class="kpi-value">{kpis.get("total_calls",0):,}</div><div class="kpi-sub">Workdays only (Sun–Thu)</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Written Inquiries</div><div class="kpi-value">{kpis.get("total_inquiries",0):,}</div><div class="kpi-sub">Apr 2024 – Sep 2025</div></div>', unsafe_allow_html=True)
    with c3:
        peak = kpis.get("peak_month","—").replace("_"," ")
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Peak Month</div><div class="kpi-value" style="font-size:1.1rem">{peak}</div><div class="kpi-sub">{kpis.get("peak_calls",0):,} calls</div></div>', unsafe_allow_html=True)
    with c4:
        mom=kpis.get("mom_change",0); color="#22c55e" if mom>=0 else "#ef4444"; arrow="↑" if mom>=0 else "↓"
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">MoM Change</div><div class="kpi-value" style="color:{color}">{arrow} {abs(mom)}%</div><div class="kpi-sub">{kpis.get("prev_month","").replace("_"," ")} → {kpis.get("last_month","").replace("_"," ")}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Charts
    st.markdown('<div class="section-title">Charts</div>', unsafe_allow_html=True)
    t1,t2,t3 = st.tabs(["📈 Monthly Volume","📊 Calls vs Inquiries","📅 Day of Week"])
    with t1:
        if "Total_Inbound_Calls" in summary.columns:
            df = summary[["Month","Total_Inbound_Calls"]].copy()
            df["Month"] = df["Month"].str.replace("_"," ")
            st.line_chart(df.rename(columns={"Total_Inbound_Calls":"Inbound Calls"}).set_index("Month"), height=240)
    with t2:
        cols_c = [c for c in ["Total_Inbound_Calls","Total_Written_Inquiries"] if c in summary.columns]
        if cols_c:
            df = summary[["Month"]+cols_c].copy()
            df["Month"] = df["Month"].str.replace("_"," ")
            df = df.rename(columns={"Total_Inbound_Calls":"Inbound Calls","Total_Written_Inquiries":"Written Inquiries"})
            st.bar_chart(df.set_index("Month"), height=240)
    with t3:
        day_data = []
        for mn in MONTH_ORDER:
            df = st.session_state.sheets.get(mn)
            if df is None or "inbound_calls" not in df.columns or "day_of_week" not in df.columns: continue
            is_weekend = df.get("is_weekend", pd.Series([False]*len(df), index=df.index))
            is_holiday = df["notes"].isin(HOLIDAY_EVENTS) if "notes" in df.columns else pd.Series([False]*len(df), index=df.index)
            work = df[~is_weekend & ~is_holiday]
            for day in ["Sunday","Monday","Tuesday","Wednesday","Thursday"]:
                vals = pd.to_numeric(work[work["day_of_week"]==day]["inbound_calls"],errors="coerce").dropna()
                if not vals.empty: day_data.append({"Day":day,"Avg Calls":round(float(vals.mean()),0)})
        if day_data:
            df = pd.DataFrame(day_data).groupby("Day")["Avg Calls"].mean().reset_index()
            df["Day"] = pd.Categorical(df["Day"],categories=["Sunday","Monday","Tuesday","Wednesday","Thursday"],ordered=True)
            st.bar_chart(df.sort_values("Day").set_index("Day"), height=240)

    st.markdown("<br>", unsafe_allow_html=True)

    # Anomalies + Agents
    col_left,col_right = st.columns([3,2])
    with col_left:
        st.markdown('<div class="section-title">Anomalies Detected</div>', unsafe_allow_html=True)
        anom = st.session_state.anomalies
        if not anom.empty:
            for _,row in anom.head(10).iterrows():
                flag=row.get("Flag",""); expl=str(row.get("Explanation",""))
                badge=(f'<span class="badge-exp">EXPECTED</span>' if "Expected" in expl
                       else f'<span class="badge-spike">SPIKE</span>' if flag=="SPIKE"
                       else f'<span class="badge-dip">DIP</span>')
                st.markdown(f"""<div style="background:#0f1117;border:1px solid #1e2130;border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:0.82rem;">
                    {badge} <span style="color:#9ca3af;font-family:'DM Mono',monospace;margin-left:8px;">{row.get("Date","")} · {row.get("Day","")}</span>
                    <span style="color:#6b7280;margin-left:8px;">{row.get("Month","").replace("_"," ")}</span><br>
                    <span style="color:#e0e0e0;font-weight:500;">{int(row.get("Inbound_Calls",0)):,} calls</span>
                    <span style="color:#6b7280;font-size:0.75rem;margin-left:6px;">avg {int(row.get("Month_Avg",0)):,} · z={row.get("Z_Score","")}</span><br>
                    <span style="color:#9ca3af;font-size:0.75rem;">{expl}</span>
                </div>""", unsafe_allow_html=True)
        else:
            st.markdown('<p style="color:#6b7280;font-size:0.85rem;">No anomalies detected.</p>', unsafe_allow_html=True)
    with col_right:
        st.markdown('<div class="section-title">Agent Performance</div>', unsafe_allow_html=True)
        agents_df = st.session_state.agents
        if not agents_df.empty:
            for _,row in agents_df.iterrows():
                cph=row.get("calls_per_hour_gross",0); rank=int(row.get("Rank",0))
                bar_w=int(min(cph/20*100,100)) if cph else 0
                medal="🥇" if rank==1 else "🥈" if rank==2 else "🥉" if rank==3 else f"#{rank}"
                st.markdown(f"""<div style="background:#0f1117;border:1px solid #1e2130;border-radius:8px;padding:10px 14px;margin-bottom:8px;">
                    <div style="display:flex;justify-content:space-between;">
                        <span style="color:#e0e0e0;font-weight:500;">{medal} {row.get("agent_name","")}</span>
                        <span style="color:#22c55e;font-family:'DM Mono',monospace;font-size:0.8rem;">{cph:.1f} calls/hr</span>
                    </div>
                    <div style="background:#1e2130;border-radius:3px;height:4px;margin-top:8px;">
                        <div style="background:#22c55e;width:{bar_w}%;height:4px;border-radius:3px;"></div>
                    </div>
                    <div style="color:#6b7280;font-size:0.72rem;margin-top:4px;">{row.get("avg_handle_time","")} avg handle time</div>
                </div>""", unsafe_allow_html=True)
        else:
            st.markdown('<p style="color:#6b7280;font-size:0.85rem;">No agent data found.</p>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Download
    if st.session_state.excel_path and os.path.exists(st.session_state.excel_path):
        with open(st.session_state.excel_path,"rb") as f: excel_bytes=f.read()
        st.download_button("⬇️  Download Analyzed Excel File", data=excel_bytes,
            file_name=f"contact_center_analyzed_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Chat ─────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">Ask the Agent</div>', unsafe_allow_html=True)

    suggestions = ["Analyze this dataset","Show me a chart of monthly calls","Chart agent performance",
                   "Why was February 2025 so high?","Compare May to April","Executive summary"]
    cols = st.columns(len(suggestions))
    for i,(col,q) in enumerate(zip(cols,suggestions)):
        with col:
            if st.button(q,key=f"sug_{i}",use_container_width=True):
                st.session_state._suggested = q

    # Chat history
    if st.session_state.messages:
        chat_html = '<div class="chat-wrap">'
        for msg in st.session_state.messages:
            if msg["role"]=="user":
                chat_html+=f'<div class="msg-label">You</div><div class="msg-user">{msg["content"]}</div>'
            elif msg["role"]=="error":
                chat_html+=f'<div class="msg-label">Agent</div><div class="msg-error">⚠️ {msg["content"]}</div>'
            else:
                chat_html+=f'<div class="msg-label">Agent</div><div class="msg-agent">{msg["content"]}</div>'
        chat_html+="</div>"
        st.markdown(chat_html, unsafe_allow_html=True)

        # Render chart from last response if any
        last_chart = st.session_state.last_chart
        if last_chart:
            chart_type, chart_df, chart_title = last_chart
            st.markdown(f'<div class="section-title">{chart_title}</div>', unsafe_allow_html=True)
            if chart_type=="line":
                st.line_chart(chart_df, height=300, use_container_width=True)
            else:
                st.bar_chart(chart_df, height=300, use_container_width=True)

    # Input
    question = st.chat_input("Ask anything — try 'show me a chart of...'")
    if hasattr(st.session_state,"_suggested"):
        question = st.session_state._suggested
        del st.session_state._suggested

    if question:
        if not api_key:
            show_error("No API Key", "Please enter your Anthropic API key in the sidebar.", "Get one free at console.anthropic.com")
        else:
            st.session_state.messages.append({"role":"user","content":question})
            with st.spinner(""):
                answer, error, chart = ask_agent_smart(
                    st.session_state.context, question, api_key,
                    st.session_state.summary, st.session_state.sheets, st.session_state.agents
                )
            if error:
                st.session_state.messages.append({"role":"error","content":error})
                st.session_state.last_chart = None
            else:
                st.session_state.messages.append({"role":"assistant","content":answer})
                st.session_state.last_chart = chart
            st.rerun()

else:
    # Empty state
    if st.session_state.get("load_error"):
        show_error("Could not load file", st.session_state.load_error, "Upload a different file to try again.")
    else:
        st.markdown("""
        <div style="text-align:center;padding:4rem 2rem;color:#6b7280;">
            <div style="font-size:3rem;margin-bottom:1rem;">📊</div>
            <div style="font-size:1.1rem;color:#9ca3af;margin-bottom:0.5rem;">Upload your Excel file to begin</div>
            <div style="font-size:0.82rem;font-family:'DM Mono',monospace;">Hebrew contact center data · Auto-translates to English</div>
            <br>
            <div style="font-size:0.78rem;color:#4b5563;font-family:'DM Mono',monospace;">
                Supported: .xlsx .xls · Max 50MB · Hebrew or English columns
            </div>
        </div>
        """, unsafe_allow_html=True)
