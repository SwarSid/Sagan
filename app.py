import streamlit as st
import pandas as pd
import json
import io
import re
import anthropic
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

st.set_page_config(
    page_title="Doctor Response Analyzer",
    page_icon="⚕️",
    layout="wide"
)

COLORS = [
    "#1a2e5a","#e5333a","#f7a826","#16a34a",
    "#7c3aed","#0891b2","#be185d","#b45309",
    "#065f46","#1d4ed8","#9d174d","#0f766e",
]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Inter:wght@400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.top-banner {
    background: #1a2e5a;
    padding: 1.8rem 2rem;
    border-radius: 14px;
    margin-bottom: 1.5rem;
    color: white;
}
.top-banner h1 { font-family: 'Syne', sans-serif; font-size: 1.8rem; font-weight: 800; margin: 0 0 4px 0; }
.top-banner p  { margin: 0; opacity: 0.6; font-size: 13px; }
.red-tag { display:inline-block; background:#e5333a; color:white; font-size:11px; font-weight:600; padding:2px 10px; border-radius:20px; margin-bottom:8px; }
.qblock { border-left:3px solid #e5333a; border-radius:0 8px 8px 0; padding:9px 13px; margin:7px 0; background:#f5f4f1; font-size:13px; font-style:italic; color:#555; line-height:1.55; }
.qbuck { font-style:normal; font-size:10px; font-weight:700; color:#e5333a; margin-top:3px; text-transform:uppercase; letter-spacing:.4px; }
.bcard { background:white; border:1px solid rgba(0,0,0,0.08); border-radius:10px; padding:13px; border-left:4px solid #ccc; margin-bottom:7px; }
.bpct  { font-family:'Syne',sans-serif; font-size:26px; font-weight:800; line-height:1; margin-bottom:4px; }
.bname { font-size:13px; font-weight:600; margin-bottom:3px; }
.btheme{ font-size:12px; color:#6b6b78; margin-bottom:7px; }
.bquote{ font-size:11px; font-style:italic; color:#888; border-left:2px solid #e0e0e8; padding-left:7px; line-height:1.45; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="top-banner">
  <div class="red-tag">⚕ AI Research Tool</div>
  <h1>Doctor Response Analyzer</h1>
  <p>Upload responses · AI bucketing · Excel + PowerPoint export · One link for your whole team</p>
</div>
""", unsafe_allow_html=True)

# ── API key ───────────────────────────────────────────────────────────────────
if "ANTHROPIC_API_KEY" in st.secrets:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
else:
    with st.sidebar:
        st.markdown("### API Key")
        api_key = st.text_input("Anthropic API key", type="password",
                                 placeholder="sk-ant-api03-...")
        if not api_key:
            st.warning("Enter API key to run analysis")

# ── Session state ─────────────────────────────────────────────────────────────
if "buckets"       not in st.session_state: st.session_state.buckets = []
if "total"         not in st.session_state: st.session_state.total = 0
if "tagged_df"     not in st.session_state: st.session_state.tagged_df = None
if "col"           not in st.session_state: st.session_state.col = None
if "bg"            not in st.session_state: st.session_state.bg = ""
if "question"      not in st.session_state: st.session_state.question = ""
if "done"          not in st.session_state: st.session_state.done = False

# ── Step 1: Context ───────────────────────────────────────────────────────────
with st.expander("**Step 1 — Study context** (recommended)", expanded=not st.session_state.done):
    c1, c2 = st.columns(2)
    with c1:
        bg = st.text_area("Study background",
            placeholder="e.g. Post-consultation survey with oncologists discussing VORANIGO prescribing barriers...",
            height=90, key="bg_in")
    with c2:
        question = st.text_input("Question asked to doctors",
            placeholder="e.g. What are the main barriers you face when prescribing VORANIGO?",
            key="q_in")

# ── Step 2: Upload ────────────────────────────────────────────────────────────
st.markdown("### 📂 Step 2 — Upload your Excel or CSV")
uploaded = st.file_uploader("Upload file", type=["xlsx","xls","csv"],
                              label_visibility="collapsed")

df = None
sel_col = None

if uploaded:
    try:
        if uploaded.name.endswith(".csv"):
            df = pd.read_csv(uploaded, dtype=str).fillna("")
        else:
            df = pd.read_excel(uploaded, dtype=str).fillna("")

        st.success(f"✓ **{uploaded.name}** — {len(df)} rows, {len(df.columns)} columns")

        kw = ["response","answer","comment","text","doctor","feedback","reply","verbatim"]
        auto = next((c for c in df.columns if any(k in c.lower() for k in kw)), df.columns[0])

        sel_col = st.selectbox("Which column has the doctor responses?",
                                df.columns.tolist(),
                                index=df.columns.tolist().index(auto))
        if sel_col and len(df) > 0:
            st.caption(f'↳ Preview: *"{str(df[sel_col].iloc[0])[:160]}…"*')
    except Exception as e:
        st.error(f"Cannot read file: {e}")

# ── Step 3: Run ───────────────────────────────────────────────────────────────
st.markdown("### 🤖 Step 3 — Run AI Analysis")

can_run = df is not None and sel_col is not None and bool(api_key)
if st.button("▶ Run AI Analysis", disabled=not can_run, type="primary"):
    responses = df[sel_col].dropna().astype(str).str.strip()
    responses = responses[responses.str.len() > 8].tolist()

    if len(responses) < 3:
        st.error("Not enough text responses. Check your column selection.")
    else:
        with st.spinner(f"Analyzing {len(responses)} responses… usually 20–40 seconds"):
            sys_p = """You are an expert qualitative researcher for medical/pharma studies.
Analyse the doctor responses and return ONLY valid JSON — no markdown, no backticks, no explanation.

{
  "totalResponses": <integer>,
  "buckets": [
    {
      "name": "<insight-forward label, 4-7 words>",
      "count": <integer>,
      "percentage": <number 1dp>,
      "theme": "<one sentence core insight>",
      "quotes": ["<verbatim 10-25 word fragment>","<another>","<third>"],
      "responseIndices": [<0-based integers>]
    }
  ]
}

Rules: 10-12 buckets, every index 0..(n-1) in exactly one bucket, counts sum to totalResponses, percentages sum to 100.0, order by count desc, insight-forward names."""

            sample = responses[:120]
            user_msg = (f"Background: {bg or 'Medical research'}\n"
                        f"Question: {question or 'Doctor experiences'}\n\n"
                        f"Responses ({len(sample)} total):\n" +
                        "\n".join(f"[{i}] {r}" for i, r in enumerate(sample)))
            try:
                client = anthropic.Anthropic(api_key=api_key)
                msg = client.messages.create(
                    model="claude-sonnet-4-5",
                    max_tokens=2000,
                    system=sys_p,
                    messages=[{"role":"user","content":user_msg}]
                )
                raw = msg.content[0].text.strip()
                raw = re.sub(r'^```json\s*','',raw,flags=re.I)
                raw = re.sub(r'^```','',raw).replace('```','').strip()
                parsed = json.loads(raw)

                buckets = parsed["buckets"]
                total   = parsed["totalResponses"]

                tagged = df.copy()
                tagged["Bucket"] = "Uncategorized"
                for b in buckets:
                    for idx in (b.get("responseIndices") or []):
                        if idx < len(tagged):
                            tagged.at[idx, "Bucket"] = b["name"]

                st.session_state.buckets   = buckets
                st.session_state.total     = total
                st.session_state.tagged_df = tagged
                st.session_state.col       = sel_col
                st.session_state.bg        = bg
                st.session_state.question  = question
                st.session_state.done      = True
                st.rerun()

            except Exception as e:
                st.error(f"Analysis failed: {e}")

# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.done and st.session_state.buckets:
    buckets   = st.session_state.buckets
    total     = st.session_state.total
    tagged_df = st.session_state.tagged_df
    sel_col   = st.session_state.col

    st.divider()
    st.markdown("## Results")

    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Total Responses", total)
    m2.metric("Buckets", len(buckets))
    m3.metric("Largest Bucket", f"{buckets[0]['percentage']}%")
    m4.metric("Major Themes ≥10%", sum(1 for b in buckets if b["percentage"]>=10))

    t1,t2,t3,t4 = st.tabs(["📋 Table","📊 Visualize","💬 Quotes","🏷️ Tagged"])

    # ── Table ─────────────────────────────────────────────────────────────────
    with t1:
        rows = [{"#":i+1,"Bucket":b["name"],"Count":b["count"],
                 "%":f"{b['percentage']}%","Theme":b["theme"],
                 "Sample quote":b["quotes"][0] if b.get("quotes") else ""}
                for i,b in enumerate(buckets)]
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # ── Visualize ─────────────────────────────────────────────────────────────
    with t2:
        viz = st.selectbox("Choose style", [
            "Horizontal bar chart + quotes",
            "Stat cards per bucket",
            "Ranked driver list",
            "Donut chart + legend",
            "Summary table"
        ], key="viz_pick")

        names  = [b["name"] for b in buckets]
        counts = [b["count"] for b in buckets]
        pcts   = [b["percentage"] for b in buckets]
        colors = [COLORS[i%len(COLORS)] for i in range(len(buckets))]

        if "bar" in viz:
            fig = go.Figure(go.Bar(
                x=counts, y=names, orientation='h',
                marker_color=colors,
                text=[f"{p}% ({c})" for c,p in zip(counts,pcts)],
                textposition='outside'
            ))
            fig.update_layout(
                height=max(300,len(buckets)*40+100),
                margin=dict(l=20,r=120,t=10,b=20),
                yaxis=dict(autorange="reversed"),
                plot_bgcolor='white', paper_bgcolor='white',
                xaxis_title="Responses"
            )
            fig.update_xaxes(showgrid=True,gridcolor='#f0f0f0')
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("**Key quotes:**")
            for b in buckets[:3]:
                if b.get("quotes"):
                    st.markdown(f'<div class="qblock">"{b["quotes"][0]}"<div class="qbuck">{b["name"]} · {b["percentage"]}%</div></div>', unsafe_allow_html=True)

        elif "cards" in viz:
            cols = st.columns(3)
            for i,b in enumerate(buckets):
                col = COLORS[i%len(COLORS)]
                q = f'<div class="bquote">"{b["quotes"][0]}"</div>' if b.get("quotes") else ""
                with cols[i%3]:
                    st.markdown(f'<div class="bcard" style="border-left-color:{col}"><div class="bpct" style="color:{col}">{b["percentage"]}%</div><div class="bname">{b["name"]}</div><div class="btheme">{b["theme"]}</div>{q}</div>', unsafe_allow_html=True)

        elif "ranked" in viz:
            labels = ["🔴 Most impactful","🟡 Moderate","🟢 Least impactful"]
            splits = [int(len(buckets)*0.3), int(len(buckets)*0.7), len(buckets)]
            prev = 0
            for lbl,end in zip(labels,splits):
                grp = buckets[prev:end]
                if grp:
                    st.markdown(f"**{lbl}**")
                    for i,b in enumerate(grp, start=prev+1):
                        col = COLORS[(i-1)%len(COLORS)]
                        q = f'<div style="font-size:11px;font-style:italic;color:#888;margin-top:5px;border-left:2px solid #e0e0e8;padding-left:7px">"{b["quotes"][0]}"</div>' if b.get("quotes") else ""
                        st.markdown(f'<div style="display:flex;align-items:flex-start;gap:12px;padding:11px 14px;background:white;border:1px solid rgba(0,0,0,0.08);border-radius:10px;margin-bottom:6px;border-left:4px solid {col}"><div style="font-family:sans-serif;font-size:18px;font-weight:800;color:{col};min-width:26px;text-align:center">{i}</div><div style="flex:1"><div style="font-size:13px;font-weight:600">{b["name"]}</div><div style="font-size:11px;color:#6b6b78">{b["theme"]}</div>{q}</div><div style="font-size:20px;font-weight:800;color:{col}">{b["percentage"]}%</div></div>', unsafe_allow_html=True)
                prev = end

        elif "donut" in viz:
            fig = go.Figure(go.Pie(
                labels=names, values=counts, hole=0.52,
                marker_colors=colors, textinfo='percent'
            ))
            fig.update_layout(height=420,margin=dict(l=20,r=20,t=10,b=20),
                               showlegend=True, paper_bgcolor='white')
            st.plotly_chart(fig, use_container_width=True)

        else:
            for i,b in enumerate(buckets):
                col = COLORS[i%len(COLORS)]
                q = f'<em>"{b["quotes"][0]}"</em>' if b.get("quotes") else "—"
                st.markdown(f'<div style="display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center;padding:10px 12px;background:white;border:1px solid rgba(0,0,0,0.07);border-radius:8px;margin-bottom:5px;border-left:4px solid {col}"><div style="font-size:14px;font-weight:800;color:{col};min-width:28px;text-align:center">{i+1}</div><div><div style="font-size:13px;font-weight:600">{b["name"]}</div><div style="font-size:11px;color:#888;margin-top:2px">{b["theme"]}</div><div style="font-size:11px;color:#aaa;margin-top:3px">{q}</div></div><div style="font-size:22px;font-weight:800;color:{col}">{b["percentage"]}%</div></div>', unsafe_allow_html=True)

    # ── Quotes ────────────────────────────────────────────────────────────────
    with t3:
        for b in buckets:
            for q in (b.get("quotes") or []):
                st.markdown(f'<div class="qblock">"{q}"<div class="qbuck">{b["name"]} · {b["percentage"]}%</div></div>', unsafe_allow_html=True)

    # ── Tagged ────────────────────────────────────────────────────────────────
    with t4:
        st.dataframe(tagged_df, use_container_width=True, height=380)

    # ── Export ────────────────────────────────────────────────────────────────
    st.divider()
    st.markdown("## 📥 Export")
    e1, e2 = st.columns(2)

    with e1:
        st.markdown("#### 📊 Excel")
        if st.button("Generate Excel", type="primary", use_container_width=True):
            with st.spinner("Building Excel…"):
                wb = Workbook()
                ws1 = wb.active
                ws1.title = "Responses (Bucketed)"
                hdrs = list(tagged_df.columns)
                for j,h in enumerate(hdrs,1):
                    c = ws1.cell(row=1,column=j,value=h)
                    c.font = Font(bold=True,color="FFFFFF",name="Calibri",size=10)
                    c.fill = PatternFill("solid",fgColor="1A2E5A")
                    c.alignment = Alignment(horizontal="center")
                for i,row in tagged_df.iterrows():
                    for j,val in enumerate(row,1):
                        cell = ws1.cell(row=i+2,column=j,value=str(val))
                        cell.font = Font(name="Calibri",size=10)
                for col_cells in ws1.columns:
                    mx = max((len(str(c.value or "")) for c in col_cells),default=10)
                    ws1.column_dimensions[get_column_letter(col_cells[0].column)].width = min(mx+2,50)

                ws2 = wb.create_sheet("Bucket Summary")
                hdr2 = ["Rank","Bucket","Count","%","Theme","Quote 1","Quote 2","Quote 3"]
                for j,h in enumerate(hdr2,1):
                    c = ws2.cell(row=1,column=j,value=h)
                    c.font = Font(bold=True,color="FFFFFF",name="Calibri",size=10)
                    c.fill = PatternFill("solid",fgColor="E5333A")
                    c.alignment = Alignment(horizontal="center")
                for i,b in enumerate(buckets):
                    qs = b.get("quotes",[])
                    row = [i+1,b["name"],b["count"],f"{b['percentage']}%",b["theme"],
                           qs[0] if len(qs)>0 else "",qs[1] if len(qs)>1 else "",qs[2] if len(qs)>2 else ""]
                    for j,v in enumerate(row,1):
                        c = ws2.cell(row=i+2,column=j,value=v)
                        c.font = Font(name="Calibri",size=10)
                        c.alignment = Alignment(wrap_text=True,vertical="top")
                ws2.column_dimensions["A"].width = 6
                ws2.column_dimensions["B"].width = 38
                ws2.column_dimensions["C"].width = 8
                ws2.column_dimensions["D"].width = 10
                ws2.column_dimensions["E"].width = 55
                for col in ["F","G","H"]:
                    ws2.column_dimensions[col].width = 50

                buf = io.BytesIO()
                wb.save(buf)
                buf.seek(0)
                st.download_button("⬇️ Download Excel", data=buf.getvalue(),
                    file_name="doctor_responses_bucketed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    with e2:
        st.markdown("#### 📑 PowerPoint")
        if st.button("Generate PowerPoint", type="primary", use_container_width=True):
            with st.spinner("Building PowerPoint…"):
                prs = Presentation()
                prs.slide_width  = Inches(13.33)
                prs.slide_height = Inches(7.5)
                blank = prs.slide_layouts[6]

                def rect(slide,x,y,w,h,fill):
                    s = slide.shapes.add_shape(1,Inches(x),Inches(y),Inches(w),Inches(h))
                    s.fill.solid()
                    s.fill.fore_color.rgb = RGBColor.from_string(fill)
                    s.line.fill.background()
                    return s

                def txt(slide,text,x,y,w,h,sz=12,bold=False,col="000000",italic=False,align=PP_ALIGN.LEFT):
                    tb = slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
                    tb.word_wrap = True
                    tf = tb.text_frame
                    tf.word_wrap = True
                    p  = tf.paragraphs[0]
                    p.alignment = align
                    run = p.add_run()
                    run.text = text
                    run.font.size   = Pt(sz)
                    run.font.bold   = bold
                    run.font.italic = italic
                    run.font.color.rgb = RGBColor.from_string(col)

                # Cover
                s1 = prs.slides.add_slide(blank)
                s1.background.fill.solid()
                s1.background.fill.fore_color.rgb = RGBColor(15,15,24)
                rect(s1,0,0,0.22,7.5,"E5333A")
                txt(s1,"DOCTOR RESPONSE ANALYSIS",0.46,1.8,9,0.4,sz=10,bold=True,col="E5333A")
                txt(s1,st.session_state.question or "Qualitative Research Insights",0.46,2.3,9.5,1.1,sz=28,bold=True,col="FFFFFF")
                if st.session_state.bg:
                    txt(s1,st.session_state.bg[:110],0.46,3.75,9,0.6,sz=12,italic=True,col="9898B8")
                rect(s1,0.46,5.0,2.0,0.85,"1A2E5A")
                txt(s1,str(total),0.46,5.0,0.9,0.85,sz=28,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)
                txt(s1,"responses",1.4,5.0,1.1,0.85,sz=10,col="8888AA")
                rect(s1,2.8,5.0,2.0,0.85,"1A2E5A")
                txt(s1,str(len(buckets)),2.8,5.0,0.9,0.85,sz=28,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)
                txt(s1,"buckets",3.7,5.0,1.1,0.85,sz=10,col="8888AA")

                # Chart slide — horizontal bars drawn as shapes
                s2 = prs.slides.add_slide(blank)
                s2.background.fill.solid()
                s2.background.fill.fore_color.rgb = RGBColor(255,255,255)
                txt(s2,"RESPONSE DISTRIBUTION",0.4,0.25,8,0.3,sz=8,bold=True,col="E5333A")
                txt(s2,st.session_state.question or "Themes ranked by frequency",0.4,0.58,12.5,0.45,sz=17,bold=True,col="0F0F18")

                n = len(buckets)
                top = 1.2
                avail_h = 5.9
                rh = avail_h / n
                max_c = max(b["count"] for b in buckets) or 1

                for i,b in enumerate(buckets):
                    ch = COLORS[i%len(COLORS)].replace("#","")
                    bw = (b["count"]/max_c)*6.2
                    yp = top + i*rh
                    lbl = b["name"][:38]+("…" if len(b["name"])>38 else "")
                    txt(s2,lbl,0.3,yp+0.01,3.6,rh-0.04,sz=8,col="333344")
                    rect(s2,4.1,yp+rh*0.22,6.5,rh*0.52,"F4F3F0")
                    if bw>0.05:
                        rect(s2,4.1,yp+rh*0.22,bw,rh*0.52,ch)
                    txt(s2,f"{b['percentage']}% ({b['count']})",4.15+bw,yp+0.01,2.0,rh-0.04,sz=8,bold=True,col=ch)

                # Quote callouts
                top_qs = [(b,b["quotes"][0]) for b in buckets[:3] if b.get("quotes")]
                rect(s2,11.1,1.2,2.0,5.9,"F4F3F0")
                txt(s2,"KEY QUOTES",11.2,1.32,1.8,0.24,sz=7,bold=True,col="E5333A")
                qy_list = [1.72,3.07,4.42]
                for idx,(b,q) in enumerate(top_qs[:3]):
                    ch = COLORS[idx%len(COLORS)].replace("#","")
                    rect(s2,11.2,qy_list[idx],0.05,0.88,ch)
                    txt(s2,f'"{q[:85]}{"…" if len(q)>85 else ""}"',11.3,qy_list[idx]+0.02,1.6,0.7,sz=8,italic=True,col="333344")
                    txt(s2,f"— {b['name'][:25]}",11.3,qy_list[idx]+0.72,1.6,0.2,sz=7,bold=True,col=ch)

                # Full table slide
                s3 = prs.slides.add_slide(blank)
                s3.background.fill.solid()
                s3.background.fill.fore_color.rgb = RGBColor(244,243,240)
                txt(s3,"ALL BUCKETS",0.4,0.25,9,0.28,sz=8,bold=True,col="E5333A")
                txt(s3,f"All {len(buckets)} themes ranked",0.4,0.56,9,0.38,sz=15,bold=True,col="0F0F18")
                cxs = [0.4,0.85,3.7,4.35,4.9]
                cws = [0.42,2.82,0.62,0.52,7.8]
                hdrs3 = ["#","Bucket","Count","%","Theme"]
                hy = 1.1
                for j,(lbl,cx,cw) in enumerate(zip(hdrs3,cxs,cws)):
                    rect(s3,cx,hy,cw,0.32,"1A2E5A")
                    txt(s3,lbl,cx+0.04,hy+0.03,cw-0.08,0.26,sz=9,bold=True,col="FFFFFF")
                trh = min(0.36,(7.5-hy-0.5)/(len(buckets)+1))
                for i,b in enumerate(buckets):
                    ry = hy+0.32+i*trh
                    bg_c = "FFFFFF" if i%2==0 else "F9F8F6"
                    ch   = COLORS[i%len(COLORS)].replace("#","")
                    for cx,cw in zip(cxs,cws):
                        rect(s3,cx,ry,cw,trh,bg_c)
                    txt(s3,str(i+1),cxs[0]+0.04,ry+0.02,0.32,trh-0.04,sz=9,col="888888")
                    txt(s3,b["name"],cxs[1]+0.04,ry+0.02,cws[1]-0.08,trh-0.04,sz=9,bold=True,col=ch)
                    txt(s3,str(b["count"]),cxs[2]+0.04,ry+0.02,cws[2]-0.08,trh-0.04,sz=9,col="333344")
                    txt(s3,f"{b['percentage']}%",cxs[3]+0.04,ry+0.02,cws[3]-0.08,trh-0.04,sz=9,bold=True,col=ch)
                    txt(s3,b["theme"][:80],cxs[4]+0.04,ry+0.02,cws[4]-0.08,trh-0.04,sz=8,italic=True,col="555566")

                buf2 = io.BytesIO()
                prs.save(buf2)
                buf2.seek(0)
                st.download_button("⬇️ Download PowerPoint", data=buf2.getvalue(),
                    file_name="doctor_response_analysis.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True)

    st.divider()
    if st.button("↺ Start new analysis"):
        for k in ["buckets","total","tagged_df","col","bg","question","done"]:
            st.session_state[k] = None
        st.session_state.done = False
        st.session_state.buckets = []
        st.rerun()

st.markdown("---")
st.markdown("<div style='text-align:center;font-size:11px;color:#aaa'>Doctor Response Analyzer · Powered by Claude · No data stored</div>", unsafe_allow_html=True)
