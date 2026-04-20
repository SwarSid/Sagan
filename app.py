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
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import math

st.set_page_config(page_title="Doctor Response Analyzer", page_icon="⚕️", layout="wide")

COLORS = ["#1a2e5a","#e5333a","#f7a826","#16a34a","#7c3aed","#0891b2","#be185d","#b45309","#065f46","#1d4ed8","#9d174d","#0f766e"]
PALH   = [c.replace("#","") for c in COLORS]

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=Inter:wght@400;500;600&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif}
.banner{background:#1a2e5a;padding:1.6rem 2rem;border-radius:12px;margin-bottom:1.5rem;color:white}
.banner h1{font-family:'Syne',sans-serif;font-size:1.7rem;font-weight:800;margin:0 0 4px}
.banner p{margin:0;opacity:.6;font-size:13px}
.rtag{display:inline-block;background:#e5333a;color:white;font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;margin-bottom:8px}
.qb{border-left:3px solid #e5333a;border-radius:0 8px 8px 0;padding:9px 13px;margin:7px 0;background:#f5f4f1;font-size:13px;font-style:italic;color:#555;line-height:1.55}
.qbk{font-style:normal;font-size:10px;font-weight:700;color:#e5333a;margin-top:3px;text-transform:uppercase;letter-spacing:.4px}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="banner">
  <div class="rtag">⚕ AI Research Tool</div>
  <h1>Doctor Response Analyzer</h1>
  <p>Upload responses · AI bucketing · Excel + PowerPoint export · One link for your whole team</p>
</div>
""", unsafe_allow_html=True)

# API key
if "ANTHROPIC_API_KEY" in st.secrets:
    api_key = st.secrets["ANTHROPIC_API_KEY"]
else:
    api_key = st.sidebar.text_input("Anthropic API key", type="password", placeholder="sk-ant-...")
    if not api_key:
        st.sidebar.warning("Add API key to run analysis")

defaults = {"buckets":[],"total":0,"tagged_df":None,"col":None,"bg":"","question":"","done":False}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k] = v

# ── Step 1 ────────────────────────────────────────────────────────────────────
with st.expander("**Step 1 — Study context**", expanded=not st.session_state.done):
    c1,c2 = st.columns(2)
    bg       = c1.text_area("Study background", height=90,
                 placeholder="e.g. Post-consultation survey with oncologists discussing VORANIGO prescribing barriers...")
    question = c2.text_input("Question asked to doctors",
                 placeholder="e.g. What are the main barriers to prescribing VORANIGO?")

# ── Step 2 ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Step 2 — Upload Excel or CSV")
uploaded = st.file_uploader("Upload file", type=["xlsx","xls","csv"], label_visibility="collapsed")

df, sel_col = None, None
if uploaded:
    try:
        df = pd.read_csv(uploaded, dtype=str).fillna("") if uploaded.name.endswith(".csv") else pd.read_excel(uploaded, dtype=str).fillna("")
        st.success(f"✓ **{uploaded.name}** — {len(df)} rows, {len(df.columns)} columns")
        kw = ["response","answer","comment","text","doctor","feedback","reply","verbatim","transcript"]
        auto = next((c for c in df.columns if any(k in c.lower() for k in kw)), df.columns[0])
        sel_col = st.selectbox("Which column has the doctor responses?", df.columns.tolist(), index=df.columns.tolist().index(auto))
        if sel_col and len(df) > 0:
            st.caption(f'↳ *"{str(df[sel_col].iloc[0])[:160]}…"*')
    except Exception as e:
        st.error(f"Cannot read file: {e}")

# ── Step 3 ────────────────────────────────────────────────────────────────────
st.markdown("### 🤖 Step 3 — Run AI Analysis")

def is_doctor_response(text):
    if not text or len(str(text).strip()) < 15:
        return False
    t = str(text).strip()
    for pat in [r'^\[GT\]\s*SPEAKER_A', r'^SPEAKER_A\s*:', r'^AI Moderator\s*:',
                r'^Moderator\s*:', r'^Interviewer\s*:', r'Hello,?\s*doctor',
                r'Thank you for taking the time', r'Our discussion (focuses|will focus)']:
        if re.search(pat, t, re.IGNORECASE):
            return False
    if re.match(r'^[\dX_\-]+$', t): return False
    if re.match(r'^Q\d+_', t): return False
    return True

if st.button("▶ Run AI Analysis", disabled=not (df is not None and sel_col and api_key), type="primary"):
    all_rows = df[sel_col].astype(str).tolist()
    doctor_indices = [(i, row) for i, row in enumerate(all_rows) if is_doctor_response(row)]
    if len(doctor_indices) < 3:
        st.error(f"Only {len(doctor_indices)} valid doctor responses found. Check your column selection.")
    else:
        responses = [r for _,r in doctor_indices]
        orig_indices = [i for i,_ in doctor_indices]
        st.info(f"Found **{len(responses)} doctor responses** (filtered out {len(all_rows)-len(responses)} moderator/metadata rows)")
        with st.spinner(f"Analyzing {len(responses)} responses… 20–40 seconds"):
            sys_p = """You are an expert qualitative researcher for medical/pharma studies.
Return ONLY valid JSON — no markdown, no backticks, no explanation.
{
  "totalResponses": <integer>,
  "buckets": [
    {"name":"<4-7 word insight-forward label>","count":<int>,"percentage":<float 1dp>,
     "theme":"<one sentence>","quotes":["<verbatim 10-25 word fragment>","<another>","<third>"],
     "responseIndices":[<0-based ints within filtered list>]}
  ]
}
CRITICAL: 10-12 buckets, every index 0..(n-1) in exactly ONE bucket, no uncategorized, counts sum to totalResponses, percentages sum to 100.0, order by count desc."""
            sample = responses[:120]
            user_msg = (f"Background: {bg or 'Medical research'}\nQuestion: {question or 'Doctor experiences'}\n\n"
                        f"Doctor responses ({len(sample)} total):\n" + "\n".join(f"[{i}] {r}" for i,r in enumerate(sample)))
            try:
                client = anthropic.Anthropic(api_key=api_key)
                msg = client.messages.create(model="claude-sonnet-4-5", max_tokens=2000,
                    system=sys_p, messages=[{"role":"user","content":user_msg}])
                raw = re.sub(r'^```json\s*','',msg.content[0].text.strip(),flags=re.I)
                raw = re.sub(r'^```','',raw).replace('```','').strip()
                parsed = json.loads(raw)
                buckets = parsed["buckets"]; total_r = parsed["totalResponses"]
                tagged = df.copy(); tagged["Bucket"] = "N/A (moderator/metadata)"
                for b in buckets:
                    for fi in (b.get("responseIndices") or []):
                        if fi < len(orig_indices): tagged.at[orig_indices[fi], "Bucket"] = b["name"]
                st.session_state.update({"buckets":buckets,"total":total_r,"tagged_df":tagged,
                    "col":sel_col,"bg":bg,"question":question,"done":True})
                st.rerun()
            except Exception as e:
                st.error(f"Failed: {e}")

# ── Results ───────────────────────────────────────────────────────────────────
if st.session_state.done and st.session_state.buckets:
    B = st.session_state.buckets
    total = st.session_state.total
    tagged_df = st.session_state.tagged_df

    st.divider()
    st.markdown("## Results")
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Doctor Responses", total)
    m2.metric("Buckets", len(B))
    m3.metric("Top Bucket", f"{B[0]['percentage']}%")
    m4.metric("Major Themes ≥10%", sum(1 for b in B if b["percentage"]>=10))

    t1,t2,t3,t4 = st.tabs(["📋 Table","📊 Visualize","💬 Quotes","🏷️ Tagged"])

    with t1:
        st.dataframe(pd.DataFrame([{"#":i+1,"Bucket":b["name"],"Count":b["count"],
            "%":f"{b['percentage']}%","Theme":b["theme"],
            "Sample quote":b["quotes"][0] if b.get("quotes") else ""} for i,b in enumerate(B)]),
            use_container_width=True, hide_index=True)

    with t2:
        viz = st.selectbox("Choose style", ["Horizontal bar chart + quotes","Stat cards per bucket",
            "Ranked driver list","Donut chart","Summary table"], key="viz_pick")
        names=[b["name"] for b in B]; counts=[b["count"] for b in B]
        pcts=[b["percentage"] for b in B]; colors=[COLORS[i%len(COLORS)] for i in range(len(B))]

        if "bar" in viz:
            fig = go.Figure(go.Bar(x=counts,y=names,orientation='h',marker_color=colors,
                text=[f"{p}% ({c})" for c,p in zip(counts,pcts)],textposition='outside'))
            fig.update_layout(height=max(300,len(B)*42+100),margin=dict(l=20,r=140,t=10,b=20),
                yaxis=dict(autorange="reversed"),plot_bgcolor='white',paper_bgcolor='white',
                xaxis_title="Number of responses",font=dict(family="Inter",size=12))
            fig.update_xaxes(showgrid=True,gridcolor='#f0f0f0')
            st.plotly_chart(fig,use_container_width=True)
            for b in B[:3]:
                if b.get("quotes"):
                    st.markdown(f'<div class="qb">"{b["quotes"][0]}"<div class="qbk">{b["name"]} · {b["percentage"]}%</div></div>',unsafe_allow_html=True)
        elif "cards" in viz:
            cols=st.columns(3)
            for i,b in enumerate(B):
                c=COLORS[i%len(COLORS)]
                q=f'<div style="font-size:11px;font-style:italic;color:#888;border-left:2px solid {c};padding-left:7px;margin-top:6px">"{b["quotes"][0]}"</div>' if b.get("quotes") else ""
                with cols[i%3]:
                    st.markdown(f'<div style="background:white;border:1px solid rgba(0,0,0,0.08);border-radius:10px;padding:13px;border-left:4px solid {c};margin-bottom:8px"><div style="font-size:26px;font-weight:800;color:{c};line-height:1;margin-bottom:4px">{b["percentage"]}%</div><div style="font-size:13px;font-weight:600;margin-bottom:3px">{b["name"]}</div><div style="font-size:12px;color:#6b6b78;margin-bottom:6px">{b["theme"]}</div>{q}</div>',unsafe_allow_html=True)
        elif "ranked" in viz:
            for i,b in enumerate(B):
                c=COLORS[i%len(COLORS)]
                impact="🔴 Most impactful" if i<int(len(B)*0.3) else "🟡 Moderate" if i<int(len(B)*0.7) else "🟢 Least impactful"
                if i==0 or i==int(len(B)*0.3) or i==int(len(B)*0.7): st.markdown(f"**{impact}**")
                q=f'<div style="font-size:11px;font-style:italic;color:#888;margin-top:5px;border-left:2px solid #e0e0e8;padding-left:7px">"{b["quotes"][0]}"</div>' if b.get("quotes") else ""
                st.markdown(f'<div style="display:flex;align-items:flex-start;gap:12px;padding:11px 14px;background:white;border:1px solid rgba(0,0,0,0.08);border-radius:10px;margin-bottom:6px;border-left:4px solid {c}"><div style="font-size:18px;font-weight:800;color:{c};min-width:26px;text-align:center">{i+1}</div><div style="flex:1"><div style="font-size:13px;font-weight:600">{b["name"]}</div><div style="font-size:11px;color:#6b6b78">{b["theme"]}</div>{q}</div><div style="font-size:20px;font-weight:800;color:{c}">{b["percentage"]}%</div></div>',unsafe_allow_html=True)
        elif "donut" in viz:
            fig=go.Figure(go.Pie(labels=names,values=counts,hole=0.52,marker_colors=colors,textinfo='percent'))
            fig.update_layout(height=420,margin=dict(l=20,r=20,t=10,b=20),showlegend=True,paper_bgcolor='white')
            st.plotly_chart(fig,use_container_width=True)
        else:
            for i,b in enumerate(B):
                c=COLORS[i%len(COLORS)]
                q=f'<em>"{b["quotes"][0]}"</em>' if b.get("quotes") else "—"
                st.markdown(f'<div style="display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center;padding:10px 12px;background:white;border:1px solid rgba(0,0,0,0.07);border-radius:8px;margin-bottom:5px;border-left:4px solid {c}"><div style="font-size:14px;font-weight:800;color:{c};min-width:28px;text-align:center">{i+1}</div><div><div style="font-size:13px;font-weight:600">{b["name"]}</div><div style="font-size:11px;color:#888;margin-top:2px">{b["theme"]}</div><div style="font-size:11px;color:#aaa;margin-top:3px">{q}</div></div><div style="font-size:22px;font-weight:800;color:{c}">{b["percentage"]}%</div></div>',unsafe_allow_html=True)

    with t3:
        for b in B:
            for q in (b.get("quotes") or []):
                st.markdown(f'<div class="qb">"{q}"<div class="qbk">{b["name"]} · {b["percentage"]}%</div></div>',unsafe_allow_html=True)

    with t4:
        na_count=(tagged_df["Bucket"]=="N/A (moderator/metadata)").sum()
        if na_count>0: st.info(f"{na_count} rows are moderator questions or metadata — not doctor responses.")
        st.dataframe(tagged_df,use_container_width=True,height=380)

    # ── Export ────────────────────────────────────────────────────────────────
    st.divider()
    st.markdown("## 📥 Export")
    e1,e2 = st.columns(2)

    with e1:
        st.markdown("#### 📊 Excel")
        if st.button("Generate Excel", type="primary", use_container_width=True):
            with st.spinner("Building Excel…"):
                wb = Workbook()
                ws1 = wb.active; ws1.title = "Responses (Bucketed)"
                for j,h in enumerate(tagged_df.columns,1):
                    c=ws1.cell(row=1,column=j,value=h)
                    c.font=Font(bold=True,color="FFFFFF",name="Calibri",size=10)
                    c.fill=PatternFill("solid",fgColor="1A2E5A")
                    c.alignment=Alignment(horizontal="center")
                for i,row in tagged_df.iterrows():
                    for j,val in enumerate(row,1):
                        cell=ws1.cell(row=i+2,column=j,value=str(val))
                        cell.font=Font(name="Calibri",size=10)
                        if j==len(tagged_df.columns):
                            bkt=next((b for b in B if b["name"]==str(val)),None)
                            if bkt:
                                idx=B.index(bkt); hx=PALH[idx%len(PALH)]
                                cell.fill=PatternFill("solid",fgColor=lighten(hx, 0.85))
                                cell.font=Font(name="Calibri",size=10,color=hx,bold=True)
                for cc in ws1.columns:
                    ws1.column_dimensions[get_column_letter(cc[0].column)].width=min(max(len(str(c.value or "")) for c in cc)+2,60)
                ws2=wb.create_sheet("Bucket Summary")
                for j,h in enumerate(["Rank","Bucket","Count","%","Theme","Quote 1","Quote 2","Quote 3"],1):
                    c=ws2.cell(row=1,column=j,value=h)
                    c.font=Font(bold=True,color="FFFFFF",name="Calibri",size=10)
                    c.fill=PatternFill("solid",fgColor="E5333A")
                    c.alignment=Alignment(horizontal="center")
                for i,b in enumerate(B):
                    qs=b.get("quotes",[])
                    for j,v in enumerate([i+1,b["name"],b["count"],f"{b['percentage']}%",b["theme"],
                        qs[0] if len(qs)>0 else "",qs[1] if len(qs)>1 else "",qs[2] if len(qs)>2 else ""],1):
                        c=ws2.cell(row=i+2,column=j,value=v)
                        c.font=Font(name="Calibri",size=10)
                        c.alignment=Alignment(wrap_text=True,vertical="top")
                for col_letter,width in zip(["A","B","C","D","E","F","G","H"],[6,38,8,10,55,50,50,50]):
                    ws2.column_dimensions[col_letter].width=width
                buf=io.BytesIO(); wb.save(buf); buf.seek(0)
                st.download_button("⬇️ Download Excel",data=buf.getvalue(),
                    file_name="doctor_responses_bucketed.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    with e2:
        st.markdown("#### 📑 PowerPoint — 4 slides auto-generated")
        st.caption("Slide 1: Big stat infographic · Slide 2: Numbered arch cards · Slide 3: Bar chart + quotes · Slide 4: Full table")
        if st.button("Generate PowerPoint", type="primary", use_container_width=True):
            with st.spinner("Building 4-slide deck…"):

                prs = Presentation()
                prs.slide_width  = Inches(13.33)
                prs.slide_height = Inches(7.5)
                blank = prs.slide_layouts[6]

                # ── helpers ──────────────────────────────────────────────────
                def bg_color(slide, hex_str):
                    slide.background.fill.solid()
                    h=hex_str[:6]
                    r,g,b=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
                    slide.background.fill.fore_color.rgb=RGBColor(r,g,b)

                def lighten(hex6, factor=0.85):
                    h=str(hex6)[:6]
                    r,g,b=int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
                    r=int(r+(255-r)*factor); g=int(g+(255-g)*factor); b=int(b+(255-b)*factor)
                    return f'{min(r,255):02X}{min(g,255):02X}{min(b,255):02X}'

                def R(slide,x,y,w,h,fill_hex):
                    hex6=str(fill_hex)[:6]
                    sp=slide.shapes.add_shape(1,Inches(x),Inches(y),Inches(w),Inches(h))
                    sp.fill.solid()
                    sp.fill.fore_color.rgb=RGBColor.from_string(hex6)
                    sp.line.fill.background()
                    return sp

                def oval(slide,x,y,w,h,fill_hex):
                    hex6=str(fill_hex)[:6]
                    sp=slide.shapes.add_shape(9,Inches(x),Inches(y),Inches(w),Inches(h))
                    sp.fill.solid()
                    sp.fill.fore_color.rgb=RGBColor.from_string(hex6)
                    sp.line.fill.background()
                    return sp

                def T(slide,text,x,y,w,h,sz=12,bold=False,col="000000",italic=False,align=PP_ALIGN.LEFT,wrap=True):
                    tb=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
                    tb.word_wrap=wrap; tf=tb.text_frame; tf.word_wrap=wrap
                    p=tf.paragraphs[0]; p.alignment=align
                    run=p.add_run(); run.text=str(text)
                    run.font.size=Pt(sz); run.font.bold=bold; run.font.italic=italic
                    run.font.color.rgb=RGBColor.from_string(col)

                def add_line(slide,x1,y1,x2,y2,col_hex,width_pt=1.5):
                    from pptx.util import Inches, Pt
                    import pptx.oxml as oxml
                    connector = slide.shapes.add_connector(1,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
                    connector.line.color.rgb = RGBColor.from_string(col_hex)
                    connector.line.width = Pt(width_pt)

                question_txt = st.session_state.question or "Key Themes from Doctor Responses"
                bg_txt = st.session_state.bg or ""

                # ════════════════════════════════════════════════════════════
                # SLIDE 1 — Bold stat infographic  (dark bg, big % cards)
                # Like the "pentagon/arch" style but as bold stat blocks
                # ════════════════════════════════════════════════════════════
                s1 = prs.slides.add_slide(blank)
                bg_color(s1,"0F0F1E")

                # Left red accent
                R(s1,0,0,0.18,7.5,"E5333A")

                # Title
                T(s1,"PHYSICIAN RESEARCH INSIGHTS",0.38,0.28,9,0.32,sz=9,bold=True,col="E5333A")
                T(s1,question_txt,0.38,0.65,9.5,0.95,sz=26,bold=True,col="FFFFFF")
                if bg_txt:
                    T(s1,bg_txt[:110]+(("…") if len(bg_txt)>110 else ""),0.38,1.7,9,0.5,sz=11,italic=True,col="8888AA")

                # Top 5 arch-style bold cards
                top5 = B[:5]
                card_colors = ["E5333A","1a2e5a","F7A826","16A34A","7C3AED"]
                cw,ch = 2.42, 2.1
                gap   = 0.14
                total_row_w = len(top5)*(cw+gap)-gap
                sx = (13.33 - total_row_w)/2

                for i,b in enumerate(top5):
                    bx = sx + i*(cw+gap)
                    by = 2.4
                    col_hex = card_colors[i%len(card_colors)]

                    # Card background
                    R(s1,bx,by,cw,ch,col_hex)

                    # Big percentage
                    T(s1,f"{b['percentage']}%",bx,by+0.1,cw,0.75,
                      sz=36,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)

                    # Count pill
                    R(s1,bx+cw/2-0.45,by+0.82,0.9,0.28,"FFFFFF")
                    T(s1,f"n = {b['count']}",bx+cw/2-0.45,by+0.82,0.9,0.28,
                      sz=9,bold=True,col=col_hex,align=PP_ALIGN.CENTER)

                    # Bucket name
                    nm = b["name"][:30]+("…" if len(b["name"])>30 else "")
                    T(s1,nm,bx+0.1,by+1.18,cw-0.2,0.68,
                      sz=10,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)

                # Remaining buckets as smaller strip
                rest = B[5:12]
                if rest:
                    rw = (13.33-0.56)/len(rest)
                    for i,b in enumerate(rest):
                        rx = 0.38 + i*rw
                        col_hex = PALH[(i+5)%len(PALH)]
                        R(s1,rx,4.72,rw-0.1,1.62,col_hex)
                        T(s1,f"{b['percentage']}%",rx,4.77,rw-0.1,0.55,
                          sz=20,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)
                        nm2 = b["name"][:20]+("…" if len(b["name"])>20 else "")
                        T(s1,nm2,rx+0.04,5.32,rw-0.18,0.78,sz=8,col="CCCCEE",align=PP_ALIGN.CENTER)

                T(s1,f"n = {total} physician responses  ·  {len(B)} themes identified",
                  0,6.9,13.33,0.3,sz=9,col="444460",align=PP_ALIGN.CENTER)

                # ════════════════════════════════════════════════════════════
                # SLIDE 2 — Numbered steps infographic
                # Circles with numbers + coloured boxes with name/% /quote
                # Inspired by the numbered arch/steps template layout
                # ════════════════════════════════════════════════════════════
                s2 = prs.slides.add_slide(blank)
                bg_color(s2,"FFFFFF")

                T(s2,"TOP THEMES — RANKED",0.5,0.2,12,0.3,sz=8,bold=True,col="E5333A")
                T(s2,question_txt,0.5,0.52,12.33,0.48,sz=18,bold=True,col="0F0F18")

                # Draw top 5 as numbered vertical steps with connector line
                top5_s2 = B[:5]
                step_colors = ["E5333A","F7A826","1A2E5A","16A34A","7C3AED"]
                circle_r = 0.42
                step_x_circle = 0.7
                step_x_box    = 1.5
                box_w         = 5.5
                box_h         = 0.95
                step_y_start  = 1.2
                step_gap      = 1.12

                for i,b in enumerate(top5_s2):
                    col_hex = step_colors[i%len(step_colors)]
                    cy = step_y_start + i*step_gap

                    # Connector line between circles (except last)
                    if i < len(top5_s2)-1:
                        add_line(s2, step_x_circle+circle_r/2,
                                    cy+circle_r,
                                    step_x_circle+circle_r/2,
                                    cy+step_gap,
                                    "DDDDDD", 1.0)

                    # Circle
                    oval(s2,step_x_circle,cy,circle_r,circle_r,col_hex)
                    T(s2,str(i+1),step_x_circle,cy,circle_r,circle_r,
                      sz=16,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)

                    # Box
                    R(s2,step_x_box,cy-0.05,box_w,box_h,lighten(col_hex, 0.88))
                    R(s2,step_x_box,cy-0.05,0.07,box_h,col_hex)

                    nm = b["name"]
                    T(s2,nm,step_x_box+0.18,cy,box_w-0.5,0.35,sz=11,bold=True,col="0F0F18")
                    T(s2,b["theme"][:65],step_x_box+0.18,cy+0.34,box_w-0.5,0.34,sz=9,italic=True,col="555566")

                    # Percentage badge on right
                    R(s2,step_x_box+box_w+0.1,cy+0.1,0.85,0.55,col_hex)
                    T(s2,f"{b['percentage']}%",step_x_box+box_w+0.1,cy+0.1,0.85,0.55,
                      sz=14,bold=True,col="FFFFFF",align=PP_ALIGN.CENTER)

                # Right panel: remaining buckets 6-12 as mini list
                rx_panel = 8.0
                T(s2,"OTHER THEMES",rx_panel,1.1,4.8,0.3,sz=8,bold=True,col="E5333A")

                rest_s2 = B[5:]
                mini_h  = min(0.62, 5.5/max(len(rest_s2),1))
                for i,b in enumerate(rest_s2):
                    col_hex = PALH[(i+5)%len(PALH)]
                    my = 1.5 + i*mini_h
                    R(s2,rx_panel,my,4.8,mini_h-0.06,lighten(col_hex, 0.92))
                    R(s2,rx_panel,my,0.06,mini_h-0.06,col_hex)
                    T(s2,b["name"][:38],rx_panel+0.15,my+0.03,3.5,mini_h-0.12,sz=9,bold=True,col="222233")
                    T(s2,f"{b['percentage']}%",rx_panel+3.65,my+0.03,1.0,mini_h-0.12,
                      sz=11,bold=True,col=col_hex,align=PP_ALIGN.RIGHT)

                # ════════════════════════════════════════════════════════════
                # SLIDE 3 — Horizontal bar chart + quote callouts
                # ════════════════════════════════════════════════════════════
                s3 = prs.slides.add_slide(blank)
                bg_color(s3,"FFFFFF")

                T(s3,"RESPONSE DISTRIBUTION",0.5,0.22,8,0.28,sz=8,bold=True,col="E5333A")
                T(s3,question_txt,0.5,0.52,12.5,0.48,sz=17,bold=True,col="0F0F18")
                T(s3,f"All {len(B)} themes  ·  n = {total} responses",
                  0.5,1.02,8,0.26,sz=9,italic=True,col="888899")

                n=len(B); rh=5.3/n; max_c=max(b["count"] for b in B) or 1
                for i,b in enumerate(B):
                    col_hex=PALH[i%len(PALH)]
                    bw=(b["count"]/max_c)*5.6
                    yp=1.38+i*rh
                    lbl=b["name"][:36]+("…" if len(b["name"])>36 else "")
                    T(s3,lbl,0.5,yp+0.01,3.4,rh-0.04,sz=8,col="222233")
                    R(s3,4.0,yp+rh*0.18,5.8,rh*0.6,"F4F3F0")
                    if bw>0.05: R(s3,4.0,yp+rh*0.18,bw,rh*0.6,col_hex)
                    T(s3,f"{b['percentage']}%  n={b['count']}",
                      4.08+bw,yp+0.01,2.2,rh-0.04,sz=8,bold=True,col=col_hex)

                # Quote panel
                R(s3,10.08,1.38,3.08,5.82,"F7F6F3")
                T(s3,"KEY QUOTES",10.22,1.5,2.8,0.24,sz=7,bold=True,col="E5333A")
                top_qs=[(b,b["quotes"][0]) for b in B[:3] if b.get("quotes")]
                for idx,(b,q) in enumerate(top_qs[:3]):
                    col_hex=PALH[idx%len(PALH)]
                    qy=1.88+idx*1.38
                    R(s3,10.22,qy,0.05,0.85,col_hex)
                    T(s3,f'"{q[:88]}{"…" if len(q)>88 else ""}"',
                      10.32,qy+0.02,2.6,0.68,sz=8,italic=True,col="333344")
                    T(s3,f"— {b['name'][:26]}",10.32,qy+0.7,2.6,0.2,sz=7,bold=True,col=col_hex)

                # ════════════════════════════════════════════════════════════
                # SLIDE 4 — Full colour-coded table
                # ════════════════════════════════════════════════════════════
                s4 = prs.slides.add_slide(blank)
                bg_color(s4,"F7F6F3")

                T(s4,"FULL BREAKDOWN",0.5,0.22,9,0.28,sz=8,bold=True,col="E5333A")
                T(s4,f"All {len(B)} themes ranked by frequency",0.5,0.54,9,0.38,sz=15,bold=True,col="0F0F18")
                T(s4,f"n = {total} physician responses",0.5,0.96,5,0.26,sz=9,italic=True,col="888899")

                cxs=[0.5,0.98,3.88,4.56,5.18]; cws=[0.44,2.86,0.64,0.58,7.54]
                hy=1.3
                for lbl,cx,cw in zip(["#","Bucket","Count","%","Core theme"],cxs,cws):
                    R(s4,cx,hy,cw,0.36,"1A2E5A")
                    T(s4,lbl,cx+0.05,hy+0.04,cw-0.1,0.28,sz=9,bold=True,col="FFFFFF")

                trh=min(0.37,(7.5-hy-0.5)/(len(B)+1))
                for i,b in enumerate(B):
                    ry=hy+0.36+i*trh
                    bg_c="FFFFFF" if i%2==0 else "F2F1EE"
                    col_hex=PALH[i%len(PALH)]
                    for cx,cw in zip(cxs,cws): R(s4,cx,ry,cw,trh,bg_c)
                    T(s4,str(i+1),cxs[0]+0.05,ry+0.03,0.34,trh-0.06,sz=9,col="888888")
                    T(s4,b["name"],cxs[1]+0.05,ry+0.03,cws[1]-0.1,trh-0.06,sz=9,bold=True,col=col_hex)
                    T(s4,str(b["count"]),cxs[2]+0.05,ry+0.03,cws[2]-0.1,trh-0.06,sz=9,col="333344")
                    T(s4,f"{b['percentage']}%",cxs[3]+0.05,ry+0.03,cws[3]-0.1,trh-0.06,sz=9,bold=True,col=col_hex)
                    T(s4,b["theme"][:85],cxs[4]+0.05,ry+0.03,cws[4]-0.1,trh-0.06,sz=8,italic=True,col="555566")

                buf2=io.BytesIO(); prs.save(buf2); buf2.seek(0)
                st.download_button("⬇️ Download PowerPoint",data=buf2.getvalue(),
                    file_name="doctor_response_analysis.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True)

    st.divider()
    if st.button("↺ New analysis"):
        for k,v in defaults.items(): st.session_state[k] = v
        st.rerun()

st.markdown("---")
st.markdown("<div style='text-align:center;font-size:11px;color:#aaa'>Doctor Response Analyzer · Powered by Claude · No data stored</div>",unsafe_allow_html=True)
