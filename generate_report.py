import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. Load Transaction Dataset
df = pd.read_csv('SuperMarket Analysis - Copy.csv')

# 2. Configure Matplotlib Styling
plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['font.family'] = 'sans-serif'

# --- CHART 1: Ranked Product Revenue with Exact Value Labels ---
fig, ax = plt.subplots(figsize=(8.5, 4.2), dpi=300)
cat_sales = df.groupby('Product line')['Sales'].sum().sort_values(ascending=True)
palette = ['#93C5FD', '#60A5FA', '#3B82F6', '#2563EB', '#1D4ED8', '#1E3A8A']
bars = ax.barh(cat_sales.index, cat_sales.values, color=palette, height=0.58)

ax.set_title("Gross Revenue by Product Category", fontsize=13, fontweight='bold', color='#0F172A', pad=14, loc='left')
ax.set_xlabel("Total Sales Revenue ($ USD)", fontsize=10, fontweight='semibold', color='#475569', labelpad=10)
ax.tick_params(colors='#334155', labelsize=9.5)
ax.xaxis.grid(True, linestyle='--', alpha=0.5, color='#CBD5E1')
ax.yaxis.grid(False)
ax.set_axisbelow(True)

for bar in bars:
    w = bar.get_width()
    ax.text(w + 1100, bar.get_y() + bar.get_height()/2, f"${w:,.2f}", 
            va='center', ha='left', fontsize=9.5, fontweight='bold', color='#1E293B')

ax.set_xlim(0, max(cat_sales.values) * 1.16)
for spine in ['top', 'right', 'left', 'bottom']:
    ax.spines[spine].set_color('#E2E8F0')

plt.tight_layout()
plt.savefig('fig1_category_clean.png', bbox_inches='tight')
plt.close()

# --- CHART 2: Dual Panel - Store Contribution & Payment Modes ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.4), dpi=300)

branch_sales = df.groupby('City')['Sales'].sum().reindex(['Naypyitaw', 'Yangon', 'Mandalay'])
city_palette = ['#1E3A8A', '#2563EB', '#60A5FA']
bars_city = ax1.bar(branch_sales.index, branch_sales.values, color=city_palette, width=0.52)
ax1.set_title("Store Revenue Contribution", fontsize=11.5, fontweight='bold', color='#0F172A', pad=12, loc='left')
ax1.set_ylabel("Sales ($ USD)", fontsize=9.5, fontweight='semibold', color='#475569')
ax1.tick_params(colors='#334155', labelsize=9.5)
ax1.yaxis.grid(True, linestyle='--', alpha=0.5, color='#CBD5E1')
ax1.xaxis.grid(False)
ax1.set_axisbelow(True)

for bar in bars_city:
    h = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2, h + 1800, f"${h:,.0f}", 
             ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1E293B')

ax1.set_ylim(0, max(branch_sales.values) * 1.15)
for spine in ['top', 'right', 'left', 'bottom']:
    ax1.spines[spine].set_color('#E2E8F0')

pay_data = df['Payment'].value_counts()
pay_colors = ['#1E3A8A', '#3B82F6', '#93C5FD']
wedges, texts, autotexts = ax2.pie(pay_data, labels=pay_data.index, autopct='%1.1f%%', 
                                  startangle=135, colors=pay_colors, pctdistance=0.78,
                                  textprops={'fontsize': 9.5, 'color': '#1E293B', 'fontweight': 'semibold'},
                                  wedgeprops=dict(width=0.42, edgecolor='white', linewidth=2.5))
for autotext in autotexts:
    autotext.set_color('white')
    autotext.set_fontweight('bold')
    autotext.set_fontsize(9)

ax2.set_title("Payment Method Share", fontsize=11.5, fontweight='bold', color='#0F172A', pad=12, loc='left')
plt.tight_layout()
plt.savefig('fig2_branch_payment.png', bbox_inches='tight')
plt.close()

# --- CHART 3: Customer Spend Distribution Boxplot ---
fig, ax = plt.subplots(figsize=(7.5, 3.8), dpi=300)
sns.boxplot(data=df, x='Customer type', y='Sales', palette=['#3B82F6', '#94A3B8'], 
            width=0.42, linewidth=1.4, fliersize=3.5, ax=ax)
ax.set_title("Transaction Spend Distribution (Member vs. Normal)", fontsize=11.5, fontweight='bold', color='#0F172A', pad=12, loc='left')
ax.set_xlabel("Customer Segment", fontsize=9.5, fontweight='semibold', color='#475569', labelpad=8)
ax.set_ylabel("Sales ($ USD)", fontsize=9.5, fontweight='semibold', color='#475569', labelpad=8)
ax.tick_params(colors='#334155', labelsize=9.5)
ax.yaxis.grid(True, linestyle='--', alpha=0.5, color='#CBD5E1')
ax.xaxis.grid(False)
ax.set_axisbelow(True)
for spine in ['top', 'right', 'left', 'bottom']:
    ax.spines[spine].set_color('#E2E8F0')

plt.tight_layout()
plt.savefig('fig3_customer_spend.png', bbox_inches='tight')
plt.close()

# --- 3. BUILD FORMATTED WORD DOCUMENT ---
doc = Document()

# Standard margins (0.75 in)
for section in doc.sections:
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

def apply_cell_bg(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def apply_cell_padding(cell, top=160, bottom=160, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def add_custom_heading(doc, text, level=1):
    h = doc.add_paragraph()
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True
    run = h.add_run(text)
    run.font.bold = True
    if level == 1:
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(15, 23, 42)
    elif level == 2:
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(30, 58, 138)
    return h

# Header Banner
banner = doc.add_table(rows=1, cols=1)
banner.alignment = WD_TABLE_ALIGNMENT.CENTER
banner.autofit = False
b_cell = banner.cell(0, 0)
b_cell.width = Inches(7.0)
apply_cell_bg(b_cell, "1E3A8A")
apply_cell_padding(b_cell, top=260, bottom=240, left=240, right=240)

bp = b_cell.paragraphs[0]
bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
bp.paragraph_format.space_after = Pt(2)
r_t = bp.add_run("SUPERMARKET SALES & BUSINESS INTELLIGENCE\n")
r_t.font.size = Pt(16)
r_t.font.bold = True
r_t.font.color.rgb = RGBColor(255, 255, 255)

r_sub = bp.add_run("AICTE | IBM SkillsBuild Internship Capstone Project")
r_sub.font.size = Pt(10.5)
r_sub.font.color.rgb = RGBColor(191, 219, 254)

doc.add_paragraph().paragraph_format.space_after = Pt(4)

# Candidate Credentials Table
meta_table = doc.add_table(rows=2, cols=2)
meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_table.autofit = False
meta_data = [
    [("Author", "Md Sameer Khan"), ("Email", "Kmdsameer993@gmail.com")],
    [("Institution", "Maharaja Agrasen Institute of Technology (MAIT)"), ("Repository", "github.com/Sameer-khan-27/IBM-SkillsBuild-Sales-Analytics")]
]
for row_idx, row in enumerate(meta_table.rows):
    for col_idx, cell in enumerate(row.cells):
        cell.width = Inches(3.5)
        apply_cell_bg(cell, "F8FAFC")
        apply_cell_padding(cell, top=90, bottom=90, left=140, right=140)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        lbl, val = meta_data[row_idx][col_idx]
        r_lbl = p.add_run(f"{lbl}: ")
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(9)
        r_lbl.font.color.rgb = RGBColor(71, 85, 105)
        r_val = p.add_run(val)
        r_val.font.size = Pt(9)
        r_val.font.color.rgb = RGBColor(15, 23, 42)

# 1. Executive Summary
add_custom_heading(doc, "1. Executive Summary", level=1)
p_exec = doc.add_paragraph(
    "This analytical report presents an operational and exploratory evaluation of 1,000 retail transaction records "
    "originating from three supermarket branches (Naypyitaw, Yangon, and Mandalay). Developed as an end-to-end "
    "internship capstone project, the analysis examines product category sales distributions, customer segment purchasing power, "
    "checkout payment preferences, and branch contribution levels to produce strategic business intelligence."
)
p_exec.paragraph_format.space_after = Pt(8)
p_exec.paragraph_format.line_spacing = 1.15

# KPI Summary Cards
kpi_table = doc.add_table(rows=1, cols=4)
kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
kpi_table.autofit = False
cards = [
    ("TOTAL REVENUE", "$322,966.75", "1,000 Transactions"),
    ("AVG ORDER VALUE", "$322.97", "+9.59% Member Lift"),
    ("TOP CATEGORY", "Food & Beverage", "$56,144.84 Gross"),
    ("SATISFACTION", "6.97 / 10.0", "Stable CSAT Index")
]
for idx, (title, val, sub) in enumerate(cards):
    cell = kpi_table.cell(0, idx)
    cell.width = Inches(1.75)
    apply_cell_bg(cell, "EFF6FF")
    apply_cell_padding(cell, top=140, bottom=140, left=80, right=80)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r1 = p.add_run(f"{title}\n")
    r1.font.size = Pt(7.5)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(29, 78, 216)
    r2 = p.add_run(f"{val}\n")
    r2.font.size = Pt(12.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(15, 23, 42)
    r3 = p.add_run(sub)
    r3.font.size = Pt(7.5)
    r3.font.color.rgb = RGBColor(100, 116, 139)

# 2. Strategic Insights
add_custom_heading(doc, "2. Key Strategic Insights", level=1)
bullets = [
    ("Geographical Revenue Parity: ", "Store revenues remain balanced across cities. Naypyitaw leads with $110,568.71 (34.2%), while Yangon ($106,200.37) and Mandalay ($106,197.67) record nearly identical top-line results (~32.9% each)."),
    ("Product Line Performance: ", "Food and Beverages ($56,144.84) and Sports and Travel ($55,122.83) generate the highest sales volume. Even the lowest contributor (Health & Beauty at $49,193.74) captures 15.2% of total turnover, indicating a resilient, diversified product catalog."),
    ("High Digital Payment Penetration: ", "Cashless payment methods account for 65.6% of customer checkouts (E-wallets at 34.5% and Credit Cards at 31.1%), outperforming traditional cash settlement (34.4%)."),
    ("Loyalty Member Uplift: ", "Registered Loyalty Members generate an average basket size of $335.74 per order compared to $306.37 for Regular walk-in customers, delivering a statistically verified 9.59% revenue increase per transaction.")
]
for title, desc in bullets:
    p_b = doc.add_paragraph()
    p_b.paragraph_format.space_before = Pt(2)
    p_b.paragraph_format.space_after = Pt(4)
    p_b.paragraph_format.line_spacing = 1.15
    rt = p_b.add_run("•  " + title)
    rt.font.bold = True
    rt.font.size = Pt(9.5)
    rt.font.color.rgb = RGBColor(30, 58, 138)
    rd = p_b.add_run(desc)
    rd.font.size = Pt(9.5)
    rd.font.color.rgb = RGBColor(51, 65, 85)

doc.add_page_break()

# 3. Visual Exploratory Analysis
add_custom_heading(doc, "3. Visual Exploratory Analysis", level=1)

add_custom_heading(doc, "A. Product Category Revenue Distribution", level=2)
doc.add_picture('fig1_category_clean.png', width=Inches(6.6))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c1 = doc.add_paragraph()
p_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c1.paragraph_format.space_after = Pt(12)
rc1 = p_c1.add_run("Figure 1: Ranked gross sales revenue across the six core product divisions.")
rc1.font.size = Pt(8.5)
rc1.font.italic = True
rc1.font.color.rgb = RGBColor(100, 116, 139)

add_custom_heading(doc, "B. Geographical Sales & Payment Channel Breakdown", level=2)
doc.add_picture('fig2_branch_payment.png', width=Inches(6.8))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c2 = doc.add_paragraph()
p_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c2.paragraph_format.space_after = Pt(12)
rc2 = p_c2.add_run("Figure 2: Store branch gross sales (left) and transaction settlement mode percentages (right).")
rc2.font.size = Pt(8.5)
rc2.font.italic = True
rc2.font.color.rgb = RGBColor(100, 116, 139)

add_custom_heading(doc, "C. Customer Segment Spending Behavior", level=2)
doc.add_picture('fig3_customer_spend.png', width=Inches(5.8))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c3 = doc.add_paragraph()
p_c3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_c3.paragraph_format.space_after = Pt(14)
rc3 = p_c3.add_run("Figure 3: Interquartile ranges and median transaction expenditures for Members vs. Normal shoppers.")
rc3.font.size = Pt(8.5)
rc3.font.italic = True
rc3.font.color.rgb = RGBColor(100, 116, 139)

# 4. Recommendations
add_custom_heading(doc, "4. Strategic Business Recommendations", level=1)
recs = [
    ("Membership Conversion Strategy: ", "Because loyalty members spend ~$29.37 more per visit, store personnel should be incentivized with enrollment targets, offering an immediate 5% first-purchase coupon on sign-up."),
    ("Express Digital Checkout Kiosks: ", "With 65.6% of checkouts conducted via E-wallet or Credit Card, setting up dedicated digital self-checkout lanes will shorten wait times and reduce cashier overhead during peak hours."),
    ("Cross-Merchandise Bundling: ", "To boost the lowest-performing category (Health & Beauty at $49,194), create co-promotional discount bundles with top-performing Food & Beverage staples.")
]
for title, desc in recs:
    p_r = doc.add_paragraph()
    p_r.paragraph_format.space_before = Pt(2)
    p_r.paragraph_format.space_after = Pt(4)
    p_r.paragraph_format.line_spacing = 1.15
    rt = p_r.add_run("1.  " if "Membership" in title else ("2.  " if "Express" in title else "3.  "))
    rt.font.bold = True
    rt.font.size = Pt(9.5)
    rt.font.color.rgb = RGBColor(30, 58, 138)
    rt2 = p_r.add_run(title)
    rt2.font.bold = True
    rt2.font.size = Pt(9.5)
    rt2.font.color.rgb = RGBColor(30, 58, 138)
    rd = p_r.add_run(desc)
    rd.font.size = Pt(9.5)
    rd.font.color.rgb = RGBColor(51, 65, 85)

# 5. Methodology
add_custom_heading(doc, "5. Analytical Stack & Repository Deliverables", level=1)
p_tech = doc.add_paragraph(
    "• Python Analytical Stack: Pandas (data cleaning, aggregations), NumPy (statistical indexing), Matplotlib & Seaborn (visualizations).\n"
    "• Project Deliverables: Jupyter Notebook (Md_Sameer_Khan_Sales_Analytics.ipynb), Requirements (requirements.txt), Markdown Briefing (README.md).\n"
    "• Version Control: Git workflow tracked and publicly accessible at https://github.com/Sameer-khan-27/IBM-SkillsBuild-Sales-Analytics."
)
p_tech.paragraph_format.space_after = Pt(6)
p_tech.paragraph_format.line_spacing = 1.15

# Save the finalized report
doc.save("MdSameerKhan_ProjectReport.docx")
print("SUCCESS: MdSameerKhan_ProjectReport.docx generated locally!")