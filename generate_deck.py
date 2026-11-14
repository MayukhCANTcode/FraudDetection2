import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path="AI_Fraud_Detection_Presentation.pptx"):
    prs = Presentation()
    # 16:9 Widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank slide

    # Minimalist Color Palette
    COLOR_BG = RGBColor(248, 250, 252)       # Slate 50 (#F8FAFC)
    COLOR_CARD = RGBColor(255, 255, 255)     # Crisp White (#FFFFFF)
    COLOR_CARD_BORDER = RGBColor(226, 232, 240) # Slate 200 (#E2E8F0)
    COLOR_PRIMARY = RGBColor(15, 23, 42)      # Deep Navy / Slate 900 (#0F172A)
    COLOR_MUTED = RGBColor(100, 116, 139)    # Muted Slate (#64748B)
    COLOR_ACCENT = RGBColor(37, 99, 235)     # Indigo/Blue 600 (#2563EB)
    COLOR_ACCENT_LIGHT = RGBColor(239, 246, 255) # Blue 50 (#EFF6FF)
    COLOR_CRIMSON = RGBColor(220, 38, 38)    # Red 600 (#DC2626)
    COLOR_GREEN = RGBColor(16, 185, 129)     # Green 600 (#10B981)
    COLOR_AMBER = RGBColor(217, 119, 6)      # Amber 600 (#D97706)

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.fill.background() # No line
        return bg

    def add_header(slide, category, title, slide_num):
        # Category Tag
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(10), Inches(0.35))
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = category.upper()
        p_tag.font.name = "Segoe UI"
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = COLOR_ACCENT

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(10.5), Inches(0.7))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.name = "Segoe UI"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_PRIMARY

        # Slide Number
        num_box = slide.shapes.add_textbox(Inches(11.8), Inches(0.6), Inches(0.8), Inches(0.4))
        tf_num = num_box.text_frame
        p_num = tf_num.paragraphs[0]
        p_num.text = f"{slide_num:02d} / 10"
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.name = "Segoe UI"
        p_num.font.size = Pt(11)
        p_num.font.color.rgb = COLOR_MUTED

    def add_card(slide, left, top, width, height, bg_color=COLOR_CARD, border_color=COLOR_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(1)
        else:
            card.line.fill.background()
        return card

    # ==========================================
    # SLIDE 1: TITLE & TEAM
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Decorative subtle accent line
    accent_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(0.9), Inches(0.08))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = COLOR_ACCENT
    accent_bar.line.fill.background()

    # Project Title & Subtitle box
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.5), Inches(2.2))
    tf1 = t_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "AI-Powered Financial Fraud Detection"
    p1.font.name = "Segoe UI"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_PRIMARY

    p2 = tf1.add_paragraph()
    p2.text = "Real-Time Machine Learning Pipeline, Risk Scoring & Decision Intelligence Dashboard"
    p2.font.name = "Segoe UI"
    p2.font.size = Pt(16)
    p2.font.color.rgb = COLOR_MUTED
    p2.space_before = Pt(8)

    # Team Members Card (Left)
    add_card(s1, Inches(0.8), Inches(4.0), Inches(5.6), Inches(2.7))
    team_box = s1.shapes.add_textbox(Inches(1.1), Inches(4.2), Inches(5.0), Inches(2.3))
    tf_team = team_box.text_frame
    tf_team.word_wrap = True
    p_th = tf_team.paragraphs[0]
    p_th.text = "PROJECT TEAM"
    p_th.font.name = "Segoe UI"
    p_th.font.size = Pt(11)
    p_th.font.bold = True
    p_th.font.color.rgb = COLOR_ACCENT

    members = [
        ("Member Name 1", "[Role / Specialization - e.g. Lead ML Engineer]"),
        ("Member Name 2", "[Role / Specialization - e.g. Data Scientist / EDA]"),
        ("Member Name 3", "[Role / Specialization - e.g. Full-Stack / UI Integration]"),
    ]
    for name, role in members:
        pn = tf_team.add_paragraph()
        pn.text = f"•  {name}  —  {role}"
        pn.font.name = "Segoe UI"
        pn.font.size = Pt(12)
        pn.font.color.rgb = COLOR_PRIMARY
        pn.space_before = Pt(8)

    # Project Metadata Card (Right)
    add_card(s1, Inches(6.8), Inches(4.0), Inches(5.7), Inches(2.7))
    meta_box = s1.shapes.add_textbox(Inches(7.1), Inches(4.2), Inches(5.1), Inches(2.3))
    tf_meta = meta_box.text_frame
    tf_meta.word_wrap = True
    p_mh = tf_meta.paragraphs[0]
    p_mh.text = "PROJECT SCOPE & TECH STACK"
    p_mh.font.name = "Segoe UI"
    p_mh.font.size = Pt(11)
    p_mh.font.bold = True
    p_mh.font.color.rgb = COLOR_ACCENT

    metas = [
        ("Core Stack", "Python • Scikit-Learn • Streamlit • SQLite • Plotly"),
        ("Dataset", "Credit Card Fraud Dataset (284,807 transactions)"),
        ("Objective", "High-recall automated fraud classification & risk triaging"),
        ("Date", "November 2026")
    ]
    for label, val in metas:
        pm = tf_meta.add_paragraph()
        pm.text = f"•  {label}: {val}"
        pm.font.name = "Segoe UI"
        pm.font.size = Pt(12)
        pm.font.color.rgb = COLOR_PRIMARY
        pm.space_before = Pt(6)

    # ==========================================
    # SLIDE 2: PROBLEM STATEMENT
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Executive Context", "The Growing Challenge of Digital Transaction Fraud", 2)

    # 3 Stat Cards Across Top
    card_w = Inches(3.64)
    c1 = add_card(s2, Inches(0.8), Inches(1.8), card_w, Inches(1.6))
    b1 = s2.shapes.add_textbox(Inches(1.0), Inches(1.95), card_w - Inches(0.4), Inches(1.3))
    tf = b1.text_frame
    p = tf.paragraphs[0]
    p.text = "$38+ Billion"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = COLOR_CRIMSON
    p = tf.add_paragraph()
    p.text = "Global annual financial losses due to credit card and transaction fraud."
    p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    c2 = add_card(s2, Inches(4.84), Inches(1.8), card_w, Inches(1.6))
    b2 = s2.shapes.add_textbox(Inches(5.04), Inches(1.95), card_w - Inches(0.4), Inches(1.3))
    tf = b2.text_frame
    p = tf.paragraphs[0]
    p.text = "0.172% Only"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = COLOR_AMBER
    p = tf.add_paragraph()
    p.text = "Extreme needle-in-a-haystack imbalance (only 492 fraud cases in 284,807 records)."
    p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    c3 = add_card(s2, Inches(8.88), Inches(1.8), card_w, Inches(1.6))
    b3 = s2.shapes.add_textbox(Inches(9.08), Inches(1.95), card_w - Inches(0.4), Inches(1.3))
    tf = b3.text_frame
    p = tf.paragraphs[0]
    p.text = "Sub-Second"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT
    p = tf.add_paragraph()
    p.text = "Need for instant latency so checkout conversion is never compromised."
    p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # Bottom 2 Detailed Problem Cards
    add_card(s2, Inches(0.8), Inches(3.7), Inches(5.7), Inches(3.2))
    pb_l = s2.shapes.add_textbox(Inches(1.1), Inches(3.9), Inches(5.1), Inches(2.8))
    tf_pl = pb_l.text_frame; tf_pl.word_wrap = True
    p = tf_pl.paragraphs[0]; p.text = "CURRENT SYSTEM SHORTCOMINGS"; p.font.bold = True; p.font.size = Pt(12); p.font.color.rgb = COLOR_PRIMARY
    bullets_l = [
        ("Legacy Rule Engines", "Static if-else threshold rules fail against evolving, sophisticated fraud vectors."),
        ("Manual Review Bottleneck", "High false alarm rates overwhelm risk review teams and delay order processing."),
        ("Customer Friction", "Blocking legitimate customer cards leads to abandoned checkouts and lost customer lifetime value.")
    ]
    for b_title, b_desc in bullets_l:
        p = tf_pl.add_paragraph()
        p.text = f"•  {b_title}: {b_desc}"
        p.font.size = Pt(11.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(8)

    add_card(s2, Inches(6.8), Inches(3.7), Inches(5.7), Inches(3.2))
    pb_r = s2.shapes.add_textbox(Inches(7.1), Inches(3.9), Inches(5.1), Inches(2.8))
    tf_pr = pb_r.text_frame; tf_pr.word_wrap = True
    p = tf_pr.paragraphs[0]; p.text = "TECHNICAL MACHINE LEARNING CHALLENGES"; p.font.bold = True; p.font.size = Pt(12); p.font.color.rgb = COLOR_PRIMARY
    bullets_r = [
        ("Extreme Class Imbalance", "Standard classifiers bias heavily toward majority class (predicting everything legitimate gives 99.8% dummy accuracy)."),
        ("Precision-Recall Tradeoff", "Catching more fraud (high recall) without inundating legitimate users (high precision)."),
        ("Operational Deployment", "Bridging the gap between a trained static model and a usable, auditable analyst dashboard.")
    ]
    for b_title, b_desc in bullets_r:
        p = tf_pr.add_paragraph()
        p.text = f"•  {b_title}: {b_desc}"
        p.font.size = Pt(11.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(8)

    # ==========================================
    # SLIDE 3: SOLUTION ARCHITECTURE
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "System Overview", "End-to-End Solution Architecture & Workflow", 3)

    pillars = [
        ("01. INGESTION & PREP", "Standardized Data Pipeline",
         ["Secure CSV Batch Upload", "Automated Null & Duplicate Checks", "StandardScaler Feature Scaling", "PCA Vector Alignments (V1-V28)"],
         COLOR_ACCENT),
        ("02. ML INFERENCE", "Ensemble Random Forest",
         ["Trained with SMOTE Oversampling", "Binary Label (Legitimate vs. Fraud)", "Continuous Fraud Probability (0-100%)", "Sub-50ms Inference Latency"],
         COLOR_PRIMARY),
        ("03. TRIAGE & UI", "Streamlit Intelligence Hub",
         ["Dynamic 3-Tier Risk Assignment", "Interactive Plotly Visualizations", "Instant CSV Results Download", "Analyst Deep-Dive Row Filters"],
         COLOR_AMBER),
        ("04. AUDIT & LOGGING", "Persistent SQLite Store",
         ["Real-Time Timestamp Recording", "Stored Historical Probabilities", "Tamper-Evident History Querying", "Model Governance & Compliance"],
         COLOR_GREEN),
    ]

    col_w = Inches(2.72)
    gap = Inches(0.3)
    for i, (tag, header, items, clr) in enumerate(pillars):
        c_left = Inches(0.8) + i * (col_w + gap)
        add_card(s3, c_left, Inches(1.8), col_w, Inches(5.1))

        # Colored header indicator
        h_bar = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left + Inches(0.2), Inches(2.0), col_w - Inches(0.4), Inches(0.35))
        h_bar.fill.solid(); h_bar.fill.fore_color.rgb = clr; h_bar.line.fill.background()
        tf = h_bar.text_frame
        p = tf.paragraphs[0]; p.text = tag; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.CENTER

        tb = s3.shapes.add_textbox(c_left + Inches(0.2), Inches(2.5), col_w - Inches(0.4), Inches(4.2))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = header; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY
        p.space_after = Pt(10)

        for it in items:
            p = tf.add_paragraph()
            p.text = f"• {it}"
            p.font.size = Pt(10.5); p.font.color.rgb = COLOR_MUTED
            p.space_before = Pt(6)

    # ==========================================
    # SLIDE 4: DATASET & EDA
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Data Exploration", "Dataset Profiling & Key Feature Insights", 4)

    # Top summary metrics
    stats = [
        ("Total Records", "284,807", "European cardholder transactions"),
        ("Legitimate Cases", "284,315 (99.83%)", "Dominate the sample distribution"),
        ("Fraudulent Cases", "492 (0.17%)", "Extreme minority target class"),
        ("Feature Dimensions", "30 Features", "Time, Amount + 28 PCA Components")
    ]
    w_box = Inches(2.72)
    for i, (lbl, val, desc) in enumerate(stats):
        bx_left = Inches(0.8) + i * (w_box + gap)
        add_card(s4, bx_left, Inches(1.8), w_box, Inches(1.5))
        t = s4.shapes.add_textbox(bx_left + Inches(0.2), Inches(1.9), w_box - Inches(0.4), Inches(1.2))
        tf = t.text_frame
        p = tf.paragraphs[0]; p.text = val; p.font.size = Pt(20); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY
        p = tf.add_paragraph(); p.text = lbl; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT; p.space_before = Pt(2)
        p = tf.add_paragraph(); p.text = desc; p.font.size = Pt(9.5); p.font.color.rgb = COLOR_MUTED

    # Bottom 2 Cards with EDA findings
    add_card(s4, Inches(0.8), Inches(3.6), Inches(5.7), Inches(3.3))
    tb_l = s4.shapes.add_textbox(Inches(1.1), Inches(3.8), Inches(5.1), Inches(2.9))
    tf = tb_l.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = "TRANSACTION AMOUNT & PATTERN DYNAMICS"; p.font.bold = True; p.font.size = Pt(12); p.font.color.rgb = COLOR_PRIMARY
    findings_l = [
        ("Mean Transaction Amount", "Overall average is $88.35; highest transaction in dataset reaches $25,691.16."),
        ("Fraud Amount Profiling", "Fraud transactions exhibit lower median values (~$9.25), often testing cards with rapid mini-charges to evade detection limits."),
        ("High-Dollar Outliers", "Occasional high-amount fraud spikes up to $2,125.81, requiring robust non-linear boundary separation."),
        ("Time Distribution", "Cyclical drops in legitimate activity during early morning hours; fraud rate percentage peaks during low-traffic windows.")
    ]
    for h, d in findings_l:
        p = tf.add_paragraph()
        p.text = f"•  {h}: {d}"
        p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(6)

    add_card(s4, Inches(6.8), Inches(3.6), Inches(5.7), Inches(3.3))
    tb_r = s4.shapes.add_textbox(Inches(7.1), Inches(3.8), Inches(5.1), Inches(2.9))
    tf = tb_r.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = "PREPROCESSING & FEATURE ENGINEERING"; p.font.bold = True; p.font.size = Pt(12); p.font.color.rgb = COLOR_PRIMARY
    findings_r = [
        ("PCA Privacy Preservation", "V1 through V28 are pre-transformed principal components, protecting confidential cardholder identities."),
        ("Robust Standardization", "Raw 'Time' and 'Amount' features exhibited massive variance; transformed via StandardScaler to zero mean and unit variance."),
        ("Zero Missingness", "Dataset has 0 null/NaN entries; duplicates removed to prevent over-optimistic evaluation."),
        ("Data Leakage Mitigation", "Scaling and oversampling fit strictly on training folds, never exposing hold-out test distributions.")
    ]
    for h, d in findings_r:
        p = tf.add_paragraph()
        p.text = f"•  {h}: {d}"
        p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(6)

    # ==========================================
    # SLIDE 5: METHODOLOGY & MODELING
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Methodology", "Handling Class Imbalance & Model Selection", 5)

    steps = [
        ("1. Data Partitioning", "Stratified 80/20 train-test split ensuring identical fraud ratios (0.17%) across both splits to avoid sampling distortion."),
        ("2. SMOTE Oversampling", "Synthetic Minority Oversampling applied strictly to training data, generating realistic synthetic fraud vectors rather than naive duplicates."),
        ("3. Model Benchmarking", "Compared 4 diverse algorithm families: Linear (Logistic Regression), Tree-based (Decision Tree), and Ensembles (Random Forest, XGBoost)."),
        ("4. Objective Metric", "Optimized specifically for F1-Score and ROC-AUC rather than misleading Raw Accuracy (which is trivially 99.8%).")
    ]
    for i, (title, desc) in enumerate(steps):
        top_pos = Inches(1.8) + i * Inches(1.28)
        add_card(s5, Inches(0.8), top_pos, Inches(11.7), Inches(1.1))
        # Icon / Number pill
        pill = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), top_pos + Inches(0.25), Inches(0.6), Inches(0.6))
        pill.fill.solid(); pill.fill.fore_color.rgb = COLOR_ACCENT_LIGHT; pill.line.fill.background()
        tf_pill = pill.text_frame
        p = tf_pill.paragraphs[0]; p.text = str(i+1); p.font.size = Pt(14); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT; p.alignment = PP_ALIGN.CENTER

        tb = s5.shapes.add_textbox(Inches(1.9), top_pos + Inches(0.15), Inches(10.3), Inches(0.8))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY
        p = tf.add_paragraph(); p.text = desc; p.font.size = Pt(11); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

    # ==========================================
    # SLIDE 6: MODEL COMPARISON & SCORES (TABLE)
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Evaluation & Benchmark", "Model Performance & Quantitative Scores", 6)

    # Top highlight card for selected model
    add_card(s6, Inches(0.8), Inches(1.8), Inches(11.7), Inches(1.1), bg_color=COLOR_ACCENT_LIGHT, border_color=COLOR_ACCENT)
    tb_hl = s6.shapes.add_textbox(Inches(1.1), Inches(1.95), Inches(11.1), Inches(0.8))
    tf = tb_hl.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = "SELECTED PRODUCTION MODEL: RANDOM FOREST CLASSIFIER"; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT
    p = tf.add_paragraph()
    p.text = "Achieved 99.92% Accuracy, 73.53% Precision, 78.95% Recall, 76.14% F1-Score, and 97.93% ROC-AUC (Cross-Validation Mean F1: 0.8335)."
    p.font.size = Pt(11); p.font.color.rgb = COLOR_PRIMARY; p.space_before = Pt(2)

    # Table of Results
    rows = 5; cols = 6
    tbl_shape = s6.shapes.add_table(rows, cols, Inches(0.8), Inches(3.2), Inches(11.7), Inches(3.2))
    tbl = tbl_shape.table
    tbl.columns[0].width = Inches(3.2)
    tbl.columns[1].width = Inches(1.7)
    tbl.columns[2].width = Inches(1.7)
    tbl.columns[3].width = Inches(1.7)
    tbl.columns[4].width = Inches(1.7)
    tbl.columns[5].width = Inches(1.7)

    headers = ["Model Architecture", "Accuracy", "Precision", "Recall", "F1 Score", "ROC-AUC"]
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = COLOR_PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(11); p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT

    table_data = [
        ("Logistic Regression (Baseline)", "97.37%", "5.30%", "87.37%", "10.00%", "96.26%"),
        ("Decision Tree Classifier", "99.75%", "35.67%", "64.21%", "45.86%", "82.00%"),
        ("Random Forest (Selected)", "99.92%", "73.53%", "78.95%", "76.14%", "97.93%"),
        ("XGBoost Classifier", "99.92%", "74.26%", "78.95%", "76.53%", "97.02%")
    ]

    for i, row in enumerate(table_data):
        is_winner = (i == 2) # Random forest
        for j, val in enumerate(row):
            cell = tbl.cell(i+1, j)
            cell.text = val
            cell.fill.solid()
            if is_winner:
                cell.fill.fore_color.rgb = RGBColor(238, 242, 255) # Light Indigo highlight
            else:
                cell.fill.fore_color.rgb = COLOR_CARD if (i % 2 == 0) else RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.bold = is_winner
            p.font.color.rgb = COLOR_ACCENT if is_winner else COLOR_PRIMARY
            p.alignment = PP_ALIGN.CENTER if j > 0 else PP_ALIGN.LEFT

    # Footnote
    fn_box = s6.shapes.add_textbox(Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.4))
    tf = fn_box.text_frame
    p = tf.paragraphs[0]
    p.text = "*Note: While Logistic Regression yields high recall (87%), its 5.3% precision produces 18 false alerts for every real fraud. Random Forest cuts false alarms by >14x."
    p.font.size = Pt(9.5); p.font.italic = True; p.font.color.rgb = COLOR_MUTED

    # ==========================================
    # SLIDE 7: VISUAL RESULTS & GRAPHS
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Analytics & Visualizations", "Dashboard Graphical Visualizations", 7)

    cards_vis = [
        ("Fraud Class Distribution", "Pie & Donut Chart",
         "Visualizes the severe class disproportion (99.83% vs 0.17%) in uploaded batch files, providing immediate visual awareness of fraud prevalence.",
         "Interactive Plotly Hole=0.45 Donut"),
        ("Prediction Breakdown", "Comparative Bar Chart",
         "Instant volume tallies of Legitimate vs. Flagged Fraud transactions from the current inference run, updating dynamically upon file upload.",
         "Real-time reactive bar rendering"),
        ("Amount Histogram & Boxplot", "Outlier & Spread Analysis",
         "Overlaid logarithmic distribution displaying transaction sizes across classes. Highlights critical clusters where high-value anomalies occur.",
         "50-bin histogram & quartile boxplots"),
        ("Top Fraudulent Transactions", "Ranked Severity Bar Chart",
         "Isolates and ranks the top 10 highest-value fraudulent attempts by dollar amount, enabling fraud teams to stop massive capital drains first.",
         "Ranked top-10 ledger visualization")
    ]

    for i, (head, sub, desc, tech) in enumerate(cards_vis):
        col = i % 2
        row = i // 2
        l = Inches(0.8) + col * (Inches(5.7) + Inches(0.3))
        t = Inches(1.8) + row * (Inches(2.4) + Inches(0.3))

        add_card(s7, l, t, Inches(5.7), Inches(2.4))
        tb = s7.shapes.add_textbox(l + Inches(0.25), t + Inches(0.2), Inches(5.2), Inches(2.0))
        tf = tb.text_frame; tf.word_wrap = True

        p = tf.paragraphs[0]; p.text = head; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY
        p = tf.add_paragraph(); p.text = sub; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT; p.space_before = Pt(2)
        p = tf.add_paragraph(); p.text = desc; p.font.size = Pt(10); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(6)

        p = tf.add_paragraph(); p.text = f"Implementation: {tech}"; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY; p.space_before = Pt(6)

    # ==========================================
    # SLIDE 8: REAL-TIME DASHBOARD & RISK TRIAGING
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Operational Dashboard", "Tiered Risk Scoring & Interactive UI", 8)

    # 3 Risk Tiers
    tiers = [
        ("HIGH RISK", "Probability >= 80%", COLOR_CRIMSON,
         ["Immediate transaction block / freeze", "Direct routing to Level-2 Fraud Investigator", "Automated SMS/Email notification to cardholder", "Full event logging in security audit table"]),
        ("MEDIUM RISK", "Probability 40% - 79%", COLOR_AMBER,
         ["Automated Step-Up Authentication (OTP/2FA)", "Biometric verification trigger on mobile app", "Temporary hold pending user confirmation", "Prevents customer cart abandonment"]),
        ("LOW RISK", "Probability < 40%", COLOR_GREEN,
         ["Frictionless, instant sub-50ms clearance", "Standard clearing house settlement", "Zero user disruption at checkout", "99%+ of all regular daily transactions"])
    ]

    tier_w = Inches(3.64)
    for i, (title, crit, clr, acts) in enumerate(tiers):
        t_left = Inches(0.8) + i * (tier_w + Inches(0.39))
        add_card(s8, t_left, Inches(1.8), tier_w, Inches(3.4))

        # Top tag
        bar = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, t_left, Inches(1.8), tier_w, Inches(0.45))
        bar.fill.solid(); bar.fill.fore_color.rgb = clr; bar.line.fill.background()
        tf = bar.text_frame
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = RGBColor(255, 255, 255); p.alignment = PP_ALIGN.CENTER

        tb = s8.shapes.add_textbox(t_left + Inches(0.2), Inches(2.35), tier_w - Inches(0.4), Inches(2.7))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = crit; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = clr; p.alignment = PP_ALIGN.CENTER
        p.space_after = Pt(10)

        for a in acts:
            p = tf.add_paragraph()
            p.text = f"• {a}"
            p.font.size = Pt(10.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(4)

    # Bottom Dashboard Features Card
    add_card(s8, Inches(0.8), Inches(5.4), Inches(11.7), Inches(1.5))
    tb_df = s8.shapes.add_textbox(Inches(1.1), Inches(5.55), Inches(11.1), Inches(1.2))
    tf = tb_df.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = "STREAMLIT DASHBOARD CAPABILITIES"; p.font.bold = True; p.font.size = Pt(11); p.font.color.rgb = COLOR_ACCENT
    p = tf.add_paragraph()
    p.text = "• Batch CSV Upload: Ingest thousands of transactions with a single drag-and-drop.\n• Live Score Generation: Computes probability and risk assignment for every row in memory.\n• One-Click CSV Export: Download enriched datasets containing raw features + predictions + risk tiers.\n• SQLite Auditing: Automatically writes top transactions to fraud.db for compliance tracing."
    p.font.size = Pt(10.5); p.font.color.rgb = COLOR_PRIMARY; p.space_before = Pt(4)

    # ==========================================
    # SLIDE 9: HOW OUR SOLUTION HELPED (IMPACT)
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Business Value", "Tangible Impact & How Our Solution Helped", 9)

    impacts = [
        ("14x Reduction in False Positives",
         "Baseline models flagged ~18 legitimate users for every 1 fraud detected. Our Random Forest tuned pipeline elevated Precision from 5.3% to 73.53%, eliminating massive customer frustration."),
        ("Near-Instantaneous Latency",
         "The optimized inference pipeline executes feature standardization and tree ensemble voting in <50ms, allowing seamless embedding into live payment gateways."),
        ("Elimination of Manual Bottlenecks",
         "Rather than manually inspecting thousands of records, fraud teams focus strictly on the <0.5% identified High-Risk queue, increasing investigative throughput by over 80%."),
        ("Actionable 3-Tier Risk Model",
         "Instead of binary accept/reject, our tiered model enables Step-Up Authentication (OTP/2FA) for borderline cases, salvaging authentic transactions that would have been rejected."),
        ("Zero-Configuration Accessibility",
         "Non-technical fraud analysts and risk officers can easily upload transaction batches through an intuitive UI without writing SQL queries or running terminal scripts.")
    ]

    for i, (title, desc) in enumerate(impacts):
        top_pos = Inches(1.8) + i * Inches(1.0)
        add_card(s9, Inches(0.8), top_pos, Inches(11.7), Inches(0.88))

        # Checkmark / Bullet indicator
        chk = s9.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), top_pos + Inches(0.22), Inches(0.45), Inches(0.45))
        chk.fill.solid(); chk.fill.fore_color.rgb = RGBColor(220, 252, 231) # Green 100
        chk.line.fill.background()
        tf_c = chk.text_frame
        p = tf_c.paragraphs[0]; p.text = "✓"; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_GREEN; p.alignment = PP_ALIGN.CENTER

        tb = s9.shapes.add_textbox(Inches(1.75), top_pos + Inches(0.12), Inches(10.5), Inches(0.68))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(12); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY
        p = tf.add_paragraph(); p.text = desc; p.font.size = Pt(10.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(2)

    # ==========================================
    # SLIDE 10: ROADMAP & CONCLUSION
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Summary & Future Work", "Strategic Roadmap & Concluding Summary", 10)

    # 3 Strategic Pillars
    future = [
        ("Explainable AI (XAI)", "Integrate SHAP (Shapley Additive exPlanations) to provide risk officers with local feature attribution explaining exactly why a transaction was flagged."),
        ("Kafka Streaming Engine", "Transition from batch CSV uploads to real-time streaming ingestion via Apache Kafka / Redis pub-sub for continuous millisecond monitoring."),
        ("Cloud Microservices", "Containerize via Docker & Kubernetes on AWS/GCP, scaling FastAPI prediction endpoints independently under high-traffic sales events.")
    ]

    w_fut = Inches(3.64)
    for i, (head, body) in enumerate(future):
        f_left = Inches(0.8) + i * (w_fut + Inches(0.39))
        add_card(s10, f_left, Inches(1.8), w_fut, Inches(2.5))
        tb = s10.shapes.add_textbox(f_left + Inches(0.25), Inches(2.0), w_fut - Inches(0.5), Inches(2.1))
        tf = tb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.text = f"PHASE {i+1}"; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = COLOR_ACCENT
        p = tf.add_paragraph(); p.text = head; p.font.size = Pt(13); p.font.bold = True; p.font.color.rgb = COLOR_PRIMARY; p.space_before = Pt(4)
        p = tf.add_paragraph(); p.text = body; p.font.size = Pt(10.5); p.font.color.rgb = COLOR_MUTED; p.space_before = Pt(6)

    # Concluding Takeaway Card
    add_card(s10, Inches(0.8), Inches(4.6), Inches(11.7), Inches(2.3), bg_color=COLOR_PRIMARY, border_color=None)
    tb_c = s10.shapes.add_textbox(Inches(1.2), Inches(4.8), Inches(10.9), Inches(1.9))
    tf = tb_c.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = "CONCLUSION & WRAP-UP"; p.font.size = Pt(11); p.font.bold = True; p.font.color.rgb = RGBColor(147, 197, 253)
    p = tf.add_paragraph()
    p.text = "This project delivers a complete, production-ready fraud intelligence solution. By coupling advanced class-imbalance mitigation (SMOTE) with a robust Random Forest classifier (97.9% ROC-AUC), it bridges the gap between machine learning research and real-world operations—safeguarding capital while ensuring legitimate customers experience zero friction."
    p.font.size = Pt(11.5); p.font.color.rgb = RGBColor(241, 245, 249); p.space_before = Pt(6)

    p_qa = tf.add_paragraph()
    p_qa.text = "Thank You! Questions & Discussion are Welcome."
    p_qa.font.size = Pt(12); p_qa.font.bold = True; p_qa.font.color.rgb = RGBColor(255, 255, 255); p_qa.space_before = Pt(10)

    # Save
    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    create_deck()
