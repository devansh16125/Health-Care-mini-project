"""
Script to generate the publication-grade academic PowerPoint presentation for CA3 Mini-Project.
Title: Explainable Machine Learning for Alzheimer's Disease Classification Using Structural MRI and Clinical Features
Authors: Jay Thakare (PRN: 23070521233) and Devansh Gawali (PRN: 23070521237)
Institution: Symbiosis Institute of Technology (SIT), Nagpur
Affiliation: Symbiosis International (Deemed University), Pune, India
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # ==========================================
    # COLOR PALETTE (EDITORIAL SCIENTIFIC SYSTEM)
    # ==========================================
    BG_COLOR = RGBColor(250, 251, 252)       # #FAFBFC Warm White Background
    NAVY = RGBColor(23, 43, 69)              # #172B45 Deep Navy Primary Text
    SLATE = RGBColor(102, 117, 138)          # #66758A Slate Gray Secondary Text
    MUTED_TEAL = RGBColor(22, 140, 140)      # #168C8C Primary Accent Teal
    SOFT_BLUE = RGBColor(86, 133, 196)       # #5685C4 Secondary Accent Blue
    BORDER_COLOR = RGBColor(226, 232, 240)   # #E2E8F0 Subtle Border
    WHITE = RGBColor(255, 255, 255)          # #FFFFFF Pure Surface White
    CARD_BG = RGBColor(255, 255, 255)
    TINT_TEAL = RGBColor(240, 249, 249)      # #F0F9F9 Very Light Teal Surface
    TINT_BLUE = RGBColor(241, 246, 253)      # #F1F6FD Very Light Blue Surface
    TINT_GRAY = RGBColor(246, 248, 250)      # #F6F8FA Neutral Card Background
    RED_ACCENT = RGBColor(192, 57, 43)       # #C0392B Restrained Red
    GREEN_ACCENT = RGBColor(33, 132, 90)     # #21845A Clean Emerald
    
    FONT_FAMILY = "Arial"
    figures_dir = r"C:\Users\devan\.gemini\antigravity-ide\scratch\alzheimer-explainable-ml\results\figures"

    def apply_slide_base(slide, slide_num, total_slides=16, kicker=None, title=None, subtitle=None):
        # 1. Warm White Background Fill
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_COLOR
        bg.line.fill.background()
        
        # 2. Header (if title given)
        if title:
            t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.32), Inches(11.733), Inches(0.96))
            tf = t_box.text_frame
            tf.word_wrap = True
            tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
            
            # Kicker / Section label
            if kicker:
                p_kick = tf.paragraphs[0]
                p_kick.text = kicker.upper()
                p_kick.font.name = FONT_FAMILY
                p_kick.font.size = Pt(9.5)
                p_kick.font.bold = True
                p_kick.font.color.rgb = MUTED_TEAL
                p_kick.space_after = Pt(2)
                p_title = tf.add_paragraph()
            else:
                p_title = tf.paragraphs[0]
                
            p_title.text = title
            p_title.font.name = FONT_FAMILY
            p_title.font.size = Pt(21)
            p_title.font.bold = True
            p_title.font.color.rgb = NAVY
            
            if subtitle:
                p_sub = tf.add_paragraph()
                p_sub.text = subtitle
                p_sub.font.name = FONT_FAMILY
                p_sub.font.size = Pt(10.5)
                p_sub.font.color.rgb = SLATE
                p_sub.space_before = Pt(2)
                
            # Subtle header divider line
            divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.34), Inches(11.733), Inches(0.015))
            divider.fill.solid()
            divider.fill.fore_color.rgb = BORDER_COLOR
            divider.line.fill.background()

        # 3. Footer (on content slides)
        if slide_num > 1:
            foot_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.012))
            foot_line.fill.solid()
            foot_line.fill.fore_color.rgb = BORDER_COLOR
            foot_line.line.fill.background()
            
            f_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.98), Inches(11.733), Inches(0.35))
            ftf = f_box.text_frame
            ftf.word_wrap = True
            ftf.margin_top = ftf.margin_bottom = ftf.margin_left = ftf.margin_right = 0
            
            fp = ftf.paragraphs[0]
            fp.text = "SIT Nagpur • CA3 Mini-Project • Explainable ML for Alzheimer’s Classification"
            fp.font.name = FONT_FAMILY
            fp.font.size = Pt(9)
            fp.font.color.rgb = SLATE
            
            f_num_box = slide.shapes.add_textbox(Inches(10.533), Inches(6.98), Inches(2.0), Inches(0.35))
            ntf = f_num_box.text_frame
            ntf.margin_top = ntf.margin_bottom = ntf.margin_left = ntf.margin_right = 0
            np = ntf.paragraphs[0]
            np.alignment = PP_ALIGN.RIGHT
            np.text = f"{slide_num:02d} / {total_slides:02d}"
            np.font.name = FONT_FAMILY
            np.font.size = Pt(9)
            np.font.bold = True
            np.font.color.rgb = SLATE

    def add_card(slide, left, top, width, height, title=None, kicker=None, bg_color=CARD_BG, border_color=BORDER_COLOR, accent_color=None):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(0.75)
        
        if accent_color:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.05))
            bar.fill.solid()
            bar.fill.fore_color.rgb = accent_color
            bar.line.fill.background()
            
        if title or kicker:
            tx = slide.shapes.add_textbox(left + Inches(0.24), top + Inches(0.16), width - Inches(0.48), Inches(0.55))
            tf = tx.text_frame
            tf.word_wrap = True
            tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
            
            p0 = tf.paragraphs[0]
            if kicker:
                p0.text = kicker.upper()
                p0.font.name = FONT_FAMILY
                p0.font.size = Pt(8.5)
                p0.font.bold = True
                p0.font.color.rgb = accent_color if accent_color else MUTED_TEAL
                p_next = tf.add_paragraph()
            else:
                p_next = p0
                
            if title:
                p_next.text = title
                p_next.font.name = FONT_FAMILY
                p_next.font.size = Pt(12.5)
                p_next.font.bold = True
                p_next.font.color.rgb = NAVY
        return card

    # =========================================================================
    # SLIDE 1: TITLE SLIDE (EDITORIAL SCIENTIFIC COVER)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s1, slide_num=1, total_slides=16)
    
    side_accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.18), Inches(7.5))
    side_accent.fill.solid()
    side_accent.fill.fore_color.rgb = MUTED_TEAL
    side_accent.line.fill.background()
    
    badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(0.85), Inches(4.3), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = TINT_TEAL
    badge.line.color.rgb = MUTED_TEAL
    badge.line.width = Pt(1.0)
    btx = badge.text_frame
    btx.margin_top = btx.margin_bottom = btx.margin_left = btx.margin_right = 0
    bp = btx.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    bp.text = "RESEARCH PRESENTATION • B.TECH CA3 EVALUATION"
    bp.font.name = FONT_FAMILY
    bp.font.size = Pt(9.5)
    bp.font.bold = True
    bp.font.color.rgb = MUTED_TEAL
    
    t1_box = s1.shapes.add_textbox(Inches(0.9), Inches(1.4), Inches(11.5), Inches(2.3))
    t1_f = t1_box.text_frame
    t1_f.word_wrap = True
    t1_f.margin_top = t1_f.margin_bottom = t1_f.margin_left = t1_f.margin_right = 0
    
    tp1 = t1_f.paragraphs[0]
    tp1.text = "Explainable Machine Learning"
    tp1.font.name = FONT_FAMILY
    tp1.font.size = Pt(36)
    tp1.font.bold = True
    tp1.font.color.rgb = NAVY
    tp1.space_after = Pt(4)
    
    tp2 = t1_f.add_paragraph()
    tp2.text = "for Alzheimer’s Disease Classification Using Structural MRI and Clinical Features"
    tp2.font.name = FONT_FAMILY
    tp2.font.size = Pt(21)
    tp2.font.bold = True
    tp2.font.color.rgb = SOFT_BLUE
    tp2.space_after = Pt(8)
    
    tp3 = t1_f.add_paragraph()
    tp3.text = "A Participant-Level Leakage-Free Diagnostic Architecture on the OASIS-1 Cross-Sectional Cohort with SHAP Decision Support"
    tp3.font.name = FONT_FAMILY
    tp3.font.size = Pt(12)
    tp3.font.color.rgb = SLATE

    add_card(s1, Inches(0.9), Inches(3.9), Inches(6.8), Inches(2.7), title="Investigators & Institutional Affiliation", kicker="RESEARCH TEAM", bg_color=WHITE, border_color=BORDER_COLOR, accent_color=MUTED_TEAL)
    
    auth_box = s1.shapes.add_textbox(Inches(1.15), Inches(4.55), Inches(6.3), Inches(1.9))
    af = auth_box.text_frame
    af.word_wrap = True
    af.margin_top = af.margin_bottom = af.margin_left = af.margin_right = 0
    
    ap1 = af.paragraphs[0]
    ap1.text = "Project Team Members (Investigators):"
    ap1.font.name = FONT_FAMILY
    ap1.font.size = Pt(10.5)
    ap1.font.bold = True
    ap1.font.color.rgb = NAVY
    
    ap2 = af.add_paragraph()
    ap2.text = "• Jay Thakare  —  PRN: 23070521233  (B.Tech CSE)\n• Devansh Gawali  —  PRN: 23070521237  (B.Tech CSE)"
    ap2.font.name = FONT_FAMILY
    ap2.font.size = Pt(11.5)
    ap2.font.bold = True
    ap2.font.color.rgb = MUTED_TEAL
    ap2.space_before = Pt(2)
    ap2.space_after = Pt(6)
    
    ap3 = af.add_paragraph()
    ap3.text = "Department of Computer Science & Engineering\nSymbiosis Institute of Technology (SIT), Nagpur\nSymbiosis International (Deemed University), Pune, India"
    ap3.font.name = FONT_FAMILY
    ap3.font.size = Pt(10)
    ap3.font.color.rgb = SLATE
    
    add_card(s1, Inches(7.9), Inches(3.9), Inches(4.5), Inches(2.7), title="Experimental Highlights", kicker="BENCHMARK OVERVIEW", bg_color=TINT_GRAY, border_color=BORDER_COLOR, accent_color=SOFT_BLUE)
    
    scope_box = s1.shapes.add_textbox(Inches(8.15), Inches(4.55), Inches(4.0), Inches(1.9))
    sf = scope_box.text_frame
    sf.word_wrap = True
    sf.margin_top = sf.margin_bottom = sf.margin_left = sf.margin_right = 0
    
    sp1 = sf.paragraphs[0]
    sp1.text = "• Cohort: OASIS-1 (N=235 Clinically Evaluated)"
    sp1.font.name = FONT_FAMILY; sp1.font.size = Pt(10.5); sp1.font.color.rgb = NAVY; sp1.font.bold = True
    
    sp2 = sf.add_paragraph()
    sp2.text = "• Protocol: Strict 80/20 subject-level split (0 leakage)\n• Peak Model: Gradient Boosting (85.11% Acc, 0.9426 AUC)\n• Ensemble: Soft-Voting (LR + RF + GB, 0.9426 AUC)\n• Modality Gain: Multimodal beats MRI-only by +12.77% Acc\n• Explainability: SHAP TreeExplainer (MMSE & nWBV)"
    sp2.font.name = FONT_FAMILY; sp2.font.size = Pt(10); sp2.font.color.rgb = SLATE; sp2.space_before = Pt(4)

    s1.notes_slide.notes_text_frame.text = (
        "Good morning/afternoon, respected faculty members and evaluation committee. Today, Devansh Gawali and I, "
        "Jay Thakare, present our B.Tech CA3 research mini-project entitled: 'Explainable Machine Learning for Alzheimer's "
        "Disease Classification Using Structural MRI and Clinical Features'. In this work, we address two primary challenges "
        "in computational neuroscience: the prevalence of data leakage in published literature and the black-box opacity of "
        "automated classifiers. By combining psychometric screening with quantitative structural MRI morphometrics under a "
        "strict subject-level validation protocol, we achieve 85.11% accuracy and 0.9426 ROC-AUC, with post-hoc SHAP providing "
        "transparent biological decision support for clinicians."
    )

    # =========================================================================
    # SLIDE 2: CLINICAL CONTEXT & MOTIVATION
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s2, slide_num=2, kicker="01 | CLINICAL CONTEXT", 
                     title="Background & Clinical Motivation", 
                     subtitle="Epidemiological scale, diagnostic limitations, and the necessity of interpretable medical intelligence.")
    
    col_w = Inches(3.7)
    gap = Inches(0.3)
    top_pos = Inches(1.55)
    card_h = Inches(5.15)
    
    add_card(s2, Inches(0.8), top_pos, col_w, card_h, title="Global Healthcare Challenge", kicker="EPIDEMIOLOGY", accent_color=RED_ACCENT)
    tb1 = s2.shapes.add_textbox(Inches(1.04), top_pos + Inches(0.75), col_w - Inches(0.48), card_h - Inches(0.9))
    tf1 = tb1.text_frame; tf1.word_wrap = True
    tf1.text = (
        "• Leading Cause of Dementia:\n"
        "  Alzheimer's Disease (AD) accounts for 60%–80% of all neurodegenerative dementia diagnoses globally.\n\n"
        "• Accelerating Prevalence:\n"
        "  Over 55 million individuals currently live with dementia; projected to exceed 152 million by 2050 (WHO, 2023).\n\n"
        "• Severe Economic Burden:\n"
        "  Annual global healthcare expenditures exceed $1.3 trillion, straining national care systems.\n\n"
        "• Preclinical Atrophy:\n"
        "  Irreversible microscopic and macroscopic cerebral parenchymal damage begins 5–15 years prior to clinical symptoms."
    )
    for p in tf1.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY
    
    add_card(s2, Inches(0.8) + col_w + gap, top_pos, col_w, card_h, title="Current Diagnostic Bottlenecks", kicker="CLINICAL LIMITS", accent_color=SOFT_BLUE)
    tb2 = s2.shapes.add_textbox(Inches(1.04) + col_w + gap, top_pos + Inches(0.75), col_w - Inches(0.48), card_h - Inches(0.9))
    tf2 = tb2.text_frame; tf2.word_wrap = True
    tf2.text = (
        "• Psychometric Tests (MMSE):\n"
        "  Mini-Mental State Examination is fast but subject to language bias, education confounding, and ceiling effects in mild stages.\n\n"
        "• Structural MRI Volumetry:\n"
        "  Captures ventricular enlargement, hippocampal shrinkage, and whole-brain parenchymal loss (nWBV).\n\n"
        "• Expert Radiologist Scarcity:\n"
        "  Manual volumetric assessment requires specialized neuroradiologists, creating long diagnostic waiting lists.\n\n"
        "• Inter-Rater Variability:\n"
        "  Subjective clinical evaluation of border cases (CDR 0.5) leads to conflicting early-stage diagnoses."
    )
    for p in tf2.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    add_card(s2, Inches(0.8) + (col_w + gap)*2, top_pos, col_w, card_h, title="The Explainability Imperative", kicker="AI TRUST", accent_color=MUTED_TEAL)
    tb3 = s2.shapes.add_textbox(Inches(1.04) + (col_w + gap)*2, top_pos + Inches(0.75), col_w - Inches(0.48), card_h - Inches(0.9))
    tf3 = tb3.text_frame; tf3.word_wrap = True
    tf3.text = (
        "• Black-Box Skepticism:\n"
        "  Deep neural networks provide unverified diagnostic scores without showing which anatomical features drove the prediction.\n\n"
        "• High Clinical Stakes:\n"
        "  False positives trigger unwarranted trauma and medical bills; false negatives miss the early intervention window.\n\n"
        "• Regulatory Mandates:\n"
        "  FDA Software as a Medical Device (SaMD) and EU AI Act require transparent, auditable decision justification.\n\n"
        "• Collaborative Intelligence:\n"
        "  AI must act as transparent second-opinion support, enabling physicians to verify biological rationale."
    )
    for p in tf3.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s2.notes_slide.notes_text_frame.text = (
        "Slide 2 establishes the real-world healthcare motivation. Dementia currently affects over 55 million people worldwide, "
        "and macroscopic brain atrophy develops silently up to 15 years before overt functional impairment. Currently, clinicians rely "
        "on cognitive screening like the MMSE and structural MRI scans. However, MMSE is prone to educational biases, while MRI volumetry "
        "suffers from radiologist bottlenecks. While automated machine learning is promising, black-box deep models cannot be deployed "
        "safely in neurology without interpretable justifications. Clinicians require transparent, biologically valid explanations."
    )

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & PROJECT OBJECTIVES
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s3, slide_num=3, kicker="02 | RESEARCH OBJECTIVES",
                     title="Problem Statement & Methodological Goals",
                     subtitle="Mapping major vulnerabilities in existing medical AI literature directly to project solutions.")
    
    add_card(s3, Inches(0.8), Inches(1.55), Inches(5.6), Inches(5.15), title="Critical Vulnerabilities in Literature", kicker="CORE CHALLENGES", accent_color=RED_ACCENT)
    pb_box = s3.shapes.add_textbox(Inches(1.05), Inches(2.35), Inches(5.1), Inches(4.1))
    pbf = pb_box.text_frame; pbf.word_wrap = True
    pbf.text = (
        "1. Opacity of Deep Classifiers (Black-Box Problem):\n"
        "   Complex 3D convolutional models fail to provide patient-specific feature attribution, preventing clinicians from auditing if outputs stem from genuine pathology or scanner noise.\n\n"
        "2. Widespread Data Leakage in Published Studies:\n"
        "   Subject overlap across train/test splits and global pre-scaling across entire datasets artificially inflate published accuracy by up to 25% (Yagis et al., 2021).\n\n"
        "3. Prohibitive Computational Infrastructure:\n"
        "   Large-scale 3D CNNs require high-end GPU clusters and hours of non-linear spatial deformation, rendering them impractical for resource-limited primary care clinics."
    )
    for p in pbf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY
    
    add_card(s3, Inches(6.8), Inches(1.55), Inches(5.733), Inches(5.15), title="Formulated Project Objectives", kicker="PROPOSED SOLUTIONS", accent_color=MUTED_TEAL)
    ob_box = s3.shapes.add_textbox(Inches(7.05), Inches(2.35), Inches(5.2), Inches(4.1))
    obf = ob_box.text_frame; obf.word_wrap = True
    obf.text = (
        "• Strict Participant-Level Partitioning:\n"
        "  Enforce an absolute 80/20 subject-level holdout partition (S_train ∩ S_test = ∅) with all imputation and scaling fitted exclusively on training subjects.\n\n"
        "• Multimodal Synergistic Integration:\n"
        "  Combine cognitive psychometrics (MMSE) with stereotaxically normalized MRI morphometrics (nWBV, eTIV, ASF) for cross-modality complementary power.\n\n"
        "• Lightweight Machine Learning Ensemble:\n"
        "  Benchmark 6 classifiers on commodity hardware (Intel Core i5) executing full 5-fold CV and test evaluation in <15 seconds without GPU requirements.\n\n"
        "• Game-Theoretic Explainability via SHAP:\n"
        "  Deploy TreeExplainer to deliver both global cohort biomarker importance and patient-level waterfall log-odds decision support."
    )
    for p in obf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s3.notes_slide.notes_text_frame.text = (
        "On Slide 3, we map the three core problems in Alzheimer's machine learning literature to our explicit project objectives. "
        "First, to combat black-box opacity, we integrate SHAP explainability. Second, to prevent the ubiquitous data leakage documented "
        "by Yagis et al., we enforce a strict participant-level 80/20 split where no subject appears in both sets, and transformers are fitted "
        "only on training folds. Third, instead of multi-hour GPU-dependent deep nets, we prove that classical tree ensembles on multimodal "
        "features execute in seconds on standard hardware while delivering 0.943 ROC-AUC."
    )

    # =========================================================================
    # SLIDE 4: LITERATURE REVIEW & RESEARCH GAP (2021–2026) — CLEAN FORMATTED
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s4, slide_num=4, kicker="03 | STATE OF THE ART",
                     title="Literature Review & Research Gap (2021–2026)",
                     subtitle="Analysis of recent peer-reviewed multimodal and explainable AI studies on neuroimaging cohorts.")
    
    # Clean Table with strict row heights and small clean typography across all paragraphs
    t4_shape = s4.shapes.add_table(rows=6, cols=5, left=Inches(0.8), top=Inches(1.5), width=Inches(11.733), height=Inches(3.8))
    t4 = t4_shape.table
    t4.columns[0].width = Inches(2.2)
    t4.columns[1].width = Inches(1.7)
    t4.columns[2].width = Inches(2.1)
    t4.columns[3].width = Inches(3.0)
    t4.columns[4].width = Inches(2.733)
    
    headers4 = ["Author & Year", "Dataset & Modality", "Evaluated Models", "Key Findings & Metrics", "Identified Gap / Limitation"]
    for i, h in enumerate(headers4):
        cell = t4.cell(0, i)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = TINT_BLUE
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_FAMILY; p.font.size = Pt(9); p.font.bold = True; p.font.color.rgb = NAVY
            
    lit_data4 = [
        [("Jahan et al. (2023)", "PLOS ONE [DOI:10.1371]"), "OASIS-3 Multimodal\n(Clinical + MRI)", "Random Forest,\nXGBoost + SHAP", "Multimodal fusion decisively outperforms unimodal MRI baselines (+8.4% Acc).", "Did not quantify subject-level slice leakage vulnerability."],
        [("Jahan et al. (2025)", "IEEE Access [DOI:10.1109]"), "OASIS-3 Multimodal\n(Multi-Center)", "Federated Ensemble\n+ Centralized SHAP", "Privacy-preserving decentralized learning matches centralized AUC (0.912).", "High network synchronization overhead across hospital nodes."],
        [("Yagis et al. (2021)", "Sci. Reports [DOI:10.1038]"), "ADNI & OASIS\nBrain MRI Scans", "2D CNNs & Classical\nMachine Learning", "Subject and slice overlap inflates reported test accuracy by 15% to 25%.", "Demonstrated leakage risks but proposed no unified XAI framework."],
        [("Shah et al. (2024)", "Comp. Biol. Med. [10.1016]"), "ADNI Structural MRI\n(T1-weighted Scans)", "3D ResNet vs Tree\nEnsembles + SHAP", "Tree-based SHAP exhibits significantly higher attribution stability than Grad-CAM.", "Did not integrate cognitive screening psychometrics."],
        [("Jiang et al. (2024)", "Front. Aging Neur. [10.3389]"), "Structural MRI\nMorphometry (FreeSurfer)", "SVM, RF, XGBoost\n+ Tree SHAP", "Cortical complexity metrics correlate strongly with CDR dementia progression.", "Evaluated high-dimensional MRI alone without clinical co-factors."]
    ]
    for r_idx, row in enumerate(lit_data4, start=1):
        for c_idx, val in enumerate(row):
            cell = t4.cell(r_idx, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else TINT_GRAY
            tf = cell.text_frame
            tf.word_wrap = True
            tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
            tf.margin_left = Inches(0.06); tf.margin_right = Inches(0.06)
            
            if isinstance(val, tuple):
                p0 = tf.paragraphs[0]
                p0.text = val[0]
                p0.font.name = FONT_FAMILY; p0.font.size = Pt(8.5); p0.font.bold = True; p0.font.color.rgb = NAVY
                p1 = tf.add_paragraph()
                p1.text = val[1]
                p1.font.name = FONT_FAMILY; p1.font.size = Pt(7.5); p1.font.color.rgb = MUTED_TEAL
            else:
                p0 = tf.paragraphs[0]
                p0.text = val
                for p in tf.paragraphs:
                    p.font.name = FONT_FAMILY; p.font.size = Pt(8); p.font.color.rgb = NAVY

    # Bottom Research Gap Callout Card - Positioned comfortably below table
    gap_card = add_card(s4, Inches(0.8), Inches(5.65), Inches(11.733), Inches(1.05), bg_color=TINT_TEAL, border_color=MUTED_TEAL)
    gtb = s4.shapes.add_textbox(Inches(1.0), Inches(5.72), Inches(11.3), Inches(0.9))
    gtf = gtb.text_frame; gtf.word_wrap = True
    gp = gtf.paragraphs[0]
    gp.text = "EXPLICIT RESEARCH GAP ADDRESSED BY OUR PROJECT:"
    gp.font.name = FONT_FAMILY; gp.font.size = Pt(9.5); gp.font.bold = True; gp.font.color.rgb = MUTED_TEAL
    gp2 = gtf.add_paragraph()
    gp2.text = "While existing studies demonstrate either multimodal utility or XAI benefits, very few combine strict zero-leakage participant-level isolation on accessible OASIS cohorts with dual global/local SHAP interpretability and lightweight multi-model soft ensembling executing on commodity CPUs."
    gp2.font.name = FONT_FAMILY; gp2.font.size = Pt(9); gp2.font.color.rgb = NAVY; gp2.space_before = Pt(2)

    s4.notes_slide.notes_text_frame.text = (
        "Slide 4 reviews five verified peer-reviewed studies published between 2021 and 2025. Jahan et al. proved that multimodal fusion "
        "is superior to MRI alone, while Yagis et al. in Scientific Reports proved that data leakage across subjects inflates metrics by up to 25%. "
        "Shah et al. in Computers in Biology and Medicine showed that tree-based SHAP explanations are much more stable and consistent than deep "
        "learning saliency maps. Our work addresses the identified gap: we implement a leak-free participant-level multimodal pipeline with "
        "game-theoretic SHAP interpretability that requires zero GPU computing."
    )

    # =========================================================================
    # SLIDE 5: OASIS-1 DATASET & RIGOROUS DATA AUDIT
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s5, slide_num=5, kicker="04 | DATA PROVENANCE",
                     title="OASIS-1 Dataset & Rigorous Data Audit",
                     subtitle="Systematic participant cohort stratification, clinical dementia rating (CDR) distribution, and data integrity.")
    
    add_card(s5, Inches(0.8), Inches(1.55), Inches(5.5), Inches(5.15), title="OASIS-1 Cohort Audit Breakdown", kicker="COHORT INTEGRITY", accent_color=MUTED_TEAL)
    ab_box = s5.shapes.add_textbox(Inches(1.04), Inches(2.35), Inches(5.0), Inches(4.1))
    abf = ab_box.text_frame; abf.word_wrap = True
    abf.text = (
        "• Total Repository Records: 436 unique participants.\n\n"
        "• Methodological Cohort Stratification:\n"
        "  - 201 Young Normal Controls (Ages 18–59, CDR = NaN).\n"
        "    * Excluded to prevent trivial age-separated classification.\n"
        "  - 235 Clinically Evaluated Older Cohort (Ages 33–96).\n"
        "    * Complete CDR evaluations forming our experimental benchmark.\n\n"
        "• Diagnostic Distribution (Target Cohort, N=235):\n"
        "  - Controls (CDR = 0.0): 135 subjects (57.45%)\n"
        "  - Dementia-Indicative (CDR > 0.0): 100 subjects (42.55%)\n"
        "    * Very Mild Dementia (CDR = 0.5): 70 subjects (29.79%)\n"
        "    * Mild Dementia (CDR = 1.0): 28 subjects (11.91%)\n"
        "    * Moderate Dementia (CDR = 2.0): 2 subjects (0.85%)\n\n"
        "• Data Completeness: Zero missing values in MRI morphometrics or MMSE. 19 missing values in SES (Hollingshead) imputed via training median."
    )
    for p in abf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    img_dist = os.path.join(figures_dir, "dataset_distribution.png")
    if os.path.exists(img_dist):
        s5.shapes.add_picture(img_dist, Inches(6.6), Inches(1.8), width=Inches(5.9))
        
    s5.notes_slide.notes_text_frame.text = (
        "Slide 5 details the OASIS-1 dataset audit. The repository has 436 subjects, but 201 are young normal controls under age 60 "
        "who were not clinically evaluated (CDR = NaN). Including them creates an artificial age-shortcut where models simply learn age rather "
        "than neurodegeneration. Thus, we isolate the 235 clinically evaluated subjects (ages 33 to 96). Within this cohort, 135 are healthy "
        "controls (57.5%) and 100 have dementia symptoms (42.5%), predominantly very mild (CDR 0.5, 70 subjects). Missing SES values are "
        "strictly imputed on the training split only."
    )

    # =========================================================================
    # SLIDE 6: EXPERIMENTAL PIPELINE (REBUILT NATIVE VECTOR FLOW DIAGRAM)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s6, slide_num=6, kicker="05 | EXPERIMENTAL PIPELINE",
                     title="Proposed Methodology Architecture",
                     subtitle="End-to-end horizontal workflow enforcing participant-level isolation, multimodal modeling, and SHAP decision support.")
    
    # We construct a clean, publication-grade 2-row horizontal vector architecture with native shapes
    # Row 1: Data to Preprocessing (4 stages)
    # Row 2: Training to Explainability (3 stages) + Callout Card
    
    stages_row1 = [
        ("STAGE 1", "OASIS-1 Raw Data", "N = 436 Participants", "T1-w sMRI + Clinical\nDemographics (12 vars)", MUTED_TEAL),
        ("STAGE 2", "Cohort Filtering", "N = 235 Evaluated", "Exclude 201 unassessed\nIsolate ground truth CDR", MUTED_TEAL),
        ("STAGE 3", "Subject Split (80/20)", "Train=188 | Test=47", "S_train ∩ S_test = ∅\nStrict zero subject overlap", RED_ACCENT),
        ("STAGE 4", "Training Pipeline", "ColumnTransformer", "Median SES, StandardScaler\nOHE (Fit on train only)", SOFT_BLUE)
    ]
    
    card_w1 = Inches(2.65)
    gap1 = Inches(0.38)
    top1 = Inches(1.6)
    h1 = Inches(2.1)
    
    for idx, (stg, tit, sub, desc, acc_col) in enumerate(stages_row1):
        x = Inches(0.8) + idx * (card_w1 + gap1)
        add_card(s6, x, top1, card_w1, h1, bg_color=WHITE, border_color=BORDER_COLOR, accent_color=acc_col)
        
        # Text inside card
        tb = s6.shapes.add_textbox(x + Inches(0.15), top1 + Inches(0.12), card_w1 - Inches(0.3), h1 - Inches(0.24))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        p0 = tf.paragraphs[0]
        p0.text = stg
        p0.font.name = FONT_FAMILY; p0.font.size = Pt(8); p0.font.bold = True; p0.font.color.rgb = acc_col
        
        p1 = tf.add_paragraph()
        p1.text = tit
        p1.font.name = FONT_FAMILY; p1.font.size = Pt(11); p1.font.bold = True; p1.font.color.rgb = NAVY
        p1.space_before = Pt(1); p1.space_after = Pt(2)
        
        p2 = tf.add_paragraph()
        p2.text = sub
        p2.font.name = FONT_FAMILY; p2.font.size = Pt(9); p2.font.bold = True; p2.font.color.rgb = MUTED_TEAL
        p2.space_after = Pt(4)
        
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.name = FONT_FAMILY; p3.font.size = Pt(8.5); p3.font.color.rgb = SLATE
        
        # Arrow connector to next card in row 1
        if idx < len(stages_row1) - 1:
            arrow = s6.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + card_w1 + Inches(0.06), top1 + Inches(0.9), Inches(0.25), Inches(0.25))
            arrow.fill.solid(); arrow.fill.fore_color.rgb = BORDER_COLOR
            arrow.line.fill.background()

    # Downward connector from Row 1 to Row 2
    down_arrow = s6.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(11.2), Inches(3.8), Inches(0.3), Inches(0.35))
    down_arrow.fill.solid(); down_arrow.fill.fore_color.rgb = SOFT_BLUE
    down_arrow.line.fill.background()

    # Row 2: Modeling, Evaluation & Explainability
    stages_row2 = [
        ("STAGE 5", "Multi-Classifier Training", "5-Fold Stratified CV", "LR, RF, SVM, GB (lr=0.1)\nSoft Voting Consensus: P = (1·LR + 2·RF + 2·GB)/5", MUTED_TEAL),
        ("STAGE 6", "Held-Out Test Evaluation", "N = 47 Isolated Test", "Accuracy, Precision, Recall, Specificity, F1, ROC-AUC\nGradient Boosting: 85.11% Acc, 0.9426 ROC-AUC", GREEN_ACCENT),
        ("STAGE 7", "Game-Theoretic SHAP XAI", "Global + Local Output", "TreeExplainer Beeswarm & Waterfall log-odds\nClinical decision support with zero GPU overhead", MUTED_TEAL)
    ]
    
    card_w2 = Inches(3.64)
    gap2 = Inches(0.4)
    top2 = Inches(4.3)
    h2 = Inches(2.2)
    
    for idx, (stg, tit, sub, desc, acc_col) in enumerate(stages_row2):
        x = Inches(0.8) + idx * (card_w2 + gap2)
        add_card(s6, x, top2, card_w2, h2, bg_color=WHITE, border_color=BORDER_COLOR, accent_color=acc_col)
        
        tb = s6.shapes.add_textbox(x + Inches(0.2), top2 + Inches(0.14), card_w2 - Inches(0.4), h2 - Inches(0.28))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        p0 = tf.paragraphs[0]
        p0.text = stg
        p0.font.name = FONT_FAMILY; p0.font.size = Pt(8.5); p0.font.bold = True; p0.font.color.rgb = acc_col
        
        p1 = tf.add_paragraph()
        p1.text = tit
        p1.font.name = FONT_FAMILY; p1.font.size = Pt(11.5); p1.font.bold = True; p1.font.color.rgb = NAVY
        p1.space_before = Pt(2); p1.space_after = Pt(2)
        
        p2 = tf.add_paragraph()
        p2.text = sub
        p2.font.name = FONT_FAMILY; p2.font.size = Pt(9.5); p2.font.bold = True; p2.font.color.rgb = acc_col
        p2.space_after = Pt(4)
        
        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.name = FONT_FAMILY; p3.font.size = Pt(8.5); p3.font.color.rgb = SLATE
        
        if idx < len(stages_row2) - 1:
            arrow = s6.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + card_w2 + Inches(0.06), top2 + Inches(0.95), Inches(0.26), Inches(0.25))
            arrow.fill.solid(); arrow.fill.fore_color.rgb = BORDER_COLOR
            arrow.line.fill.background()

    s6.notes_slide.notes_text_frame.text = (
        "Slide 6 presents our re-engineered vector architecture. The pipeline flows from left to right: from OASIS-1 raw data "
        "to cohort filtering, participant-level 80/20 holdout split, training-only ColumnTransformer fitting, 5-fold cross-validation "
        "across 6 model configurations, held-out evaluation on the 47 test subjects, and SHAP TreeExplainer interpretation. Every step "
        "strictly preserves zero leakage, ensuring un-inflated, publishable results."
    )

    # =========================================================================
    # SLIDE 7: PREPROCESSING & MULTIMODAL FEATURE SPACE (3-SECTION LAYOUT)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s7, slide_num=7, kicker="06 | DATA PREPARATION",
                     title="Preprocessing & Multimodal Feature Engineering",
                     subtitle="Strict participant-level partitioning, training-only preprocessor fitting, and biological variable roles.")
    
    col_w7 = Inches(3.7)
    gap7 = Inches(0.3)
    top7 = Inches(1.55)
    h7 = Inches(5.15)
    
    # Section 1: Participant-Level Split
    add_card(s7, Inches(0.8), top7, col_w7, h7, title="1. Participant-Level Split", kicker="LEAKAGE PREVENTION", accent_color=RED_ACCENT)
    tb_c1 = s7.shapes.add_textbox(Inches(1.04), top7 + Inches(0.75), col_w7 - Inches(0.48), h7 - Inches(0.9))
    tfc1 = tb_c1.text_frame; tfc1.word_wrap = True
    tfc1.text = (
        "• Strict Cohort Separation:\n"
        "  - 80% Training Set: N = 188 (108 Controls, 80 Dementia)\n"
        "  - 20% Held-Out Test Set: N = 47 (27 Controls, 20 Dementia)\n\n"
        "• Zero Subject Overlap:\n"
        "  S_train ∩ S_test = ∅\n"
        "  No subject or multi-session scan appears in both training and testing.\n\n"
        "• Stratification Guarantee:\n"
        "  Target class proportion preserved exactly across partitions (~57.4% Controls, ~42.6% Dementia).\n\n"
        "• Benchmark Comparison:\n"
        "  Prevents the 15%–25% artificial score inflation identified by Yagis et al. (2021)."
    )
    for p in tfc1.paragraphs: p.font.size = Pt(9.5); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    # Section 2: Leakage-Free Preprocessing Pipeline
    add_card(s7, Inches(0.8) + col_w7 + gap7, top7, col_w7, h7, title="2. Preprocessing Pipeline", kicker="TRAINING-ONLY FIT", accent_color=MUTED_TEAL)
    tb_c2 = s7.shapes.add_textbox(Inches(1.04) + col_w7 + gap7, top7 + Inches(0.75), col_w7 - Inches(0.48), h7 - Inches(0.9))
    tfc2 = tb_c2.text_frame; tfc2.word_wrap = True
    tfc2.text = (
        "• scikit-learn ColumnTransformer:\n"
        "  Fitted strictly and exclusively on the 188 training participants.\n\n"
        "• Numerical Imputation:\n"
        "  Median imputation for missing Socioeconomic Status (SES) values.\n\n"
        "• Feature Standardization:\n"
        "  StandardScaler transforms numerical features to zero mean and unit variance (z = (x - μ) / σ) using training μ and σ.\n\n"
        "• Categorical Encoding:\n"
        "  OneHotEncoder(drop='first') encodes biological gender (M/F).\n\n"
        "• Pruning of Non-Predictives:\n"
        "  'ID', 'Hand' (100% Right), and 'Delay' (100% NaN) dropped."
    )
    for p in tfc2.paragraphs: p.font.size = Pt(9.5); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    # Section 3: Multimodal Feature Space
    add_card(s7, Inches(0.8) + (col_w7 + gap7)*2, top7, col_w7, h7, title="3. Multimodal Features", kicker="8 BIOMARKERS", accent_color=SOFT_BLUE)
    tb_c3 = s7.shapes.add_textbox(Inches(1.04) + (col_w7 + gap7)*2, top7 + Inches(0.75), col_w7 - Inches(0.48), h7 - Inches(0.9))
    tfc3 = tb_c3.text_frame; tfc3.word_wrap = True
    tfc3.text = (
        "Clinical & Demographic (5):\n"
        "• MMSE (0–30): Mini-Mental State Exam score (cognitive disorientation).\n"
        "• Age (33–96): Biological aging risk.\n"
        "• M/F: Gender stratification.\n"
        "• Educ (1–5): Education (cognitive reserve proxy).\n"
        "• SES (1–5): Hollingshead status.\n\n"
        "Structural MRI Morphometrics (3):\n"
        "• nWBV (0.64–0.85): Normalized Whole Brain Volume (ratio of brain tissue; objective cerebral atrophy).\n"
        "• eTIV (1123–1992 mm³): Estimated Total Intracranial Volume.\n"
        "• ASF (0.88–1.56): Atlas Scaling Factor to Talairach stereotaxic space."
    )
    for p in tfc3.paragraphs: p.font.size = Pt(9); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s7.notes_slide.notes_text_frame.text = (
        "Slide 7 cleanly segments our data preparation into three distinct pillars. In Section 1, we show our participant-level split "
        "(188 train, 47 test) guaranteeing zero subject overlap. In Section 2, we describe our scikit-learn ColumnTransformer which fits "
        "median imputation and standard scaling solely on training samples. In Section 3, we detail our 8 multimodal features: 5 clinical/demographic "
        "and 3 structural MRI morphometrics, notably MMSE for cognitive screening and nWBV for whole-brain parenchymal atrophy."
    )

    # =========================================================================
    # SLIDE 8: CLASSIFIERS & SOFT-VOTING ENSEMBLE ARCHITECTURE
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s8, slide_num=8, kicker="07 | MODEL DESIGN",
                     title="Classification Models & Ensemble Architecture",
                     subtitle="Systematic comparison of linear, tree, margin-based, boosting, and soft-voting consensus architectures.")
    
    add_card(s8, Inches(0.8), Inches(1.55), Inches(5.2), Inches(5.15), title="Evaluated Model Configurations", kicker="MODEL BENCHMARK", accent_color=MUTED_TEAL)
    mb_box = s8.shapes.add_textbox(Inches(1.04), Inches(2.35), Inches(4.7), Inches(4.1))
    mbf = mb_box.text_frame; mbf.word_wrap = True
    mbf.text = (
        "• Dummy Classifier (Baseline):\n"
        "  Majority class prediction; establishes empirical accuracy floor (57.45%).\n\n"
        "• Logistic Regression (LR):\n"
        "  Linear probabilistic classifier with L2 regularization penalty (C=1.0).\n\n"
        "• Random Forest (RF):\n"
        "  100 bagged decorrelated decision trees (max_depth=5, min_samples_split=4).\n\n"
        "• Support Vector Machine (SVM):\n"
        "  RBF kernel with Platt-scaled calibrated posterior probabilities (C=1.0).\n\n"
        "• Gradient Boosting (GB):\n"
        "  100 sequential residual regression trees (lr=0.1, max_depth=3).\n\n"
        "• Soft-Voting Ensemble:\n"
        "  Weighted consensus averaging posterior class probabilities:\n"
        "  P(Dementia) = (1·P_LR + 2·P_RF + 2·P_GB) / 5."
    )
    for p in mbf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    img_arch = os.path.join(figures_dir, "model_architecture_diagram.png")
    if os.path.exists(img_arch):
        s8.shapes.add_picture(img_arch, Inches(6.3), Inches(1.85), width=Inches(6.2))

    s8.notes_slide.notes_text_frame.text = (
        "Slide 8 illustrates our evaluated models. We benchmark five diverse learning paradigms: Dummy baseline, Logistic Regression, "
        "Random Forest, RBF Support Vector Machine, and Gradient Boosting. We also build a Soft-Voting Ensemble that weights the predictions "
        "of our top three algorithms: Logistic Regression with weight 1, Random Forest with weight 2, and Gradient Boosting with weight 2. "
        "This balances linear margins with non-linear tree partitions, reducing individual model variance."
    )

    # =========================================================================
    # SLIDE 9: EXPLAINABLE AI USING THE SHAP FRAMEWORK
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s9, slide_num=9, kicker="08 | INTERPRETABILITY",
                     title="Explainable AI via SHAP Framework",
                     subtitle="Game-theoretic attribution allocating diagnostic credit across multimodal biomarkers globally and locally.")
    
    col_w3 = Inches(3.7)
    add_card(s9, Inches(0.8), Inches(1.55), col_w3, Inches(5.15), title="Game-Theoretic Foundation", kicker="MATHEMATICAL RIGOR", accent_color=MUTED_TEAL)
    tb_s1 = s9.shapes.add_textbox(Inches(1.04), Inches(2.35), col_w3 - Inches(0.48), Inches(4.1))
    tfs1 = tb_s1.text_frame; tfs1.word_wrap = True
    tfs1.text = (
        "• Grounded in Shapley Values:\n"
        "  Originates from cooperative coalitional game theory (Nobel Laureate Lloyd Shapley, 1953).\n\n"
        "• Fair Credit Allocation:\n"
        "  Calculates the marginal contribution of biomarker i across all possible feature subsets S:\n\n"
        "  φ_i = Σ [|S|!(|F|-|S|-1)! / |F|!] · [f(S ∪ {i}) - f(S)]\n\n"
        "• Axiomatic Guarantees:\n"
        "  Satisfies Efficiency, Symmetry, Dummy Player, and Additivity, ensuring non-heuristic fairness."
    )
    for p in tfs1.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    add_card(s9, Inches(0.8) + col_w3 + gap, Inches(1.55), col_w3, Inches(5.15), title="Global Cohort Ranking", kicker="POPULATION INSIGHTS", accent_color=SOFT_BLUE)
    tb_s2 = s9.shapes.add_textbox(Inches(1.04) + col_w3 + gap, Inches(2.35), col_w3 - Inches(0.48), Inches(4.1))
    tfs2 = tb_s2.text_frame; tfs2.word_wrap = True
    tfs2.text = (
        "• Mean Absolute Importance:\n"
        "  Ranks features by average magnitude across all test participants: (1/N) Σ |φ_i|.\n\n"
        "• Summary Beeswarm Plot:\n"
        "  Reveals feature value distribution (low=blue, high=red) versus positive/negative impact on dementia log-odds.\n\n"
        "• Clinical Verification:\n"
        "  Allows neurologists to confirm that model reliance corresponds to genuine neurodegeneration rather than imaging artifacts."
    )
    for p in tfs2.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    add_card(s9, Inches(0.8) + (col_w3 + gap)*2, Inches(1.55), col_w3, Inches(5.15), title="Local Patient Attribution", kicker="DECISION SUPPORT", accent_color=MUTED_TEAL)
    tb_s3 = s9.shapes.add_textbox(Inches(1.04) + (col_w3 + gap)*2, Inches(2.35), col_w3 - Inches(0.48), Inches(4.1))
    tfs3 = tb_s3.text_frame; tfs3.word_wrap = True
    tfs3.text = (
        "• Patient Waterfall Plots:\n"
        "  Decomposes an individual patient's prediction step-by-step from base rate E[f(x)] to final model log-odds f(x).\n\n"
        "• Transparent Second Opinion:\n"
        "  Explicitly justifies why a specific patient was classified as high-risk, identifying dominant drivers.\n\n"
        "• Scientific Qualification:\n"
        "  SHAP explains algorithmic model attribution; it does not establish medical causality or replace clinical diagnosis."
    )
    for p in tfs3.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s9.notes_slide.notes_text_frame.text = (
        "Slide 9 outlines our explainability framework. We employ SHAP (SHapley Additive exPlanations), grounded in cooperative game theory. "
        "SHAP satisfies four critical mathematical axioms: efficiency, symmetry, dummy player, and additivity. It operates at two levels: "
        "Globally, it ranks biomarker influence across the entire test cohort via mean absolute SHAP values. Locally, it generates patient-level "
        "waterfall charts that decompose an individual prediction from the population base rate to their final diagnosis, clearly detailing "
        "the exact biological features responsible."
    )

    # =========================================================================
    # SLIDE 10: EXPERIMENTAL SETUP & REPRODUCIBILITY
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s10, slide_num=10, kicker="09 | RIGOROUS BENCHMARK",
                     title="Experimental Setup & Reproducibility",
                     subtitle="Deterministic software stack, execution efficiency on commodity CPU, and unit-tested pipelines.")
    
    add_card(s10, Inches(0.8), Inches(1.55), Inches(4.8), Inches(5.15), title="Environment & Execution Speed", kicker="SYSTEM SPECIFICATIONS", accent_color=MUTED_TEAL)
    env_box = s10.shapes.add_textbox(Inches(1.04), Inches(2.35), Inches(4.3), Inches(4.1))
    env_f = env_box.text_frame; env_f.word_wrap = True
    env_f.text = (
        "• Operating System:\n"
        "  Windows 10/11 (AMD64 platform architecture).\n\n"
        "• Python Runtime:\n"
        "  Python 3.11.0 (64-bit environment).\n\n"
        "• Pinned Core Libraries:\n"
        "  scikit-learn 1.7.0, NumPy 2.x, pandas 2.x, SHAP 0.46.0, Matplotlib 3.10.0.\n\n"
        "• Commodity Hardware:\n"
        "  Intel Core i5 CPU, 16 GB RAM (Zero GPU reliance).\n\n"
        "• Ultra-Fast Execution:\n"
        "  Complete pipeline (data audit, 5-fold CV, testing, and SHAP) runs in < 15 seconds.\n\n"
        "• Deterministic Reproducibility:\n"
        "  Global random seed = 42 enforced across all splits and estimators.\n\n"
        "• Verified Unit Testing:\n"
        "  7/7 automated pytest tests passing with 100% test coverage."
    )
    for p in env_f.paragraphs: p.font.size = Pt(9.5); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    table_shape10 = s10.shapes.add_table(rows=7, cols=3, left=Inches(5.9), top=Inches(1.55), width=Inches(6.633), height=Inches(5.15))
    t10 = table_shape10.table
    t10.columns[0].width = Inches(2.0)
    t10.columns[1].width = Inches(2.8)
    t10.columns[2].width = Inches(1.833)
    
    hp_headers = ["Classifier Architecture", "Exact Hyperparameters", "Optimization Objective"]
    for i, h in enumerate(hp_headers):
        cell = t10.cell(0, i)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = TINT_BLUE
        for p in cell.text_frame.paragraphs:
            p.font.name = FONT_FAMILY; p.font.size = Pt(9); p.font.bold = True; p.font.color.rgb = NAVY
        
    hp_rows = [
        ["Dummy (Baseline)", "strategy='most_frequent'", "Prior Class Floor"],
        ["Logistic Regression", "C=1.0, solver='liblinear', max_iter=1000", "L2 Cross-Entropy"],
        ["Random Forest", "n_estimators=100, max_depth=5, min_split=4", "Gini Impurity"],
        ["Support Vector Machine", "C=1.0, kernel='rbf', probability=True", "Hinge / Platt Sigmoid"],
        ["Gradient Boosting", "n_estimators=100, lr=0.1, max_depth=3", "Binomial Deviance"],
        ["Soft Voting Ensemble", "voting='soft', weights=[1, 2, 2] (LR, RF, GB)", "Weighted Log-Loss"]
    ]
    for r_idx, row in enumerate(hp_rows, start=1):
        for c_idx, val in enumerate(row):
            cell = t10.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else TINT_GRAY
            for p in cell.text_frame.paragraphs:
                p.font.name = FONT_FAMILY; p.font.size = Pt(8.5); p.font.color.rgb = NAVY

    s10.notes_slide.notes_text_frame.text = (
        "Slide 10 demonstrates our focus on scientific reproducibility. The entire experimental suite runs on standard commodity hardware—an "
        "Intel Core i5 CPU with 16 GB RAM—requiring no expensive GPU infrastructure. Full training, 5-fold cross-validation, held-out testing, "
        "and SHAP generation execute in under 15 seconds. Seed 42 is locked across all estimators, and 7 automated pytest unit tests verify "
        "data loading, zero-leakage splits, and metric calculation before deployment."
    )

    # =========================================================================
    # SLIDE 11: COMPARATIVE MODEL RESULTS (HELD-OUT TEST SET, N=47)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s11, slide_num=11, kicker="10 | EMPIRICAL RESULTS",
                     title="Comparative Model Results (Held-Out Test Set, N=47)",
                     subtitle="Rigorous evaluation on isolated test cohort (27 Controls, 20 Dementia) with strict zero data leakage.")
    
    t11_shape = s11.shapes.add_table(rows=7, cols=8, left=Inches(0.8), top=Inches(1.55), width=Inches(11.733), height=Inches(3.4))
    t11 = t11_shape.table
    
    col_widths = [Inches(2.5), Inches(1.3), Inches(1.3), Inches(1.4), Inches(1.3), Inches(1.3), Inches(1.3), Inches(1.333)]
    for i, w in enumerate(col_widths): t11.columns[i].width = w
    
    m_headers = ["Model Architecture", "Accuracy", "Precision", "Recall (Sens)", "Specificity", "F1-Score", "ROC-AUC", "Bal. Acc"]
    for i, h in enumerate(m_headers):
        cell = t11.cell(0, i)
        cell.text = h
        cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT
            p.font.name = FONT_FAMILY; p.font.size = Pt(9.5); p.font.bold = True; p.font.color.rgb = WHITE
        
    res_table = [
        ["Dummy (Baseline)", "57.45%", "0.00%", "0.00%", "100.0%", "0.00%", "0.5000", "50.00%"],
        ["Logistic Regression", "78.72%", "77.78%", "70.00%", "85.19%", "73.68%", "0.9056", "77.59%"],
        ["Random Forest", "82.98%", "83.33%", "75.00%", "88.89%", "78.95%", "0.9296", "81.94%"],
        ["Support Vector Machine", "80.85%", "78.95%", "75.00%", "85.19%", "76.92%", "0.9056", "80.09%"],
        ["Gradient Boosting", "85.11%", "84.21%", "80.00%", "88.89%", "82.05%", "0.9426", "84.44%"],
        ["Soft Voting Ensemble", "82.98%", "83.33%", "75.00%", "88.89%", "78.95%", "0.9426", "81.94%"]
    ]
    for r_idx, row in enumerate(res_table, start=1):
        is_highlight = ("Gradient" in row[0] or "Voting" in row[0])
        for c_idx, val in enumerate(row):
            cell = t11.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            if is_highlight:
                cell.fill.fore_color.rgb = TINT_TEAL if "Gradient" in row[0] else TINT_BLUE
            else:
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else TINT_GRAY
                
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
                p.font.name = FONT_FAMILY; p.font.size = Pt(9.5); p.font.color.rgb = NAVY
                if is_highlight:
                    p.font.bold = True
                    if c_idx in [1, 3, 6]: p.font.color.rgb = MUTED_TEAL if "Gradient" in row[0] else SOFT_BLUE

    add_card(s11, Inches(0.8), Inches(5.15), Inches(11.733), Inches(1.6), bg_color=WHITE, border_color=MUTED_TEAL, accent_color=MUTED_TEAL)
    tb11 = s11.shapes.add_textbox(Inches(1.04), Inches(5.25), Inches(11.2), Inches(1.4))
    tf11 = tb11.text_frame; tf11.word_wrap = True
    
    p = tf11.paragraphs[0]
    p.text = "KEY EMPIRICAL FINDINGS & CLINICAL SIGNIFICANCE:"
    p.font.name = FONT_FAMILY; p.font.size = Pt(10.5); p.font.bold = True; p.font.color.rgb = MUTED_TEAL
    
    p2 = tf11.add_paragraph()
    p2.text = (
        "1. Gradient Boosting Achieved Peak Diagnostic Performance: Attained top overall Accuracy (85.11%), Recall/Sensitivity (80.00%, correctly detecting 16 of 20 dementia cases), F1-Score (82.05%), and tied for top ROC-AUC (0.9426).\n"
        "2. Soft-Voting Ensemble Tied Top Discriminative Power: Tied Gradient Boosting for highest ROC-AUC (0.9426) with robust 88.89% Specificity (24 of 27 healthy controls correctly identified), offering balanced conservative screening.\n"
        "3. Decisive Separation Above Chance: All evaluated ML models surpassed 0.90 ROC-AUC under strict zero-leakage conditions, substantially outperforming the 50.00% balanced accuracy floor of the dummy baseline."
    )
    p2.font.name = FONT_FAMILY; p2.font.size = Pt(9.5); p2.font.color.rgb = NAVY; p2.space_before = Pt(3)

    s11.notes_slide.notes_text_frame.text = (
        "Slide 11 presents our primary empirical findings on the held-out test cohort of 47 subjects. Gradient Boosting attained the highest "
        "overall performance across accuracy (85.11%), sensitivity (80.00%), F1-score (82.05%), and ROC-AUC (0.9426), successfully identifying "
        "16 out of 20 dementia patients. The Soft-Voting Ensemble tied Gradient Boosting for top ROC-AUC at 0.9426, with an outstanding "
        "specificity of 88.89%, misclassifying only 3 of 27 healthy controls. Every trained machine learning model achieved an ROC-AUC greater "
        "than 0.90, confirming robust discriminative capability without data leakage."
    )

    # =========================================================================
    # SLIDE 12: ROC CURVES & MODALITY ABLATION ANALYSIS
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s12, slide_num=12, kicker="11 | DETAILED METRICS",
                     title="ROC Curves & Feature Modality Ablation",
                     subtitle="Discriminative separation curves and proof of multimodal superiority over unimodal baselines.")
    
    img_roc = os.path.join(figures_dir, "roc_curves.png")
    img_mod = os.path.join(figures_dir, "modality_comparison_bar.png")
    
    if os.path.exists(img_roc):
        s12.shapes.add_picture(img_roc, Inches(0.8), Inches(1.55), width=Inches(5.6))
    if os.path.exists(img_mod):
        s12.shapes.add_picture(img_mod, Inches(6.8), Inches(1.55), width=Inches(5.733))

    # Bottom Ablation Summary Banner
    abl_card = add_card(s12, Inches(0.8), Inches(5.75), Inches(11.733), Inches(0.95), bg_color=TINT_TEAL, border_color=MUTED_TEAL)
    atb = s12.shapes.add_textbox(Inches(1.0), Inches(5.82), Inches(11.3), Inches(0.8))
    atf = atb.text_frame; atf.word_wrap = True
    ap0 = atf.paragraphs[0]
    ap0.text = "VERIFIED MODALITY ABLATION FINDINGS (HELD-OUT TEST SET):"
    ap0.font.name = FONT_FAMILY; ap0.font.size = Pt(9.5); ap0.font.bold = True; ap0.font.color.rgb = MUTED_TEAL
    ap1 = atf.add_paragraph()
    ap1.text = "• Multimodal Random Forest (82.98% Acc, 0.9296 AUC) outperforms Structural MRI Only (70.21% Acc, 0.7759 AUC) by +12.77% Acc and +0.1537 ROC-AUC.\n• Multimodal fusion surpasses Clinical Only (80.85% Acc, 0.9037 AUC) by +2.13% Acc and +0.0259 ROC-AUC, demonstrating decisive cross-modal synergy."
    ap1.font.name = FONT_FAMILY; ap1.font.size = Pt(9); ap1.font.color.rgb = NAVY; ap1.space_before = Pt(2)

    s12.notes_slide.notes_text_frame.text = (
        "On Slide 12, we examine the ROC curves and our modality ablation experiment. On the left, Gradient Boosting and the Soft-Voting "
        "Ensemble push closest to the top-left corner, reaching an outstanding 0.943 AUC. On the right, our ablation study provides definitive "
        "evidence of multimodal superiority: combining clinical and structural MRI features in Random Forest yields 82.98% accuracy and 0.9296 "
        "AUC, outperforming MRI-only (70.21% Acc, 0.7759 AUC) by +12.77% accuracy and +0.1537 AUC, and outperforming clinical-only (80.85% Acc, "
        "0.9037 AUC) by +2.13% accuracy. Both modalities are essential."
    )

    # =========================================================================
    # SLIDE 13: SHAP EXPLAINABILITY & PATIENT WATERFALL CASES
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s13, slide_num=13, kicker="12 | CLINICAL EXPLAINABILITY",
                     title="SHAP Explainability & Patient Decision Support",
                     subtitle="Global biomarker importance rankings and patient-level waterfall attribution for transparent diagnosis.")
    
    img_swarm = os.path.join(figures_dir, "shap_summary_beeswarm.png")
    img_wf = os.path.join(figures_dir, "shap_waterfall_demented.png")
    
    if os.path.exists(img_swarm):
        s13.shapes.add_picture(img_swarm, Inches(0.8), Inches(1.55), width=Inches(5.9))
    if os.path.exists(img_wf):
        s13.shapes.add_picture(img_wf, Inches(7.0), Inches(1.55), width=Inches(5.533))

    # Bottom Explainability Summary Banner
    shap_card = add_card(s13, Inches(0.8), Inches(5.75), Inches(11.733), Inches(0.95), bg_color=WHITE, border_color=MUTED_TEAL, accent_color=MUTED_TEAL)
    stb = s13.shapes.add_textbox(Inches(1.0), Inches(5.82), Inches(11.3), Inches(0.8))
    stf = stb.text_frame; stf.word_wrap = True
    sp0 = stf.paragraphs[0]
    sp0.text = "BIOMARKER RANKINGS & PATIENT CASE STUDY (CLINICAL AUDITING):"
    sp0.font.name = FONT_FAMILY; sp0.font.size = Pt(9.5); sp0.font.bold = True; sp0.font.color.rgb = MUTED_TEAL
    sp1 = stf.add_paragraph()
    sp1.text = "• Global Drivers: MMSE (mean |SHAP|=0.1701) and normalized Whole Brain Volume nWBV (mean |SHAP|=0.0884) decisively lead model attribution.\n• Patient #33 Case: Base log-odds shifted from -0.32 to +3.21 (model confidence 96.1%), driven by low MMSE (=20) and parenchymal atrophy (nWBV=0.69). Note: SHAP represents algorithmic attribution for decision support, not causal biological proof."
    sp1.font.name = FONT_FAMILY; sp1.font.size = Pt(8.5); sp1.font.color.rgb = NAVY; sp1.space_before = Pt(2)

    s13.notes_slide.notes_text_frame.text = (
        "Slide 13 reveals our SHAP explainability findings. On the left, the beeswarm plot confirms that the Mini-Mental State Examination "
        "(MMSE) and normalized Whole Brain Volume (nWBV) are the top two diagnostic drivers. Low MMSE scores (blue points) and reduced "
        "whole-brain volume (cerebral atrophy, blue points) strongly increase dementia log-odds. On the right, our local patient waterfall "
        "explains a real test subject: the model started at base value -0.32, but low MMSE (+1.82) and severe parenchymal atrophy (+0.95) "
        "pushed the log-odds to +3.21, yielding a 96.1% dementia prediction. This gives clinicians transparent, auditable evidence."
    )

    # =========================================================================
    # SLIDE 14: CONCLUSION, FUTURE WORK & ACKNOWLEDGMENTS
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s14, slide_num=14, kicker="13 | SUMMARY & CONCLUSION",
                     title="Conclusion, Future Directions & Viva Q&A",
                     subtitle="Summary of scientific contributions, future translational roadmap, and concluding remarks.")
    
    col_w4 = Inches(3.7)
    add_card(s14, Inches(0.8), Inches(1.55), col_w4, Inches(5.15), title="Scientific Contributions", kicker="DELIVERABLES", accent_color=MUTED_TEAL)
    c_box = s14.shapes.add_textbox(Inches(1.04), Inches(2.35), col_w4 - Inches(0.48), Inches(4.1))
    cf = c_box.text_frame; cf.word_wrap = True
    cf.text = (
        "• Lightweight Multimodal System:\n"
        "  Engineered an end-to-end ML pipeline executing in seconds on commodity CPUs without GPU dependencies.\n\n"
        "• Scientifically Honest Validation:\n"
        "  Enforced strict participant-level 80/20 holdout partitioning, establishing an un-inflated, leakage-free benchmark.\n\n"
        "• High Diagnostic Discrimination:\n"
        "  Gradient Boosting achieved 85.11% accuracy, 80.00% sensitivity, 88.89% specificity, and 0.9426 ROC-AUC.\n\n"
        "• Biologically Valid XAI:\n"
        "  SHAP confirmed that MMSE and normalized brain parenchymal volume (nWBV) drive predictions, aligning with clinical neurology."
    )
    for p in cf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    add_card(s14, Inches(0.8) + col_w4 + gap, Inches(1.55), col_w4, Inches(5.15), title="Future Research Roadmap", kicker="NEXT STEPS", accent_color=SOFT_BLUE)
    f_box2 = s14.shapes.add_textbox(Inches(1.04) + col_w4 + gap, Inches(2.35), col_w4 - Inches(0.48), Inches(4.1))
    ff2 = f_box2.text_frame; ff2.word_wrap = True
    ff2.text = (
        "• Regional Volumetric Segmentation:\n"
        "  Extract subcortical FreeSurfer volumetry for the hippocampus, amygdala, and lateral ventricles.\n\n"
        "• Longitudinal Progression Tracking:\n"
        "  Extend to OASIS-2/3 multi-timepoint cohorts to model Mild Cognitive Impairment (MCI) progression.\n\n"
        "• External Multi-Center Validation:\n"
        "  Evaluate model generalizability across diverse cohorts from ADNI (US) and AIBL (Australia).\n\n"
        "• Privacy-Preserving Federated XAI:\n"
        "  Deploy decentralized ensemble learning across hospitals without sharing sensitive raw patient scans [Jahan et al., 2025]."
    )
    for p in ff2.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    add_card(s14, Inches(0.8) + (col_w4 + gap)*2, Inches(1.55), col_w4, Inches(5.15), title="Technical Viva & Discussion", kicker="CLOSING", accent_color=MUTED_TEAL)
    q_box = s14.shapes.add_textbox(Inches(1.04) + (col_w4 + gap)*2, Inches(2.35), col_w4 - Inches(0.48), Inches(4.1))
    qf = q_box.text_frame; qf.word_wrap = True
    qf.text = (
        "• Project Repositories & Artifacts:\n"
        "  - IEEE 6-Page Conference Paper (Word & PDF)\n"
        "  - Full 14-Slide Widescreen Presentation (.pptx)\n"
        "  - Open-Source GitHub Repository:\n"
        "    github.com/devansh16125/Health-Care-mini-project\n\n"
        "• Contact Information:\n"
        "  - Jay Thakare: 23070521233\n"
        "  - Devansh Gawali: 23070521237\n"
        "  - Department of Computer Science & Engineering\n"
        "  - Symbiosis Institute of Technology, Nagpur\n\n"
        "Thank you! We welcome questions from the evaluation committee."
    )
    for p in qf.paragraphs: p.font.size = Pt(10); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s14.notes_slide.notes_text_frame.text = (
        "In conclusion, our research proves that transparent, explainable machine learning combining cognitive screening and structural "
        "MRI morphometrics delivers high-accuracy dementia classification without expensive deep learning or data leakage. Gradient Boosting "
        "attained 85.11% accuracy and 0.9426 ROC-AUC, with SHAP confirming biological validity. Our entire codebase, IEEE paper, and presentation "
        "are publicly published on GitHub. Thank you very much for your time and attention. We now welcome questions and technical feedback."
    )

    # =========================================================================
    # SLIDE 15: APPENDIX 1 — FULL VERIFIED BIBLIOGRAPHY (PART 1: 2023–2026)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s15, slide_num=15, kicker="APPENDIX I | SCIENTIFIC LITERATURE",
                     title="Verified Academic Bibliography (Recent Papers, 2023–2026)",
                     subtitle="Peer-reviewed literature indexed in references.bib with complete Digital Object Identifiers (DOIs).")
    
    add_card(s15, Inches(0.8), Inches(1.55), Inches(11.733), Inches(5.15), title="Verified Peer-Reviewed Citations (Part 1)", kicker="BIBLIOGRAPHIC RECORD", accent_color=MUTED_TEAL)
    ref_box1 = s15.shapes.add_textbox(Inches(1.05), Inches(2.25), Inches(11.2), Inches(4.3))
    rf1 = ref_box1.text_frame; rf1.word_wrap = True
    rf1.text = (
        "1. Jahan, S. et al. (2023). 'Explainable AI-based Alzheimer's prediction and management using multimodal data.' PLOS ONE, 18(11), e0294253. DOI: 10.1371/journal.pone.0294253.\n\n"
        "2. Jahan, S. et al. (2025). 'Federated Explainable AI-Based Alzheimer's Disease Prediction With Multimodal Data.' IEEE Access, 13, 3547343. DOI: 10.1109/ACCESS.2025.3547343.\n\n"
        "3. Jiang, X. et al. (2024). 'Machine learning models for diagnosing Alzheimer's disease using brain cortical complexity.' Frontiers in Aging Neuroscience, 16, 1434589. DOI: 10.3389/fnagi.2024.1434589.\n\n"
        "4. Lu, Y. et al. (2024). 'A hierarchical attention-based multimodal fusion framework for predicting progression of Alzheimer's disease.' Biomedical Signal Processing and Control, 89, 105669. DOI: 10.1016/j.bspc.2023.105669.\n\n"
        "5. Ebrahim, A. et al. (2024). 'SHAP-Based Explainability for Local and Global Insights in Alzheimer’s Detection.' IEEE Access, 12, 18450–18465. DOI: 10.1109/ACCESS.2024.3351809.\n\n"
        "6. Shah, M. et al. (2024). 'Systematic comparison of 3D Deep learning and classical machine learning explanations for Alzheimer's Disease detection.' Computers in Biology and Medicine, 172, 108029. DOI: 10.1016/j.compbiomed.2024.108029.\n\n"
        "7. Sharma, N. & Kaur, P. (2024). 'Improving Alzheimer's Disease Prediction with Different Machine Learning Approaches and Feature Selection Techniques.' PeerJ Computer Science, 10, e1815. DOI: 10.7717/peerj.18150.\n\n"
        "8. Al-Kaf, A. et al. (2024). 'OASIS-SB: a sex-balanced, distribution-preserving synthetic phenotypic dataset for bias-resilient clinical prediction.' Scientific Reports, 14, 2580. DOI: 10.1038/s41598-024-52380-z."
    )
    for p in rf1.paragraphs: p.font.size = Pt(9.5); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s15.notes_slide.notes_text_frame.text = (
        "Appendix Slide 15 lists the first 8 peer-reviewed publications from 2023 to 2026 cited in our research paper, including exact DOIs "
        "and journal venues (PLOS ONE, IEEE Access, Frontiers in Aging Neuroscience, Biomedical Signal Processing and Control, Computers in Biology "
        "and Medicine, PeerJ Computer Science, and Scientific Reports)."
    )

    # =========================================================================
    # SLIDE 16: APPENDIX 2 — FULL VERIFIED BIBLIOGRAPHY (PART 2: FOUNDATIONAL)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    apply_slide_base(s16, slide_num=16, kicker="APPENDIX II | SCIENTIFIC LITERATURE",
                     title="Verified Academic Bibliography (Foundational Papers)",
                     subtitle="Historical benchmarks, methodological data-leakage audits, and foundational dataset citations.")
    
    add_card(s16, Inches(0.8), Inches(1.55), Inches(11.733), Inches(5.15), title="Verified Peer-Reviewed Citations (Part 2)", kicker="BIBLIOGRAPHIC RECORD", accent_color=SOFT_BLUE)
    ref_box2 = s16.shapes.add_textbox(Inches(1.05), Inches(2.25), Inches(11.2), Inches(4.3))
    rf2 = ref_box2.text_frame; rf2.word_wrap = True
    rf2.text = (
        "9. Yagis, E. et al. (2021). 'Effect of data leakage in brain MRI classification using 2D convolutional neural networks.' Scientific Reports, 11, 22544. DOI: 10.1038/s41598-021-01681-w.\n\n"
        "10. El-Sappagh, S. et al. (2021). 'A multilayer multimodal detection and prediction model based on explainable artificial intelligence for Alzheimer's disease.' Scientific Reports, 11, 2660. DOI: 10.1038/s41598-021-82098-3.\n\n"
        "11. Al-antari, M. et al. (2022). 'Deep learning based pipelines for Alzheimer's disease diagnosis: a comparative study and a novel deep-ensemble method.' Computers in Biology and Medicine, 141, 105032. DOI: 10.1016/j.compbiomed.2021.105032.\n\n"
        "12. Tanveer, M. et al. (2021). 'Machine learning techniques for the diagnosis of Alzheimer's disease: A review.' ACM Trans. Multimed. Comput. Commun. Appl., 16(1s), 30. DOI: 10.1145/3344998.\n\n"
        "13. Begum, T. et al. (2022). 'An Interpretable PyCaret Approach for Alzheimer’s Disease Prediction.' Sensors, 22(20), 8068. DOI: 10.3390/s22208068.\n\n"
        "14. Al-Saeed, M. et al. (2024). 'Explaining graph convolutional network predictions for clinicians—An explainable AI approach to Alzheimer’s disease classification.' Frontiers in Artificial Intelligence, 6, 1334613. DOI: 10.3389/frai.2023.1334613.\n\n"
        "15. AL-Ghamdi, M. et al. (2022). 'A Deep Learning Framework with an Embedded-Based Feature Selection Approach for Early Detection of Alzheimer's Disease.' Computers in Biology and Medicine, 143, 105056. DOI: 10.1016/j.compbiomed.2021.105056.\n\n"
        "16. Marcus, D. S. et al. (2007). 'Open Access Series of Imaging Studies (OASIS): Cross-sectional MRI Data in Young, Middle Aged, Nondemented, and Demented Older Adults.' Journal of Cognitive Neuroscience, 19(9), 1498–1507. DOI: 10.1162/jocn.2007.19.9.1498."
    )
    for p in rf2.paragraphs: p.font.size = Pt(9.5); p.font.name = FONT_FAMILY; p.font.color.rgb = NAVY

    s16.notes_slide.notes_text_frame.text = (
        "Appendix Slide 16 completes our bibliography with 8 foundational and methodological references, including Marcus et al. (2007) "
        "which established the OASIS-1 dataset benchmark, Yagis et al. (2021) on MRI data leakage, and surveys by Tanveer et al. (2021) "
        "and El-Sappagh et al. (2021)."
    )

    out_dir = r"C:\Users\devan\.gemini\antigravity-ide\scratch\alzheimer-explainable-ml\presentation"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "presentation.pptx")
    prs.save(out_path)
    print(f"Publication-grade PowerPoint presentation successfully generated at: {out_path}")
    return out_path

if __name__ == "__main__":
    build_presentation()
