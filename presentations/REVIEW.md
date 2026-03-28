# Presentation Review — DSO_DIIM_Beamer_Style.pptx

48 slides analysed. Findings grouped by category, ordered by impact.

---

## A. FORMATTING ISSUES

### A1. Footer says "Native PowerPoint" (all content slides)
Every content slide footer reads **"DIIM FYP · Native PowerPoint"**. The "Native PowerPoint" part is a build note — it tells the audience nothing and looks unprofessional. Change to something like **"DIIM FYP · Yeoh Yu Yong"** or **"DIIM FYP · NTU 2025"**.

### A2. Inconsistent body-text font sizes
Bullet-list font sizes vary across slides with no clear rationale:
| Slide(s) | Font size | Context |
|-----------|-----------|---------|
| S2 | 16 pt | Main three-bullet intro |
| S17–S19, S21–S23 | 13 pt | Annotations beside charts |
| S38 | 15 pt | Discussion bullets |
| S40 | 13 pt | Limitations / future work |
| S41 | (implied ~13 pt) | Conclusion regime boxes |

**Recommendation:** Standardise to **14 pt** for primary bullet lists and **12 pt** for side annotations. Currently the Introduction slides (S2) look much "bigger" than the Discussion slides (S38) even though both carry equally important content.

### A3. Narrow annotation boxes cause text overflow (S8, S16, S30)
Several slides use small label+description boxes where the text exceeds the box width:
- **S8** (From Leontief to DIIM): five equation-component boxes (A\*, K, c\*, q(0)) are only ~2 in wide. Labels like "Normalized interdependency matrix" and "Maps IO structure into inoperability space" overflow.
- **S16** (COVID setup): timeline annotation boxes ~1.85 in wide; "Recovery until DORSCON returned to Yellow assumptions" overflows.
- **S30** (Monte Carlo framework): step-description boxes only ~1.20 in wide; every step label overflows.

**Recommendation:** Either widen these boxes by 0.5–1.0 in, or reduce the descriptive text to fit. The current shrink-to-fit (normAutofit) will make text tiny on screen.

### A4. Appendix slide 46 missing title
Slides 44, 45, 47, and 48 all have a Georgia 22 pt title (TextBox 3). **Slide 46** (PCA top-5 table) has no title textbox — only a table and a footer callout about Construction. Add a title such as **"PCA-weighted top-5 breakdown"**.

### A5. Page-number gaps at section dividers
Content slides are numbered 2–5, 7–14, 16–19, 21–23, 25–28, 30–36, 38–41, 44–48. Section dividers (6, 15, 20, 24, 29, 37, 43) and the title/thank-you slides have no page number, creating visible jumps (e.g., 5 → 7). This is a minor issue — acceptable in Beamer style — but worth noting if the examiner references slide numbers.

### A6. Appendix nav bar is simplified
Content slides have the full 8-section nav bar. Appendix slides (44–48) only show a single "Appendix" pill. This is intentional but creates a visual break — consider adding the full nav bar with "Appendix" appended, or at least keeping the same horizontal stripe for consistency.

---

## B. CONTENT ISSUES

### B1. Title slide (S1) is missing key information
- **No date or academic year** (e.g., "AY 2025/26")
- **No school/department** (e.g., "School of Physical and Mathematical Sciences")
- **Supervisor's name is incomplete**: "Assoc Prof Santos" — add full name (e.g., "Assoc Prof Joost R. Santos")
- Consider adding DSO logo if this is a DSO-partnered FYP

### B2. Slide 4 subtitle is meta-commentary
The subtitle reads: *"The deck follows the report closely, but in presentation form."* This doesn't tell the audience anything about the objectives. Replace with something that actually summarises the contributions, e.g., *"Two Singapore disruption scenarios, five simplified ranking methods, and a decision framework."*

### B3. Slide 27 title is vague
"Method rankings at the original data point" — *which* original data point? Clearer: **"Method rankings at observed disruption levels"** or **"Single-point comparison: DIIM vs simplified methods"**.

### B4. Slide 40 future work is thin
Only 2 future-work bullets visible:
1. Finer sector breakdown and external validation
2. Correlated or evolving q(0) shocks

For an FYP defence, examiners often probe this section. Consider adding:
- Real-time data integration (dynamic IO updates)
- Multi-hazard scenarios (e.g., overlapping disruptions)
- Validation against actual post-COVID recovery data

### B5. Mathematical notation is plain-text throughout
`q_i`, `x_i`, `A*`, `K`, `c*`, `q(0)` appear as plain text. While this is a PowerPoint limitation, you can improve readability with:
- Unicode subscripts: q₁, xᵢ
- Unicode superscripts: A*
- Or embed small LaTeX-rendered equation images for the key formulas (especially the DIIM state equation on S8)

### B6. Slide 8 (From Leontief to DIIM) is dense but lacks the actual equation
This is the core methodology slide, but it only has five label boxes for the components (A\*, K, c\*, q(0)) without showing the DIIM equation itself:
> q(t+1) = A\* q(t) + K [q(t) − q\*] + c\*(t)

Consider adding the equation as a centrepiece, even as a simple text rendering.

### B7. No "key takeaway" callout on results slides (S31–S36)
The Monte Carlo section spans 7 slides (S30–S36) with dense charts and robustness checks. Each slide has a subtitle, but none has a bold **takeaway box** like the amber "Key question" / "Interpretation" labels used in earlier slides (S2, S19, S21). Adding a consistent amber callout to each results slide would help the audience track the narrative.

### B8. Missing transition between Manpower (S23) and Methods (S25)
The deck jumps from manpower economic loss (S23) directly to a section divider for "Simplified Ranking Methods" (S24). There's no bridge slide explaining *why* we now shift from scenario analysis to method comparison. A one-liner on the section divider subtitle would help, e.g., *"With both scenarios established, the question becomes: can we approximate DIIM rankings without running the full model?"*

### B9. Slide 38 "What the results mean" — the chart subtitle is weak
The source note reads: *"Across both scenarios, xi-weighted methods stay surprisingly close."* This is the key finding but it's in tiny 9 pt italic at the bottom. Promote it to a visible callout.

---

## C. DESIGN STRENGTHS (keep these)

- **Consistent Beamer-inspired template**: nav bar, title/subtitle/footer grid, section dividers with numbered headings — all well executed.
- **Color palette** is professional: navy (#17253A), gray (#637282), amber (#D97706), blue (#1E40AF).
- **Progressive disclosure**: dense detail in appendix, main deck stays at the right level of abstraction.
- **Section dividers** effectively break the narrative into digestible chunks.
- **Decision framework (S39)** is a strong practical payoff slide.
- **Chart annotations** beside images (S17–S19, S22–S23) are a good pattern.

---

## D. PRIORITY FIX LIST

| # | Fix | Impact | Effort |
|---|-----|--------|--------|
| 1 | Change footer to remove "Native PowerPoint" | High | 5 min |
| 2 | Add date, school, full supervisor name to title slide | High | 5 min |
| 3 | Standardise bullet font sizes (14 pt primary, 12 pt annotation) | Medium | 20 min |
| 4 | Widen narrow annotation boxes on S8, S16, S30 | Medium | 15 min |
| 5 | Fix S4 subtitle (remove meta-commentary) | Medium | 2 min |
| 6 | Add title to appendix S46 | Low | 2 min |
| 7 | Add DIIM equation to S8 | Medium | 10 min |
| 8 | Add amber takeaway callouts to S31–S36 | Medium | 15 min |
| 9 | Expand future work on S40 | Low | 5 min |
| 10 | Improve S38 chart subtitle visibility | Low | 5 min |
