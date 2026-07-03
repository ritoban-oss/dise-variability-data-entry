# DISE Variability Study — Clinical Data Entry App

CRF data entry for: *Assessment of Intra-individual and Inter-observer Variability in Drug-Induced Sleep Endoscopy Among Patients with Obstructive Sleep Apnoea* — Institute of Sleep Science, Kolkata.

---

## Files

| File | Purpose |
|---|---|
| `index.html` | Full 11-step CRF entry app |
| `config.js` | **Edit this** — paste your Apps Script URL here |
| `apps-script.gs` | Google Apps Script to deploy in your Sheet |

---

## Setup (same as OSA GA Study app)

1. Open a new Google Sheet → **Extensions → Apps Script**
2. Paste the contents of `apps-script.gs`
3. **Deploy → New Deployment** → Web App → Execute as: Me → Access: Anyone
4. Authorize and copy the `/exec` URL
5. Paste it into `config.js` → `SCRIPT_URL`
6. Push to GitHub → enable Pages → share the URL with your team

Health check: open the `/exec` URL in a browser. You should see:
```json
{ "status": "ready", "message": "DISE Variability Study collector is active." }
```

---

## CRF Structure (11 steps)

| Step | Section | Key fields |
|---|---|---|
| 1 | Screening & Eligibility | Study ID, AHI, ASA grade, referral indication, eligibility flags |
| 2 | Demographics | Height, weight, BMI (auto), neck circumference, ethnicity |
| 3 | Clinical & Questionnaire | ESS score, comorbidities, medications, smoking/alcohol |
| 4 | Diagnostic PSG | AHI, ODI, T90, positional/stage AHI distribution |
| 5 | DISE Session 1 — Procedure | Sedation agent/dose, BIS target/achieved, adverse events, positions |
| 6 | DISE Session 1 — VOTE | Per-structure × per-position degree (0/1/2) + pattern; video metadata |
| 7 | DISE Session 2 — Procedure | Same as Session 1 + interval from Session 1 (days) |
| 8 | DISE Session 2 — VOTE | Same VOTE grid as Session 1 |
| 9 | Observer Scoring | Two blinded observers — VOTE scores for both sessions each |
| 10 | Consensus & Treatment | Adjudicator VOTE, treatment recommendations, concordance flag |
| 11 | Review & Submit | Full summary before submission |

---

## VOTE Classification

Degree: **0** = no obstruction · **1** = partial · **2** = complete  
Pattern: **AP** (anteroposterior) · **Lateral** · **Concentric** · **N/A**

Structures: Velum (V), Oropharynx (O), Tongue base (T), Epiglottis (E)  
Positions: Supine + Lateral (where examined)

---

## Google Sheet columns

The script writes **~140 columns** per patient record. The header row is auto-written on the first submission and frozen. If you need to reset the column structure, clear row 1 before the next submission.

---

## Design decisions embedded in this CRF

- **VOTE entry: per structure per position** (not "worst position") — captures the full dataset for both intra-individual and inter-observer analysis.
- **Session interval** captured as a field (days between S1 and S2), not hardcoded — supports later stratified analysis by interval length.
- **Observer scoring** embedded in the same record — one row per patient covers all observers and both sessions, simplifying the kappa computation dataset.
- **Derived variables** (Cohen's kappa, ICC, reproducibility index) are computed post-hoc from the Sheet data, not entered here.
