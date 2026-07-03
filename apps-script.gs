/**
 * DISE Variability Study — Google Apps Script
 *
 * Writes to TWO tabs in the same Google Sheet:
 *   Tab 1 "Patient CRF"    — one row per patient (all fields except observer scores)
 *   Tab 2 "Observer Scores" — one row per observer per patient, linked by Study ID
 *
 * Deploy as a Web App from inside your Google Sheet:
 *   Extensions → Apps Script → paste this → Save → Deploy → New Deployment
 *   Type: Web App | Execute as: Me | Who has access: Anyone
 *
 * Copy the /exec URL into config.js → SCRIPT_URL
 */

// ── Tab names ────────────────────────────────────────────────
var CRF_TAB   = 'Patient CRF';
var OBS_TAB   = 'Observer Scores';

// ── CRF headers (Tab 1) ──────────────────────────────────────
var CRF_HEADERS = [
  "Timestamp",
  // Screening
  "Study ID","Consent Date","Age","Sex","DOB","Pregnancy Status",
  "Screen AHI","PSG Date","Symptom Status","ASA Grade","Referral Indication",
  "Prior UAW Surgery","UAW Surgery Details","Current URTI",
  "Sedative Allergy","Sedative Allergy Details",
  // Demographics
  "Height (cm)","Weight (kg)","BMI","Neck Circumference (cm)",
  "Waist-Hip Ratio","Ethnicity",
  // Clinical
  "ESS Score","Comorbidities","Comorbidity Other","Medications","Smoking","Alcohol",
  // PSG
  "PSG AHI","PSG ODI","PSG T90 (%)","PSG to DISE (days)",
  "Supine AHI","Non-supine AHI","REM AHI","NREM AHI",
  // Session 1 — Procedure
  "S1 Date","S1 Time","S1 Sedative Agent","S1 Sedative Dose",
  "S1 Target BIS","S1 Achieved BIS","S1 Time to Target (min)","S1 Duration (min)",
  "S1 Adverse Desat","S1 Adverse Airway","S1 Adverse Haemo",
  "S1 Supine","S1 Supine Duration (min)","S1 Lateral","S1 Lateral Duration (min)",
  // Session 1 VOTE
  "S1 V Sup Degree","S1 V Sup Pattern","S1 V Lat Degree","S1 V Lat Pattern",
  "S1 O Sup Degree","S1 O Sup Pattern","S1 O Lat Degree","S1 O Lat Pattern",
  "S1 T Sup Degree","S1 T Sup Pattern","S1 T Lat Degree","S1 T Lat Pattern",
  "S1 E Sup Degree","S1 E Sup Pattern","S1 E Lat Degree","S1 E Lat Pattern",
  "S1 Additional Structures","S1 Video ID","S1 Video Quality","S1 Timestamps",
  // Session 2 — Procedure
  "S2 Date","S2 Time","S2 Interval (days)","S2 Sedative Agent","S2 Sedative Dose",
  "S2 Target BIS","S2 Achieved BIS","S2 Time to Target (min)","S2 Duration (min)",
  "S2 Adverse Desat","S2 Adverse Airway","S2 Adverse Haemo",
  "S2 Supine","S2 Supine Duration (min)","S2 Lateral","S2 Lateral Duration (min)",
  // Session 2 VOTE
  "S2 V Sup Degree","S2 V Sup Pattern","S2 V Lat Degree","S2 V Lat Pattern",
  "S2 O Sup Degree","S2 O Sup Pattern","S2 O Lat Degree","S2 O Lat Pattern",
  "S2 T Sup Degree","S2 T Sup Pattern","S2 T Lat Degree","S2 T Lat Pattern",
  "S2 E Sup Degree","S2 E Sup Pattern","S2 E Lat Degree","S2 E Lat Pattern",
  "S2 Additional Structures","S2 Video ID","S2 Video Quality","S2 Timestamps",
  // Observer count (summary only — full scores in Observer Scores tab)
  "Observer Count",
  // Consensus
  "Con S1 Velum","Con S1 Oropharynx","Con S1 Tongue","Con S1 Epiglottis","Tx S1",
  "Con S2 Velum","Con S2 Oropharynx","Con S2 Tongue","Con S2 Epiglottis","Tx S2",
  "Tx Concordance","Concordance Notes"
];

// ── Observer Scores headers (Tab 2) ──────────────────────────
var OBS_HEADERS = [
  // Link back to patient
  "Study ID","Patient Submission Timestamp",
  // Observer metadata
  "Observer Number","Observer ID","Scoring Datetime",
  "Confidence Rating","Time to Score (min)",
  // Session 1 VOTE scores
  "S1 V Sup Degree","S1 V Sup Pattern","S1 V Lat Degree","S1 V Lat Pattern",
  "S1 O Sup Degree","S1 O Sup Pattern","S1 O Lat Degree","S1 O Lat Pattern",
  "S1 T Sup Degree","S1 T Sup Pattern","S1 T Lat Degree","S1 T Lat Pattern",
  "S1 E Sup Degree","S1 E Sup Pattern","S1 E Lat Degree","S1 E Lat Pattern",
  // Session 2 VOTE scores
  "S2 V Sup Degree","S2 V Sup Pattern","S2 V Lat Degree","S2 V Lat Pattern",
  "S2 O Sup Degree","S2 O Sup Pattern","S2 O Lat Degree","S2 O Lat Pattern",
  "S2 T Sup Degree","S2 T Sup Pattern","S2 T Lat Degree","S2 T Lat Pattern",
  "S2 E Sup Degree","S2 E Sup Pattern","S2 E Lat Degree","S2 E Lat Pattern"
];

// ── Helpers ───────────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Main handler ──────────────────────────────────────────────
function doPost(e) {
  var lock = LockService.getPublicLock();
  lock.waitLock(10000);

  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var d   = JSON.parse(e.postData.contents);

    // ── Tab 1: Patient CRF ──────────────────────────────────
    var crfSheet = getOrCreateSheet(ss, CRF_TAB, CRF_HEADERS);

    var observers = [];
    try { observers = JSON.parse(d.observers_json || '[]'); } catch(ex) {}

    crfSheet.appendRow([
      d.timestamp,
      d.study_id, d.consent_date, d.age, d.sex, d.dob, d.pregnancy_status,
      d.screen_ahi, d.psg_date, d.symptom_status, d.asa_grade, d.referral_indication,
      d.prior_uaw_surgery, d.uaw_surgery_details, d.current_urti,
      d.sedative_allergy, d.sed_allergy_details,
      d.height_cm, d.weight_kg, d.bmi, d.neck_circumference_cm,
      d.waist_hip_ratio, d.ethnicity,
      d.ess_score, d.comorbidities, d.comorbidity_other, d.medications, d.smoking, d.alcohol,
      d.psg_ahi, d.psg_odi, d.psg_t90, d.psg_to_dise_days,
      d.supine_ahi, d.nonsupine_ahi, d.rem_ahi, d.nrem_ahi,
      d.s1_date, d.s1_time, d.s1_sedative_agent, d.s1_sedative_dose,
      d.s1_target_bis, d.s1_achieved_bis, d.s1_time_to_target, d.s1_duration,
      d.s1_adverse_desat, d.s1_adverse_airway, d.s1_adverse_haemo,
      d.s1_supine, d.s1_supine_duration, d.s1_lateral, d.s1_lateral_duration,
      d.s1_v_sup_deg, d.s1_v_sup_pat, d.s1_v_lat_deg, d.s1_v_lat_pat,
      d.s1_o_sup_deg, d.s1_o_sup_pat, d.s1_o_lat_deg, d.s1_o_lat_pat,
      d.s1_t_sup_deg, d.s1_t_sup_pat, d.s1_t_lat_deg, d.s1_t_lat_pat,
      d.s1_e_sup_deg, d.s1_e_sup_pat, d.s1_e_lat_deg, d.s1_e_lat_pat,
      d.s1_additional, d.s1_video_id, d.s1_video_quality, d.s1_timestamps,
      d.s2_date, d.s2_time, d.s2_interval_days, d.s2_sedative_agent, d.s2_sedative_dose,
      d.s2_target_bis, d.s2_achieved_bis, d.s2_time_to_target, d.s2_duration,
      d.s2_adverse_desat, d.s2_adverse_airway, d.s2_adverse_haemo,
      d.s2_supine, d.s2_supine_duration, d.s2_lateral, d.s2_lateral_duration,
      d.s2_v_sup_deg, d.s2_v_sup_pat, d.s2_v_lat_deg, d.s2_v_lat_pat,
      d.s2_o_sup_deg, d.s2_o_sup_pat, d.s2_o_lat_deg, d.s2_o_lat_pat,
      d.s2_t_sup_deg, d.s2_t_sup_pat, d.s2_t_lat_deg, d.s2_t_lat_pat,
      d.s2_e_sup_deg, d.s2_e_sup_pat, d.s2_e_lat_deg, d.s2_e_lat_pat,
      d.s2_additional, d.s2_video_id, d.s2_video_quality, d.s2_timestamps,
      observers.length,
      d.con_s1_velum, d.con_s1_oropharynx, d.con_s1_tongue, d.con_s1_epiglottis, d.tx_s1,
      d.con_s2_velum, d.con_s2_oropharynx, d.con_s2_tongue, d.con_s2_epiglottis, d.tx_s2,
      d.tx_concordance, d.concordance_notes
    ]);

    // ── Tab 2: Observer Scores (one row per observer) ───────
    if (observers.length > 0) {
      var obsSheet = getOrCreateSheet(ss, OBS_TAB, OBS_HEADERS);

      observers.forEach(function(obs) {
        var s1 = obs.session1_vote || {};
        var s2 = obs.session2_vote || {};
        obsSheet.appendRow([
          d.study_id,
          d.timestamp,
          obs.observer_number,
          obs.id,
          obs.datetime,
          obs.confidence,
          obs.time_to_score,
          // Session 1 VOTE
          s1.v_sup_deg, s1.v_sup_pat, s1.v_lat_deg, s1.v_lat_pat,
          s1.o_sup_deg, s1.o_sup_pat, s1.o_lat_deg, s1.o_lat_pat,
          s1.t_sup_deg, s1.t_sup_pat, s1.t_lat_deg, s1.t_lat_pat,
          s1.e_sup_deg, s1.e_sup_pat, s1.e_lat_deg, s1.e_lat_pat,
          // Session 2 VOTE
          s2.v_sup_deg, s2.v_sup_pat, s2.v_lat_deg, s2.v_lat_pat,
          s2.o_sup_deg, s2.o_sup_pat, s2.o_lat_deg, s2.o_lat_pat,
          s2.t_sup_deg, s2.t_sup_pat, s2.t_lat_deg, s2.t_lat_pat,
          s2.e_sup_deg, s2.e_sup_pat, s2.e_lat_deg, s2.e_lat_pat
        ]);
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        crf_row: crfSheet.getLastRow(),
        observer_rows: observers.length
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ready',
      message: 'DISE Variability Study collector is active.',
      tabs: [CRF_TAB, OBS_TAB]
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
