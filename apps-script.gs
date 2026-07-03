/**
 * DISE Variability Study — Google Apps Script
 *
 * Three tabs:
 *   "Patient & DISE"  — baseline + single DISE session (Module 1)
 *   "Observer Scores" — one row per observer (Module 2)
 *   "Consensus"       — treatment recommendations (Module 3)
 *
 * Deploy as Web App: Execute as Me, Who has access: Anyone
 * Copy the /exec URL into config.js → SCRIPT_URL
 */

var TABS = {
  baseline: 'DISE',
  observers: 'Observer Scores',
  consensus: 'Consensus'
};

var BASELINE_HEADERS = [
  'Timestamp', 'Study ID',
  // Screening & Eligibility
  'Consent Date', 'Age', 'Sex', 'DOB', 'Pregnancy Status',
  'Screen AHI', 'PSG Date', 'Symptom Status', 'ASA Grade', 'Referral Indication',
  'Prior UAW Surgery', 'UAW Surgery Details', 'Current URTI',
  'Sedative Allergy', 'Sedative Allergy Details',
  // Demographics
  'Height (cm)', 'Weight (kg)', 'BMI', 'Neck Circumference (cm)', 'Waist-Hip Ratio', 'Ethnicity',
  // Clinical
  'ESS Score', 'Comorbidities', 'Comorbidity Other', 'Medications', 'Smoking', 'Alcohol',
  // PSG
  'PSG AHI', 'PSG ODI', 'PSG T90 (%)', 'PSG to DISE (days)',
  'Supine AHI', 'Non-supine AHI', 'REM AHI', 'NREM AHI',
  // DISE Session — Procedure
  'DISE Date', 'DISE Time', 'Sedative Agent', 'Sedative Dose',
  'Target BIS', 'Achieved BIS', 'Time to Target (min)', 'DISE Duration (min)',
  'Adverse Desat', 'Adverse Airway Obstruction', 'Adverse Haemo. Instability',
  'Supine Examined', 'Supine Duration (min)', 'Lateral Examined', 'Lateral Duration (min)',
  // DISE VOTE
  'V Sup Degree', 'V Sup Pattern', 'V Lat Degree', 'V Lat Pattern',
  'O Sup Degree', 'O Sup Pattern', 'O Lat Degree', 'O Lat Pattern',
  'T Sup Degree', 'T Sup Pattern', 'T Lat Degree', 'T Lat Pattern',
  'E Sup Degree', 'E Sup Pattern', 'E Lat Degree', 'E Lat Pattern',
  'Additional Structures', 'Video ID', 'Video Quality', 'Timestamp Markers'
];

var OBSERVER_HEADERS = [
  'Timestamp', 'Study ID', 'Observer Number', 'Observer ID', 'Scoring Datetime',
  'Confidence Rating', 'Time to Score (min)',
  'V Sup Degree', 'V Sup Pattern', 'V Lat Degree', 'V Lat Pattern',
  'O Sup Degree', 'O Sup Pattern', 'O Lat Degree', 'O Lat Pattern',
  'T Sup Degree', 'T Sup Pattern', 'T Lat Degree', 'T Lat Pattern',
  'E Sup Degree', 'E Sup Pattern', 'E Lat Degree', 'E Lat Pattern'
];

var CONSENSUS_HEADERS = [
  'Timestamp', 'Study ID',
  'Con Velum', 'Con Oropharynx', 'Con Tongue Base', 'Con Epiglottis',
  'Treatment Recommendation',
  'Concordance with DISE Findings', 'Concordance Notes'
];

// ── Helpers ───────────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

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

// Study ID is always column B (index 2) in every tab.
function idExistsInTab(ss, tabName, studyId) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  return ids.map(String).includes(String(studyId));
}

function countIdInTab(ss, tabName, studyId) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues()
    .flat().map(String).filter(function(id) { return id === String(studyId); }).length;
}

// ── GET: validation ───────────────────────────────────────────

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p  = e.parameter;

  if (p.action === 'exists') {
    return jsonResponse({ exists: idExistsInTab(ss, TABS.baseline, p.study_id) });
  }
  if (p.action === 'duplicate') {
    var tab = TABS[p.tab];
    if (!tab) return jsonResponse({ error: 'Unknown tab: ' + p.tab });
    if (p.tab === 'observers') {
      return jsonResponse({ count: countIdInTab(ss, TABS.observers, p.study_id) });
    }
    return jsonResponse({ duplicate: idExistsInTab(ss, tab, p.study_id) });
  }
  if (p.action === 'collision') {
    return jsonResponse({ collision: idExistsInTab(ss, TABS.baseline, p.study_id) });
  }
  return jsonResponse({ status: 'ready', message: 'DISE Variability Study collector is active.', tabs: Object.values(TABS) });
}

// ── POST: submission ──────────────────────────────────────────

function doPost(e) {
  var lock = LockService.getPublicLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var d  = JSON.parse(e.postData.contents);
    switch (d._action) {
      case 'baseline_dise': return submitBaseline(ss, d);
      case 'observers':     return submitObservers(ss, d);
      case 'consensus':     return submitConsensus(ss, d);
      default: return jsonResponse({ status: 'error', message: 'Unknown action: ' + d._action });
    }
  } catch(err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function submitBaseline(ss, d) {
  if (idExistsInTab(ss, TABS.baseline, d.study_id)) {
    return jsonResponse({ status: 'duplicate', message: 'Study ID already exists.' });
  }
  var sheet = getOrCreateSheet(ss, TABS.baseline, BASELINE_HEADERS);
  sheet.appendRow([
    d.timestamp, d.study_id,
    d.consent_date, d.age, d.sex, d.dob, d.pregnancy_status,
    d.screen_ahi, d.psg_date, d.symptom_status, d.asa_grade, d.referral_indication,
    d.prior_uaw_surgery, d.uaw_surgery_details, d.current_urti,
    d.sedative_allergy, d.sed_allergy_details,
    d.height_cm, d.weight_kg, d.bmi, d.neck_circumference_cm, d.waist_hip_ratio, d.ethnicity,
    d.ess_score, d.comorbidities, d.comorbidity_other, d.medications, d.smoking, d.alcohol,
    d.psg_ahi, d.psg_odi, d.psg_t90, d.psg_to_dise_days,
    d.supine_ahi, d.nonsupine_ahi, d.rem_ahi, d.nrem_ahi,
    d.dise_date, d.dise_time, d.dise_sedative_agent, d.dise_sedative_dose,
    d.dise_target_bis, d.dise_achieved_bis, d.dise_time_to_target, d.dise_duration,
    d.dise_adverse_desat, d.dise_adverse_airway, d.dise_adverse_haemo,
    d.dise_supine, d.dise_supine_duration, d.dise_lateral, d.dise_lateral_duration,
    d.v_sup_deg, d.v_sup_pat, d.v_lat_deg, d.v_lat_pat,
    d.o_sup_deg, d.o_sup_pat, d.o_lat_deg, d.o_lat_pat,
    d.t_sup_deg, d.t_sup_pat, d.t_lat_deg, d.t_lat_pat,
    d.e_sup_deg, d.e_sup_pat, d.e_lat_deg, d.e_lat_pat,
    d.additional_structures, d.video_id, d.video_quality, d.timestamp_markers
  ]);
  return jsonResponse({ status: 'success', study_id: d.study_id });
}

function submitObservers(ss, d) {
  var sheet = getOrCreateSheet(ss, TABS.observers, OBSERVER_HEADERS);
  var observers = JSON.parse(d.observers_json || '[]');
  observers.forEach(function(obs) {
    var v = obs.vote || {};
    sheet.appendRow([
      d.timestamp, d.study_id,
      obs.observer_number, obs.id, obs.datetime, obs.confidence, obs.time_to_score,
      v.v_sup_deg, v.v_sup_pat, v.v_lat_deg, v.v_lat_pat,
      v.o_sup_deg, v.o_sup_pat, v.o_lat_deg, v.o_lat_pat,
      v.t_sup_deg, v.t_sup_pat, v.t_lat_deg, v.t_lat_pat,
      v.e_sup_deg, v.e_sup_pat, v.e_lat_deg, v.e_lat_pat
    ]);
  });
  return jsonResponse({ status: 'success', study_id: d.study_id, rows: observers.length });
}

function submitConsensus(ss, d) {
  if (idExistsInTab(ss, TABS.consensus, d.study_id)) {
    return jsonResponse({ status: 'duplicate', message: 'Consensus data already submitted for this Study ID.' });
  }
  var sheet = getOrCreateSheet(ss, TABS.consensus, CONSENSUS_HEADERS);
  sheet.appendRow([
    d.timestamp, d.study_id,
    d.con_velum, d.con_oropharynx, d.con_tongue, d.con_epiglottis,
    d.treatment_recommendation,
    d.concordance, d.concordance_notes
  ]);
  return jsonResponse({ status: 'success', study_id: d.study_id });
}
