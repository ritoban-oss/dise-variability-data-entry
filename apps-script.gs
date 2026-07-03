/**
 * DISE Variability Study — Google Apps Script
 *
 * Writes to FIVE tabs. GET endpoint validates Study IDs.
 * Deploy as Web App: Execute as Me, Who has access: Anyone
 */

var TABS = {
  baseline:  'Patient Baseline',
  session1:  'Session 1',
  session2:  'Session 2',
  observers: 'Observer Scores',
  consensus: 'Consensus'
};

// ── Header definitions ────────────────────────────────────────

var BASELINE_HEADERS = [
  "Timestamp","Study ID",
  "Consent Date","Age","Sex","DOB","Pregnancy Status",
  "Screen AHI","PSG Date","Symptom Status","ASA Grade","Referral Indication",
  "Prior UAW Surgery","UAW Surgery Details","Current URTI",
  "Sedative Allergy","Sedative Allergy Details",
  "Height (cm)","Weight (kg)","BMI","Neck Circumference (cm)","Waist-Hip Ratio","Ethnicity",
  "ESS Score","Comorbidities","Comorbidity Other","Medications","Smoking","Alcohol",
  "PSG AHI","PSG ODI","PSG T90 (%)","PSG to DISE (days)",
  "Supine AHI","Non-supine AHI","REM AHI","NREM AHI"
];

var SESSION_HEADERS = function(n) { return [
  "Timestamp","Study ID",
  "S"+n+" Date","S"+n+" Time","S"+n+" Sedative Agent","S"+n+" Sedative Dose",
  "S"+n+" Target BIS","S"+n+" Achieved BIS","S"+n+" Time to Target (min)","S"+n+" Duration (min)",
  "S"+n+" Adverse Desat","S"+n+" Adverse Airway","S"+n+" Adverse Haemo",
  "S"+n+" Supine","S"+n+" Supine Duration (min)","S"+n+" Lateral","S"+n+" Lateral Duration (min)",
  "S"+n+" V Sup Degree","S"+n+" V Sup Pattern","S"+n+" V Lat Degree","S"+n+" V Lat Pattern",
  "S"+n+" O Sup Degree","S"+n+" O Sup Pattern","S"+n+" O Lat Degree","S"+n+" O Lat Pattern",
  "S"+n+" T Sup Degree","S"+n+" T Sup Pattern","S"+n+" T Lat Degree","S"+n+" T Lat Pattern",
  "S"+n+" E Sup Degree","S"+n+" E Sup Pattern","S"+n+" E Lat Degree","S"+n+" E Lat Pattern",
  "S"+n+" Additional Structures","S"+n+" Video ID","S"+n+" Video Quality","S"+n+" Timestamps"
];};

var OBSERVER_HEADERS = [
  "Timestamp","Study ID","Observer Number","Observer ID","Scoring Datetime",
  "Confidence Rating","Time to Score (min)",
  "S1 V Sup Degree","S1 V Sup Pattern","S1 V Lat Degree","S1 V Lat Pattern",
  "S1 O Sup Degree","S1 O Sup Pattern","S1 O Lat Degree","S1 O Lat Pattern",
  "S1 T Sup Degree","S1 T Sup Pattern","S1 T Lat Degree","S1 T Lat Pattern",
  "S1 E Sup Degree","S1 E Sup Pattern","S1 E Lat Degree","S1 E Lat Pattern",
  "S2 V Sup Degree","S2 V Sup Pattern","S2 V Lat Degree","S2 V Lat Pattern",
  "S2 O Sup Degree","S2 O Sup Pattern","S2 O Lat Degree","S2 O Lat Pattern",
  "S2 T Sup Degree","S2 T Sup Pattern","S2 T Lat Degree","S2 T Lat Pattern",
  "S2 E Sup Degree","S2 E Sup Pattern","S2 E Lat Degree","S2 E Lat Pattern"
];

var CONSENSUS_HEADERS = [
  "Timestamp","Study ID",
  "Con S1 Velum","Con S1 Oropharynx","Con S1 Tongue","Con S1 Epiglottis","Tx S1",
  "Con S2 Velum","Con S2 Oropharynx","Con S2 Tongue","Con S2 Epiglottis","Tx S2",
  "Tx Concordance","Concordance Notes"
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
  var ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  return ids.map(String).filter(function(id) { return id === String(studyId); }).length;
}

// ── GET: validation endpoint ──────────────────────────────────

function doGet(e) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var p   = e.parameter;

  if (p.action === 'exists') {
    // Does this Study ID exist in Patient Baseline?
    var exists = idExistsInTab(ss, TABS.baseline, p.study_id);
    return jsonResponse({ exists: exists, study_id: p.study_id });
  }

  if (p.action === 'duplicate') {
    // Is this Study ID already in the specified module tab?
    var tabName = TABS[p.tab];
    if (!tabName) return jsonResponse({ error: 'Unknown tab: ' + p.tab });
    if (p.tab === 'observers') {
      // Observers: soft check — return count, not bool
      var count = countIdInTab(ss, TABS.observers, p.study_id);
      return jsonResponse({ count: count, study_id: p.study_id });
    }
    var dup = idExistsInTab(ss, tabName, p.study_id);
    return jsonResponse({ duplicate: dup, study_id: p.study_id });
  }

  if (p.action === 'collision') {
    // Module 1 only: does auto-generated ID already exist?
    var col = idExistsInTab(ss, TABS.baseline, p.study_id);
    return jsonResponse({ collision: col, study_id: p.study_id });
  }

  return jsonResponse({
    status: 'ready',
    message: 'DISE Variability Study collector is active.',
    tabs: Object.values(TABS)
  });
}

// ── POST: submission handlers ─────────────────────────────────

function doPost(e) {
  var lock = LockService.getPublicLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var d  = JSON.parse(e.postData.contents);
    switch (d._action) {
      case 'baseline':  return submitBaseline(ss, d);
      case 'session1':  return submitSession(ss, d, 1);
      case 'session2':  return submitSession(ss, d, 2);
      case 'observers': return submitObservers(ss, d);
      case 'consensus': return submitConsensus(ss, d);
      default: return jsonResponse({ status: 'error', message: 'Unknown action: ' + d._action });
    }
  } catch(err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function submitBaseline(ss, d) {
  // Server-side duplicate guard
  if (idExistsInTab(ss, TABS.baseline, d.study_id)) {
    return jsonResponse({ status: 'duplicate', message: 'Study ID already exists in Patient Baseline.' });
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
    d.supine_ahi, d.nonsupine_ahi, d.rem_ahi, d.nrem_ahi
  ]);
  return jsonResponse({ status: 'success', study_id: d.study_id });
}

function submitSession(ss, d, n) {
  var tabKey = 'session' + n;
  var tabName = TABS[tabKey];
  if (idExistsInTab(ss, tabName, d.study_id)) {
    return jsonResponse({ status: 'duplicate', message: 'Session ' + n + ' data already exists for this Study ID.' });
  }
  var headers = SESSION_HEADERS(n);
  var sheet = getOrCreateSheet(ss, tabName, headers);
  var p = 's' + n + '_';
  sheet.appendRow([
    d.timestamp, d.study_id,
    d[p+'date'], d[p+'time'], d[p+'sedative_agent'], d[p+'sedative_dose'],
    d[p+'target_bis'], d[p+'achieved_bis'], d[p+'time_to_target'], d[p+'duration'],
    d[p+'adverse_desat'], d[p+'adverse_airway'], d[p+'adverse_haemo'],
    d[p+'supine'], d[p+'supine_duration'], d[p+'lateral'], d[p+'lateral_duration'],
    d[p+'v_sup_deg'], d[p+'v_sup_pat'], d[p+'v_lat_deg'], d[p+'v_lat_pat'],
    d[p+'o_sup_deg'], d[p+'o_sup_pat'], d[p+'o_lat_deg'], d[p+'o_lat_pat'],
    d[p+'t_sup_deg'], d[p+'t_sup_pat'], d[p+'t_lat_deg'], d[p+'t_lat_pat'],
    d[p+'e_sup_deg'], d[p+'e_sup_pat'], d[p+'e_lat_deg'], d[p+'e_lat_pat'],
    d[p+'additional'], d[p+'video_id'], d[p+'video_quality'], d[p+'timestamps']
  ]);
  return jsonResponse({ status: 'success', study_id: d.study_id });
}

function submitObservers(ss, d) {
  var sheet = getOrCreateSheet(ss, TABS.observers, OBSERVER_HEADERS);
  var observers = JSON.parse(d.observers_json || '[]');
  observers.forEach(function(obs) {
    var s1 = obs.session1_vote || {}, s2 = obs.session2_vote || {};
    sheet.appendRow([
      d.timestamp, d.study_id,
      obs.observer_number, obs.id, obs.datetime, obs.confidence, obs.time_to_score,
      s1.v_sup_deg, s1.v_sup_pat, s1.v_lat_deg, s1.v_lat_pat,
      s1.o_sup_deg, s1.o_sup_pat, s1.o_lat_deg, s1.o_lat_pat,
      s1.t_sup_deg, s1.t_sup_pat, s1.t_lat_deg, s1.t_lat_pat,
      s1.e_sup_deg, s1.e_sup_pat, s1.e_lat_deg, s1.e_lat_pat,
      s2.v_sup_deg, s2.v_sup_pat, s2.v_lat_deg, s2.v_lat_pat,
      s2.o_sup_deg, s2.o_sup_pat, s2.o_lat_deg, s2.o_lat_pat,
      s2.t_sup_deg, s2.t_sup_pat, s2.t_lat_deg, s2.t_lat_pat,
      s2.e_sup_deg, s2.e_sup_pat, s2.e_lat_deg, s2.e_lat_pat
    ]);
  });
  return jsonResponse({ status: 'success', study_id: d.study_id, rows: observers.length });
}

function submitConsensus(ss, d) {
  if (idExistsInTab(ss, TABS.consensus, d.study_id)) {
    return jsonResponse({ status: 'duplicate', message: 'Consensus data already exists for this Study ID.' });
  }
  var sheet = getOrCreateSheet(ss, TABS.consensus, CONSENSUS_HEADERS);
  sheet.appendRow([
    d.timestamp, d.study_id,
    d.con_s1_velum, d.con_s1_oropharynx, d.con_s1_tongue, d.con_s1_epiglottis, d.tx_s1,
    d.con_s2_velum, d.con_s2_oropharynx, d.con_s2_tongue, d.con_s2_epiglottis, d.tx_s2,
    d.tx_concordance, d.concordance_notes
  ]);
  return jsonResponse({ status: 'success', study_id: d.study_id });
}
