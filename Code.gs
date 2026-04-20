/**
 * Endangered Species Rescue Mission — Version 2.0
 * Google Apps Script server code
 */

var APP = {
  scriptProps: {
    spreadsheetIdKey: 'ENDANGERED_SPECIES_SPREADSHEET_ID'
  },
  projectTitle: 'Endangered Species Rescue Mission',
  outputFolderId: '1BESzdmXkmswTBApTqgiAeBZxLHgQo2mY',
  sheetNames: {
    species: 'SPECIES_MASTER',
    submissions: 'STUDENT_SUBMISSIONS',
    settings: 'SETTINGS'
  },
  stages: {
    login: 'login',
    species: 'species',
    briefing: 'briefing',
    ecology: 'ecology',
    threats: 'threats',
    actions: 'actions',
    final: 'final',
    submitted: 'submitted'
  },
  posterImageBox: {
    x: 14,
    y: 82,
    width: 176,
    height: 130
  },
  dropdowns: {
    habitatTypes: [
      'Coral reef',
      'Open ocean',
      'Coastal waters',
      'Rainforest',
      'Temperate forest',
      'Grassland',
      'Prairie',
      'Desert',
      'Wetland',
      'Mountains'
    ],
    dietTypes: [
      'Herbivore',
      'Carnivore',
      'Omnivore',
      'Filter feeder',
      'Insectivore',
      'Nectar feeder',
      'Seed eater'
    ],
    adaptationTypes: [
      'Camouflage',
      'Speed',
      'Shell / armor',
      'Long limbs',
      'Strong teeth / beak',
      'Special roots / leaves',
      'Migration',
      'Thick fur / skin',
      'Water storage'
    ],
    ecosystemRoles: [
      'Predator',
      'Prey',
      'Grazer',
      'Seed disperser',
      'Pollinator',
      'Scavenger',
      'Habitat helper',
      'Top consumer'
    ],
    threats: [
      'Habitat loss',
      'Climate change',
      'Pollution',
      'Poaching / overhunting',
      'Overfishing',
      'Invasive species',
      'Disease',
      'Human development',
      'Wildfire / natural disasters'
    ],
    actions: [
      'Protected areas / habitat restoration',
      'Breeding programs',
      'Anti-poaching / stronger laws',
      'Pollution reduction',
      'Research and monitoring',
      'Education / public awareness',
      'Captive breeding and reintroduction',
      'Climate action',
      'Community conservation programs'
    ],
    conservationStatuses: [
      'Critically Endangered',
      'Endangered',
      'Threatened',
      'Vulnerable',
      'Not Sure Yet'
    ]
  }
};

/* ===== Test function — run from GAS editor to verify poster generation ===== */

function testPosterGeneration() {
  var fakeSubmission = {
    first_name: 'Test', last_name: 'Student', hour: '1',
    species_id: 'hawksbill-turtle',
    common_name: 'Hawksbill Turtle', scientific_name: 'Eretmochelys imbricata',
    biome: 'Marine', identified_status: 'Critically Endangered',
    identified_scientific_name: 'Eretmochelys imbricata',
    habitat_type: 'Coral reef', diet_type: 'Carnivore',
    adaptation_type: 'Camouflage', ecosystem_role: 'Predator',
    ecology_explanation: 'Lives in coral reef environments and feeds on sponges, which helps keep the reef balanced.',
    threat_1: 'Habitat loss', threat_1_reason: 'Coral bleaching from climate change destroys their nesting and feeding areas.',
    threat_2: 'Poaching / overhunting', threat_2_reason: 'Hunted for their shells and eggs despite international protection.',
    action_1: 'Protected areas / habitat restoration', action_2: 'Anti-poaching / stronger laws',
    action_explanation: 'Protected nesting beaches and stronger anti-poaching enforcement give populations time to recover.',
    why_it_matters: 'Hawksbill turtles keep coral reefs healthy by controlling sponge populations, which allows corals to thrive.',
    selected_image_url: ''
  };
  try {
    var result = generatePosterForSubmission_(fakeSubmission);
    Logger.log('SUCCESS: ' + JSON.stringify(result));
    return 'SUCCESS — check Logger and your Drive folder.';
  } catch (e) {
    Logger.log('FAILED: ' + e + '\n' + e.stack);
    return 'FAILED: ' + e;
  }
}

/* ===== Web app entry points ===== */

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.appTitle = APP.projectTitle;
  return template
    .evaluate()
    .setTitle(APP.projectTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ===== Setup ===== */

function setupProjectSheets() {
  var ss = getOrCreateSpreadsheet_();

  var speciesHeaders = [
    'species_id', 'common_name', 'scientific_name', 'biome', 'status', 'region', 'briefing_text',
    'habitat_hint', 'diet_hint', 'ecosystem_role_hint',
    'image_option_1_url', 'image_option_2_url', 'image_option_3_url',
    'research_link_1', 'research_link_2'
  ];

  var submissionHeaders = [
    'timestamp_start', 'timestamp_submit', 'first_name', 'last_name', 'hour', 'student_key',
    'species_id', 'common_name', 'scientific_name', 'biome', 'mission_stage',
    'identified_common_name', 'identified_scientific_name', 'identified_status',
    'habitat_type', 'diet_type', 'adaptation_type', 'ecosystem_role', 'ecology_explanation',
    'threat_1', 'threat_1_reason', 'threat_2', 'threat_2_reason', 'threat_3', 'threat_3_reason',
    'action_1', 'action_2', 'action_3', 'action_explanation',
    'why_it_matters', 'selected_image_url',
    'score_out_of_100', 'submit_type', 'submission_message',
    'poster_file_id', 'poster_slide_url', 'pdf_file_id', 'poster_pdf_url',
    'tile_pdf_file_id', 'tile_pdf_url',
    'submission_status'
  ];

  var settingsHeaders = ['key', 'value'];
  var defaultSettings = [
    ['poster_template_file_id', 'PASTE_TEMPLATE_FILE_ID_HERE'],
    ['output_folder_id', APP.outputFolderId],
    ['teacher_email', 'PASTE_TEACHER_EMAIL_HERE'],
    ['project_title', APP.projectTitle],
    ['share_output_with_link', 'TRUE'],
    ['show_student_output_links', 'TRUE'],
    ['use_template_poster', 'FALSE']
  ];

  ensureSheetHeaders_(ss, APP.sheetNames.species, speciesHeaders);
  ensureSheetHeaders_(ss, APP.sheetNames.submissions, submissionHeaders);
  ensureSettingsSheet_(ss, APP.sheetNames.settings, settingsHeaders, defaultSettings);

  formatHeaderRow_(ss.getSheetByName(APP.sheetNames.species));
  formatHeaderRow_(ss.getSheetByName(APP.sheetNames.submissions));
  formatHeaderRow_(ss.getSheetByName(APP.sheetNames.settings));

  SpreadsheetApp.flush();
  return 'Project sheets are ready in: ' + ss.getUrl();
}

function resetProjectSheetsDangerously() {
  var ss = getOrCreateSpreadsheet_();
  var names = [APP.sheetNames.species, APP.sheetNames.submissions, APP.sheetNames.settings];
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) ss.deleteSheet(sheet);
  }
  return setupProjectSheets();
}

/* ===== Client-facing API functions ===== */

function getInitialConfig() {
  try {
    getSpreadsheet_();
  } catch (e) {
    setupProjectSheets();
  }
  seedDefaultSpeciesMasterIfNeeded_();

  var settings = getSettingsMap_();
  var species = getSpeciesData_();

  return {
    appTitle: settings.project_title || APP.projectTitle,
    dropdowns: APP.dropdowns,
    species: species,
    stages: APP.stages
  };
}

function startMission(student) {
  validateRequired_(student, ['firstName', 'lastName']);

  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.submissions);
  var headers = getHeaders_(sheet);
  var studentKey = Utilities.getUuid();

  var rowObject = {
    timestamp_start: new Date(),
    timestamp_submit: '',
    first_name: trim_(student.firstName),
    last_name: trim_(student.lastName),
    hour: '',
    student_key: studentKey,
    species_id: '',
    common_name: '',
    scientific_name: '',
    biome: '',
    mission_stage: APP.stages.species,
    identified_common_name: '',
    identified_scientific_name: '',
    identified_status: '',
    habitat_type: '',
    diet_type: '',
    adaptation_type: '',
    ecosystem_role: '',
    ecology_explanation: '',
    threat_1: '',
    threat_1_reason: '',
    threat_2: '',
    threat_2_reason: '',
    threat_3: '',
    threat_3_reason: '',
    action_1: '',
    action_2: '',
    action_3: '',
    action_explanation: '',
    why_it_matters: '',
    selected_image_url: '',
    score_out_of_100: 0,
    submit_type: '',
    submission_message: '',
    poster_file_id: '',
    poster_slide_url: '',
    pdf_file_id: '',
    poster_pdf_url: '',
    tile_pdf_file_id: '',
    tile_pdf_url: '',
    submission_status: 'Started'
  };

  appendObjectRow_(sheet, headers, rowObject);
  return { studentKey: studentKey, currentStage: APP.stages.species };
}

function saveSpeciesChoice(studentKey, speciesId) {
  validateString_(studentKey, 'studentKey');
  validateString_(speciesId, 'speciesId');

  var species = findSpeciesById_(speciesId);
  if (!species) throw new Error('Species not found: ' + speciesId);

  updateSubmissionByStudentKey_(studentKey, {
    species_id: species.species_id,
    common_name: species.common_name,
    scientific_name: species.scientific_name,
    biome: species.biome,
    mission_stage: APP.stages.briefing,
    submission_status: 'In Progress',
    score_out_of_100: calculateSubmissionScoreByStudentKey_(studentKey, {
      species_id: species.species_id,
      common_name: species.common_name,
      scientific_name: species.scientific_name,
      biome: species.biome
    })
  });

  return { currentStage: APP.stages.briefing, species: species };
}

function saveEcology(studentKey, payload) {
  validateString_(studentKey, 'studentKey');
  validateRequired_(payload, ['commonNameAnswer', 'scientificNameAnswer', 'statusAnswer', 'habitatType', 'dietType', 'adaptationType', 'ecosystemRole', 'ecologyExplanation']);

  if (trim_(payload.ecologyExplanation).length < 8) {
    throw new Error('Please write a little more for the ecology explanation.');
  }

  var updates = {
    identified_common_name: trim_(payload.commonNameAnswer),
    identified_scientific_name: trim_(payload.scientificNameAnswer),
    identified_status: trim_(payload.statusAnswer),
    habitat_type: payload.habitatType,
    diet_type: payload.dietType,
    adaptation_type: payload.adaptationType,
    ecosystem_role: payload.ecosystemRole,
    ecology_explanation: trim_(payload.ecologyExplanation),
    mission_stage: APP.stages.threats
  };

  updates.score_out_of_100 = calculateSubmissionScoreByStudentKey_(studentKey, updates);
  updateSubmissionByStudentKey_(studentKey, updates);
  return { currentStage: APP.stages.threats };
}

function saveThreats(studentKey, payload) {
  validateString_(studentKey, 'studentKey');
  validateRequired_(payload, ['threat1', 'threat1Reason', 'threat2', 'threat2Reason']);

  if (trim_(payload.threat1).toLowerCase() === trim_(payload.threat2).toLowerCase()) {
    throw new Error('Please choose 2 different threats.');
  }

  var reasonKeys = ['threat1Reason', 'threat2Reason'];
  for (var r = 0; r < reasonKeys.length; r++) {
    if (trim_(payload[reasonKeys[r]]).length < 5) {
      throw new Error('Each threat needs a short explanation.');
    }
  }

  var updates = {
    threat_1: payload.threat1,
    threat_1_reason: trim_(payload.threat1Reason),
    threat_2: payload.threat2,
    threat_2_reason: trim_(payload.threat2Reason),
    mission_stage: APP.stages.actions
  };
  updates.score_out_of_100 = calculateSubmissionScoreByStudentKey_(studentKey, updates);
  updateSubmissionByStudentKey_(studentKey, updates);
  return { currentStage: APP.stages.actions };
}

function saveActions(studentKey, payload) {
  validateString_(studentKey, 'studentKey');
  validateRequired_(payload, ['action1', 'action2', 'actionExplanation']);

  if (trim_(payload.action1).toLowerCase() === trim_(payload.action2).toLowerCase()) {
    throw new Error('Please choose 2 different conservation actions.');
  }

  if (trim_(payload.actionExplanation).length < 8) {
    throw new Error('Please explain how the actions help.');
  }

  var updates = {
    action_1: trim_(payload.action1),
    action_2: trim_(payload.action2),
    action_explanation: trim_(payload.actionExplanation),
    mission_stage: APP.stages.final
  };
  updates.score_out_of_100 = calculateSubmissionScoreByStudentKey_(studentKey, updates);
  updateSubmissionByStudentKey_(studentKey, updates);
  return { currentStage: APP.stages.final };
}

function finishMission(studentKey, payload) {
  validateString_(studentKey, 'studentKey');
  validateRequired_(payload, ['whyItMatters']);

  if (trim_(payload.whyItMatters).length < 8) {
    throw new Error('Please write a short explanation for why the species matters.');
  }

  var prePosterUpdates = {
    why_it_matters: trim_(payload.whyItMatters),
    selected_image_url: trim_(payload.selectedImageUrl || ''),
    mission_stage: APP.stages.submitted,
    submission_status: 'Generating Poster',
    submit_type: 'Final'
  };
  prePosterUpdates.score_out_of_100 = calculateSubmissionScoreByStudentKey_(studentKey, prePosterUpdates);
  updateSubmissionByStudentKey_(studentKey, prePosterUpdates);

  var submission = getSubmissionByStudentKey_(studentKey);
  var output = { slideUrl: '', pdfUrl: '', tilePdfUrl: '', posterFileId: '', pdfFileId: '', tilePdfFileId: '', warning: '', studentCanViewFiles: false };

  try {
    output = generatePosterForSubmission_(submission);
  } catch (error) {
    output.warning = 'Your work was submitted, but the poster file could not be created automatically. Tell your teacher.';
    Logger.log('Poster generation failed for ' + studentKey + ': ' + error + '\n' + error.stack);
  }

  var finalUpdates = {
    timestamp_submit: new Date(),
    poster_file_id: output.posterFileId || '',
    poster_slide_url: output.slideUrl || '',
    pdf_file_id: output.pdfFileId || '',
    poster_pdf_url: output.pdfUrl || '',
    tile_pdf_file_id: output.tilePdfFileId || '',
    tile_pdf_url: output.tilePdfUrl || '',
    mission_stage: APP.stages.submitted,
    submission_status: output.warning ? 'Submitted - Poster Error' : 'Submitted',
    submission_message: output.warning || 'Submitted successfully.',
    submit_type: 'Final',
    score_out_of_100: calculateSubmissionScoreByStudentKey_(studentKey, {
      why_it_matters: trim_(payload.whyItMatters),
      selected_image_url: trim_(payload.selectedImageUrl || '')
    })
  };
  updateSubmissionByStudentKey_(studentKey, finalUpdates);

  return {
    currentStage: APP.stages.submitted,
    submissionComplete: true,
    slideUrl: output.slideUrl || '',
    pdfUrl: output.pdfUrl || '',
    tilePdfUrl: output.tilePdfUrl || '',
    studentCanViewFiles: output.studentCanViewFiles || false,
    warning: output.warning || ''
  };
}

function emergencySubmitMission(studentKey) {
  validateString_(studentKey, 'studentKey');

  var submission = getSubmissionByStudentKey_(studentKey);
  var score = calculateCompletionScore_(submission);

  updateSubmissionByStudentKey_(studentKey, {
    timestamp_submit: new Date(),
    score_out_of_100: score,
    mission_stage: APP.stages.submitted,
    submission_status: 'Emergency Submitted',
    submission_message: 'Emergency submit used before mission was fully complete.',
    submit_type: 'Emergency'
  });

  return {
    currentStage: APP.stages.submitted,
    submissionComplete: true,
    scoreOutOf100: score,
    warning: 'Emergency submit saved your current work. Tell your teacher.',
    studentCanViewFiles: false,
    slideUrl: '',
    pdfUrl: '',
    tilePdfUrl: ''
  };
}

function getSubmissionDebug(studentKey) {
  return getSubmissionByStudentKey_(studentKey);
}

/* ===== Poster drawing primitives ===== */

function hexToRgb_(hex) {
  var r = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return r ? { r: parseInt(r[1], 16), g: parseInt(r[2], 16), b: parseInt(r[3], 16) } : { r: 255, g: 255, b: 255 };
}

function pRect_(slide, x, y, w, h, fillHex, borderHex, borderPt, rounded) {
  var shape = slide.insertShape(
    rounded ? SlidesApp.ShapeType.ROUND_RECTANGLE : SlidesApp.ShapeType.RECTANGLE,
    x, y, w, h
  );
  var fc = hexToRgb_(fillHex);
  shape.getFill().setSolidFill(fc.r, fc.g, fc.b);
  var bc = borderHex ? hexToRgb_(borderHex) : fc;
  shape.getBorder().getLineFill().setSolidFill(bc.r, bc.g, bc.b);
  shape.getBorder().setWeight(borderHex ? (borderPt || 1) : 1);
  return shape;
}

function pText_(slide, text, x, y, w, h, opts) {
  var tb = slide.insertTextBox(text || '', x, y, w, h);
  var ts = tb.getText().getTextStyle();
  ts.setFontFamily(opts.font || 'Arial');
  ts.setFontSize(opts.size || 9);
  if (opts.bold) ts.setBold(true);
  if (opts.italic) ts.setItalic(true);
  if (opts.color) ts.setForegroundColor(opts.color);
  if (opts.align) {
    var paras = tb.getText().getParagraphs();
    for (var i = 0; i < paras.length; i++) {
      paras[i].getParagraphStyle().setParagraphAlignment(opts.align);
    }
  }
  return tb;
}

function pLine_(slide, x, y, w, colorHex) {
  return pRect_(slide, x, y, w, 1, colorHex, null, 0, false);
}

/* ===== Product A poster section drawers ===== */

function posterA_drawHeader_(slide, submission, studentName) {
  pRect_(slide, 0, 4, 720, 64, '#153324', null, 0, false);
  pRect_(slide, 0, 4, 6, 64, '#2e7d5b', null, 0, false);

  pText_(slide, 'ENDANGERED SPECIES RESCUE REPORT', 22, 16, 520, 18,
    { font: 'Georgia', size: 11, bold: true, color: '#c8e6c9' });
  pText_(slide, submission.common_name || 'Species Rescue', 22, 31, 500, 28,
    { font: 'Georgia', size: 20, bold: true, color: '#ffffff' });
  pText_(slide, submission.scientific_name || '', 22, 55, 400, 16,
    { font: 'Georgia', size: 10, italic: true, color: '#81c784' });

  var attribution = 'Student: ' + studentName;
  if (trim_(submission.hour)) attribution += '    Period ' + trim_(submission.hour);
  pText_(slide, attribution, 490, 24, 220, 15,
    { font: 'Arial', size: 9, color: '#a5d6a7', align: SlidesApp.ParagraphAlignment.END });

  var status = trim_(submission.identified_status) || 'Endangered';
  pRect_(slide, 620, 48, 90, 18, '#c62828', null, 0, true);
  pText_(slide, status.toUpperCase(), 620, 48, 90, 18,
    { font: 'Arial', size: 8, bold: true, color: '#ffcdd2', align: SlidesApp.ParagraphAlignment.CENTER });
}

function posterA_drawPhotoCol_(slide, submission) {
  pRect_(slide, 8, 76, 188, 192, '#1a3a25', '#2e7d5b', 1, true);
  pRect_(slide, 14, 82, 176, 130, '#0d2b18', null, 0, true);

  pText_(slide, 'QUICK FACTS', 14, 219, 176, 13,
    { font: 'Arial', size: 8, bold: true, color: '#81c784' });
  pLine_(slide, 14, 232, 176, '#2e7d5b');

  var facts = [
    ['Habitat:', submission.habitat_type || '\u2014'],
    ['Diet:', submission.diet_type || '\u2014'],
    ['Adaptation:', submission.adaptation_type || '\u2014']
  ];
  var fy = 237;
  for (var fi = 0; fi < facts.length; fi++) {
    pText_(slide, facts[fi][0], 14, fy, 44, 13, { font: 'Arial', size: 8, color: '#c8e6c9' });
    pText_(slide, facts[fi][1], 60, fy, 130, 13, { font: 'Arial', size: 8, color: '#a5d6a7' });
    fy += 13;
  }
}

function posterA_drawEcologyCol_(slide, submission) {
  pRect_(slide, 204, 76, 250, 192, '#1a2c3a', '#1565c0', 1, true);
  pRect_(slide, 204, 76, 250, 22, '#1a3a5c', null, 0, true);
  pRect_(slide, 204, 90, 250, 8, '#1a3a5c', null, 0, false);
  pText_(slide, 'ECOLOGY', 214, 79, 230, 16,
    { font: 'Arial', size: 9, bold: true, color: '#90caf9' });

  pRect_(slide, 212, 103, 234, 38, '#0d2137', '#1565c0', 1, true);
  pText_(slide, 'ECOSYSTEM ROLE', 219, 104, 220, 13,
    { font: 'Arial', size: 8, bold: true, color: '#64b5f6' });
  pText_(slide, submission.ecosystem_role || '', 219, 117, 220, 20,
    { font: 'Arial', size: 8, color: '#bbdefb' });

  pText_(slide, 'ECOLOGICAL EXPLANATION', 212, 145, 234, 13,
    { font: 'Arial', size: 8, bold: true, color: '#64b5f6' });
  pLine_(slide, 212, 159, 234, '#1565c0');
  pText_(slide, submission.ecology_explanation || '', 212, 161, 234, 104,
    { font: 'Arial', size: 8, color: '#bbdefb' });
}

function posterA_drawThreatsCol_(slide, submission) {
  pRect_(slide, 462, 76, 250, 192, '#2d1515', '#c62828', 1, true);
  pRect_(slide, 462, 76, 250, 22, '#3e1a1a', null, 0, true);
  pRect_(slide, 462, 90, 250, 8, '#3e1a1a', null, 0, false);
  pText_(slide, 'THREATS', 472, 79, 230, 16,
    { font: 'Arial', size: 9, bold: true, color: '#ef9a9a' });

  var t1 = trim_(submission.threat_1) || 'Threat 1';
  pRect_(slide, 470, 103, 234, 80, '#3e1111', '#c62828', 1, true);
  pText_(slide, 'THREAT 1 \u2014 ' + t1.toUpperCase(), 477, 104, 222, 14,
    { font: 'Arial', size: 8, bold: true, color: '#ef5350' });
  pText_(slide, trim_(submission.threat_1_reason) || '', 477, 119, 222, 60,
    { font: 'Arial', size: 8, color: '#ffcdd2' });

  var t2 = trim_(submission.threat_2) || 'Threat 2';
  pRect_(slide, 470, 188, 234, 78, '#3e1111', '#c62828', 1, true);
  pText_(slide, 'THREAT 2 \u2014 ' + t2.toUpperCase(), 477, 189, 222, 14,
    { font: 'Arial', size: 8, bold: true, color: '#ef5350' });
  pText_(slide, trim_(submission.threat_2_reason) || '', 477, 204, 222, 60,
    { font: 'Arial', size: 8, color: '#ffcdd2' });
}

function posterA_drawConservation_(slide, submission) {
  pRect_(slide, 8, 276, 350, 114, '#1e1a2d', '#6a1b9a', 1, true);
  pRect_(slide, 8, 276, 350, 22, '#2e1f4a', null, 0, true);
  pRect_(slide, 8, 290, 350, 8, '#2e1f4a', null, 0, false);
  pText_(slide, 'CONSERVATION ACTIONS', 18, 279, 320, 16,
    { font: 'Arial', size: 9, bold: true, color: '#ce93d8' });

  var a1 = trim_(submission.action_1) || '';
  var a2 = trim_(submission.action_2) || '';
  pRect_(slide, 18, 305, 155, 20, '#311b47', '#7b1fa2', 1, true);
  pText_(slide, a1, 18, 305, 155, 20,
    { font: 'Arial', size: 8, bold: true, color: '#e1bee7', align: SlidesApp.ParagraphAlignment.CENTER });
  pRect_(slide, 181, 305, 165, 20, '#311b47', '#7b1fa2', 1, true);
  pText_(slide, a2, 181, 305, 165, 20,
    { font: 'Arial', size: 8, bold: true, color: '#e1bee7', align: SlidesApp.ParagraphAlignment.CENTER });

  pText_(slide, 'HOW THESE ACTIONS HELP', 18, 328, 320, 13,
    { font: 'Arial', size: 8, bold: true, color: '#ba68c8' });
  pLine_(slide, 18, 342, 330, '#6a1b9a');
  pText_(slide, trim_(submission.action_explanation) || '', 18, 344, 330, 42,
    { font: 'Arial', size: 8, color: '#e1bee7' });
}

function posterA_drawWhyMatters_(slide, submission) {
  pRect_(slide, 366, 276, 346, 114, '#1a2a1a', '#2e7d5b', 1, true);
  pRect_(slide, 366, 276, 346, 22, '#1e3d23', null, 0, true);
  pRect_(slide, 366, 290, 346, 8, '#1e3d23', null, 0, false);
  pText_(slide, 'WHY THIS SPECIES MATTERS', 376, 279, 326, 16,
    { font: 'Arial', size: 9, bold: true, color: '#a5d6a7' });
  pText_(slide, trim_(submission.why_it_matters) || '', 376, 302, 330, 84,
    { font: 'Arial', size: 8, color: '#c8e6c9' });
}

/* ===== Product B tile section drawers ===== */

function tileA_draw_(slide, submission, studentName) {
  var bg = hexToRgb_('#0d1f14');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  pRect_(slide, 0, 0, 720, 5, '#2e7d5b', null, 0, false);
  pRect_(slide, 0, 5, 720, 22, '#153324', null, 0, false);
  pText_(slide, 'ENDANGERED SPECIES RESCUE', 0, 9, 720, 14,
    { font: 'Arial', size: 8, bold: true, color: '#81c784', align: SlidesApp.ParagraphAlignment.CENTER });

  pText_(slide, submission.common_name || '', 0, 26, 720, 32,
    { font: 'Georgia', size: 22, bold: true, color: '#ffffff', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, submission.scientific_name || '', 0, 59, 720, 16,
    { font: 'Georgia', size: 9, italic: true, color: '#81c784', align: SlidesApp.ParagraphAlignment.CENTER });

  var status = trim_(submission.identified_status) || 'Endangered';
  pRect_(slide, 300, 77, 120, 16, '#c62828', null, 0, true);
  pText_(slide, status.toUpperCase(), 300, 77, 120, 16,
    { font: 'Arial', size: 8, bold: true, color: '#ffcdd2', align: SlidesApp.ParagraphAlignment.CENTER });

  pLine_(slide, 10, 99, 700, '#c62828');
  pText_(slide, 'THREATS', 0, 102, 720, 14,
    { font: 'Arial', size: 8, bold: true, color: '#ef9a9a', align: SlidesApp.ParagraphAlignment.CENTER });

  var t1 = trim_(submission.threat_1) || 'Threat 1';
  pRect_(slide, 10, 119, 700, 80, '#2d0e0e', '#c62828', 1, true);
  pText_(slide, 'THREAT 1 \u2014 ' + t1.toUpperCase(), 18, 121, 684, 14,
    { font: 'Arial', size: 8, bold: true, color: '#ef5350' });
  pText_(slide, trim_(submission.threat_1_reason) || '', 18, 136, 684, 60,
    { font: 'Arial', size: 8, color: '#ffcdd2' });

  var t2 = trim_(submission.threat_2) || 'Threat 2';
  pRect_(slide, 10, 204, 700, 80, '#2d0e0e', '#c62828', 1, true);
  pText_(slide, 'THREAT 2 \u2014 ' + t2.toUpperCase(), 18, 206, 684, 14,
    { font: 'Arial', size: 8, bold: true, color: '#ef5350' });
  pText_(slide, trim_(submission.threat_2_reason) || '', 18, 221, 684, 60,
    { font: 'Arial', size: 8, color: '#ffcdd2' });

  pText_(slide, studentName, 0, 370, 720, 14,
    { font: 'Arial', size: 8, color: '#a5d6a7', align: SlidesApp.ParagraphAlignment.CENTER });
  pRect_(slide, 0, 390, 720, 5, '#2e7d5b', null, 0, false);
}

function tileB_draw_(slide, submission) {
  var bg = hexToRgb_('#0d1826');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  pRect_(slide, 0, 0, 720, 5, '#1565c0', null, 0, false);
  pRect_(slide, 0, 5, 720, 22, '#1a3a5c', null, 0, false);
  pText_(slide, 'ECOLOGY', 0, 9, 720, 14,
    { font: 'Arial', size: 9, bold: true, color: '#90caf9', align: SlidesApp.ParagraphAlignment.CENTER });
  pLine_(slide, 10, 28, 700, '#1565c0');

  var rows = [
    ['Habitat', submission.habitat_type || '\u2014'],
    ['Diet', submission.diet_type || '\u2014'],
    ['Key Adaptation', submission.adaptation_type || '\u2014'],
    ['Ecosystem Role', submission.ecosystem_role || '\u2014']
  ];
  var ry = 36;
  for (var ri = 0; ri < rows.length; ri++) {
    pText_(slide, rows[ri][0], 20, ry, 340, 13, { font: 'Arial', size: 8, bold: true, color: '#64b5f6' });
    pText_(slide, rows[ri][1], 20, ry + 13, 680, 13, { font: 'Arial', size: 8, color: '#bbdefb' });
    ry += 30;
  }

  pLine_(slide, 10, ry + 2, 700, '#1565c0');
  pText_(slide, 'Ecological Explanation', 20, ry + 6, 680, 13,
    { font: 'Arial', size: 8, bold: true, color: '#64b5f6' });
  pText_(slide, trim_(submission.ecology_explanation) || '', 20, ry + 20, 680, 80,
    { font: 'Arial', size: 8, italic: true, color: '#90caf9' });

  pRect_(slide, 0, 390, 720, 5, '#1565c0', null, 0, false);
}

function tileC_draw_(slide, submission, imageUrl) {
  var bg = hexToRgb_('#0a1a10');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  pRect_(slide, 0, 0, 720, 5, '#2e7d5b', null, 0, false);
  pRect_(slide, 8, 10, 704, 280, '#071410', '#2e7d5b', 1, true);

  if (imageUrl) {
    try {
      var resp = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
        slide.insertImage(resp.getBlob(), 8, 10, 704, 280);
      } else {
        try { slide.insertImage(imageUrl, 8, 10, 704, 280); } catch (e2) {}
      }
    } catch (imgErr) {
      try { slide.insertImage(imageUrl, 8, 10, 704, 280); } catch (e2) {}
    }
  }

  pLine_(slide, 8, 294, 704, '#2e7d5b');
  pText_(slide, submission.common_name || '', 0, 299, 720, 22,
    { font: 'Georgia', size: 13, bold: true, color: '#ffffff', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, submission.scientific_name || '', 0, 322, 720, 14,
    { font: 'Georgia', size: 9, italic: true, color: '#81c784', align: SlidesApp.ParagraphAlignment.CENTER });
  var status = trim_(submission.identified_status) || 'Endangered';
  pText_(slide, 'Conservation Status: ' + status.toUpperCase(), 0, 337, 720, 14,
    { font: 'Arial', size: 8, color: '#a5d6a7', align: SlidesApp.ParagraphAlignment.CENTER });
  pRect_(slide, 0, 390, 720, 5, '#2e7d5b', null, 0, false);
}

function tileD_draw_(slide, submission) {
  var bg = hexToRgb_('#160d26');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  pRect_(slide, 0, 0, 720, 5, '#7b1fa2', null, 0, false);
  pRect_(slide, 0, 5, 720, 22, '#2e1f4a', null, 0, false);
  pText_(slide, 'CONSERVATION ACTIONS', 0, 9, 720, 14,
    { font: 'Arial', size: 9, bold: true, color: '#ce93d8', align: SlidesApp.ParagraphAlignment.CENTER });
  pLine_(slide, 10, 28, 700, '#7b1fa2');

  pRect_(slide, 10, 34, 700, 22, '#2a0f40', '#7b1fa2', 1, true);
  pText_(slide, trim_(submission.action_1) || '', 10, 34, 700, 22,
    { font: 'Arial', size: 9, bold: true, color: '#e1bee7', align: SlidesApp.ParagraphAlignment.CENTER });
  pRect_(slide, 10, 61, 700, 22, '#2a0f40', '#7b1fa2', 1, true);
  pText_(slide, trim_(submission.action_2) || '', 10, 61, 700, 22,
    { font: 'Arial', size: 9, bold: true, color: '#e1bee7', align: SlidesApp.ParagraphAlignment.CENTER });

  pLine_(slide, 10, 90, 700, '#7b1fa2');
  pText_(slide, 'How They Help', 20, 94, 680, 13,
    { font: 'Arial', size: 8, bold: true, color: '#ba68c8' });
  pText_(slide, trim_(submission.action_explanation) || '', 20, 108, 680, 80,
    { font: 'Arial', size: 8, color: '#e1bee7' });

  pRect_(slide, 0, 390, 720, 5, '#7b1fa2', null, 0, false);
}

function tileE_draw_(slide, submission, studentName) {
  var bg = hexToRgb_('#0d1f14');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  pRect_(slide, 0, 0, 720, 5, '#2e7d5b', null, 0, false);
  pRect_(slide, 0, 5, 720, 22, '#1e3d23', null, 0, false);
  pText_(slide, 'WHY THIS SPECIES MATTERS', 0, 9, 720, 14,
    { font: 'Arial', size: 9, bold: true, color: '#a5d6a7', align: SlidesApp.ParagraphAlignment.CENTER });
  pLine_(slide, 10, 28, 700, '#2e7d5b');

  pText_(slide, '\u201c', 20, 34, 40, 36, { font: 'Georgia', size: 30, color: '#2e7d5b' });
  pText_(slide, trim_(submission.why_it_matters) || '', 30, 44, 660, 130,
    { font: 'Georgia', size: 9, italic: true, color: '#c8e6c9' });
  pText_(slide, '\u201d', 660, 160, 40, 36, { font: 'Georgia', size: 30, color: '#2e7d5b' });

  pLine_(slide, 10, 300, 700, '#2e7d5b');
  pText_(slide, 'student-authored', 0, 305, 720, 13,
    { font: 'Arial', size: 8, italic: true, color: '#81c784', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, studentName, 0, 320, 720, 14,
    { font: 'Arial', size: 9, color: '#a5d6a7', align: SlidesApp.ParagraphAlignment.CENTER });
  pRect_(slide, 0, 390, 720, 5, '#2e7d5b', null, 0, false);
}

/* ===== Product B builder ===== */

function buildTilePdf_(outputFolder, baseName, submission) {
  var pres = SlidesApp.create(baseName + ' - Wall Tiles');
  var file = DriveApp.getFileById(pres.getId());
  var studentName = ([submission.first_name || '', submission.last_name || '']).join(' ').trim();
  var imageUrl = pickBestPosterImageUrl_(submission);

  var slides = pres.getSlides();
  while (slides.length < 5) { pres.appendSlide(); slides = pres.getSlides(); }

  var clearSlide = function(s) {
    var els = s.getPageElements();
    for (var i = els.length - 1; i >= 0; i--) { try { els[i].remove(); } catch (e) {} }
  };
  for (var si = 0; si < 5; si++) { clearSlide(slides[si]); }

  try { tileA_draw_(slides[0], submission, studentName); } catch (e) { Logger.log('tileA failed: ' + e); }
  try { tileB_draw_(slides[1], submission); } catch (e) { Logger.log('tileB failed: ' + e); }
  try { tileC_draw_(slides[2], submission, imageUrl); } catch (e) { Logger.log('tileC failed: ' + e); }
  try { tileD_draw_(slides[3], submission); } catch (e) { Logger.log('tileD failed: ' + e); }
  try { tileE_draw_(slides[4], submission, studentName); } catch (e) { Logger.log('tileE failed: ' + e); }

  pres.saveAndClose();
  try { outputFolder.addFile(file); } catch (e) {}
  try { DriveApp.getRootFolder().removeFile(file); } catch (e) {}
  return { file: file, warning: '' };
}

/* ===== Poster generation ===== */

function generatePosterForSubmission_(submission) {
  var settings = getSettingsMap_();
  var warnings = [];

  var folderResult = getOutputFolderSafe_(settings.output_folder_id);
  var outputFolder = folderResult.folder;
  if (folderResult.warning) warnings.push(folderResult.warning);

  var studentName = [submission.first_name, submission.last_name].join(' ').trim();
  var posterName = sanitizeFileName_(['Rescue Poster', submission.common_name || 'Species', studentName || 'Student'].join(' - '));

  var posterFile = null;
  var buildResult = null;
  var templateId = settings.poster_template_file_id;
  var canUseTemplate = hasConfiguredDriveId_(templateId) && isTrueSetting_(settings.use_template_poster);

  if (canUseTemplate) {
    try {
      buildResult = buildPosterFromTemplate_(templateId, outputFolder, posterName, submission);
      posterFile = buildResult.file;
      if (buildResult.warning) warnings.push(buildResult.warning);
    } catch (templateError) {
      Logger.log('Template poster build failed. Falling back. ' + templateError);
      warnings.push('Template poster failed, so a simple poster was created instead.');
    }
  }

  if (!posterFile) {
    buildResult = buildFallbackPoster_(outputFolder, posterName, submission);
    posterFile = buildResult.file;
    if (buildResult.warning) warnings.push(buildResult.warning);
  }

  if (!posterFile) throw new Error('Poster file could not be created.');

  var pdfFile = null;
  try {
    pdfFile = exportSlidesFileAsPdfWithRetry_(posterFile.getId(), posterName, outputFolder);
  } catch (pdfError) {
    Logger.log('PDF export failed: ' + pdfError);
    warnings.push('Poster was created, but PDF export failed.');
  }

  var tileName = sanitizeFileName_(['Wall Tiles', submission.common_name || 'Species', studentName || 'Student'].join(' - '));
  var tileFile = null;
  var tilePdfFile = null;
  try {
    var tileResult = buildTilePdf_(outputFolder, tileName, submission);
    tileFile = tileResult.file;
    if (tileResult.warning) warnings.push(tileResult.warning);
  } catch (tileError) {
    Logger.log('Tile build failed: ' + tileError);
    warnings.push('Poster was created, but the wall-tile file could not be generated.');
  }

  if (tileFile) {
    try {
      tilePdfFile = exportSlidesFileAsPdfWithRetry_(tileFile.getId(), tileName, outputFolder);
    } catch (tilePdfError) {
      Logger.log('Tile PDF export failed: ' + tilePdfError);
    }
  }

  try {
    posterFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (pdfFile) pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (tileFile) tileFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (tilePdfFile) tilePdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareError) {
    Logger.log('Sharing update failed: ' + shareError);
  }

  return {
    slideUrl: posterFile.getUrl(),
    pdfUrl: pdfFile ? pdfFile.getUrl() : '',
    posterFileId: posterFile.getId(),
    pdfFileId: pdfFile ? pdfFile.getId() : '',
    tilePdfUrl: tilePdfFile ? tilePdfFile.getUrl() : '',
    tilePdfFileId: tilePdfFile ? tilePdfFile.getId() : '',
    warning: warnings.join(' '),
    studentCanViewFiles: !!posterFile.getId()
  };
}

function buildPosterFromTemplate_(templateId, outputFolder, posterName, submission) {
  var templateFile = DriveApp.getFileById(templateId);
  var copiedFile = templateFile.makeCopy(posterName, outputFolder);
  var presentation = SlidesApp.openById(copiedFile.getId());
  var slides = presentation.getSlides();

  if (!slides || !slides.length) throw new Error('Poster template has no slides.');

  var habitatEcologyText =
    (submission.habitat_type || '') + '\n' +
    (submission.diet_type || '') + '\n' +
    (submission.adaptation_type || '') + '\n' +
    (submission.ecosystem_role || '') + '\n\n' +
    (submission.ecology_explanation || '');

  var majorThreatsText =
    '1. ' + (submission.threat_1 || '') + ': ' + (submission.threat_1_reason || '') + '\n\n' +
    '2. ' + (submission.threat_2 || '') + ': ' + (submission.threat_2_reason || '');

  var actionLines = [];
  if (submission.action_1) actionLines.push('1. ' + submission.action_1);
  if (submission.action_2) actionLines.push((actionLines.length + 1) + '. ' + submission.action_2);
  var conservationActionsText = actionLines.join('\n') + '\n\n' + (submission.action_explanation || '');

  var replacements = {
    '{{STUDENT_NAME}}': [submission.first_name, submission.last_name].join(' ').trim(),
    '{{HOUR}}': submission.hour || '',
    '{{COMMON_NAME}}': submission.common_name || '',
    '{{SCIENTIFIC_NAME}}': submission.scientific_name || '',
    '{{HABITAT_ECOLOGY}}': habitatEcologyText,
    '{{MAJOR_THREATS}}': majorThreatsText,
    '{{CONSERVATION_ACTIONS}}': conservationActionsText,
    '{{WHY_IT_MATTERS}}': submission.why_it_matters || ''
  };

  for (var placeholder in replacements) {
    if (replacements.hasOwnProperty(placeholder)) {
      presentation.replaceAllText(placeholder, replacements[placeholder]);
    }
  }

  presentation.saveAndClose();
  return { file: copiedFile, warning: '' };
}

function buildFallbackPoster_(outputFolder, posterName, submission) {
  var pres = SlidesApp.create(posterName);
  var file = DriveApp.getFileById(pres.getId());
  var slide = pres.getSlides()[0];

  var els = slide.getPageElements();
  for (var i = els.length - 1; i >= 0; i--) {
    try { els[i].remove(); } catch (e) {}
  }

  var bg = hexToRgb_('#0d1f14');
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  var studentName = ([submission.first_name || '', submission.last_name || '']).join(' ').trim();

  var sections = [
    ['topBar',       function() { pRect_(slide, 0, 0, 720, 4, '#2e7d5b', null, 0, false); }],
    ['header',       function() { posterA_drawHeader_(slide, submission, studentName); }],
    ['photoCol',     function() { posterA_drawPhotoCol_(slide, submission); }],
    ['ecologyCol',   function() { posterA_drawEcologyCol_(slide, submission); }],
    ['threatsCol',   function() { posterA_drawThreatsCol_(slide, submission); }],
    ['conservation', function() { posterA_drawConservation_(slide, submission); }],
    ['whyMatters',   function() { posterA_drawWhyMatters_(slide, submission); }],
    ['bottomBar',    function() { pRect_(slide, 0, 401, 720, 4, '#2e7d5b', null, 0, false); }]
  ];

  for (var si = 0; si < sections.length; si++) {
    try {
      sections[si][1]();
    } catch (sectionErr) {
      Logger.log('Poster section "' + sections[si][0] + '" failed: ' + sectionErr);
    }
  }

  var imageUrl = pickBestPosterImageUrl_(submission);
  if (imageUrl) {
    try { insertSlideImageFromUrl_(slide, imageUrl); } catch (imgErr) {
      Logger.log('Image insert failed: ' + imgErr);
    }
  }

  pres.saveAndClose();
  try { outputFolder.addFile(file); } catch (addErr) {}
  try { DriveApp.getRootFolder().removeFile(file); } catch (removeErr) {}
  return { file: file, warning: '' };
}

/* ===== Image and folder helpers ===== */

function pickBestPosterImageUrl_(submission) {
  if (isUsableSpeciesImageUrl_(submission.selected_image_url)) {
    return trim_(submission.selected_image_url);
  }
  var species = submission.species_id ? findSpeciesById_(submission.species_id) : null;
  if (!species) return '';
  var candidates = [species.image_option_1_url, species.image_option_2_url, species.image_option_3_url];
  for (var i = 0; i < candidates.length; i++) {
    if (isUsableSpeciesImageUrl_(candidates[i])) return trim_(candidates[i]);
  }
  return '';
}

function getOutputFolderSafe_(folderId) {
  var idsToTry = [];
  if (hasConfiguredDriveId_(folderId)) idsToTry.push(folderId);
  if (hasConfiguredDriveId_(APP.outputFolderId)) idsToTry.push(APP.outputFolderId);

  for (var t = 0; t < idsToTry.length; t++) {
    try {
      var folder = DriveApp.getFolderById(idsToTry[t]);
      Logger.log('Using output folder: ' + folder.getName() + ' (' + idsToTry[t] + ')');
      return { folder: folder, warning: '' };
    } catch (folderError) {
      Logger.log('Output folder ' + idsToTry[t] + ' unavailable: ' + folderError);
    }
  }
  return {
    folder: DriveApp.getRootFolder(),
    warning: 'Output folder was not configured correctly, so files were saved in Drive root.'
  };
}

function hasConfiguredDriveId_(value) {
  var text = trim_(value);
  if (!text) return false;
  if (text.indexOf('PASTE_') === 0) return false;
  return true;
}

function insertSlideImageFromUrl_(slide, imageUrl) {
  if (!imageUrl) return false;
  try {
    var response = UrlFetchApp.fetch(imageUrl, {
      muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      slide.insertImage(response.getBlob(), APP.posterImageBox.x, APP.posterImageBox.y, APP.posterImageBox.width, APP.posterImageBox.height);
      return true;
    }
  } catch (error) {
    Logger.log('UrlFetch image insert failed. Trying direct URL. ' + error);
  }
  try {
    slide.insertImage(imageUrl, APP.posterImageBox.x, APP.posterImageBox.y, APP.posterImageBox.width, APP.posterImageBox.height);
    return true;
  } catch (fallbackError) {
    Logger.log('Image insert failed for URL: ' + imageUrl + ' :: ' + fallbackError);
  }
  return false;
}

function exportSlidesFileAsPdfWithRetry_(presentationId, posterName, outputFolder) {
  var attempts = 3;
  var lastError = null;
  Utilities.sleep(1000);

  for (var i = 0; i < attempts; i++) {
    try {
      var file = DriveApp.getFileById(presentationId);
      var pdfBlob = file.getAs(MimeType.PDF).setName(posterName + '.pdf');
      return outputFolder.createFile(pdfBlob);
    } catch (drivePdfError) {
      lastError = drivePdfError;
      Logger.log('Drive PDF export attempt ' + (i + 1) + ' failed: ' + drivePdfError);
    }

    try {
      var exportUrl = 'https://docs.google.com/presentation/d/' + presentationId + '/export/pdf';
      var response = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true, followRedirects: true
      });
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        var exportBlob = response.getBlob().setName(posterName + '.pdf');
        return outputFolder.createFile(exportBlob);
      }
      lastError = new Error('Slides export HTTP ' + response.getResponseCode());
    } catch (urlFetchError) {
      lastError = urlFetchError;
      Logger.log('Slides export attempt ' + (i + 1) + ' UrlFetch failed: ' + urlFetchError);
    }
    Utilities.sleep(1200);
  }

  throw lastError || new Error('PDF export failed after retries.');
}

/* ===== Scoring ===== */

function calculateSubmissionScoreByStudentKey_(studentKey, pendingUpdates) {
  var submission = getSubmissionByStudentKey_(studentKey);
  return calculateCompletionScore_(mergeObjects_(submission, pendingUpdates || {}));
}

function calculateCompletionScore_(submission) {
  var score = 0;
  if (isFilled_(submission.species_id)) score += 10;

  if (isFilled_(submission.identified_common_name) && isFilled_(submission.identified_status) &&
      isFilled_(submission.habitat_type) && isFilled_(submission.diet_type) &&
      isFilled_(submission.adaptation_type) && isFilled_(submission.ecosystem_role) &&
      trim_(submission.ecology_explanation).length >= 8) {
    score += 25;
  }

  var threatValues = [submission.threat_1, submission.threat_2];
  if (areAllFilled_(threatValues) && uniqueCount_(threatValues) === 2 &&
      trim_(submission.threat_1_reason).length >= 5 && trim_(submission.threat_2_reason).length >= 5) {
    score += 25;
  }

  var actionValues = [submission.action_1, submission.action_2];
  var filledActions = actionValues.filter(function(v) { return isFilled_(v); });
  if (filledActions.length >= 2 && uniqueCount_(filledActions) === filledActions.length &&
      trim_(submission.action_explanation).length >= 8) {
    score += 20;
  }

  if (trim_(submission.why_it_matters).length >= 8) score += 20;

  if (score > 100) score = 100;
  if (score < 0) score = 0;
  return score;
}

/* ===== Default species catalog ===== */

function getDefaultSpeciesCatalog_() {
  return [
    {
      species_id: 'hawksbill-turtle', common_name: 'Hawksbill Turtle',
      scientific_name: 'Eretmochelys imbricata', biome: 'Marine',
      status: 'Critically Endangered', region: 'Tropical oceans and coral reefs',
      briefing_text: 'Your mission is to investigate the hawksbill turtle. Find out where it lives, what role it plays in reef ecosystems, what threats are causing its decline, and what people can do to help protect it.',
      habitat_hint: 'Coral reef', diet_hint: 'Carnivore', ecosystem_role_hint: 'Consumer in reef food webs',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/sea-turtle/hawksbill-turtle',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'whale-shark', common_name: 'Whale Shark',
      scientific_name: 'Rhincodon typus', biome: 'Marine',
      status: 'Endangered', region: 'Warm tropical oceans',
      briefing_text: 'Your mission is to investigate the whale shark. Learn where it lives, how it survives, what is putting it at risk, and which conservation actions could help protect it.',
      habitat_hint: 'Open ocean', diet_hint: 'Filter feeder', ecosystem_role_hint: 'Large marine consumer',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/shark/whale-shark',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'blue-whale', common_name: 'Blue Whale',
      scientific_name: 'Balaenoptera musculus', biome: 'Marine',
      status: 'Endangered', region: 'Oceans worldwide',
      briefing_text: 'Your mission is to investigate the blue whale. Study where it lives, how it survives, what threats it faces, and how conservation efforts can help it recover.',
      habitat_hint: 'Open ocean', diet_hint: 'Carnivore', ecosystem_role_hint: 'Large marine consumer',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/blue-whale',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'bornean-orangutan', common_name: 'Bornean Orangutan',
      scientific_name: 'Pongo pygmaeus', biome: 'Rainforest / Forest',
      status: 'Critically Endangered', region: 'Borneo rainforests',
      briefing_text: 'Your mission is to investigate the Bornean orangutan. Find out how it survives in the rainforest, what threats are reducing its population, and what humans can do to help.',
      habitat_hint: 'Rainforest', diet_hint: 'Omnivore', ecosystem_role_hint: 'Seed disperser',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/orangutan/bornean-orangutan',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'red-panda', common_name: 'Red Panda',
      scientific_name: 'Ailurus fulgens', biome: 'Rainforest / Forest',
      status: 'Endangered', region: 'Eastern Himalayas and southwestern China',
      briefing_text: 'Your mission is to investigate the red panda. Learn about its forest habitat, what it eats, the threats to its survival, and how people can protect it.',
      habitat_hint: 'Temperate forest', diet_hint: 'Omnivore', ecosystem_role_hint: 'Consumer in forest food webs',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/red-panda',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'african-forest-elephant', common_name: 'African Forest Elephant',
      scientific_name: 'Loxodonta cyclotis', biome: 'Rainforest / Forest',
      status: 'Critically Endangered', region: 'Central and West African forests',
      briefing_text: 'Your mission is to investigate the African forest elephant. Explore its habitat, its role in the ecosystem, the threats it faces, and the actions that could help protect it.',
      habitat_hint: 'Rainforest', diet_hint: 'Herbivore', ecosystem_role_hint: 'Seed disperser',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/elephant/african-elephant/african-forest-elephant',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'african-wild-dog', common_name: 'African Wild Dog',
      scientific_name: 'Lycaon pictus', biome: 'Desert / Grassland',
      status: 'Endangered', region: 'Sub-Saharan grasslands and savannas',
      briefing_text: 'Your mission is to investigate the African wild dog. Find out where it lives, how it survives, what is causing its decline, and how conservation could help its pack populations recover.',
      habitat_hint: 'Grassland', diet_hint: 'Carnivore', ecosystem_role_hint: 'Predator',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/african-wild-dog',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'black-footed-ferret', common_name: 'Black-footed Ferret',
      scientific_name: 'Mustela nigripes', biome: 'Desert / Grassland',
      status: 'Endangered', region: 'North American prairies and grasslands',
      briefing_text: 'Your mission is to investigate the black-footed ferret. Learn about its prairie habitat, how it survives, why it is endangered, and what can help increase its population.',
      habitat_hint: 'Prairie', diet_hint: 'Carnivore', ecosystem_role_hint: 'Predator',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species/black-footed-ferret',
      research_link_2: 'https://www.worldwildlife.org/species'
    },
    {
      species_id: 'addax', common_name: 'Addax',
      scientific_name: 'Addax nasomaculatus', biome: 'Desert / Grassland',
      status: 'Critically Endangered', region: 'Sahara Desert',
      briefing_text: 'Your mission is to investigate the addax. Study how it survives in desert conditions, what is causing its population decline, and what people can do to help prevent extinction.',
      habitat_hint: 'Desert', diet_hint: 'Herbivore', ecosystem_role_hint: 'Grazer',
      image_option_1_url: '', image_option_2_url: '', image_option_3_url: '',
      research_link_1: 'https://www.worldwildlife.org/species',
      research_link_2: 'https://www.saharaconservation.org/species-recovery/restoring-the-addax/'
    }
  ];
}

/* ===== Image URL helpers ===== */

function createPlaceholderImageDataUri_(label, hexColor) {
  var safeLabel = String(label || 'Species').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="800" height="520">' +
    '<rect width="100%" height="100%" fill="' + (hexColor || '#d9d0bf') + '"/>' +
    '<rect x="26" y="26" width="748" height="468" rx="24" fill="#fffaf0" opacity="0.72"/>' +
    '<text x="400" y="230" text-anchor="middle" font-family="Arial" font-size="56" font-weight="700" fill="#2f4f3f">Poster Image</text>' +
    '<text x="400" y="310" text-anchor="middle" font-family="Arial" font-size="36" fill="#2f4f3f">' + safeLabel + '</text>' +
    '</svg>';
  return 'data:image/svg+xml;charset=UTF-8,' + encodeURIComponent(svg);
}

function wikiFileUrl_(fileName) {
  return 'https://commons.wikimedia.org/wiki/Special:FilePath/' + encodeURIComponent(fileName);
}

function getCuratedImageSetForSpecies_(speciesId) {
  var fileNameMap = {
    'hawksbill-turtle': ['Hawksbill Turtle.jpg', 'Hawksbill sea turtle.jpg', 'Hawksbill turtle off the coast of Saba.jpg'],
    'whale-shark': ['Whale Shark AdF.jpg', 'Whale shark.JPG', 'Whale Shark (Rhincodon typus) with open mouth in La Paz, Mexico.jpg'],
    'blue-whale': ['Balaenoptera musculus (blue whale) 1 (31068434305).jpg', 'Bluewhale877.jpg', 'Blue Whale 003 noaa blowholes.jpg'],
    'bornean-orangutan': ['Bornean Orangutan.jpg', 'Bornean orangutan (Pongo pygmaeus), Tanjung Putting National Park 05.jpg', 'Male Bornean Orangutan - Big Cheeks.jpg'],
    'red-panda': ['Red Panda.JPG', 'Red Panda (24986761703).jpg', 'Wiki loves red panda.jpg'],
    'african-forest-elephant': ['Loxodontacyclotis.jpg', 'Forest elephant.jpg', 'African Forest Elephant - Mole National Part Wildlife, Northern Ghana.jpg'],
    'african-wild-dog': ['African Wild Dog at Working with Wildlife.jpg', 'African wild dog (25934660276).jpg', 'African wild dog (7660880326).jpg'],
    'black-footed-ferret': ['Black footed ferret.jpg', 'Black-footed Ferret Learning to Hunt.jpg', 'Black-footed ferret (6001739395).jpg'],
    'addax': ['A big male Addax showing as the power of his horns.jpg', 'Addax nasomaculatus.jpg', 'Addax (Addax nasomaculatus) adult male and juvenile.jpg']
  };
  var fileNames = fileNameMap[speciesId] || [];
  return fileNames.map(function(fn) { return wikiFileUrl_(fn); });
}

function getFallbackImageSetForSpecies_(speciesId, commonName) {
  var curated = getCuratedImageSetForSpecies_(speciesId);
  if (curated && curated.length === 3) return curated;
  var colorMap = {
    'hawksbill-turtle': ['#cfe7ee', '#d8e9d2', '#efe0c8'],
    'whale-shark': ['#d2e6f8', '#d9ecf6', '#cfe3f2'],
    'blue-whale': ['#d7e6fb', '#dbeefe', '#e1f1fc'],
    'bornean-orangutan': ['#f0dccb', '#e7d4c1', '#f2e4d3'],
    'red-panda': ['#efd2c0', '#f4dacd', '#edd8cc'],
    'african-forest-elephant': ['#ddd7d0', '#e5ded6', '#dcd5cc'],
    'african-wild-dog': ['#ead7c5', '#efe1d4', '#f0dfcf'],
    'black-footed-ferret': ['#e9e2d2', '#ece6d8', '#efe9dc'],
    'addax': ['#efe5d2', '#f3ead8', '#f5efe0']
  };
  var colors = colorMap[speciesId] || ['#d9d0bf', '#e5dac5', '#ece6da'];
  return colors.map(function(c) { return createPlaceholderImageDataUri_(commonName, c); });
}

function shouldBackfillImageValue_(value) {
  var text = trim_(value).toLowerCase();
  if (!text) return true;
  if (text.indexOf('data:image/svg+xml') === 0) return true;
  if (text.indexOf('poster image') !== -1) return true;
  if (text.indexOf('species image') !== -1) return true;
  return false;
}

function isUsableSpeciesImageUrl_(url) {
  var text = trim_(url);
  if (!text) return false;
  if (shouldBackfillImageValue_(text)) return false;
  return /^https?:\/\//i.test(text);
}

function resolveSpeciesImageSet_(speciesId, commonName, rawUrls) {
  var curated = getCuratedImageSetForSpecies_(speciesId);
  if (curated && curated.length === 3) return curated;
  var validRaw = rawUrls.filter(function(u) { return isUsableSpeciesImageUrl_(u); }).map(function(u) { return trim_(u); });
  while (validRaw.length && validRaw.length < 3) validRaw.push(validRaw[validRaw.length - 1]);
  if (validRaw.length >= 3) return validRaw.slice(0, 3);
  return getFallbackImageSetForSpecies_(speciesId, commonName);
}

function seedDefaultSpeciesMasterIfNeeded_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.species);
  var existing = getObjectsFromSheet_(sheet);
  if (existing.length) return;

  var defaults = getDefaultSpeciesCatalog_();
  var headers = getHeaders_(sheet);
  for (var i = 0; i < defaults.length; i++) {
    var rowObject = mergeObjects_(defaults[i], {});
    var fallbackImages = getFallbackImageSetForSpecies_(rowObject.species_id, rowObject.common_name);
    rowObject.image_option_1_url = rowObject.image_option_1_url || fallbackImages[0];
    rowObject.image_option_2_url = rowObject.image_option_2_url || fallbackImages[1];
    rowObject.image_option_3_url = rowObject.image_option_3_url || fallbackImages[2];
    appendObjectRow_(sheet, headers, rowObject);
  }
}

function patchMissingSpeciesImageData_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.species);
  var data = getObjectsFromSheet_(sheet);
  var headers = getHeaders_(sheet);
  if (!data.length) return;

  var imageHeaders = ['image_option_1_url', 'image_option_2_url', 'image_option_3_url'];
  for (var i = 0; i < data.length; i++) {
    if (!trim_(data[i].species_id)) continue;
    var resolvedImages = resolveSpeciesImageSet_(data[i].species_id, data[i].common_name, [
      data[i].image_option_1_url, data[i].image_option_2_url, data[i].image_option_3_url
    ]);
    for (var j = 0; j < imageHeaders.length; j++) {
      var colIndex = headers.indexOf(imageHeaders[j]);
      if (colIndex === -1) continue;
      if (trim_(data[i][imageHeaders[j]]) !== resolvedImages[j]) {
        sheet.getRange(i + 2, colIndex + 1).setValue(resolvedImages[j]);
      }
    }
  }
  SpreadsheetApp.flush();
}

/* ===== Data access ===== */

function getSpeciesData_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.species);
  var rows = getObjectsFromSheet_(sheet);
  var defaults = getDefaultSpeciesCatalog_();
  var defaultMap = {};
  var out = [];

  for (var d = 0; d < defaults.length; d++) defaultMap[defaults[d].species_id] = defaults[d];

  for (var i = 0; i < rows.length; i++) {
    if (rows[i].species_id && rows[i].common_name) {
      var base = defaultMap[rows[i].species_id] || {};
      var fallbackImages = getFallbackImageSetForSpecies_(rows[i].species_id, rows[i].common_name);
      out.push({
        species_id: rows[i].species_id,
        common_name: rows[i].common_name || base.common_name || '',
        scientific_name: rows[i].scientific_name || base.scientific_name || '',
        biome: rows[i].biome || base.biome || '',
        status: rows[i].status || base.status || '',
        region: rows[i].region || base.region || '',
        briefing_text: rows[i].briefing_text || base.briefing_text || '',
        habitat_hint: rows[i].habitat_hint || base.habitat_hint || '',
        diet_hint: rows[i].diet_hint || base.diet_hint || '',
        ecosystem_role_hint: rows[i].ecosystem_role_hint || base.ecosystem_role_hint || '',
        image_option_1_url: rows[i].image_option_1_url || base.image_option_1_url || fallbackImages[0],
        image_option_2_url: rows[i].image_option_2_url || base.image_option_2_url || fallbackImages[1],
        image_option_3_url: rows[i].image_option_3_url || base.image_option_3_url || fallbackImages[2],
        research_link_1: rows[i].research_link_1 || base.research_link_1 || '',
        research_link_2: rows[i].research_link_2 || base.research_link_2 || ''
      });
    }
  }

  if (!out.length) {
    for (var j = 0; j < defaults.length; j++) {
      var fi = getFallbackImageSetForSpecies_(defaults[j].species_id, defaults[j].common_name);
      out.push(mergeObjects_(defaults[j], {
        image_option_1_url: defaults[j].image_option_1_url || fi[0],
        image_option_2_url: defaults[j].image_option_2_url || fi[1],
        image_option_3_url: defaults[j].image_option_3_url || fi[2]
      }));
    }
  }
  return out;
}

function findSpeciesById_(speciesId) {
  var species = getSpeciesData_();
  for (var i = 0; i < species.length; i++) {
    if (species[i].species_id === speciesId) return species[i];
  }
  return null;
}

function getSubmissionByStudentKey_(studentKey) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.submissions);
  var objects = getObjectsFromSheet_(sheet);
  for (var i = 0; i < objects.length; i++) {
    if (String(objects[i].student_key) === String(studentKey)) return objects[i];
  }
  throw new Error('Submission not found for studentKey: ' + studentKey);
}

function updateSubmissionByStudentKey_(studentKey, updates) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.submissions);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var headers = values[0];

  if (values.length < 2) throw new Error('No student submissions found to update.');

  var studentKeyColumnIndex = headers.indexOf('student_key');
  if (studentKeyColumnIndex === -1) throw new Error('student_key column missing from STUDENT_SUBMISSIONS.');

  for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (String(values[rowIndex][studentKeyColumnIndex]) === String(studentKey)) {
      var row = values[rowIndex].slice();
      for (var key in updates) {
        if (updates.hasOwnProperty(key)) {
          var colIndex = headers.indexOf(key);
          if (colIndex !== -1) row[colIndex] = updates[key];
        }
      }
      sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
      SpreadsheetApp.flush();
      return;
    }
  }
  throw new Error('Unable to update submission; studentKey not found: ' + studentKey);
}

function getSettingsMap_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.settings);
  var values = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var key = trim_(values[i][0]);
    if (!key) continue;
    map[key] = values[i][1];
  }
  return map;
}

/* ===== Spreadsheet helpers ===== */

function getSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) { setSpreadsheetId_(active.getId()); return active; }
  var storedId = PropertiesService.getScriptProperties().getProperty(APP.scriptProps.spreadsheetIdKey);
  if (storedId) return SpreadsheetApp.openById(storedId);
  throw new Error('No spreadsheet is configured yet. Run setupProjectSheets() once from the Apps Script editor.');
}

function getOrCreateSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) { setSpreadsheetId_(active.getId()); return active; }
  var props = PropertiesService.getScriptProperties();
  var storedId = props.getProperty(APP.scriptProps.spreadsheetIdKey);
  if (storedId) return SpreadsheetApp.openById(storedId);
  var ss = SpreadsheetApp.create(APP.projectTitle + ' Data');
  setSpreadsheetId_(ss.getId());
  return ss;
}

function setSpreadsheetId_(spreadsheetId) {
  PropertiesService.getScriptProperties().setProperty(APP.scriptProps.spreadsheetIdKey, spreadsheetId);
}

function ensureSheetHeaders_(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    autoResizeColumns_(sheet, headers.length);
    return sheet;
  }

  var existingHeaders = getHeaders_(sheet);
  var changed = false;

  if (existingHeaders.length > headers.length) {
    if (sheet.getLastRow() > 1) {
      throw new Error('Sheet "' + sheetName + '" has extra columns. Remove or rename them manually before continuing.');
    }
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    existingHeaders = headers.slice();
    changed = true;
  }

  for (var i = 0; i < headers.length; i++) {
    if (existingHeaders[i] === headers[i]) continue;
    if (i >= existingHeaders.length) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
      existingHeaders.push(headers[i]);
      changed = true;
      continue;
    }
    if (sheet.getLastRow() > 1) {
      throw new Error('Sheet "' + sheetName + '" already has data and a header does not match. Fix headers manually.');
    }
    sheet.getRange(1, i + 1).setValue(headers[i]);
    existingHeaders[i] = headers[i];
    changed = true;
  }

  sheet.setFrozenRows(1);
  autoResizeColumns_(sheet, headers.length);
  if (changed) SpreadsheetApp.flush();
  return sheet;
}

function ensureSettingsSheet_(ss, sheetName, headers, defaultRows) {
  var sheet = ensureSheetHeaders_(ss, sheetName, headers);
  var existing = getObjectsFromSheet_(sheet);
  var existingMap = {};
  for (var i = 0; i < existing.length; i++) existingMap[trim_(existing[i].key)] = i + 2;

  for (var r = 0; r < defaultRows.length; r++) {
    var key = defaultRows[r][0];
    var value = defaultRows[r][1];
    if (existingMap[key]) {
      if (trim_(sheet.getRange(existingMap[key], 2).getValue()) === '') {
        sheet.getRange(existingMap[key], 2).setValue(value);
      }
    } else {
      sheet.appendRow([key, value]);
    }
  }
  autoResizeColumns_(sheet, headers.length);
  return sheet;
}

function appendObjectRow_(sheet, headers, rowObject) {
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(headers[i] in rowObject ? rowObject[headers[i]] : '');
  }
  sheet.appendRow(row);
}

function getHeaders_(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) throw new Error('Sheet has no headers: ' + sheet.getName());
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
}

function getObjectsFromSheet_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var output = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = values[r][c];
    output.push(obj);
  }
  return output;
}

function autoResizeColumns_(sheet, count) {
  for (var i = 1; i <= count; i++) sheet.autoResizeColumn(i);
}

function formatHeaderRow_(sheet) {
  if (!sheet || sheet.getLastColumn() === 0) return;
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold').setBackground('#2f4f3f').setFontColor('#ffffff').setWrap(true);
  sheet.setFrozenRows(1);
}

/* ===== Utility helpers ===== */

function validateRequired_(object, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (!(keys[i] in object) || trim_(object[keys[i]]) === '') {
      throw new Error('Missing required field: ' + keys[i]);
    }
  }
}

function validateString_(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(label + ' must be a non-empty string.');
  }
}

function trim_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function sanitizeFileName_(name) {
  return String(name || 'Poster').replace(/[\\\/:*?"<>|#\[\]]/g, '').replace(/\s+/g, ' ').trim();
}

function isFilled_(value) { return trim_(value) !== ''; }

function areAllFilled_(values) {
  for (var i = 0; i < values.length; i++) { if (!isFilled_(values[i])) return false; }
  return true;
}

function uniqueCount_(values) {
  var seen = {};
  for (var i = 0; i < values.length; i++) seen[trim_(values[i]).toLowerCase()] = true;
  return Object.keys(seen).length;
}

function mergeObjects_(base, updates) {
  var merged = {};
  var key;
  for (key in base) { if (base.hasOwnProperty(key)) merged[key] = base[key]; }
  for (key in updates) { if (updates.hasOwnProperty(key)) merged[key] = updates[key]; }
  return merged;
}

function isTrueSetting_(value) {
  return String(value || '').toUpperCase() === 'TRUE';
}
