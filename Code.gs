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
    x: 24,
    y: 102,
    width: 232,
    height: 142
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

function runThreeEndToEndTests() {
  var timestamp = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyyMMddHHmm');
  var cases = [
    {
      first_name: 'Test', last_name: 'AddaxA_' + timestamp, hour: '1',
      species_id: 'addax',
      common_name: 'Addax', scientific_name: 'Addax nasomaculatus',
      biome: 'Desert', identified_status: 'Critically Endangered',
      habitat_type: 'Desert', diet_type: 'Herbivore',
      adaptation_type: 'Water conservation', ecosystem_role: 'Grazer',
      ecosystem_role_text: 'Addax eat desert plants and help keep plant growth balanced in dry areas.',
      ecology_explanation: 'Addax matter in the desert because they eat tough plants, move around seeds, and are part of the food web in a place where not many big animals can survive.',
      threat_1: 'Habitat loss',
      threat_1_reason: 'Roads, drilling areas, and other human activity can break up the open desert land addax need to travel and find food. When their habitat gets split apart, it is harder for them to survive.',
      threat_2: 'Poaching / overhunting',
      threat_2_reason: 'People have hunted addax for meat and horns, and since there are already so few left, even a small amount of hunting can hurt the population a lot.',
      action_1: 'Protected areas / habitat restoration', action_2: 'Anti-poaching / stronger laws',
      action_explanation: 'Large protected desert areas and stricter anti-poaching patrols would give addax a better chance to recover.',
      why_it_matters: 'Addax are one of the few large animals built for the Sahara, so saving them helps protect a really unique desert ecosystem.',
      selected_image_url: ''
    },
    {
      first_name: 'Test', last_name: 'FerretB_' + timestamp, hour: '2',
      species_id: 'black-footed-ferret',
      common_name: 'Black-Footed Ferret', scientific_name: 'Mustela nigripes',
      biome: 'Grassland', identified_status: 'Endangered',
      habitat_type: 'Prairie', diet_type: 'Carnivore',
      adaptation_type: 'Camouflage', ecosystem_role: 'Predator',
      ecology_explanation: 'Black-footed ferrets help keep prairie dog numbers balanced, and they depend on healthy prairie ecosystems to survive.',
      threat_1: 'Habitat loss',
      threat_1_reason: 'A lot of prairie land has been turned into farms or developed, so ferrets have fewer places to live and fewer prairie dog colonies to hunt in.',
      threat_2: 'Disease',
      threat_2_reason: 'Diseases like plague can kill prairie dogs and also hurt ferrets, which is a big problem because ferrets rely on prairie dogs for food and shelter.',
      action_1: 'Protected areas / habitat restoration', action_2: 'Captive breeding / reintroduction',
      action_explanation: 'Breeding ferrets in captivity, releasing them into safe prairie areas, and protecting prairie dog towns can help rebuild their population.',
      why_it_matters: 'Saving black-footed ferrets also helps protect prairie grasslands and the animals that live there.',
      selected_image_url: ''
    },
    {
      first_name: 'Test', last_name: 'BlueWhaleC_' + timestamp, hour: '3',
      species_id: 'blue-whale',
      common_name: 'Blue Whale', scientific_name: 'Balaenoptera musculus',
      biome: 'Marine', identified_status: 'Endangered',
      habitat_type: 'Open ocean', diet_type: 'Carnivore',
      adaptation_type: 'Blubber', ecosystem_role: 'Consumer',
      ecology_explanation: 'Blue whales are important in ocean food webs, and even their waste helps move nutrients around the water so other life can grow.',
      threat_1: 'Habitat loss',
      threat_1_reason: 'Busy shipping routes and loud boat traffic can disturb blue whales in places where they feed and migrate. Some whales also get hit by ships.',
      threat_2: 'Climate change',
      threat_2_reason: 'Warming oceans can change where krill live, and krill are the main food blue whales eat. If the food moves or drops, whales have a harder time finding enough to eat.',
      action_1: 'Protected areas / habitat restoration', action_2: 'Pollution reduction',
      action_explanation: 'Safer shipping lanes, slower ship speeds, and cleaner oceans can lower the risks blue whales face.',
      why_it_matters: 'Blue whales are the biggest animals on Earth, and protecting them helps keep ocean ecosystems healthier.',
      selected_image_url: ''
    }
  ];

  var results = [];
  for (var i = 0; i < cases.length; i++) {
    try {
      var out = generatePosterForSubmission_(cases[i]);
      results.push({
        case: i + 1,
        species: cases[i].common_name,
        ok: true,
        slideUrl: out.slideUrl,
        pdfUrl: out.pdfUrl,
        warning: out.warning
      });
      Logger.log('Test ' + (i + 1) + ' (' + cases[i].common_name + ') OK: ' + out.slideUrl + ' | ' + out.pdfUrl);
    } catch (err) {
      results.push({ case: i + 1, species: cases[i].common_name, ok: false, error: String(err) });
      Logger.log('Test ' + (i + 1) + ' FAILED: ' + err);
    }
  }
  return results;
}

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

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'runE2ETests' &&
      e.parameter.key === 'codex-e2e-2026') {
    var results = runThreeEndToEndTests();
    return ContentService
      .createTextOutput(JSON.stringify(results, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

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
    ['use_template_poster', 'FALSE'],
    // Poster renderer: "fallback_v2" (default) or "template".
    ['poster_render_mode', 'fallback_v2']
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
  patchMissingSpeciesImageData_();

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

function getEmbeddableImagePreview(url) {
  var text = trim_(url);
  if (!text) return '';
  return fetchImagePreviewDataUri_(text);
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
  var output = emptyPosterOutput_();

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
    tile_pdf_file_id: '',
    tile_pdf_url: '',
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
    pdfUrl: output.pdfUrl || '',
    studentCanViewFiles: canStudentOpenPosterPdf_(output),
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
    pdfUrl: ''
  };
}

function emptyPosterOutput_() {
  return {
    slideUrl: '',
    pdfUrl: '',
    posterFileId: '',
    pdfFileId: '',
    warning: '',
    studentCanViewFiles: false
  };
}

function canStudentOpenPosterPdf_(output) {
  return !!(output && output.studentCanViewFiles && output.pdfUrl);
}

function getSubmissionDebug(studentKey) {
  return getSubmissionByStudentKey_(studentKey);
}

/* ===== Poster drawing primitives ===== */

function hexToRgb_(hex) {
  var r = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return r ? { r: parseInt(r[1], 16), g: parseInt(r[2], 16), b: parseInt(r[3], 16) } : { r: 255, g: 255, b: 255 };
}

function normalizeSlideFontFamily_(fontName) {
  var font = trim_(fontName);
  if (!font) return 'Arial';

  // Slides PDF export has been dropping body/title copy for some web fonts
  // in this project, so we normalize to export-safe families.
  var safeFonts = {
    'montserrat': 'Arial',
    'lato': 'Arial',
    'merriweather': 'Georgia'
  };

  return safeFonts[font.toLowerCase()] || font;
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
  opts = opts || {};
  var tb = slide.insertTextBox(String(text || ''), x, y, w, h);
  var range = tb.getText();

  try {
    var ts = range.getTextStyle();
    try { ts.setFontFamily(normalizeSlideFontFamily_(opts.font || 'Arial')); } catch (fontErr) { Logger.log('pText font failed: ' + fontErr); }
    try { ts.setFontSize(opts.size || 9); } catch (sizeErr) { Logger.log('pText size failed: ' + sizeErr); }
    if (opts.bold) {
      try { ts.setBold(true); } catch (boldErr) { Logger.log('pText bold failed: ' + boldErr); }
    }
    if (opts.italic) {
      try { ts.setItalic(true); } catch (italicErr) { Logger.log('pText italic failed: ' + italicErr); }
    }
    if (opts.color) {
      try { ts.setForegroundColor(opts.color); } catch (colorErr) { Logger.log('pText color failed: ' + colorErr); }
    }
  } catch (styleErr) {
    Logger.log('pText style block failed: ' + styleErr);
  }

  if (opts.align) {
    try {
      var paras = range.getParagraphs();
      for (var i = 0; i < paras.length; i++) {
        try {
          paras[i].getParagraphStyle().setParagraphAlignment(opts.align);
        } catch (alignErr) {
          Logger.log('pText align failed: ' + alignErr);
        }
      }
    } catch (paraErr) {
      Logger.log('pText paragraph block failed: ' + paraErr);
    }
  }
  return tb;
}

function pLine_(slide, x, y, w, colorHex) {
  return pRect_(slide, x, y, w, 1, colorHex, null, 0, false);
}

function pCircle_(slide, x, y, size, fillHex, borderHex, borderPt) {
  var shape = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, size, size);
  var fc = hexToRgb_(fillHex || '#ffffff');
  shape.getFill().setSolidFill(fc.r, fc.g, fc.b);
  if (borderHex) {
    var bc = hexToRgb_(borderHex);
    shape.getBorder().getLineFill().setSolidFill(bc.r, bc.g, bc.b);
    shape.getBorder().setWeight(borderPt || 1);
  } else {
    shape.getBorder().setTransparent();
  }
  return shape;
}

function pBadge_(slide, x, y, size, fillHex, borderHex, borderPt, label, textOpts) {
  pCircle_(slide, x, y, size, fillHex, borderHex, borderPt);
  pText_(slide, label || '', x, y + Math.max(0, Math.floor(size * 0.18)), size, Math.max(12, Math.floor(size * 0.52)), textOpts || {});
}

function svgBlobFromMarkup_(svgMarkup, name) {
  return Utilities.newBlob(svgMarkup, 'image/svg+xml', name || 'poster-icon.svg');
}

function insertSvgMarkup_(slide, svgMarkup, x, y, width, height, sendBack) {
  try {
    var image = slide.insertImage(svgBlobFromMarkup_(svgMarkup), x, y, width, height);
    if (sendBack) {
      try { image.sendToBack(); } catch (sendErr) {}
    }
    return image;
  } catch (svgErr) {
    Logger.log('SVG insert failed: ' + svgErr);
    return null;
  }
}

function buildPosterIconSvg_(type, strokeHex, fillHex, strokeWidth) {
  var stroke = strokeHex || '#ffffff';
  var fill = fillHex || 'none';
  var width = strokeWidth || 3;
  var svgStart = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" fill="none" stroke="' + stroke + '" stroke-width="' + width + '" stroke-linecap="round" stroke-linejoin="round">';
  var svgEnd = '</svg>';
  var body = '';

  switch (type) {
    case 'turtle':
      body = '<ellipse cx="32" cy="31" rx="15" ry="18" fill="' + fill + '"/>' +
        '<path d="M24 18c3-4 13-4 16 0"/>' +
        '<path d="M22 28c4-2 16-2 20 0"/>' +
        '<path d="M24 40c3 3 13 3 16 0"/>' +
        '<circle cx="32" cy="11.5" r="4"/>' +
        '<path d="M20 22l-6-6"/><path d="M44 22l6-6"/><path d="M20 40l-6 6"/><path d="M44 40l6 6"/>';
      break;
    case 'leaf':
      body = '<path d="M47 16C31 16 18 25 16 42c15 4 30-4 33-26Z" fill="' + fill + '"/>' +
        '<path d="M20 40c8-7 14-11 24-18"/><path d="M28 31c2 3 4 6 5 11"/>';
      break;
    case 'heart':
      body = '<path d="M32 50 14 32c-5-5-5-14 0-19 5-5 13-5 18 0 5-5 13-5 18 0 5 5 5 14 0 19L32 50Z" fill="' + fill + '"/>';
      break;
    case 'coral':
      body = '<path d="M18 48c3-9 5-17 4-26"/><path d="M22 30c-4-3-6-7-7-12"/><path d="M23 24c5-1 9-5 11-11"/>' +
        '<path d="M30 48c1-9 1-18 0-28"/><path d="M30 26c4-2 8-7 10-12"/><path d="M30 33c-4-2-8-6-11-10"/>' +
        '<path d="M40 48c-1-8-2-14-1-22"/><path d="M39 29c4-2 7-5 9-10"/>';
      break;
    case 'waves':
      body = '<path d="M10 22c5 0 5 4 10 4s5-4 10-4 5 4 10 4 5-4 10-4"/>' +
        '<path d="M10 32c5 0 5 4 10 4s5-4 10-4 5 4 10 4 5-4 10-4"/>' +
        '<path d="M10 42c5 0 5 4 10 4s5-4 10-4 5 4 10 4 5-4 10-4"/>';
      break;
    case 'alert':
      body = '<path d="M32 11 53 49H11L32 11Z" fill="' + fill + '"/><path d="M32 25v12"/><circle cx="32" cy="43" r="2" fill="' + stroke + '" stroke="none"/>';
      break;
    case 'thermometer':
      body = '<path d="M32 16v22"/><path d="M27 16a5 5 0 0 1 10 0v22a10 10 0 1 1-10 0Z" fill="' + fill + '"/><circle cx="32" cy="46" r="5" fill="' + fill + '"/>';
      break;
    case 'shield':
      body = '<path d="M32 12 47 18v12c0 10-6 18-15 22-9-4-15-12-15-22V18l15-6Z" fill="' + fill + '"/>' +
        '<path d="M32 23c-5 0-9 3-10 9 7 2 15-2 18-10-3 1-5 1-8 1Z"/><path d="M27 33c3-2 6-4 9-8"/>';
      break;
    case 'community':
      body = '<circle cx="23" cy="24" r="5"/><circle cx="41" cy="24" r="5"/>' +
        '<path d="M15 44c1-8 6-12 11-12s10 4 11 12"/><path d="M29 44c1-8 6-12 11-12s10 4 11 12"/>';
      break;
    case 'scales':
      body = '<path d="M32 14v32"/><path d="M20 20h24"/><path d="M22 20 16 31h12l-6-11Z" fill="' + fill + '"/><path d="M42 20 36 31h12l-6-11Z" fill="' + fill + '"/><path d="M24 46h16"/>';
      break;
    case 'pin':
      body = '<path d="M32 51c8-8 12-15 12-22a12 12 0 1 0-24 0c0 7 4 14 12 22Z" fill="' + fill + '"/><circle cx="32" cy="28" r="4"/>';
      break;
    case 'ruler':
      body = '<rect x="14" y="23" width="36" height="18" rx="3" fill="' + fill + '"/>' +
        '<path d="M20 23v7"/><path d="M26 23v4"/><path d="M32 23v7"/><path d="M38 23v4"/><path d="M44 23v7"/>';
      break;
    case 'weight':
      body = '<path d="M24 24a8 8 0 0 1 16 0"/><path d="M18 28h28l-4 22H22l-4-22Z" fill="' + fill + '"/>';
      break;
    case 'info':
      body = '<circle cx="32" cy="18" r="3" fill="' + stroke + '" stroke="none"/><path d="M32 28v18"/>';
      break;
    case 'turtle-coral':
      body = '<ellipse cx="39" cy="26" rx="11" ry="13" fill="' + fill + '"/><path d="M33 18c2-3 10-3 12 0"/><path d="M31 26c3-1 12-1 16 0"/><path d="M34 34c3 2 10 2 12 0"/><circle cx="39" cy="11" r="3.5"/><path d="M29 21l-5-4"/><path d="M48 21l5-4"/><path d="M29 33l-5 4"/><path d="M48 33l5 4"/>' +
        '<path d="M14 49c2-7 4-13 3-20"/><path d="M17 34c-3-2-5-5-5-9"/><path d="M17 29c4-1 7-4 8-8"/><path d="M22 49c1-7 1-13 0-21"/><path d="M22 31c3-1 6-4 8-8"/>';
      break;
    case 'waves-decor':
      body = '<path stroke-opacity=".24" d="M2 16c8 0 8 6 16 6s8-6 16-6 8 6 16 6 8-6 16-6"/>' +
        '<path stroke-opacity=".18" d="M4 28c10 0 10 7 20 7s10-7 20-7 10 7 20 7"/>' +
        '<path stroke-opacity=".14" d="M8 40c9 0 9 5 18 5s9-5 18-5 9 5 18 5"/>';
      break;
    case 'coral-decor':
      body = '<path stroke-opacity=".24" d="M10 56c4-11 6-22 5-34"/><path stroke-opacity=".24" d="M15 36c-5-4-7-9-8-16"/><path stroke-opacity=".24" d="M16 28c6-1 10-6 13-14"/>' +
        '<path stroke-opacity=".18" d="M28 56c2-12 2-24 0-36"/><path stroke-opacity=".18" d="M28 31c5-2 10-8 13-15"/><path stroke-opacity=".18" d="M28 38c-5-2-10-7-14-13"/>' +
        '<path stroke-opacity=".14" d="M42 56c-1-10-2-18-1-28"/><path stroke-opacity=".14" d="M41 31c5-2 9-7 11-13"/>';
      break;
    default:
      body = '<circle cx="32" cy="32" r="12" fill="' + fill + '"/>';
  }

  return svgStart + body + svgEnd;
}

function pIconBadge_(slide, type, x, y, size, fillHex, borderHex, borderPt, iconHex, inset) {
  pCircle_(slide, x, y, size, fillHex, borderHex, borderPt);
  var pad = inset == null ? Math.max(3, Math.round(size * 0.22)) : inset;
  var inserted = insertSvgMarkup_(
    slide,
    buildPosterIconSvg_(type, iconHex || '#ffffff', 'none', Math.max(2, Math.round(size * 0.075))),
    x + pad, y + pad, size - pad * 2, size - pad * 2
  );
  var label = posterIconFallbackLabel_(type);
  if (!inserted && label) {
    pText_(slide, posterIconFallbackLabel_(type), x, y + Math.max(0, Math.floor(size * 0.17)), size, Math.max(10, Math.floor(size * 0.5)),
      {
        font: 'Arial',
        size: Math.max(6, Math.floor(size * 0.26)),
        bold: true,
        color: iconHex || '#ffffff',
        align: SlidesApp.ParagraphAlignment.CENTER
      });
  }
}

function pIcon_(slide, type, x, y, size, strokeHex, strokeWidth) {
  var inserted = insertSvgMarkup_(slide, buildPosterIconSvg_(type, strokeHex || '#1f5f51', 'none', strokeWidth || 3), x, y, size, size);
  var label = posterIconFallbackLabel_(type);
  if (!inserted && label) {
    pText_(slide, posterIconFallbackLabel_(type), x, y + Math.max(0, Math.floor(size * 0.12)), size, Math.max(10, Math.floor(size * 0.56)),
      {
        font: 'Arial',
        size: Math.max(6, Math.floor(size * 0.28)),
        bold: true,
        color: strokeHex || '#1f5f51',
        align: SlidesApp.ParagraphAlignment.CENTER
      });
  }
  return inserted;
}

function pDecorIcon_(slide, type, x, y, width, height, strokeHex) {
  return insertSvgMarkup_(slide, buildPosterIconSvg_(type, strokeHex || '#2c6b5b', 'none', 2), x, y, width, height, true);
}

function posterIconFallbackLabel_(type) {
  var labels = {
    'turtle': '',
    'leaf': '',
    'heart': '',
    'coral': '',
    'waves': '',
    'alert': '!',
    'thermometer': '',
    'shield': '',
    'community': '',
    'scales': '',
    'pin': '',
    'ruler': '',
    'weight': '',
    'info': 'i',
    'turtle-coral': ''
  };
  return labels[type] || '•';
}

function pDotDivider_(slide, x, y, width, colorHex, dotSize, gap) {
  var size = dotSize || 2;
  var step = size + (gap || 5);
  var count = Math.max(3, Math.floor(width / step));
  for (var i = 0; i < count; i++) {
    pCircle_(slide, x + i * step, y, size, colorHex || '#9dc8c5', null, 0);
  }
}

function posterActionCaption_(text, targetBreak, maxChars) {
  var clean = trim_(text).replace(/\s+/g, ' ');
  if (!clean) return '';
  var clipped = clipText_(clean, maxChars || 34);
  if (clipped.length <= (targetBreak || 14)) return clipped;
  var pivot = clipped.lastIndexOf(' ', targetBreak || 14);
  if (pivot < 8) pivot = clipped.indexOf(' ', targetBreak || 14);
  if (pivot < 0) return clipped;
  return clipped.substring(0, pivot) + '\n' + clipped.substring(pivot + 1);
}

function posterThreatLabelShort_(text) {
  var clean = trim_(text).toLowerCase();
  var map = {
    'habitat loss': 'HABITAT LOSS',
    'climate change': 'CLIMATE CHANGE',
    'poaching / overhunting': 'POACHING',
    'poaching': 'POACHING',
    'overhunting': 'OVERHUNTING',
    'pollution': 'POLLUTION',
    'disease': 'DISEASE'
  };
  if (map[clean]) return map[clean];
  return clipText_(trim_(text).toUpperCase(), 18);
}

function posterActionLabelShort_(text, fallback) {
  var clean = trim_(text).toLowerCase();
  var map = {
    'protected areas / habitat restoration': 'Protected Areas',
    'protected areas': 'Protected Areas',
    'habitat restoration': 'Habitat Restoration',
    'pollution reduction': 'Reduce Pollution',
    'reduce pollution': 'Reduce Pollution',
    'anti-poaching / stronger laws': 'Anti-Poaching',
    'anti-poaching': 'Anti-Poaching',
    'stronger laws': 'Stronger Laws',
    'community conservation programs': 'Community Support',
    'support community conservation': 'Community Support',
    'enforce laws against poaching': 'Enforce Laws'
  };
  if (map[clean]) return map[clean];
  if (fallback) return fallback;
  return clipText_(trim_(text), 18);
}

function posterStatusPalette_(status) {
  var value = String(status || '').toLowerCase();
  if (value.indexOf('critical') >= 0) {
    return { fill: '#8f3430', border: '#d98177', text: '#fff1ee' };
  }
  if (value.indexOf('vulnerable') >= 0 || value.indexOf('watch') >= 0) {
    return { fill: '#8a6a27', border: '#d6bf7e', text: '#fff6db' };
  }
  return { fill: '#2f6b39', border: '#8dbc8f', text: '#eef7ea' };
}

/* Override icon fallback labels so failed icon inserts degrade to empty circles
 * instead of noisy placeholder bullets/letters.
 */
function posterIconFallbackLabel_(type) {
  var labels = {
    'turtle': '',
    'leaf': '',
    'heart': '',
    'coral': '',
    'waves': '',
    'alert': '!',
    'thermometer': '',
    'shield': '',
    'community': '',
    'scales': '',
    'pin': '',
    'ruler': '',
    'weight': '',
    'info': 'i',
    'turtle-coral': ''
  };
  return Object.prototype.hasOwnProperty.call(labels, type) ? labels[type] : '';
}

function clipText_(text, maxChars) {
  var clean = trim_(text).replace(/\s+/g, ' ');
  if (!clean) return '';
  if (!maxChars || clean.length <= maxChars) return clean;
  return clean.substring(0, Math.max(0, maxChars - 3)).trim() + '...';
}

function getPosterFactPack_(submission) {
  var bySpecies = {
    'hawksbill-turtle': { length: '70-95 cm (28-37 in)', weight: '45-70 kg (99-154 lbs)', location: 'Tropical and subtropical oceans', diet: 'Omnivore', habitat: 'Coral reefs' },
    'whale-shark': { length: '9-12 m (30-40 ft)', weight: '15-20 t (33k-44k lbs)', location: 'Warm tropical oceans', diet: 'Filter feeder', habitat: 'Open ocean' },
    'blue-whale': { length: '24-30 m (79-98 ft)', weight: '90-150 t (198k-330k lbs)', location: 'Oceans worldwide', diet: 'Filter feeder', habitat: 'Open ocean' },
    'bornean-orangutan': { length: '1.2-1.5 m (4-5 ft)', weight: '30-100 kg (66-220 lbs)', location: 'Borneo rainforests', diet: 'Omnivore', habitat: 'Rainforest canopy' },
    'red-panda': { length: '50-64 cm (20-25 in)', weight: '3-6 kg (7-13 lbs)', location: 'Eastern Himalayas', diet: 'Omnivore', habitat: 'Temperate forest' },
    'african-forest-elephant': { length: '2.4-3 m (8-10 ft)', weight: '2-5 t (4k-11k lbs)', location: 'Central and West African forests', diet: 'Herbivore', habitat: 'Tropical forest' },
    'african-wild-dog': { length: '75-110 cm (30-43 in)', weight: '18-36 kg (40-79 lbs)', location: 'Sub-Saharan savannas', diet: 'Carnivore', habitat: 'Grassland and savanna' },
    'black-footed-ferret': { length: '38-61 cm (15-24 in)', weight: '0.6-1.4 kg (1-3 lbs)', location: 'North American prairies', diet: 'Carnivore', habitat: 'Prairie grassland' },
    'addax': { length: '1-1.3 m (3-4 ft)', weight: '60-125 kg (132-275 lbs)', location: 'Sahara Desert', diet: 'Herbivore', habitat: 'Desert' }
  };

  var key = trim_(submission.species_id).toLowerCase();
  var defaults = bySpecies[key] || {};

  return {
    habitat: trim_(submission.habitat_type) || defaults.habitat || trim_(submission.biome) || 'Varies by species',
    diet: trim_(submission.diet_type) || defaults.diet || 'Varies by species',
    length: defaults.length || 'See species guide',
    weight: defaults.weight || 'See species guide',
    location: defaults.location || trim_(submission.biome) || 'Multiple regions'
  };
}

/* ===== Product A poster section drawers =====
 *
 * Grid (poster is 720 x 405 pt):
 *   header band      : y 0  .. 72
 *   left column      : x 12 .. 260   (photo on top, quick facts below)
 *   middle column    : x 268 .. 476  (ecology on top, conservation below)
 *   right column     : x 484 .. 712  (threats on top, why-matters below)
 *   footer           : y 382 .. 402
 *
 * Palette matches the target rescue-report layout:
 *   - deep forest green background  #07281f
 *   - cream quick-facts card        #f4ecd6
 *   - mint ecology card             #dfeceb
 *   - blush threats card            #fbe5d8
 *   - lavender conservation card    #ebe4f5
 *   - sage why-matters card         #dfe8d4
 */

var POSTER_A = {
  bg:        '#07281f',
  bgBorder:  '#0f5448',
  headerBg:  '#061f19',
  accent:    '#9fce6b',
  cream:     '#f4ecd6',
  mint:      '#dfeceb',
  blush:     '#fbe5d8',
  lavender:  '#ebe4f5',
  sage:      '#dfe8d4',
  body:      '#2b3b36',
  mutedLine: '#c5cbbf'
};

function posterA_drawHeader_(slide, submission, studentName) {
  pRect_(slide, 0, 0, 720, 405, POSTER_A.bg, POSTER_A.bg, 0, false);
  pRect_(slide, 6, 4, 708, 397, '#08261d', '#124d41', 1, true);
  pRect_(slide, 10, 8, 700, 389, POSTER_A.bg, '#0f5448', 1, true);
  pRect_(slide, 12, 10, 696, 82, POSTER_A.headerBg, null, 0, true);
  pRect_(slide, 12, 88, 696, 2, POSTER_A.bgBorder, null, 0, false);

  pDecorIcon_(slide, 'waves-decor', 430, 28, 152, 46, '#2a6a5d');
  pDecorIcon_(slide, 'turtle', 520, 24, 62, 62, '#1d574b');
  pDecorIcon_(slide, 'coral-decor', 596, 18, 92, 64, '#275f52');

  pCircle_(slide, 18, 14, 52, '#0f5f53', '#cfceb2', 1.2);
  pCircle_(slide, 22, 18, 44, '#093a31', '#cfceb2', 1);
  pText_(slide, 'ESR', 22, 31, 44, 12,
    { font: 'Montserrat', size: 11, bold: true, color: '#f4ecd6', align: SlidesApp.ParagraphAlignment.CENTER });

  pText_(slide, 'ENDANGERED SPECIES RESCUE REPORT', 86, 14, 430, 14,
    { font: 'Montserrat', size: 10, bold: true, color: '#cfd9a8' });
  pText_(slide, submission.common_name || 'Species Rescue', 86, 22, 430, 40,
    { font: 'Merriweather', size: 30, bold: true, color: '#ffffff' });
  pText_(slide, submission.scientific_name || '', 86, 62, 430, 18,
    { font: 'Merriweather', size: 14, italic: true, color: POSTER_A.accent });

  var attribution = 'Student: ' + studentName;
  if (trim_(submission.hour)) attribution += '    Period ' + trim_(submission.hour);
  pText_(slide, attribution, 498, 18, 196, 12,
    { font: 'Lato', size: 7, color: '#cadb9d', align: SlidesApp.ParagraphAlignment.END });
}

function posterA_drawPhotoCol_(slide, submission) {
  pRect_(slide, 16, 94, 248, 156, '#efe7d4', '#d7d7ba', 1, true);
  pRect_(slide, 20, 98, 240, 148, '#0c2f28', '#7ba394', 0.8, true);

  var facts = getPosterFactPack_(submission);
  pRect_(slide, 16, 254, 248, 128, POSTER_A.cream, '#d6d4bb', 1, true);
  pRect_(slide, 16, 254, 248, 20, '#0b5144', null, 0, true);
  pCircle_(slide, 24, 257, 14, '#2f7868', '#9ad1b0', 1);
  pText_(slide, 'QUICK FACTS', 44, 258, 126, 12,
    { font: 'Montserrat', size: 9, bold: true, color: '#f2f8e6' });

  var status = trim_(submission.identified_status);
  if (status) {
    var statusTone = posterStatusPalette_(status);
    pRect_(slide, 176, 258, 78, 11, statusTone.fill, statusTone.border, 0.7, true);
    pText_(slide, status.toUpperCase(), 180, 259, 70, 8,
      { font: 'Montserrat', size: 5, bold: true, color: statusTone.text, align: SlidesApp.ParagraphAlignment.CENTER });
  }

  var rows = [
    ['leaf', 'Habitat', facts.habitat],
    ['waves', 'Diet', facts.diet],
    ['ruler',  'Length', facts.length],
    ['weight', 'Weight', facts.weight],
    ['pin', 'Location', facts.location]
  ];
  var ry = 282;
  for (var i = 0; i < rows.length; i++) {
    pIcon_(slide, rows[i][0], 28, ry + 1, 12, '#19584b', 2.1);
    pText_(slide, rows[i][1], 50, ry, 72, 14,
      { font: 'Montserrat', size: 8, bold: true, color: '#174a3f' });
    pText_(slide, clipText_(rows[i][2], 38), 126, ry, 124, 14,
      { font: 'Lato', size: 8, color: POSTER_A.body });
    if (i < rows.length - 1) pLine_(slide, 28, ry + 14, 216, '#d4d0bb');
    ry += 22;
  }
}

function posterA_drawEcologyCol_(slide, submission) {
  pRect_(slide, 278, 96, 192, 170, POSTER_A.mint, '#9bc3c7', 1, true);
  pIconBadge_(slide, 'leaf', 286, 100, 22, '#0f7f79', '#0f6a63', 1, '#eefbf9', 4);
  pText_(slide, 'ECOLOGY', 318, 104, 140, 18,
    { font: 'Montserrat', size: 12, bold: true, color: '#166f74' });

  pRect_(slide, 288, 128, 172, 128, '#f6faf8', '#95bcc0', 1, true);
  var roleText = trim_(submission.ecosystem_role_text) || trim_(submission.ecosystem_role) || '';
  pIconBadge_(slide, 'coral', 298, 140, 24, '#f7ffff', '#6aaab0', 1, '#157077', 4);
  pText_(slide, 'ECOSYSTEM ROLE', 332, 142, 116, 12,
    { font: 'Montserrat', size: 8, bold: true, color: '#12656a' });
  pText_(slide, clipText_(roleText, 148), 332, 154, 116, 30,
    { font: 'Lato', size: 8, color: POSTER_A.body });

  pText_(slide, '·····································', 300, 195, 150, 8,
    { font: 'Arial', size: 6, color: '#79bac2', align: SlidesApp.ParagraphAlignment.CENTER });

  pRect_(slide, 288, 202, 176, 56, '#f5f8f5', '#95bcc0', 1, true);
  pIconBadge_(slide, 'waves', 294, 208, 18, '#ffffff', '#6aaab0', 1, '#157077', 3);
  pText_(slide, 'ECOLOGICAL EXPLANATION', 318, 208, 136, 12,
    { font: 'Montserrat', size: 8, bold: true, color: '#12656a' });
  pText_(slide, clipText_(submission.ecology_explanation, 170), 318, 220, 136, 34,
    { font: 'Lato', size: 8, color: POSTER_A.body });
}

function posterA_drawThreatsCol_(slide, submission) {
  pRect_(slide, 476, 96, 226, 168, POSTER_A.blush, '#efb7a7', 1, true);
  pIconBadge_(slide, 'alert', 482, 100, 22, '#ec856d', '#df765f', 1, '#fff5f2', 3);
  pText_(slide, 'THREATS', 514, 104, 166, 18,
    { font: 'Montserrat', size: 12, bold: true, color: '#c43d2f' });

  var t1 = trim_(submission.threat_1) || 'Threat 1';
  pRect_(slide, 486, 126, 208, 68, '#fff7f2', '#efab99', 1, true);
  pIcon_(slide, 'coral', 496, 144, 24, '#ad2c23', 2.1);
  pText_(slide, 'THREAT 1 \u2013 ' + t1.toUpperCase(), 530, 132, 154, 12,
    { font: 'Montserrat', size: 8, bold: true, color: '#c53e31' });
  pText_(slide, clipText_(submission.threat_1_reason, 170), 530, 146, 154, 40,
    { font: 'Lato', size: 8, color: '#4a2b26' });

  pText_(slide, '·····································', 520, 198, 150, 8,
    { font: 'Arial', size: 6, color: '#f0aa97', align: SlidesApp.ParagraphAlignment.CENTER });

  var t2 = trim_(submission.threat_2) || 'Threat 2';
  pRect_(slide, 486, 206, 208, 50, '#fff7f2', '#efab99', 1, true);
  pIcon_(slide, 'thermometer', 496, 220, 24, '#ad2c23', 2.1);
  pText_(slide, 'THREAT 2 \u2013 ' + t2.toUpperCase(), 530, 212, 154, 12,
    { font: 'Montserrat', size: 8, bold: true, color: '#c53e31' });
  pText_(slide, clipText_(submission.threat_2_reason, 150), 530, 226, 154, 24,
    { font: 'Lato', size: 8, color: '#4a2b26' });
}

function posterA_drawConservation_(slide, submission) {
  pRect_(slide, 280, 270, 190, 108, POSTER_A.lavender, '#b9acd9', 1, true);
  pIconBadge_(slide, 'shield', 286, 274, 20, '#6b58a9', '#5e4b99', 1, '#f4efff', 3);
  pText_(slide, 'CONSERVATION ACTIONS', 314, 276, 150, 14,
    { font: 'Montserrat', size: 9, bold: true, color: '#4f3b8f' });

  var speciesLabel = (trim_(submission.common_name) || 'this species').toLowerCase();
  pText_(slide, 'Actions we can take to help protect ' + speciesLabel + ':', 288, 298, 170, 12,
    { font: 'Lato', size: 7, italic: true, color: '#4a3f78' });

  var a1 = clipText_(trim_(submission.action_1) || '', 28);
  var a2 = clipText_(trim_(submission.action_2) || '', 30);
  var actions = [
    { glyph: 'leaf', label: a1 || 'Protected areas' },
    { glyph: 'turtle', label: a2 || 'Reduce pollution' },
    { glyph: 'community', label: 'Support conservation education' },
    { glyph: 'scales', label: 'Enforce laws against poaching' }
  ];
  var baseX = 292;
  var cellW = 42;
  for (var i = 0; i < actions.length; i++) {
    var cx = baseX + i * cellW;
    pCircle_(slide, cx, 318, 26, '#f6f2fd', '#a391cf', 1);
    pIcon_(slide, actions[i].glyph, cx + 5, 323, 16, '#5e4b99', 1.9);
    pText_(slide, clipText_(actions[i].label, 22), cx - 5, 346, 36, 26,
      { font: 'Lato', size: 5, color: '#3f3366', align: SlidesApp.ParagraphAlignment.CENTER });
  }
}

function posterA_drawWhyMatters_(slide, submission) {
  pRect_(slide, 476, 270, 226, 108, POSTER_A.sage, '#adc4a2', 1, true);
  pIconBadge_(slide, 'leaf', 482, 274, 20, '#3a7a46', '#2f693b', 1, '#eff9ef', 3);
  pText_(slide, 'WHY THIS SPECIES MATTERS', 510, 276, 184, 14,
    { font: 'Montserrat', size: 9, bold: true, color: '#2f6e38' });

  pCircle_(slide, 492, 302, 48, '#f8fbf3', '#85ab80', 1);
  pIcon_(slide, 'turtle-coral', 502, 312, 28, '#356c3e', 1.7);

  pText_(slide, clipText_(trim_(submission.why_it_matters) || '', 220), 548, 304, 144, 66,
    { font: 'Lato', size: 8, color: '#28412a' });
}

function posterA_drawFooter_(slide, submission) {
  pRect_(slide, 12, 384, 696, 18, POSTER_A.headerBg, '#1a6556', 1, true);
  pIconBadge_(slide, 'leaf', 22, 386, 14, '#0f5448', '#8cc8a8', 1, '#dbeedc', 3);
  pDecorIcon_(slide, 'coral-decor', 640, 384, 56, 18, '#255e51');

  var speciesName = (trim_(submission.common_name) || 'this species').toLowerCase();
  var prefix = 'A healthy ecosystem depends on every species. ';
  var highlight = 'Protect ' + speciesName + ' today ';
  var tail = 'for a stronger, more resilient tomorrow.';
  pText_(slide, prefix, 42, 389, 258, 12,
    { font: 'Merriweather', size: 8, color: '#dbeedc' });
  pText_(slide, highlight, 292, 389, 178, 12,
    { font: 'Merriweather', size: 8, italic: true, bold: true, color: '#b1d26c' });
  pText_(slide, tail, 460, 389, 220, 12,
    { font: 'Merriweather', size: 8, color: '#dbeedc' });
}

/* Refined Poster A section drawers override the earlier versions above. */
function posterA_drawHeader_(slide, submission, studentName) {
  pRect_(slide, 0, 0, 720, 405, POSTER_A.bg, POSTER_A.bg, 0, false);
  pRect_(slide, 6, 4, 708, 397, '#08261d', '#124d41', 1, true);
  pRect_(slide, 10, 8, 700, 389, POSTER_A.bg, '#0f5448', 1, true);
  pRect_(slide, 12, 10, 696, 82, POSTER_A.headerBg, null, 0, true);
  pRect_(slide, 12, 88, 696, 2, POSTER_A.bgBorder, null, 0, false);

  pDecorIcon_(slide, 'waves-decor', 430, 28, 152, 46, '#2a6a5d');
  pDecorIcon_(slide, 'turtle', 520, 24, 62, 62, '#1d574b');
  pDecorIcon_(slide, 'coral-decor', 596, 18, 92, 64, '#275f52');

  pCircle_(slide, 18, 14, 52, '#0f5f53', '#cfceb2', 1.2);
  pCircle_(slide, 22, 18, 44, '#0a4338', '#cfceb2', 1);
  pCircle_(slide, 29, 25, 30, '#0f5f53', '#80b59d', 0.8);
  pText_(slide, 'ESR', 22, 24, 44, 16,
    { font: 'Montserrat', size: 12, bold: true, color: '#f4ecd6', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'CONSERVATION HQ', 22, 42, 44, 8,
    { font: 'Montserrat', size: 3, bold: true, color: '#cadb9d', align: SlidesApp.ParagraphAlignment.CENTER });

  pText_(slide, 'ENDANGERED SPECIES RESCUE REPORT', 86, 14, 430, 14,
    { font: 'Montserrat', size: 9, bold: true, color: '#cfd9a8' });
  pText_(slide, submission.common_name || 'Species Rescue', 86, 22, 430, 40,
    { font: 'Merriweather', size: 31, bold: true, color: '#ffffff' });
  pText_(slide, submission.scientific_name || '', 86, 62, 430, 18,
    { font: 'Merriweather', size: 15, italic: true, color: POSTER_A.accent });

  var attribution = 'Student: ' + studentName;
  if (trim_(submission.hour)) attribution += '    Period ' + trim_(submission.hour);
  pText_(slide, attribution, 498, 18, 196, 12,
    { font: 'Lato', size: 7, color: '#cadb9d', align: SlidesApp.ParagraphAlignment.END });
}

function posterA_drawPhotoCol_(slide, submission) {
  pRect_(slide, 16, 94, 248, 156, '#efe7d4', '#d7d7ba', 1, true);
  pRect_(slide, 20, 98, 240, 148, '#0c2f28', '#7ba394', 0.8, true);
  pRect_(slide, 24, 102, 232, 140, '#0a2620', '#efe7d4', 0.8, true);

  var facts = getPosterFactPack_(submission);
  pRect_(slide, 16, 254, 248, 128, POSTER_A.cream, '#d6d4bb', 1, true);
  pRect_(slide, 16, 254, 248, 20, '#0b5144', null, 0, true);
  pCircle_(slide, 24, 257, 14, '#2f7868', '#9ad1b0', 1);
  pText_(slide, 'Q', 24, 260, 14, 8,
    { font: 'Montserrat', size: 6, bold: true, color: '#e7f4e7', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'QUICK FACTS', 44, 258, 126, 12,
    { font: 'Montserrat', size: 9, bold: true, color: '#f2f8e6' });

  var status = trim_(submission.identified_status);
  if (status) {
    var statusTone = posterStatusPalette_(status);
    pRect_(slide, 176, 258, 78, 11, statusTone.fill, statusTone.border, 0.7, true);
    pText_(slide, status.toUpperCase(), 180, 259, 70, 8,
      { font: 'Montserrat', size: 5, bold: true, color: statusTone.text, align: SlidesApp.ParagraphAlignment.CENTER });
  }

  var rows = [
    ['leaf', 'Habitat', facts.habitat],
    ['waves', 'Diet', facts.diet],
    ['ruler', 'Length', facts.length],
    ['weight', 'Weight', facts.weight],
    ['pin', 'Location', facts.location]
  ];
  var ry = 280;
  for (var i = 0; i < rows.length; i++) {
    pCircle_(slide, 28, ry + 4, 8, '#2f7868', '#9ec8b8', 0.7);
    pText_(slide, rows[i][1], 48, ry, 72, 12,
      { font: 'Montserrat', size: 8, bold: true, color: '#174a3f' });
    pText_(slide, clipText_(rows[i][2], 38), 126, ry, 124, 12,
      { font: 'Lato', size: 8, color: POSTER_A.body });
    if (i < rows.length - 1) pLine_(slide, 28, ry + 16, 216, '#d4d0bb');
    ry += 19;
  }
}

function posterA_drawEcologyCol_(slide, submission) {
  pRect_(slide, 278, 96, 192, 170, POSTER_A.mint, '#9bc3c7', 1, true);
  pCircle_(slide, 286, 100, 22, '#0f7f79', '#0f6a63', 1);
  pText_(slide, 'E', 286, 106, 22, 10,
    { font: 'Montserrat', size: 8, bold: true, color: '#eefbf9', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'ECOLOGY', 318, 104, 140, 18,
    { font: 'Montserrat', size: 11, bold: true, color: '#166f74' });

  pRect_(slide, 288, 128, 172, 128, '#f7fbfa', '#95bcc0', 1, true);

  var roleText = trim_(submission.ecosystem_role_text) || trim_(submission.ecosystem_role) || '';
  pRect_(slide, 300, 142, 5, 24, '#1a8d86', null, 0, true);
  pText_(slide, 'ECOSYSTEM ROLE', 314, 142, 134, 12,
    { font: 'Montserrat', size: 7, bold: true, color: '#12656a' });
  pText_(slide, clipText_(roleText, 138), 314, 155, 134, 23,
    { font: 'Lato', size: 7, color: POSTER_A.body });

  pDotDivider_(slide, 314, 190, 118, '#79bac2', 2, 5);

  pRect_(slide, 300, 202, 5, 24, '#1a8d86', null, 0, true);
  pText_(slide, 'ECOLOGICAL EXPLANATION', 314, 200, 136, 12,
    { font: 'Montserrat', size: 7, bold: true, color: '#12656a' });
  pText_(slide, clipText_(submission.ecology_explanation, 138), 314, 217, 134, 29,
    { font: 'Lato', size: 7, color: POSTER_A.body });
}

function posterA_drawThreatsCol_(slide, submission) {
  pRect_(slide, 476, 96, 226, 170, POSTER_A.blush, '#efb7a7', 1, true);
  pCircle_(slide, 482, 100, 22, '#ec856d', '#df765f', 1);
  pText_(slide, '!', 482, 104, 22, 10,
    { font: 'Montserrat', size: 9, bold: true, color: '#fff5f2', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'THREATS', 514, 104, 166, 18,
    { font: 'Montserrat', size: 11, bold: true, color: '#c43d2f' });

  var t1 = trim_(submission.threat_1) || 'Threat 1';
  var t2 = trim_(submission.threat_2) || 'Threat 2';
  pRect_(slide, 486, 128, 206, 128, '#fff8f4', '#efab99', 1, true);

  pRect_(slide, 498, 142, 6, 24, '#e46e58', null, 0, true);
  pText_(slide, 'THREAT 1 - ' + posterThreatLabelShort_(t1), 514, 142, 166, 12,
    { font: 'Montserrat', size: 7, bold: true, color: '#c53e31' });
  pText_(slide, clipText_(submission.threat_1_reason, 138), 514, 156, 166, 28,
    { font: 'Lato', size: 7, color: '#4a2b26' });

  pDotDivider_(slide, 514, 194, 146, '#f0aa97', 2, 5);

  pRect_(slide, 498, 206, 6, 24, '#e46e58', null, 0, true);
  pText_(slide, 'THREAT 2 - ' + posterThreatLabelShort_(t2), 514, 206, 166, 12,
    { font: 'Montserrat', size: 7, bold: true, color: '#c53e31' });
  pText_(slide, clipText_(submission.threat_2_reason, 118), 514, 220, 166, 24,
    { font: 'Lato', size: 7, color: '#4a2b26' });
}

function posterA_drawConservation_(slide, submission) {
  pRect_(slide, 278, 270, 192, 108, POSTER_A.lavender, '#b9acd9', 1, true);
  pCircle_(slide, 286, 274, 20, '#6b58a9', '#5e4b99', 1);
  pText_(slide, 'C', 286, 279, 20, 10,
    { font: 'Montserrat', size: 8, bold: true, color: '#f4efff', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'CONSERVATION ACTIONS', 314, 276, 150, 14,
    { font: 'Montserrat', size: 8, bold: true, color: '#4f3b8f' });

  var speciesLabel = (trim_(submission.common_name) || 'this species').toLowerCase();
  pText_(slide, 'Actions we can take to help protect ' + speciesLabel + ':', 288, 292, 172, 12,
    { font: 'Lato', size: 6, color: '#4a3f78' });

  var a1 = clipText_(trim_(submission.action_1) || '', 30);
  var a2 = clipText_(trim_(submission.action_2) || '', 30);
  var actions = [
    { tone: '#6d59af', label: posterActionLabelShort_(a1 || 'Protected areas', 'Protected Areas') },
    { tone: '#7b67bb', label: posterActionLabelShort_(a2 || 'Reduce pollution', 'Reduce Pollution') },
    { tone: '#846fc6', label: 'Community Support' },
    { tone: '#8c77cf', label: 'Enforce Laws' }
  ];
  var positions = [
    { x: 288, y: 312 },
    { x: 376, y: 312 },
    { x: 288, y: 340 },
    { x: 376, y: 340 }
  ];
  for (var i = 0; i < actions.length; i++) {
    var box = positions[i];
    pRect_(slide, box.x, box.y, 78, 20, '#fbf9ff', '#a391cf', 1, true);
    pRect_(slide, box.x + 6, box.y + 4, 4, 12, actions[i].tone, null, 0, true);
    pText_(slide, actions[i].label, box.x + 16, box.y + 5, 56, 9,
      { font: 'Montserrat', size: 5, bold: true, color: '#4f3b8f', align: SlidesApp.ParagraphAlignment.CENTER });
  }
}

function posterA_drawWhyMatters_(slide, submission) {
  pRect_(slide, 476, 270, 226, 108, POSTER_A.sage, '#adc4a2', 1, true);
  pCircle_(slide, 482, 274, 20, '#3a7a46', '#2f693b', 1);
  pText_(slide, 'W', 482, 279, 20, 10,
    { font: 'Montserrat', size: 8, bold: true, color: '#eff9ef', align: SlidesApp.ParagraphAlignment.CENTER });
  pText_(slide, 'WHY THIS SPECIES MATTERS', 510, 276, 184, 14,
    { font: 'Montserrat', size: 8, bold: true, color: '#2f6e38' });

  pRect_(slide, 486, 302, 206, 62, '#f8fbf3', '#9fc297', 1, true);
  pRect_(slide, 498, 314, 6, 34, '#5b8a55', null, 0, true);
  pText_(slide, clipText_(trim_(submission.why_it_matters) || '', 160), 514, 314, 168, 34,
    { font: 'Lato', size: 7, color: '#28412a' });
}

function posterA_drawFooter_(slide, submission) {
  pRect_(slide, 12, 382, 696, 20, POSTER_A.headerBg, '#1a6556', 1, true);
  pCircle_(slide, 22, 384, 15, '#0f5448', '#8cc8a8', 1);
  pText_(slide, 'R', 22, 387, 15, 8,
    { font: 'Montserrat', size: 6, bold: true, color: '#dbeedc', align: SlidesApp.ParagraphAlignment.CENTER });
  pDecorIcon_(slide, 'coral-decor', 638, 382, 58, 20, '#255e51');

  var speciesName = (trim_(submission.common_name) || 'this species').toLowerCase();
  var prefix = 'A healthy ecosystem depends on every species. ';
  var highlight = 'Protect ' + speciesName + ' today ';
  var tail = 'for a stronger, more resilient tomorrow.';
  pText_(slide, prefix, 44, 386, 238, 14,
    { font: 'Merriweather', size: 7, color: '#dbeedc' });
  pText_(slide, highlight, 282, 386, 180, 14,
    { font: 'Merriweather', size: 7, italic: true, bold: true, color: '#b1d26c' });
  pText_(slide, tail, 448, 386, 228, 14,
    { font: 'Merriweather', size: 7, color: '#dbeedc' });
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
      insertSlideImageIntoBox_(slide, imageUrl, 8, 10, 704, 280);
    } catch (imgErr) {
      Logger.log('tileC image insert failed: ' + imgErr);
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
  var canUseTemplate = shouldUseTemplatePoster_(settings, templateId);

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

  try {
    posterFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (pdfFile) pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareError) {
    Logger.log('Sharing update failed: ' + shareError);
  }

  return {
    slideUrl: posterFile.getUrl(),
    pdfUrl: pdfFile ? pdfFile.getUrl() : '',
    posterFileId: posterFile.getId(),
    pdfFileId: pdfFile ? pdfFile.getId() : '',
    warning: warnings.join(' '),
    studentCanViewFiles: !!pdfFile
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

  var bg = hexToRgb_(POSTER_A.bg);
  slide.getBackground().setSolidFill(bg.r, bg.g, bg.b);

  var studentName = ([submission.first_name || '', submission.last_name || '']).join(' ').trim();

  var sections = [
    ['header',       function() { posterA_drawHeader_(slide, submission, studentName); }],
    ['photoCol',     function() { posterA_drawPhotoCol_(slide, submission); }],
    ['ecologyCol',   function() { posterA_drawEcologyCol_(slide, submission); }],
    ['threatsCol',   function() { posterA_drawThreatsCol_(slide, submission); }],
    ['conservation', function() { posterA_drawConservation_(slide, submission); }],
    ['whyMatters',   function() { posterA_drawWhyMatters_(slide, submission); }],
    ['footer',       function() { posterA_drawFooter_(slide, submission); }]
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

function shouldUseTemplatePoster_(settings, templateId) {
  var mode = trim_(settings.poster_render_mode).toLowerCase();
  if (mode !== 'template') return false;
  return hasConfiguredDriveId_(templateId) && isTrueSetting_(settings.use_template_poster);
}

function insertSlideImageFromUrl_(slide, imageUrl) {
  return insertSlideImageIntoBox_(
    slide,
    imageUrl,
    APP.posterImageBox.x,
    APP.posterImageBox.y,
    APP.posterImageBox.width,
    APP.posterImageBox.height
  );
}

function insertSlideImageIntoBox_(slide, imageUrl, x, y, width, height) {
  if (!imageUrl) return false;
  var blob = getImageBlobFromSource_(imageUrl);
  if (blob) {
    var blobImage = slide.insertImage(blob, x, y, width, height);
    try { blobImage.replace(blob, true); } catch (blobCropErr) {
      Logger.log('Image crop replace failed for blob: ' + blobCropErr);
    }
    return true;
  }

  var fallbackUrl = normalizeImageUrlForBrowser_(imageUrl) || trim_(imageUrl);
  try {
    var remoteImage = slide.insertImage(fallbackUrl, x, y, width, height);
    try { remoteImage.replace(fallbackUrl, true); } catch (urlCropErr) {
      Logger.log('Image crop replace failed for URL: ' + urlCropErr);
    }
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

function driveThumbnailUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1600';
}

function extractDriveFileIdFromUrl_(url) {
  var text = trim_(url);
  if (!text) return '';
  var patterns = [
    /[?&]id=([a-zA-Z0-9_-]{10,})/i,
    /\/d\/([a-zA-Z0-9_-]{10,})/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match && match[1]) return match[1];
  }
  return '';
}

function extractWikimediaFileName_(url) {
  var text = trim_(url);
  if (!text) return '';
  var match = text.match(/(?:commons\.wikimedia|[a-z-]+\.wikipedia)\.org\/wiki\/(?:Special:FilePath\/|File:)([^?#]+)/i);
  if (!match || !match[1]) return '';
  try {
    return decodeURIComponent(match[1]).replace(/_/g, ' ');
  } catch (e) {
    return match[1].replace(/_/g, ' ');
  }
}

function normalizeImageUrlForBrowser_(url) {
  var text = trim_(url);
  if (!text) return '';
  if (/^data:image\//i.test(text)) return text;

  var driveFileId = extractDriveFileIdFromUrl_(text);
  if (driveFileId) return driveThumbnailUrl_(driveFileId);

  var wikiFileName = extractWikimediaFileName_(text);
  if (wikiFileName) return wikiFileUrl_(wikiFileName);

  if (/dropbox\.com/i.test(text)) {
    if (/[?&]raw=1\b/i.test(text)) return text;
    if (/[?&]dl=0\b/i.test(text)) {
      return text.replace(/([?&])dl=0\b/i, '$1raw=1');
    }
    return text + (text.indexOf('?') === -1 ? '?raw=1' : '&raw=1');
  }

  return text;
}

function getImageBlobFromDriveFileId_(fileId) {
  if (!fileId) return null;
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    return isImageBlob_(blob) ? blob : null;
  } catch (error) {
    Logger.log('Drive image blob lookup failed for file ID ' + fileId + ': ' + error);
    return null;
  }
}

function fetchImageBlobFromHttpUrl_(url) {
  if (!/^https?:\/\//i.test(trim_(url))) return null;
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return null;

    var blob = response.getBlob();
    return isImageBlob_(blob) ? blob : null;
  } catch (error) {
    Logger.log('HTTP image fetch failed for URL: ' + url + ' :: ' + error);
    return null;
  }
}

function isImageBlob_(blob) {
  if (!blob) return false;
  var contentType = trim_(blob.getContentType() || '').toLowerCase();
  return !contentType || contentType.indexOf('image/') === 0;
}

function getImageBlobFromSource_(url) {
  var text = trim_(url);
  if (!text) return null;
  if (/^data:image\//i.test(text)) return null;

  var driveFileId = extractDriveFileIdFromUrl_(text);
  if (driveFileId) {
    var driveBlob = getImageBlobFromDriveFileId_(driveFileId);
    if (driveBlob) return driveBlob;
  }

  var normalizedUrl = normalizeImageUrlForBrowser_(text);
  var normalizedBlob = fetchImageBlobFromHttpUrl_(normalizedUrl);
  if (normalizedBlob) return normalizedBlob;

  if (normalizedUrl !== text) return fetchImageBlobFromHttpUrl_(text);
  return null;
}

function fetchImagePreviewDataUri_(url) {
  var normalizedUrl = normalizeImageUrlForBrowser_(url);
  if (!normalizedUrl) return '';
  if (/^data:image\//i.test(normalizedUrl)) return normalizedUrl;

  var blob = getImageBlobFromSource_(url);
  if (!blob) return '';

  var contentType = trim_(blob.getContentType() || '') || 'image/jpeg';
  return 'data:' + contentType + ';base64,' + Utilities.base64Encode(blob.getBytes());
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
  var text = normalizeImageUrlForBrowser_(url);
  if (!text) return false;
  if (shouldBackfillImageValue_(text)) return false;
  return /^https?:\/\//i.test(text);
}

function resolveSpeciesImageSet_(speciesId, commonName, rawUrls) {
  var validRaw = rawUrls
    .map(function(u) { return normalizeImageUrlForBrowser_(u); })
    .filter(function(u) { return isUsableSpeciesImageUrl_(u); })
    .map(function(u) { return trim_(u); });
  while (validRaw.length && validRaw.length < 3) validRaw.push(validRaw[validRaw.length - 1]);
  if (validRaw.length >= 3) return validRaw.slice(0, 3);

  var curated = getCuratedImageSetForSpecies_(speciesId).map(function(u) { return normalizeImageUrlForBrowser_(u); });
  while (curated.length && curated.length < 3) curated.push(curated[curated.length - 1]);
  if (curated.length >= 3) return curated.slice(0, 3);

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
    var resolvedImages = resolveSpeciesImageSet_(rowObject.species_id, rowObject.common_name, [
      rowObject.image_option_1_url,
      rowObject.image_option_2_url,
      rowObject.image_option_3_url
    ]);
    rowObject.image_option_1_url = resolvedImages[0];
    rowObject.image_option_2_url = resolvedImages[1];
    rowObject.image_option_3_url = resolvedImages[2];
    appendObjectRow_(sheet, headers, rowObject);
  }
}

function patchMissingSpeciesImageData_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(APP.sheetNames.species);
  var data = getObjectsFromSheet_(sheet);
  if (!data.length) return;

  var imageColumnStart = getHeaders_(sheet).indexOf('image_option_1_url') + 1;
  if (!imageColumnStart) return;

  var pendingUpdates = [];
  for (var i = 0; i < data.length; i++) {
    if (!trim_(data[i].species_id)) continue;
    var resolvedImages = resolveSpeciesImageSet_(data[i].species_id, data[i].common_name, [
      data[i].image_option_1_url, data[i].image_option_2_url, data[i].image_option_3_url
    ]);
    var currentImages = [
      trim_(data[i].image_option_1_url),
      trim_(data[i].image_option_2_url),
      trim_(data[i].image_option_3_url)
    ];
    if (currentImages.join('\n') !== resolvedImages.join('\n')) {
      pendingUpdates.push({ row: i + 2, values: [resolvedImages] });
    }
  }

  for (var u = 0; u < pendingUpdates.length; u++) {
    sheet.getRange(pendingUpdates[u].row, imageColumnStart, 1, 3).setValues(pendingUpdates[u].values);
  }
  if (pendingUpdates.length) SpreadsheetApp.flush();
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
      var resolvedImages = resolveSpeciesImageSet_(
        rows[i].species_id,
        rows[i].common_name || base.common_name || '',
        [
          rows[i].image_option_1_url || base.image_option_1_url || '',
          rows[i].image_option_2_url || base.image_option_2_url || '',
          rows[i].image_option_3_url || base.image_option_3_url || ''
        ]
      );
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
        image_option_1_url: resolvedImages[0],
        image_option_2_url: resolvedImages[1],
        image_option_3_url: resolvedImages[2],
        research_link_1: rows[i].research_link_1 || base.research_link_1 || '',
        research_link_2: rows[i].research_link_2 || base.research_link_2 || ''
      });
    }
  }

  if (!out.length) {
    for (var j = 0; j < defaults.length; j++) {
      var fi = resolveSpeciesImageSet_(defaults[j].species_id, defaults[j].common_name, [
        defaults[j].image_option_1_url,
        defaults[j].image_option_2_url,
        defaults[j].image_option_3_url
      ]);
      out.push(mergeObjects_(defaults[j], {
        image_option_1_url: fi[0],
        image_option_2_url: fi[1],
        image_option_3_url: fi[2]
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
