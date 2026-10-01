/**
 * Open House Needs Sheet
 *
 * One roster of people, one set of email templates, and a menu that turns
 * the two into Gmail drafts. Nothing sends automatically -- every email
 * lands in your Gmail Drafts folder for you to read and send.
 */

var ROSTER = 'Roster';
var TEMPLATES = 'Templates';
var GROUPS = ['Parent', 'Student', 'Musician', 'Teacher'];

var ROSTER_HEADERS = ['Name', 'Email', 'Group', 'Notes', 'Emailed?'];
var TEMPLATE_HEADERS = ['Group', 'Subject', 'Body'];

// Starter wording, written into the Templates tab by "Set up / repair sheet".
// Edit it on the tab, not here -- the tab is what the drafts are built from.
var TEMPLATE_SEED = {
  Parent: {
    subject: "Open House - we'd love to see you",
    body: "Hi {{name}},\n\nWe're hosting our Open House and would love to have you there.\n\nDate:\nTime:\nLocation:\n\nPlease let me know if you can make it.\n\nThank you,\n{{sender}}"
  },
  Student: {
    subject: "Open House - you're invited",
    body: "Hi {{name}},\n\nYou're invited to our Open House. Come see the work we've been doing.\n\nDate:\nTime:\nLocation:\n\nSee you there,\n{{sender}}"
  },
  Musician: {
    subject: 'Open House - performance details',
    body: "Hi {{name}},\n\nThank you for playing at our Open House. Here are the details:\n\nDate:\nTime:\nCall time:\nLocation:\nWhat to bring:\n\nLet me know if you have any questions.\n\nThank you,\n{{sender}}"
  },
  Teacher: {
    subject: 'Open House - setup and staffing',
    body: "Hi {{name}},\n\nHere are the Open House details and what we need from you:\n\nDate:\nTime:\nYour station:\nSetup time:\n\nThanks for pitching in.\n\n{{sender}}"
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Open House')
    .addItem('Set up / repair sheet', 'setUpSheet')
    .addSeparator()
    .addItem('Create a draft for each person...', 'createIndividualDrafts')
    .addItem('Create one group draft (BCC)...', 'createGroupDraft')
    .addSeparator()
    .addItem('Copy email addresses...', 'showEmailList')
    .addItem('Check for problems', 'validateRoster')
    .addToUi();
}

/* ------------------------------------------------------------------ setup */

function setUpSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var roster = sheetNamed(ss, ROSTER, ROSTER_HEADERS);
  var templates = sheetNamed(ss, TEMPLATES, TEMPLATE_HEADERS);

  // Group column becomes a dropdown so names stay consistent.
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(GROUPS, true)
    .setAllowInvalid(false)
    .setHelpText('Pick one: ' + GROUPS.join(', '))
    .build();
  roster.getRange(2, 3, Math.max(roster.getMaxRows() - 1, 1), 1).setDataValidation(rule);

  roster.setColumnWidth(1, 180);
  roster.setColumnWidth(2, 240);
  roster.setColumnWidth(4, 260);

  // One example row per group on a brand new sheet, so the shape is obvious.
  if (roster.getLastRow() < 2) {
    roster.getRange(2, 1, 4, 4).setValues([
      ['Jane Doe', 'jane.doe@example.com', 'Parent', 'Example row - delete me'],
      ['Sam Doe', 'sam.doe@example.com', 'Student', 'Example row - delete me'],
      ['Alex Rivera', 'alex.rivera@example.com', 'Musician', 'Cello'],
      ['Pat Chen', 'pat.chen@example.com', 'Teacher', 'Grade 4']
    ]);
  }

  // Seed a template row for any group that doesn't have one yet.
  var have = {};
  readRows(templates).forEach(function (r) { have[String(r[0]).trim()] = true; });
  GROUPS.forEach(function (g) {
    if (!have[g]) {
      templates.appendRow([g, TEMPLATE_SEED[g].subject, TEMPLATE_SEED[g].body]);
    }
  });
  templates.getRange(2, 3, Math.max(templates.getLastRow() - 1, 1), 1).setWrap(true);
  templates.setColumnWidth(2, 280);
  templates.setColumnWidth(3, 480);

  SpreadsheetApp.getUi().alert('Sheet is ready.\n\nFill in the Roster tab, edit the wording on the Templates tab, then use the Open House menu.');
}

function sheetNamed(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var head = sheet.getRange(1, 1, 1, headers.length);
  head.setValues([headers]).setFontWeight('bold').setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
  return sheet;
}

/* ------------------------------------------------------------------ email */

function createIndividualDrafts() {
  var group = askForGroup('Create a draft for each person');
  if (!group) return;

  var people = peopleIn(group);
  if (!people.length) {
    alert('No one on the Roster tab is in the group "' + group + '" with an email address.');
    return;
  }

  var template = templateFor(group);
  if (!template) {
    alert('No template for "' + group + '" on the Templates tab. Run "Set up / repair sheet" to add one.');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    'Create ' + people.length + ' draft' + (people.length === 1 ? '' : 's') + '?',
    people.length + ' draft(s) will appear in your Gmail Drafts folder. Nothing is sent.',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  var sender = senderName();
  people.forEach(function (person) {
    GmailApp.createDraft(person.email, fill(template.subject, person, sender), fill(template.body, person, sender));
    markEmailed(person.row);
  });

  alert(people.length + ' draft(s) created in Gmail. Open Drafts to review and send.');
}

function createGroupDraft() {
  var group = askForGroup('Create one group draft');
  if (!group) return;

  var people = peopleIn(group);
  if (!people.length) {
    alert('No one on the Roster tab is in the group "' + group + '" with an email address.');
    return;
  }

  var template = templateFor(group);
  if (!template) {
    alert('No template for "' + group + '" on the Templates tab. Run "Set up / repair sheet" to add one.');
    return;
  }

  var sender = senderName();
  // No {{name}} to substitute in a group mail, so greet the group instead.
  var placeholder = { name: group === 'Parent' ? 'everyone' : 'all', email: '' };

  // Gmail will not accept a draft with an empty To field, so the draft is
  // addressed to you and the group goes in BCC -- nobody sees anyone else.
  var me = Session.getEffectiveUser().getEmail();
  if (!me) {
    alert('Could not work out your own email address, which the group draft needs in the To field. Use "Create a draft for each person" instead.');
    return;
  }

  GmailApp.createDraft(me, fill(template.subject, placeholder, sender), fill(template.body, placeholder, sender), {
    bcc: people.map(function (p) { return p.email; }).join(',')
  });

  people.forEach(function (p) { markEmailed(p.row); });
  alert('One draft created: addressed to you, with ' + people.length + ' BCC recipient(s).\n\nReview it in Gmail before sending.');
}

function showEmailList() {
  var group = askForGroup('Copy email addresses');
  if (!group) return;

  var people = peopleIn(group);
  if (!people.length) {
    alert('No one on the Roster tab is in the group "' + group + '" with an email address.');
    return;
  }

  var list = people.map(function (p) { return p.email; }).join(', ');
  var html = HtmlService
    .createHtmlOutput('<p style="font:13px Arial">' + people.length + ' address(es) for <b>' + group +
      '</b>. Select all and copy:</p><textarea style="width:100%;height:220px;font:12px monospace">' +
      list.replace(/&/g, '&amp;').replace(/</g, '&lt;') + '</textarea>')
    .setWidth(460).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, group + ' emails');
}

/* ------------------------------------------------------------------ checks */

function validateRoster() {
  var problems = [];
  var seen = {};

  readRows(rosterSheet()).forEach(function (row, i) {
    var line = i + 2;
    var name = String(row[0] || '').trim();
    var email = String(row[1] || '').trim();
    var group = String(row[2] || '').trim();

    if (!name && !email && !group) return; // blank row, ignore
    if (!name) problems.push('Row ' + line + ': no name.');
    if (!email) problems.push('Row ' + line + ': no email address.');
    else if (!/^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(email)) problems.push('Row ' + line + ': "' + email + '" does not look like an email address.');
    else if (seen[email.toLowerCase()]) problems.push('Row ' + line + ': ' + email + ' is already on row ' + seen[email.toLowerCase()] + '.');
    else seen[email.toLowerCase()] = line;

    if (!group) problems.push('Row ' + line + ': no group.');
    else if (GROUPS.indexOf(group) === -1) problems.push('Row ' + line + ': group "' + group + '" is not one of ' + GROUPS.join(', ') + '.');
  });

  alert(problems.length ? problems.join('\n') : 'No problems found.');
}

/* ------------------------------------------------------------------ guts */

function rosterSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROSTER);
  if (!sheet) throw new Error('No "' + ROSTER + '" tab. Run "Set up / repair sheet" first.');
  return sheet;
}

function readRows(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

function peopleIn(group) {
  var people = [];
  readRows(rosterSheet()).forEach(function (row, i) {
    var email = String(row[1] || '').trim();
    if (!email) return;
    if (String(row[2] || '').trim().toLowerCase() !== group.toLowerCase()) return;
    people.push({ name: String(row[0] || '').trim() || 'there', email: email, row: i + 2 });
  });
  return people;
}

function templateFor(group) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATES);
  if (!sheet) return null;
  var match = null;
  readRows(sheet).forEach(function (row) {
    if (String(row[0] || '').trim().toLowerCase() === group.toLowerCase()) {
      match = { subject: String(row[1] || 'Open House'), body: String(row[2] || '') };
    }
  });
  return match;
}

function fill(text, person, sender) {
  return String(text)
    .replace(/\{\{\s*name\s*\}\}/gi, person.name)
    .replace(/\{\{\s*email\s*\}\}/gi, person.email)
    .replace(/\{\{\s*sender\s*\}\}/gi, sender);
}

function senderName() {
  var saved = PropertiesService.getDocumentProperties().getProperty('senderName');
  if (saved) return saved;

  var ui = SpreadsheetApp.getUi();
  var answer = ui.prompt('Sign your emails', 'What name should {{sender}} be replaced with?', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return '';
  var name = answer.getResponseText().trim();
  if (name) PropertiesService.getDocumentProperties().setProperty('senderName', name);
  return name;
}

function askForGroup(title) {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.prompt(title, 'Which group? ' + GROUPS.join(' / '), ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return null;

  var typed = answer.getResponseText().trim();
  for (var i = 0; i < GROUPS.length; i++) {
    if (GROUPS[i].toLowerCase() === typed.toLowerCase() ||
        GROUPS[i].toLowerCase() + 's' === typed.toLowerCase()) {
      return GROUPS[i];
    }
  }
  alert('"' + typed + '" is not a group. Use one of: ' + GROUPS.join(', '));
  return null;
}

function markEmailed(row) {
  rosterSheet().getRange(row, 5).setValue(new Date());
}

function alert(message) {
  SpreadsheetApp.getUi().alert(message);
}
