/**
 * Open House Needs Sheet
 *
 * Five tabs -- Roster, Team, Tasks, Rooms, Templates -- and an "Open House"
 * menu that turns them into email.
 *
 * Two different things happen, on purpose:
 *   - Emails to parents, students, musicians and teachers become Gmail DRAFTS.
 *     You read them and press Send yourself. Nothing goes out on its own.
 *   - Task reminders to your own team SEND immediately, because that is the
 *     point of a reminder, and because the optional daily reminder runs while
 *     you are not at the computer.
 */

var ROSTER = 'Roster';
var TEAM = 'Team';
var TASKS = 'Tasks';
var ROOMS = 'Rooms';
var TEMPLATES = 'Templates';

var GROUPS = ['Parent', 'Student', 'Musician', 'Teacher'];
var STATUSES = ['Not started', 'In progress', 'Blocked', 'Done'];

var ROSTER_HEADERS = ['Name', 'Email', 'Group', 'Notes', 'Emailed?'];
var TEAM_HEADERS = ['Name', 'Email', 'Role'];
var TASK_HEADERS = ['Task', 'Details', 'Owner', 'Due date', 'Status', 'Last reminded'];
var ROOM_HEADERS = ['Room', 'Location', 'Activity', 'Teachers', 'Students', 'Notes'];
var TEMPLATE_HEADERS = ['Group', 'Subject', 'Body'];

// "Combined" is the wording used when one email goes to more than one group.
var TEMPLATE_KEYS = GROUPS.concat(['Combined']);

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
    body: "Hi {{name}},\n\nHere are the Open House details and what we need from you:\n\nDate:\nTime:\nYour room: {{room}}\nSetup time:\n\nThanks for pitching in.\n\n{{sender}}"
  },
  Combined: {
    subject: 'Open House - details',
    body: "Hi {{name}},\n\nHere are the details for our Open House.\n\nDate:\nTime:\nLocation:\n\nWe hope to see you there.\n\nThank you,\n{{sender}}"
  }
};

/* ------------------------------------------------------------------- menu */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Open House')
    .addItem('Set up / repair sheet', 'setUpSheet')
    .addSeparator()
    .addItem('Send emails...', 'showSendDialog')
    .addItem('Email room assignments to teachers', 'emailRoomAssignments')
    .addSeparator()
    .addItem('Send task reminders now', 'sendTaskRemindersNow')
    .addItem('Turn on daily task reminders', 'enableDailyReminders')
    .addItem('Turn off daily task reminders', 'disableDailyReminders')
    .addSeparator()
    .addItem('Check for problems', 'checkForProblems')
    .addToUi();
}

/* ------------------------------------------------------------------ setup */

function setUpSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var roster = tab(ss, ROSTER, ROSTER_HEADERS);
  dropdown(roster, 3, GROUPS, 'Pick one: ' + GROUPS.join(', '));
  widths(roster, [180, 240, 110, 260, 130]);
  if (roster.getLastRow() < 2) {
    roster.getRange(2, 1, 4, 4).setValues([
      ['Jane Doe', 'jane.doe@example.com', 'Parent', 'Example row - delete me'],
      ['Sam Doe', 'sam.doe@example.com', 'Student', 'Example row - delete me'],
      ['Alex Rivera', 'alex.rivera@example.com', 'Musician', 'Cello'],
      ['Pat Chen', 'pat.chen@example.com', 'Teacher', 'Grade 4']
    ]);
  }

  var team = tab(ss, TEAM, TEAM_HEADERS);
  widths(team, [180, 240, 200]);
  if (team.getLastRow() < 2) {
    team.getRange(2, 1, 1, 3).setValues([[me() || 'You', me() || 'you@example.com', 'Organizer']]);
  }

  var tasks = tab(ss, TASKS, TASK_HEADERS);
  dropdown(tasks, 5, STATUSES, 'Pick one: ' + STATUSES.join(', '));
  widths(tasks, [260, 300, 160, 110, 120, 140]);
  tasks.getRange(2, 4, Math.max(tasks.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  // Owner is a dropdown of whoever is on the Team tab.
  var teamNames = readTab(TEAM).map(function (r) { return r.Name; }).filter(String);
  if (teamNames.length) {
    dropdown(tasks, 3, teamNames, 'Someone from the Team tab');
  }
  if (tasks.getLastRow() < 2) {
    tasks.getRange(2, 1, 2, 5).setValues([
      ['Book the auditorium', 'Confirm with the front office', teamNames[0] || '', '', 'Not started'],
      ['Print programs', '50 copies, double sided', teamNames[0] || '', '', 'Not started']
    ]);
  }

  var rooms = tab(ss, ROOMS, ROOM_HEADERS);
  widths(rooms, [120, 160, 200, 240, 300, 220]);
  rooms.getRange(2, 4, Math.max(rooms.getMaxRows() - 1, 1), 2).setWrap(true);
  if (rooms.getLastRow() < 2) {
    rooms.getRange(2, 1, 1, 6).setValues([
      ['101', 'First floor', 'Science projects', 'Pat Chen', 'Sam Doe', 'Example row - delete me']
    ]);
  }

  var templates = tab(ss, TEMPLATES, TEMPLATE_HEADERS);
  var have = {};
  readTab(TEMPLATES).forEach(function (r) { have[String(r.Group).trim()] = true; });
  TEMPLATE_KEYS.forEach(function (key) {
    if (!have[key]) {
      templates.appendRow([key, TEMPLATE_SEED[key].subject, TEMPLATE_SEED[key].body]);
    }
  });
  widths(templates, [120, 300, 520]);
  templates.getRange(2, 3, Math.max(templates.getLastRow() - 1, 1), 1).setWrap(true);

  alert('Sheet is ready.\n\nFill in the Roster, Team, Tasks and Rooms tabs, edit the wording on Templates, then use the Open House menu.\n\nRe-run this any time you add people to the Team tab, so the Tasks owner dropdown picks them up.');
}

function tab(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers]).setFontWeight('bold').setBackground('#f1f3f4');
  sheet.setFrozenRows(1);
  return sheet;
}

function dropdown(sheet, column, values, help) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).setHelpText(help).build();
  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function widths(sheet, list) {
  for (var i = 0; i < list.length; i++) sheet.setColumnWidth(i + 1, list[i]);
}

/* ------------------------------------------------------- reading the tabs */

/**
 * Reads a tab into objects keyed by its own header row, so reordering or
 * renaming a column in the sheet doesn't break the script. Each object also
 * carries _row, the sheet row it came from.
 */
function readTab(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var width = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();

  var out = [];
  rows.forEach(function (row, i) {
    var blank = row.every(function (c) { return c === '' || c === null; });
    if (blank) return;
    var obj = { _row: i + 2 };
    headers.forEach(function (h, c) {
      var key = String(h).trim();
      if (key) obj[key] = row[c] instanceof Date ? row[c] : String(row[c]).trim();
    });
    out.push(obj);
  });
  return out;
}

function sheetFor(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('No "' + name + '" tab. Run "Set up / repair sheet" first.');
  return sheet;
}

/* --------------------------------------------- recipient group combinations */

/**
 * Every non-empty combination of the groups: each on its own, each pair, each
 * trio, and all of them. Four groups gives 15 options, ordered shortest first.
 */
function recipientCombinations() {
  var combos = [];
  for (var mask = 1; mask < (1 << GROUPS.length); mask++) {
    var picked = [];
    for (var i = 0; i < GROUPS.length; i++) {
      if (mask & (1 << i)) picked.push(GROUPS[i]);
    }
    combos.push(picked);
  }

  combos.sort(function (a, b) {
    if (a.length !== b.length) return a.length - b.length;
    for (var i = 0; i < a.length; i++) {
      var d = GROUPS.indexOf(a[i]) - GROUPS.indexOf(b[i]);
      if (d) return d;
    }
    return 0;
  });

  return combos.map(function (groups) {
    return {
      key: groups.join('|'),
      groups: groups,
      label: groups.length === GROUPS.length
        ? 'Everyone (' + groups.map(plural).join(' + ') + ')'
        : groups.map(plural).join(' + ')
    };
  });
}

function plural(group) {
  return group + 's';
}

/** Everyone in any of the given groups, one entry per email address. */
function peopleIn(groups) {
  var wanted = {};
  groups.forEach(function (g) { wanted[g.toLowerCase()] = true; });

  var seen = {};
  var people = [];
  readTab(ROSTER).forEach(function (r) {
    var email = String(r.Email || '').trim();
    var group = String(r.Group || '').trim();
    if (!email || !wanted[group.toLowerCase()]) return;

    var key = email.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;

    people.push({
      name: String(r.Name || '').trim() || 'there',
      email: email,
      group: group,
      row: r._row
    });
  });
  return people;
}

/* ------------------------------------------------------------ send dialog */

function showSendDialog() {
  var combos = recipientCombinations();
  var counts = combos.map(function (c) { return peopleIn(c.groups).length; });

  var recipientOptions = combos.map(function (c, i) {
    return '<option value="' + esc(c.key) + '">' + esc(c.label) +
      ' — ' + counts[i] + ' ' + (counts[i] === 1 ? 'person' : 'people') + '</option>';
  }).join('');

  var wordingOptions = ['<option value="__own__">Each person’s own group wording</option>']
    .concat(readTab(TEMPLATES).map(function (t) {
      return '<option value="' + esc(t.Group) + '">' + esc(t.Group) + ' wording</option>';
    })).join('');

  var html = [
    '<style>',
    'body{font:13px/1.5 Arial,sans-serif;margin:0;padding:14px;color:#202124}',
    'label{display:block;font-weight:bold;margin:12px 0 4px}',
    'select{width:100%;padding:6px;font:13px Arial}',
    '.btns{margin-top:18px;display:flex;gap:8px;flex-wrap:wrap}',
    'button{padding:8px 12px;font:13px Arial;cursor:pointer}',
    '#status{margin-top:14px;min-height:34px;white-space:pre-wrap}',
    '#out{width:100%;height:110px;font:12px monospace;margin-top:8px;display:none}',
    '.hint{color:#5f6368;font-size:12px;margin-top:4px}',
    '</style>',
    '<label for="groups">Send to</label>',
    '<select id="groups" onchange="syncWording()">', recipientOptions, '</select>',
    '<label for="wording">Wording</label>',
    '<select id="wording">', wordingOptions, '</select>',
    '<div class="hint">A BCC draft goes to several people at once, so it always uses one shared wording.</div>',
    '<div class="btns">',
    '<button onclick="go(\'each\')">One draft per person</button>',
    '<button onclick="go(\'bcc\')">One BCC draft</button>',
    '<button onclick="go(\'list\')">Just show addresses</button>',
    '</div>',
    '<div id="status"></div>',
    '<textarea id="out" readonly></textarea>',
    '<script>',
    'function sel(id){return document.getElementById(id);}',
    'function buttons(disabled){',
    '  var b=document.getElementsByTagName("button");',
    '  for(var i=0;i<b.length;i++)b[i].disabled=disabled;',
    '}',
    // One group selected -> default to that group's own wording.
    'function syncWording(){',
    '  var picked=sel("groups").value.split("|");',
    '  var w=sel("wording");',
    '  var want=picked.length===1?picked[0]:"__own__";',
    '  for(var i=0;i<w.options.length;i++){if(w.options[i].value===want){w.selectedIndex=i;return;}}',
    '}',
    'function go(mode){',
    '  sel("status").textContent="Working…";',
    '  sel("out").style.display="none";',
    '  buttons(true);',
    '  google.script.run.withSuccessHandler(done).withFailureHandler(fail).runSend({',
    '    groups:sel("groups").value, wording:sel("wording").value, mode:mode });',
    '}',
    'function done(r){',
    '  buttons(false);',
    '  sel("status").textContent=r.message;',
    '  if(r.addresses){var o=sel("out");o.style.display="block";o.value=r.addresses;o.focus();o.select();}',
    '}',
    'function fail(e){ buttons(false); sel("status").textContent="Error: "+e.message; }',
    'syncWording();',
    '<\/script>'
  ].join('');

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(430).setHeight(480), 'Send Open House emails');
}

/** Called from the dialog. Returns {message, addresses}. */
function runSend(payload) {
  var groups = String(payload.groups || '').split('|').filter(String);
  if (!groups.length) return { message: 'Pick who the email is going to.' };

  var people = peopleIn(groups);
  if (!people.length) {
    return { message: 'Nobody on the Roster tab is in ' + groups.map(plural).join(' or ') + ' with an email address.' };
  }

  if (payload.mode === 'list') {
    return {
      message: people.length + ' address' + (people.length === 1 ? '' : 'es') + '. Copy them from the box below.',
      addresses: people.map(function (p) { return p.email; }).join(', ')
    };
  }

  var sender = senderName();
  var rooms = roomsByPerson();

  if (payload.mode === 'bcc') {
    // One message to many people, so per-person wording cannot apply.
    var key = payload.wording === '__own__' ? 'Combined' : payload.wording;
    var template = templateFor(key);
    if (!template) return { message: 'No "' + key + '" row on the Templates tab.' };

    var inbox = me();
    if (!inbox) return { message: 'Could not work out your own email address, which the To field of a BCC draft needs. Use "One draft per person" instead.' };

    var greeting = { name: groups.length === 1 ? plural(groups[0]).toLowerCase() : 'everyone', email: '', room: '' };
    GmailApp.createDraft(inbox, fill(template.subject, greeting, sender), fill(template.body, greeting, sender), {
      bcc: people.map(function (p) { return p.email; }).join(',')
    });
    people.forEach(function (p) { stampEmailed(p.row); });

    return { message: 'One draft created, addressed to you with ' + people.length + ' BCC recipient(s), using the ' + key + ' wording.\n\nReview it in Gmail before sending.' };
  }

  // One personalized draft per person.
  var missing = {};
  var made = 0;
  people.forEach(function (person) {
    var key = payload.wording === '__own__' ? person.group : payload.wording;
    var template = templateFor(key);
    if (!template) { missing[key] = true; return; }

    person.room = rooms[person.name.toLowerCase()] || '';
    GmailApp.createDraft(person.email, fill(template.subject, person, sender), fill(template.body, person, sender));
    stampEmailed(person.row);
    made++;
  });

  var message = made + ' draft' + (made === 1 ? '' : 's') + ' created in Gmail. Open Drafts to review and send.';
  var gaps = Object.keys(missing);
  if (gaps.length) {
    message += '\n\nSkipped anyone needing a "' + gaps.join('" or "') + '" template -- no such row on the Templates tab.';
  }
  return { message: message };
}

/* ------------------------------------------------------- room assignments */

/** Lowercased person name -> the room they are in, for the {{room}} placeholder. */
function roomsByPerson() {
  var map = {};
  readTab(ROOMS).forEach(function (r) {
    var room = String(r.Room || '').trim();
    if (!room) return;
    namesIn(r.Teachers).concat(namesIn(r.Students)).forEach(function (n) {
      var key = n.toLowerCase();
      map[key] = map[key] ? map[key] + ', ' + room : room;
    });
  });
  return map;
}

/** Splits a "Pat Chen, Alex Rivera" cell into names. Commas or newlines. */
function namesIn(cell) {
  return String(cell || '').split(/[,\n;]/).map(function (n) { return n.trim(); }).filter(String);
}

function emailRoomAssignments() {
  var rooms = readTab(ROOMS);
  if (!rooms.length) {
    alert('The Rooms tab is empty. Add your classrooms first.');
    return;
  }

  var emails = emailLookup();
  var sender = senderName();
  var made = 0;
  var unknown = [];

  rooms.forEach(function (room) {
    var roomName = String(room.Room || '').trim();
    if (!roomName) return;

    var teachers = namesIn(room.Teachers);
    var students = namesIn(room.Students);

    teachers.forEach(function (teacher) {
      var address = emails[teacher.toLowerCase()];
      if (!address) { unknown.push(teacher + ' (room ' + roomName + ')'); return; }

      var body = [
        'Hi ' + teacher + ',',
        '',
        'Here is your room for the Open House.',
        '',
        'Room: ' + roomName,
        'Location: ' + (room.Location || 'TBC'),
        'Activity: ' + (room.Activity || 'TBC'),
        '',
        teachers.length > 1 ? 'With you: ' + teachers.filter(function (t) { return t !== teacher; }).join(', ') : '',
        students.length ? 'Students in your room (' + students.length + '): ' + students.join(', ') : 'No students listed yet.',
        '',
        room.Notes ? 'Notes: ' + room.Notes : '',
        '',
        'Thank you,',
        sender
      ].filter(function (line) { return line !== ''; }).join('\n');

      GmailApp.createDraft(address, 'Open House - you are in room ' + roomName, body);
      made++;
    });
  });

  var message = made + ' room assignment draft' + (made === 1 ? '' : 's') + ' created in Gmail.';
  if (unknown.length) {
    message += '\n\nNo email address found for:\n- ' + unknown.join('\n- ') +
      '\n\nAdd them to the Roster or Team tab and run this again.';
  }
  if (!made && !unknown.length) message = 'No teachers are listed in the Teachers column of the Rooms tab.';
  alert(message);
}

/** Lowercased name -> email, from the Roster and Team tabs. */
function emailLookup() {
  var map = {};
  readTab(ROSTER).concat(readTab(TEAM)).forEach(function (r) {
    var name = String(r.Name || '').trim();
    var email = String(r.Email || '').trim();
    if (name && email && !map[name.toLowerCase()]) map[name.toLowerCase()] = email;
  });
  return map;
}

/* ------------------------------------------------------------- task email */

function sendTaskRemindersNow() {
  var ui = SpreadsheetApp.getUi();
  var open = openTasks();
  if (!open.length) {
    alert('No open tasks on the Tasks tab. Nothing to remind anyone about.');
    return;
  }

  var confirm = ui.alert('Send reminders now?',
    open.length + ' open task(s) will be emailed to the people responsible, with a copy to you.\n\n' +
    'Unlike the Open House emails, these SEND straight away rather than becoming drafts.',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  alert(sendTaskReminders());
}

/**
 * Trigger-safe: no dialogs, returns what it did as a string. One digest per
 * owner so nobody gets five separate emails, plus a full summary to you.
 */
function sendTaskReminders() {
  var open = openTasks();
  if (!open.length) return 'No open tasks. No reminders sent.';

  var inbox = me();
  var emails = emailLookup();
  var sender = senderName();
  var today = midnight(new Date());

  var byOwner = {};
  var unassigned = [];
  open.forEach(function (task) {
    if (task.owner) {
      (byOwner[task.owner] = byOwner[task.owner] || []).push(task);
    } else {
      unassigned.push(task);
    }
  });

  var sent = 0;
  var unknown = [];

  Object.keys(byOwner).forEach(function (owner) {
    var address = emails[owner.toLowerCase()];
    if (!address) { unknown.push(owner); return; }

    var mine = byOwner[owner];
    var body = ['Hi ' + owner + ',', '', 'Open House tasks that still need you:', '']
      .concat(mine.map(function (t) { return ' - ' + describe(t, today); }))
      .concat(['', 'The full list lives on the Tasks tab of the Open House sheet.', '', 'Thank you,', sender])
      .join('\n');

    var options = {};
    // Copy yourself in, unless you are the owner.
    if (inbox && inbox.toLowerCase() !== address.toLowerCase()) options.cc = inbox;

    GmailApp.sendEmail(address,
      'Open House: ' + mine.length + ' task' + (mine.length === 1 ? '' : 's') + ' for you', body, options);

    mine.forEach(function (t) { stampReminded(t.row); });
    sent++;
  });

  // A single summary to you, covering everything including what nobody owns.
  if (inbox) {
    var lines = ['Open House task summary - ' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE d MMMM'), ''];
    Object.keys(byOwner).forEach(function (owner) {
      lines.push(owner + ':');
      byOwner[owner].forEach(function (t) { lines.push('  - ' + describe(t, today)); });
      lines.push('');
    });
    if (unassigned.length) {
      lines.push('Nobody assigned yet:');
      unassigned.forEach(function (t) { lines.push('  - ' + describe(t, today)); });
      lines.push('');
    }
    if (unknown.length) {
      lines.push('No email address on the Team or Roster tab for: ' + unknown.join(', '));
      lines.push('Those people were not emailed.');
    }
    GmailApp.sendEmail(inbox, 'Open House: ' + open.length + ' open task' + (open.length === 1 ? '' : 's'), lines.join('\n'));
  }

  var report = 'Reminders sent to ' + sent + ' ' + (sent === 1 ? 'person' : 'people') +
    ' covering ' + open.length + ' open task(s).';
  if (inbox) report += '\n\nA full summary went to ' + inbox + '.';
  if (unassigned.length) report += '\n\n' + unassigned.length + ' task(s) have no owner yet, so only you were told about them.';
  if (unknown.length) report += '\n\nNo email address found for: ' + unknown.join(', ') + '. Add them to the Team tab.';
  return report;
}

function openTasks() {
  return readTab(TASKS).map(function (r) {
    return {
      row: r._row,
      task: String(r.Task || '').trim(),
      details: String(r.Details || '').trim(),
      owner: String(r.Owner || '').trim(),
      due: r['Due date'] instanceof Date ? r['Due date'] : null,
      status: String(r.Status || '').trim()
    };
  }).filter(function (t) {
    return t.task && t.status.toLowerCase() !== 'done';
  });
}

function describe(task, today) {
  var parts = [task.task];
  if (task.due) {
    var days = Math.round((midnight(task.due) - today) / 86400000);
    var when = Utilities.formatDate(task.due, Session.getScriptTimeZone(), 'EEE d MMM');
    if (days < 0) parts.push('OVERDUE, was due ' + when);
    else if (days === 0) parts.push('due TODAY');
    else if (days === 1) parts.push('due tomorrow, ' + when);
    else parts.push('due ' + when + ', in ' + days + ' days');
  } else {
    parts.push('no due date');
  }
  if (task.status && task.status.toLowerCase() !== 'not started') parts.push(task.status);
  var line = parts.join(' — ');
  return task.details ? line + '\n      ' + task.details : line;
}

function midnight(date) {
  var d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

/* ---------------------------------------------------------- daily trigger */

function enableDailyReminders() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.prompt('Daily task reminders',
    'What hour should they go out? Enter 0-23 (7 means around 7am).', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  var hour = parseInt(answer.getResponseText().trim(), 10);
  if (isNaN(hour) || hour < 0 || hour > 23) {
    alert('"' + answer.getResponseText() + '" is not an hour between 0 and 23.');
    return;
  }

  disableDailyReminders(true);
  ScriptApp.newTrigger('sendTaskReminders').timeBased().everyDays(1).atHour(hour).create();
  alert('Daily reminders are on. Google will run them each day somewhere in the hour starting at ' +
    hour + ':00, and email everyone with an open task, copying you.\n\n' +
    'Turn them off from the Open House menu whenever you like.');
}

function disableDailyReminders(quiet) {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'sendTaskReminders') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  if (!quiet) {
    alert(removed ? 'Daily reminders are off.' : 'Daily reminders were not on.');
  }
}

/* ------------------------------------------------------------ the checker */

function checkForProblems() {
  var problems = [];
  var emails = emailLookup();

  // Roster
  var seen = {};
  readTab(ROSTER).forEach(function (r) {
    var at = ROSTER + ' row ' + r._row + ': ';
    var email = String(r.Email || '').trim();
    var group = String(r.Group || '').trim();

    if (!String(r.Name || '').trim()) problems.push(at + 'no name.');
    if (!email) problems.push(at + 'no email address.');
    else if (!looksLikeEmail(email)) problems.push(at + '"' + email + '" does not look like an email address.');
    else if (seen[email.toLowerCase()]) problems.push(at + email + ' is already on row ' + seen[email.toLowerCase()] + '.');
    else seen[email.toLowerCase()] = r._row;

    if (!group) problems.push(at + 'no group.');
    else if (GROUPS.indexOf(group) === -1) problems.push(at + 'group "' + group + '" is not one of ' + GROUPS.join(', ') + '.');
  });

  // Team
  readTab(TEAM).forEach(function (r) {
    var at = TEAM + ' row ' + r._row + ': ';
    if (!String(r.Name || '').trim()) problems.push(at + 'no name.');
    var email = String(r.Email || '').trim();
    if (!email) problems.push(at + 'no email address, so this person cannot be reminded.');
    else if (!looksLikeEmail(email)) problems.push(at + '"' + email + '" does not look like an email address.');
  });

  // Tasks
  readTab(TASKS).forEach(function (r) {
    var at = TASKS + ' row ' + r._row + ': ';
    var owner = String(r.Owner || '').trim();
    var status = String(r.Status || '').trim();

    if (!String(r.Task || '').trim()) problems.push(at + 'no task described.');
    if (!owner) problems.push(at + 'nobody is responsible yet.');
    else if (!emails[owner.toLowerCase()]) problems.push(at + 'no email address for "' + owner + '" on the Team or Roster tab.');
    if (status && STATUSES.indexOf(status) === -1) problems.push(at + 'status "' + status + '" is not one of ' + STATUSES.join(', ') + '.');
    if (r['Due date'] && !(r['Due date'] instanceof Date)) problems.push(at + 'the due date is text, not a date. Retype it as a date.');
  });

  // Rooms, and anyone booked into two of them at once.
  var placed = {};
  readTab(ROOMS).forEach(function (r) {
    var at = ROOMS + ' row ' + r._row + ': ';
    var room = String(r.Room || '').trim();
    if (!room) { problems.push(at + 'no room name.'); return; }
    if (!namesIn(r.Teachers).length) problems.push(at + 'room ' + room + ' has no teacher.');

    namesIn(r.Teachers).concat(namesIn(r.Students)).forEach(function (name) {
      var key = name.toLowerCase();
      if (!emails[key]) problems.push(at + '"' + name + '" is not on the Roster or Team tab.');
      if (placed[key] && placed[key] !== room) {
        problems.push(at + '"' + name + '" is also in room ' + placed[key] + '.');
      } else {
        placed[key] = room;
      }
    });
  });

  // Templates
  var keys = {};
  readTab(TEMPLATES).forEach(function (r) { keys[String(r.Group).trim()] = true; });
  TEMPLATE_KEYS.forEach(function (key) {
    if (!keys[key]) problems.push(TEMPLATES + ': no "' + key + '" row. Run "Set up / repair sheet".');
  });

  alert(problems.length
    ? problems.length + ' thing(s) to look at:\n\n' + problems.join('\n')
    : 'No problems found.');
}

function looksLikeEmail(value) {
  return /^[^@\s]+@[^@\s.]+\.[^@\s]+$/.test(value);
}

/* ------------------------------------------------------------------ guts */

function templateFor(key) {
  var match = null;
  readTab(TEMPLATES).forEach(function (r) {
    if (String(r.Group || '').trim().toLowerCase() === String(key).toLowerCase()) {
      match = { subject: String(r.Subject || 'Open House'), body: String(r.Body || '') };
    }
  });
  return match;
}

function fill(text, person, sender) {
  return String(text)
    .replace(/\{\{\s*name\s*\}\}/gi, person.name || '')
    .replace(/\{\{\s*email\s*\}\}/gi, person.email || '')
    .replace(/\{\{\s*room\s*\}\}/gi, person.room || 'TBC')
    .replace(/\{\{\s*sender\s*\}\}/gi, sender || '');
}

/** Your name for {{sender}}. Asked once, then remembered. */
function senderName() {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty('senderName');
  if (saved) return saved;

  try {
    var ui = SpreadsheetApp.getUi();
    var answer = ui.prompt('Sign your emails', 'What name should {{sender}} be replaced with?', ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() === ui.Button.OK) {
      var name = answer.getResponseText().trim();
      if (name) { props.setProperty('senderName', name); return name; }
    }
  } catch (e) {
    // Running from the daily trigger, where there is no one to ask.
  }
  return '';
}

function me() {
  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/** Finds a column by its header, so reordering columns doesn't misplace writes. */
function columnOf(sheetName, header) {
  var sheet = sheetFor(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === header.toLowerCase()) return i + 1;
  }
  return 0;
}

function stamp(sheetName, row, header) {
  var column = columnOf(sheetName, header);
  if (column) sheetFor(sheetName).getRange(row, column).setValue(new Date());
}

function stampEmailed(row) {
  stamp(ROSTER, row, 'Emailed?');
}

function stampReminded(row) {
  stamp(TASKS, row, 'Last reminded');
}

/** Shows a message, or logs it when running unattended from the trigger. */
function alert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

function esc(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
