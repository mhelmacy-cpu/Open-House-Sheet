/**
 * Open House Needs Sheet
 *
 * One tab per kind of person, so each can carry the columns it actually needs:
 *
 *   Students   name, email, grade, and both parents' names and emails
 *   Parents    name, email, child(ren), and LS / MS / LS-MS division
 *   Musicians  name and email
 *   Teachers   name, email, grade or subject
 *   Team       the people running the event, who can own tasks
 *   Tasks      the to-do list
 *   Rooms      the classrooms in use, and who is in each
 *   Templates  the wording of every email
 *
 * Two different things happen when email is generated, on purpose:
 *   - Emails to parents, students, musicians and teachers become Gmail DRAFTS.
 *     You read them and press Send yourself. Nothing goes out on its own.
 *   - Task reminders to your own team SEND immediately, because that is the
 *     point of a reminder, and because the optional daily reminder runs while
 *     you are not at the computer.
 */

var TEAM = 'Team';
var TASKS = 'Tasks';
var ROOMS = 'Rooms';
var PANEL = 'Panel';
var TEMPLATES = 'Templates';
var STUDENTS = 'Students';
var PARENTS = 'Parents';

var GROUPS = ['Parent', 'Student', 'Musician', 'Teacher'];

// SENT is on every group tab and is stamped by the script when a draft is
// generated. ATTENDING is yours to fill in as answers come back; it is on every
// tab but Teachers, and on Students it is what the borough count keys off.
var SENT = 'Email confirmation sent';
var ATTENDING = 'Confirmed attending';
var PARENT_SENT = 'Parent confirmation sent';

// LREI red for the header row of every tab. If the school's exact hex is
// different, change it here and re-run "Set up / repair sheet" -- this one
// value colours every header on every tab.
var HEADER_FILL = '#c8102e';
var HEADER_TEXT = '#ffffff';

// Students and musicians are listed with their names in two columns, so the
// tabs sort by surname. Parents and teachers keep one column, and so do the
// Parent 1 / Parent 2 columns on the Students tab.
var NAME = 'Name';
var FIRST = 'First name';
var LAST = 'Last name';

var DIVISIONS = ['LS', 'MS', 'LS/MS'];

// Where a family travels in from. Not strictly boroughs -- the last three are
// not -- but it is one question with one answer, so it is one column.
// The rows of the grade summary, and the suffix written beside each child's
// name. Zeros are informative here -- they show which grades are missing.
var GRADES = ['K', '1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th',
  '9th', '10th', '11th', '12th'];

var BOROUGHS = [
  'M - Manhattan',
  'B - Brooklyn',
  'Q - Queens',
  'X - The Bronx',
  'J - New Jersey',
  'LI - Long Island',
  'CT - Connecticut'
];
var ATTENDING_VALUES = ['Yes', 'No', 'Maybe'];

/** Each group lives on its own tab, with its own columns. */
var GROUP_TABS = {
  Parent: {
    tab: PARENTS,
    headers: [NAME, 'Email', 'Child(ren)', 'Division', 'Borough', 'Notes', SENT, ATTENDING]
  },
  Student: {
    tab: STUDENTS,
    headers: [FIRST, LAST, 'Email', 'Grade', 'Borough',
      'Parent 1 name', 'Parent 2 name', 'Parent 1 email', 'Parent 2 email',
      'Notes', SENT, ATTENDING, PARENT_SENT]
  },
  Musician: {
    tab: 'Musicians',
    headers: [FIRST, LAST, 'Email', 'Notes', SENT, ATTENDING]
  },
  Teacher: {
    tab: 'Teachers',
    headers: [NAME, 'Role', 'Division', 'Grade', 'Subject', 'Notes', 'Email', SENT]
  }
};

/**
 * Columns that used to be called something else. Setup renames them in place
 * rather than adding a second column, so dates already stamped are kept.
 */
var RENAMED_COLUMNS = {};
RENAMED_COLUMNS[SENT] = ['Emailed?'];
RENAMED_COLUMNS[PARENT_SENT] = ['Parents emailed?'];
// On the two tabs that now split the name, the old single column becomes the
// first-name column and keeps whatever is in it; splitFullNames then moves
// each surname across.
RENAMED_COLUMNS[FIRST] = [NAME];
RENAMED_COLUMNS['Grade'] = ['Grade / subject'];

var STATUSES = ['Not started', 'In progress', 'Blocked', 'Done'];

var TEAM_HEADERS = [NAME, 'Email', 'Role'];
var TASK_HEADERS = ['Task', 'Details', 'Owner', 'Due date', 'Status', 'Last reminded'];
var ROOM_HEADERS = ['Room', 'Location', 'Activity', 'Teachers', 'Students', 'Notes'];

// The Panel tab holds two blocks side by side: the panellists as a normal
// table in column A onwards, then one empty column, then the questions. That
// empty column is load-bearing -- a table is taken to end at its first blank
// header, so it is what keeps the questions out of the script's reach: they
// are never reordered, renamed or cleared.
var PANEL_HEADERS = [FIRST, LAST, 'Grade', 'Notes'];
var QUESTION_HEADERS = ['#', 'Question', 'Notes'];
var TEMPLATE_HEADERS = ['Group', 'Subject', 'Body'];

// "Combined" is the wording used when one email goes to more than one group.
var TEMPLATE_KEYS = GROUPS.concat(['Combined']);

// Starter wording, written into the Templates tab by "Set up / repair sheet".
// Edit it on the tab, not here -- the tab is what the drafts are built from.
// The Templates tab starts blank: one row per group, with the subject and body
// yours to write. Nothing here puts words in your mouth, and a row left empty
// stops a send rather than producing empty drafts.
var TEMPLATE_SEED = {
  Parent: { subject: '', body: '' },
  Student: { subject: '', body: '' },
  Musician: { subject: '', body: '' },
  Teacher: { subject: '', body: '' },
  Combined: { subject: '', body: '' }
};

/* ------------------------------------------------------------------- menu */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Open House')
    .addItem('Set up / repair sheet', 'setUpSheet')
    .addItem('Fill in children on the Parents tab', 'fillInChildren')
    .addItem('Blank out the wording...', 'restoreTemplates')
    .addSeparator()
    .addItem('Send emails...', 'showSendDialog')
    .addSeparator()
    .addItem('Assign people to a room...', 'showRoomDialog')
    .addItem('Email room assignments to teachers', 'emailRoomAssignments')
    .addItem('Update who has been emailed', 'updateSentColumn')
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
  var notes = [];

  var students = tab(ss, GROUP_TABS.Student.tab, GROUP_TABS.Student.headers);
  var studentsSplit = splitFullNames(STUDENTS);
  widths(STUDENTS, [[FIRST, 140], [LAST, 140], ['Email', 230], ['Grade', 70],
    ['Borough', 135], ['Parent 1 name', 170], ['Parent 2 name', 170],
    ['Parent 1 email', 230], ['Parent 2 email', 230], ['Notes', 220]]);
  boroughDropdown(STUDENTS);
  attendingDropdown(STUDENTS);
  if (students.getLastRow() < 2) {
    students.getRange(2, 1, 2, 9).setValues([
      ['Sam', 'Doe', 'sam.doe@example.com', '7', 'B - Brooklyn',
        'Jane Doe', 'Chris Doe', 'jane.doe@example.com', 'chris.doe@example.com'],
      ['Alex', 'Rivera', 'alex.rivera@example.com', '9', 'M - Manhattan',
        'Mia Rivera', '', 'mia.rivera@example.com', '']
    ]);
  }

  var parents = tab(ss, GROUP_TABS.Parent.tab, GROUP_TABS.Parent.headers);
  widths(PARENTS, [[NAME, 170], ['Email', 230], ['Child(ren)', 240],
    ['Division', 95], ['Borough', 135], ['Notes', 220]]);
  dropdown(PARENTS, 'Division', DIVISIONS, 'Lower School, Middle School, or both');
  boroughDropdown(PARENTS);
  attendingDropdown(PARENTS);
  if (parents.getLastRow() < 2) {
    parents.getRange(2, 1, 2, 2).setValues([
      ['Jane Doe', 'jane.doe@example.com'],
      ['Mia Rivera', 'mia.rivera@example.com']
    ]);
  }

  var musicians = tab(ss, GROUP_TABS.Musician.tab, GROUP_TABS.Musician.headers);
  // The Instrument column is gone, but only drop it if nothing was typed in.
  var leftover = dropEmptyColumn(GROUP_TABS.Musician.tab, 'Instrument');
  if (leftover) notes.push(leftover);
  var musiciansSplit = splitFullNames(GROUP_TABS.Musician.tab);
  widths(GROUP_TABS.Musician.tab, [[FIRST, 140], [LAST, 140], ['Email', 230], ['Notes', 220]]);
  attendingDropdown(GROUP_TABS.Musician.tab);
  if (musicians.getLastRow() < 2) {
    musicians.getRange(2, 1, 1, 3).setValues([['Alex', 'Rivera', 'alex.rivera@example.com']]);
  }

  var teachers = tab(ss, GROUP_TABS.Teacher.tab, GROUP_TABS.Teacher.headers);
  widths(GROUP_TABS.Teacher.tab, [[NAME, 170], ['Role', 150], ['Division', 95],
    ['Grade', 90], ['Subject', 160], ['Notes', 220], ['Email', 230]]);
  dropdown(GROUP_TABS.Teacher.tab, 'Division', DIVISIONS,
    'Lower School, Middle School, or both');
  if (teachers.getLastRow() < 2) {
    teachers.getRange(2, 1, 1, 7).setValues([
      ['Pat Chen', 'Head teacher', 'LS', '4', 'Science', '', 'pat.chen@example.com']
    ]);
  }

  var team = tab(ss, TEAM, TEAM_HEADERS);
  widths(TEAM, [[NAME, 180], ['Email', 240], ['Role', 200]]);
  if (team.getLastRow() < 2) {
    team.getRange(2, 1, 1, 3).setValues([[me() || 'You', me() || 'you@example.com', 'Organizer']]);
  }

  var tasks = tab(ss, TASKS, TASK_HEADERS);
  widths(TASKS, [['Task', 260], ['Details', 300], ['Owner', 160],
    ['Due date', 110], ['Status', 120], ['Last reminded', 140]]);
  dropdown(TASKS, 'Status', STATUSES, 'Pick one: ' + STATUSES.join(', '));
  numberFormat(TASKS, 'Due date', 'yyyy-mm-dd');
  var teamNames = readTab(TEAM).map(personName).filter(String);
  if (teamNames.length) dropdown(TASKS, 'Owner', teamNames, 'Someone from the Team tab');
  if (tasks.getLastRow() < 2) {
    tasks.getRange(2, 1, 2, 5).setValues([
      ['Book the auditorium', 'Confirm with the front office', teamNames[0] || '', '', 'Not started'],
      ['Print programs', '50 copies, double sided', teamNames[0] || '', '', 'Not started']
    ]);
  }

  var rooms = tab(ss, ROOMS, ROOM_HEADERS);
  widths(ROOMS, [['Room', 120], ['Location', 160], ['Activity', 200],
    ['Teachers', 240], ['Students', 300], ['Notes', 220]]);
  wrap(ROOMS, 'Teachers');
  wrap(ROOMS, 'Students');
  if (rooms.getLastRow() < 2) {
    rooms.getRange(2, 1, 1, 6).setValues([
      ['101', 'First floor', 'Science projects', 'Pat Chen', 'Sam Doe', 'Example row - delete me']
    ]);
  }

  var panel = tab(ss, PANEL, PANEL_HEADERS);
  widths(PANEL, [[FIRST, 140], [LAST, 140], ['Grade', 70], ['Notes', 220]]);
  buildPanelQuestions(panel);

  var templates = tab(ss, TEMPLATES, TEMPLATE_HEADERS);
  var have = {};
  readTab(TEMPLATES).forEach(function (r) { have[str(r.Group)] = true; });
  TEMPLATE_KEYS.forEach(function (key) {
    if (!have[key]) templates.appendRow([key, TEMPLATE_SEED[key].subject, TEMPLATE_SEED[key].body]);
  });
  widths(TEMPLATES, [['Group', 120], ['Subject', 300], ['Body', 520]]);
  wrap(TEMPLATES, 'Body');

  buildParentSummary();
  buildStudentSummary();

  var nameSplit = (studentsSplit || 0) + (musiciansSplit || 0);
  if (nameSplit) {
    notes.push('Split ' + nameSplit + ' full name(s) on the Students and Musicians tabs into ' +
      'First name and Last name. The split happens at the first space, so a two-word surname ' +
      'stays whole but a two-word given name does not -- worth a quick look down those columns.');
  }

  notes = notes.concat(migrateOldRoster(ss));

  alert('Sheet is ready.\n\n' + notes.join('\n\n') +
    (notes.length ? '\n\n' : '') +
    'Re-run this whenever you add someone to the Team tab, so the Tasks owner dropdown picks them up.');
}

/**
 * Earlier versions of this script kept everyone on one "Roster" tab with a
 * Group column. If that tab is still here, move its rows onto the new tabs
 * rather than leaving them stranded. The old tab is renamed, never deleted.
 */
function migrateOldRoster(ss) {
  var old = ss.getSheetByName('Roster');
  if (!old) return [];

  var rows = readTab('Roster').filter(function (r) {
    // Skip the example rows the old version seeded.
    return str(r.Email) && str(r.Email).indexOf('@example.com') === -1;
  });

  if (!rows.length) {
    ss.setActiveSheet(old);
    ss.renameActiveSheet('Roster (old, empty)');
    return ['Your old Roster tab held no real rows, so nothing was moved. It is now called "Roster (old, empty)" and you can delete it.'];
  }

  var moved = 0;
  var skipped = [];
  rows.forEach(function (r) {
    var group = str(r.Group);
    var spec = GROUP_TABS[group];
    if (!spec) { skipped.push('row ' + r._row + ' (group "' + group + '")'); return; }

    var sheet = ss.getSheetByName(spec.tab);
    if (alreadyListed(spec.tab, str(r.Email))) return;

    // Written by header, not by position: the name is one column on some tabs
    // and two on others, so column 2 is not always the email address.
    sheet.appendRow([]);
    var row = sheet.getLastRow();
    var whole = str(r.Name);
    if (columnOf(spec.tab, NAME)) {
      sheet.getRange(row, columnOf(spec.tab, NAME)).setValue(whole);
    } else {
      var gap = whole.indexOf(' ');
      sheet.getRange(row, columnOf(spec.tab, FIRST))
        .setValue(gap === -1 ? whole : whole.slice(0, gap));
      if (gap !== -1) {
        sheet.getRange(row, columnOf(spec.tab, LAST)).setValue(whole.slice(gap + 1).trim());
      }
    }
    sheet.getRange(row, columnOf(spec.tab, 'Email')).setValue(str(r.Email));
    var notesColumn = columnOf(spec.tab, 'Notes');
    if (notesColumn && str(r.Notes)) {
      sheet.getRange(row, notesColumn).setValue(str(r.Notes));
    }
    moved++;
  });

  ss.setActiveSheet(old);
  ss.renameActiveSheet('Roster (old)');

  var note = 'Moved ' + moved + ' person/people off your old Roster tab onto the new Students, Parents, Musicians and Teachers tabs.';
  if (skipped.length) note += '\n\nNot moved, because the group was not recognised: ' + skipped.join(', ') + '.';
  note += '\n\nThe old tab is still here, renamed "Roster (old)", so you can check nothing was lost before deleting it. Students moved across will need their grade and parent details filled in.';
  return [note];
}

function alreadyListed(tabName, email) {
  if (!email) return false;
  return readTab(tabName).some(function (r) {
    return str(r.Email).toLowerCase() === email.toLowerCase();
  });
}

function tab(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    ensureColumns(sheet, headers);
    orderColumns(sheet, headers);
  }
  sheet.getRange(1, 1, 1, Math.max(tableWidth(sheet), headers.length))
    .setFontWeight('bold')
    .setBackground(HEADER_FILL)
    .setFontColor(HEADER_TEXT);
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Brings an existing tab up to date without disturbing what is already on it.
 * A column that has been renamed is renamed in place; a genuinely new one is
 * inserted at its proper position, which shifts the columns after it along
 * with their data. Columns you added yourself are left where they are.
 */
function ensureColumns(sheet, headers) {
  for (var i = 0; i < headers.length; i++) {
    var width = Math.max(tableWidth(sheet), 1);
    var current = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) {
      return String(h).trim().toLowerCase();
    });
    var want = headers[i].toLowerCase();
    if (current.indexOf(want) !== -1) continue;

    // Renamed rather than new? Relabel the old column and keep its contents.
    var oldNames = RENAMED_COLUMNS[headers[i]] || [];
    var found = -1;
    oldNames.forEach(function (name) {
      if (found === -1) found = current.indexOf(name.toLowerCase());
    });
    if (found !== -1) {
      sheet.getRange(1, found + 1).setValue(headers[i]);
      continue;
    }

    // Place it just after the column it is meant to follow, wherever that
    // has ended up, rather than at its raw index -- otherwise a new last
    // column lands in front of the one before it.
    var after = i > 0 ? current.indexOf(headers[i - 1].toLowerCase()) + 1 : 0;
    var position = after > 0 ? after + 1 : i + 1;

    if (position > sheet.getLastColumn()) {
      sheet.getRange(1, position).setValue(headers[i]);
    } else {
      sheet.insertColumnBefore(position);
      sheet.getRange(1, position).setValue(headers[i]);
    }
  }
}

/**
 * Puts the table's columns into the order the script lists them in, carrying
 * each column's contents along with it. Only ever moves a column leftwards:
 * by the time it looks at position n, positions before it are already right,
 * so the column it wants can only be further along.
 *
 * This is what makes "Set up / repair sheet" able to change a layout rather
 * than only add to it. Columns you added yourself sit after the ones the
 * script knows about, and are left in their own order.
 */
function orderColumns(sheet, headers) {
  for (var target = 0; target < headers.length; target++) {
    var width = tableWidth(sheet);
    if (!width) return;

    var current = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) {
      return String(h).trim().toLowerCase();
    });

    var at = current.indexOf(headers[target].toLowerCase());
    if (at === -1 || at === target) continue;

    sheet.moveColumns(sheet.getRange(1, at + 1, sheet.getMaxRows(), 1), target + 1);
  }
}

/**
 * Removes a column we no longer use, but only when it is empty -- deleting
 * one with anything in it would throw away whatever was typed there.
 */
function dropEmptyColumn(sheetName, header) {
  var sheet = sheetFor(sheetName);
  var column = columnOf(sheetName, header);
  if (!column) return '';

  var last = sheet.getLastRow();
  if (last > 1) {
    var used = sheet.getRange(2, column, last - 1, 1).getValues().some(function (row) {
      return str(row[0]) !== '' && str(row[0]) !== 'Cello';
    });
    if (used) {
      return 'The ' + sheetName + ' tab still has entries in its "' + header +
        '" column, so it was left in place. Delete the column yourself once you are sure you do not need them.';
    }
  }
  sheet.deleteColumn(column);
  return '';
}

function dropdown(sheetName, header, values, help) {
  var column = columnOf(sheetName, header);
  if (!column) return;
  var sheet = sheetFor(sheetName);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).setHelpText(help).build();
  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function attendingDropdown(sheetName) {
  dropdown(sheetName, ATTENDING, ATTENDING_VALUES, 'Did they say they are coming?');
}

function boroughDropdown(sheetName) {
  dropdown(sheetName, 'Borough', BOROUGHS, 'Where they travel in from');
}

function widths(sheetName, pairs) {
  var sheet = sheetFor(sheetName);
  pairs.forEach(function (pair) {
    var column = columnOf(sheetName, pair[0]);
    if (column) sheet.setColumnWidth(column, pair[1]);
  });
}

function wrap(sheetName, header) {
  var column = columnOf(sheetName, header);
  if (!column) return;
  var sheet = sheetFor(sheetName);
  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setWrap(true);
}

function numberFormat(sheetName, header, format) {
  var column = columnOf(sheetName, header);
  if (!column) return;
  var sheet = sheetFor(sheetName);
  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat(format);
}

/* ------------------------------------------------------- the panel questions */

/**
 * Writes the questions block's header beside the panellists, leaving one empty
 * column between them. Only the header is ever written, so questions already
 * typed in are never touched; on a first run the numbers 1-10 are filled in to
 * give the list some shape.
 */
function buildPanelQuestions(sheet) {
  var start = questionsStart(sheet);
  var needed = start + QUESTION_HEADERS.length - 1;
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }

  sheet.getRange(1, start, 1, QUESTION_HEADERS.length).setValues([QUESTION_HEADERS])
    .setFontWeight('bold').setBackground(HEADER_FILL).setFontColor(HEADER_TEXT);

  sheet.setColumnWidth(start - 1, 30);        // the gap
  sheet.setColumnWidth(start, 40);            // #
  sheet.setColumnWidth(start + 1, 420);       // Question
  sheet.setColumnWidth(start + 2, 200);       // Notes

  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, start + 1, rows, 1).setWrap(true);
  sheet.getRange(2, start, rows, 1).setHorizontalAlignment('center');

  // Number the list on a first run, when nothing has been written yet.
  var depth = Math.min(10, rows);
  var already = sheet.getRange(2, start, depth, 2).getValues().some(function (row) {
    return str(row[0]) !== '' || str(row[1]) !== '';
  });
  if (!already) {
    var numbers = [];
    for (var i = 1; i <= depth; i++) numbers.push([i]);
    sheet.getRange(2, start, depth, 1).setValues(numbers);
  }
}

/**
 * Which column the questions start in. An existing block is found by its "#"
 * header rather than assumed, so adding a panellist column -- which shifts the
 * block right -- does not strand it and grow a second one.
 */
function questionsStart(sheet) {
  var width = tableWidth(sheet);
  var last = sheet.getLastColumn();

  if (last > width) {
    var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
    for (var i = width; i < headers.length; i++) {
      if (String(headers[i]).trim() === QUESTION_HEADERS[0]) return i + 1;
    }
  }
  return width + 2;
}

/* --------------------------------------------------------- the summaries */

/**
 * Both summaries are built out of spreadsheet formulas rather than written as
 * numbers, so they recalculate the moment you type -- no command to run and
 * nothing to wait for. They sit to the right of each tab's data with one blank
 * column between; that gap is what keeps them out of the script's way.
 *
 * Everything from the gap rightwards belongs to the summary and is rewritten
 * whenever setup runs, so don't keep your own notes over there.
 */
function buildParentSummary() {
  var sheet = sheetFor(PARENTS);
  var children = columnLetter(columnOf(PARENTS, 'Child(ren)'));
  var borough = columnLetter(columnOf(PARENTS, 'Borough'));
  var names = columnLetter(columnOf(PARENTS, NAME));
  if (!children || !borough || !names) return;

  var start = summaryStart(sheet);
  var labelColumn = columnLetter(start);
  var valueColumn = columnLetter(start + 1);
  var tab = "'" + PARENTS + "'!";

  var rows = [];
  var bold = [];
  function add(label, formula) {
    rows.push([label === undefined ? '' : label, formula === undefined ? '' : formula]);
    return rows.length;          // the sheet row this landed on
  }

  bold.push(add('PARENTS SUMMARY'));
  add('Counts itself as you type');
  add('');

  var representedRow = add('Grades represented');
  bold.push(representedRow);
  bold.push(add('Grade', 'Parents'));

  var firstGradeRow = rows.length + 1;
  GRADES.forEach(function (grade) {
    var row = rows.length + 1;
    // Matches "Maren (9th)" inside the Child(ren) cell, however many children
    // are listed there. The criteria reads the label beside it, so renaming a
    // grade row re-points its own count.
    add(grade, '=COUNTIF(' + tab + '$' + children + '$2:$' + children +
      ', "*(" & $' + labelColumn + row + ' & ")*")');
  });
  var lastGradeRow = rows.length;

  rows[representedRow - 1][1] =
    '=COUNTIF(' + valueColumn + firstGradeRow + ':' + valueColumn + lastGradeRow + ', ">0")';

  add('');
  bold.push(add('Borough', 'Parents'));
  BOROUGHS.forEach(function (name) {
    var row = rows.length + 1;
    add(name, '=COUNTIF(' + tab + '$' + borough + '$2:$' + borough +
      ', $' + labelColumn + row + ')');
  });
  add('No borough yet', '=COUNTA(' + tab + '$' + names + '$2:$' + names +
    ') - COUNTA(' + tab + '$' + borough + '$2:$' + borough + ')');
  bold.push(add('Total parents', '=COUNTA(' + tab + '$' + names + '$2:$' + names + ')'));

  paintSummary(sheet, start, rows, bold);
}

function buildStudentSummary() {
  var sheet = sheetFor(STUDENTS);
  var borough = columnLetter(columnOf(STUDENTS, 'Borough'));
  var attending = columnLetter(columnOf(STUDENTS, ATTENDING));
  var names = columnLetter(columnOf(STUDENTS, FIRST));
  if (!borough || !attending || !names) return;

  var start = summaryStart(sheet);
  var labelColumn = columnLetter(start);
  var tab = "'" + STUDENTS + "'!";

  var rows = [];
  var bold = [];
  function add(label, formula) {
    rows.push([label === undefined ? '' : label, formula === undefined ? '' : formula]);
    return rows.length;
  }

  bold.push(add('STUDENTS SUMMARY'));
  add('Confirmed attending only');
  add('');

  bold.push(add('Borough', 'Confirmed'));
  BOROUGHS.forEach(function (name) {
    var row = rows.length + 1;
    // A student counts towards their borough only once Confirmed attending
    // says Yes.
    add(name, '=COUNTIFS(' + tab + '$' + borough + '$2:$' + borough +
      ', $' + labelColumn + row +
      ', ' + tab + '$' + attending + '$2:$' + attending + ', "Yes")');
  });

  add('');
  bold.push(add('Confirmed', '=COUNTIF(' + tab + '$' + attending + '$2:$' + attending + ', "Yes")'));
  add('Not coming', '=COUNTIF(' + tab + '$' + attending + '$2:$' + attending + ', "No")');
  add('No answer yet', '=COUNTA(' + tab + '$' + names + '$2:$' + names +
    ') - COUNTA(' + tab + '$' + attending + '$2:$' + attending + ')');
  bold.push(add('Total students', '=COUNTA(' + tab + '$' + names + '$2:$' + names + ')'));

  paintSummary(sheet, start, rows, bold);
}

/** One blank column past the table, and never left of column T. */
function summaryStart(sheet) {
  return Math.max(tableWidth(sheet) + 2, 20);
}

/** Clears whatever summary was there, then writes this one. */
function paintSummary(sheet, start, rows, boldRows) {
  if (sheet.getMaxColumns() < start + 1) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), start + 1 - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < rows.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rows.length - sheet.getMaxRows());
  }

  // Everything right of the gap is ours, so a block left stranded by an
  // inserted column is cleared along with the old one.
  var firstOurs = tableWidth(sheet) + 2;
  if (sheet.getMaxColumns() >= firstOurs) {
    sheet.getRange(1, firstOurs, sheet.getMaxRows(), sheet.getMaxColumns() - firstOurs + 1).clear();
  }

  sheet.getRange(1, start, rows.length, 2).setValues(rows);
  sheet.setColumnWidth(start, 175);
  sheet.setColumnWidth(start + 1, 95);

  sheet.getRange(1, start, 1, 2).setBackground(HEADER_FILL).setFontColor(HEADER_TEXT);
  sheet.getRange(2, start, 1, 2).setFontSize(9).setFontColor('#777777');
  boldRows.forEach(function (row) {
    sheet.getRange(row, start, 1, 2).setFontWeight('bold');
  });
  sheet.getRange(1, start + 1, rows.length, 1).setHorizontalAlignment('right');
}

/* ---------------------------------------------------------- the wording */

/**
 * "Set up / repair sheet" only adds template rows that are missing, so it
 * never overwrites what you have written. This is the way to wipe a row back
 * to blank and start again -- and it does overwrite.
 */
function restoreTemplates() {
  var ui = SpreadsheetApp.getUi();

  var answer = ui.prompt('Blank out the wording',
    'Which one? ' + TEMPLATE_KEYS.join(' / ') + '\n\n' +
    'Or type all for every row. The subject and body of whatever you name are ' +
    'emptied so you can start again. The other rows are left alone.',
    ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  var typed = answer.getResponseText().trim();
  if (!typed) return;

  var wanted;
  if (typed.toLowerCase() === 'all') {
    wanted = TEMPLATE_KEYS.slice();
  } else {
    wanted = TEMPLATE_KEYS.filter(function (key) {
      return key.toLowerCase() === typed.toLowerCase() ||
        key.toLowerCase() + 's' === typed.toLowerCase();
    });
    if (!wanted.length) {
      alert('"' + typed + '" is not one of ' + TEMPLATE_KEYS.join(', ') + ', and is not "all".');
      return;
    }
  }

  var sheet = sheetFor(TEMPLATES);
  var subjectColumn = columnOf(TEMPLATES, 'Subject');
  var bodyColumn = columnOf(TEMPLATES, 'Body');
  if (!subjectColumn || !bodyColumn) {
    alert('The Templates tab is missing its Subject or Body column. Run "Set up / repair sheet".');
    return;
  }

  var existing = {};
  readTab(TEMPLATES).forEach(function (r) { existing[str(r.Group)] = r._row; });

  wanted.forEach(function (key) {
    if (existing[key]) {
      sheet.getRange(existing[key], subjectColumn).setValue(TEMPLATE_SEED[key].subject);
      sheet.getRange(existing[key], bodyColumn).setValue(TEMPLATE_SEED[key].body);
    } else {
      sheet.appendRow([key, TEMPLATE_SEED[key].subject, TEMPLATE_SEED[key].body]);
    }
  });

  wrap(TEMPLATES, 'Body');
  alert('Blanked: ' + wanted.join(', ') + '.\n\nOpen the Templates tab to write the new wording.');
}

/* ------------------------------------------------- children on the Parents tab */

/**
 * Looks every parent up against the Students tab and writes their children's
 * names into the Child(ren) column. Plain text rather than a formula, so it
 * survives column edits and you can still type a name in by hand.
 */
function fillInChildren(quiet) {
  var sheet = sheetFor(PARENTS);
  var column = columnOf(PARENTS, 'Child(ren)');
  if (!column) {
    if (!quiet) alert('The Parents tab has no "Child(ren)" column. Run "Set up / repair sheet" first.');
    return;
  }

  var byParentName = {};
  var byParentEmail = {};
  readTab(STUDENTS).forEach(function (s) {
    // "Maren (9th)" -- the given name, and the grade the summary counts.
    var given = personFirstName(s);
    if (!given) return;
    var grade = gradeLabel(s.Grade);
    var child = grade ? given + ' (' + grade + ')' : given;

    [['Parent 1 name', 'Parent 1 email'], ['Parent 2 name', 'Parent 2 email']].forEach(function (pair) {
      var name = str(s[pair[0]]);
      var email = str(s[pair[1]]);
      if (name) (byParentName[name.toLowerCase()] = byParentName[name.toLowerCase()] || []).push(child);
      if (email) (byParentEmail[email.toLowerCase()] = byParentEmail[email.toLowerCase()] || []).push(child);
    });
  });

  var filled = 0;
  var noMatch = [];
  readTab(PARENTS).forEach(function (p) {
    var name = personName(p);
    if (!name) return;

    // Match on email first -- two parents can share a name, not an address.
    var children = byParentEmail[str(p.Email).toLowerCase()] || byParentName[name.toLowerCase()];
    if (!children || !children.length) { noMatch.push(name); return; }

    var unique = children.filter(function (c, i) { return children.indexOf(c) === i; });
    var joined = unique.join(', ');
    // Only write when it actually differs, so an automatic refill after every
    // edit does not churn the sheet.
    if (str(p['Child(ren)']) !== joined) sheet.getRange(p._row, column).setValue(joined);
    filled++;
  });

  if (quiet) return;

  var message = 'Filled in children for ' + filled + ' parent(s).';
  if (noMatch.length) {
    message += '\n\nNo child found on the Students tab for:\n- ' + noMatch.join('\n- ') +
      '\n\nEither their child is not working, or the name or email on the Students tab does not match. ' +
      'Anything already typed in by hand was left alone.';
  }
  alert(message);
}

/**
 * Google runs this by itself after any hand edit, so the Child(ren) column --
 * and therefore the grade count that reads it -- keeps up as students and
 * parents are added. Script-made edits do not fire it, so there is no loop.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var name = sheet.getName();
    if (name !== STUDENTS && name !== PARENTS) return;
    if (e.range.getLastRow() < 2) return;

    // Only the columns that decide which child belongs to which parent.
    var watched = (name === STUDENTS
      ? [FIRST, LAST, 'Grade', 'Parent 1 name', 'Parent 2 name', 'Parent 1 email', 'Parent 2 email']
      : [NAME, 'Email'])
      .map(function (header) { return columnOf(name, header); })
      .filter(function (column) { return column > 0; });

    var from = e.range.getColumn();
    var to = e.range.getLastColumn();
    var touched = watched.some(function (column) { return column >= from && column <= to; });
    if (!touched) return;

    fillInChildren(true);
  } catch (err) {
    // An edit must never fail because of this.
  }
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

  var width = tableWidth(sheet);
  if (!width) return [];
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();

  var out = [];
  rows.forEach(function (row, i) {
    if (row.every(function (c) { return c === '' || c === null; })) return;
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

function str(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

/**
 * A person's full name, whichever shape their tab uses: one Name column, or
 * First name and Last name. Everything that matches people by name -- room
 * assignments, task owners, the Rooms tab -- goes through this.
 */
function personName(row) {
  var single = str(row[NAME]);
  if (single) return single;
  return [str(row[FIRST]), str(row[LAST])].filter(String).join(' ');
}

/** 7 -> "7th", 1 -> "1st", and anything not a number (K, Pre-K) unchanged. */
function gradeLabel(grade) {
  var whole = str(grade);
  if (!whole) return '';
  var n = parseInt(whole, 10);
  if (isNaN(n) || String(n) !== whole) return whole;

  var teens = n % 100;
  if (teens >= 11 && teens <= 13) return n + 'th';
  if (n % 10 === 1) return n + 'st';
  if (n % 10 === 2) return n + 'nd';
  if (n % 10 === 3) return n + 'rd';
  return n + 'th';
}

/** 1 -> "A", 27 -> "AA", for building the summary's formulas. */
function columnLetter(index) {
  var letters = '';
  while (index > 0) {
    var remainder = (index - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    index = Math.floor((index - 1) / 26);
  }
  return letters;
}

/** The first word of a full name, for greeting someone by it. */
function givenName(whole) {
  return str(whole).split(/\s+/)[0] || str(whole);
}

/** Just the given name, for greeting someone as "Hi Sam," not "Hi Sam Doe,". */
function personFirstName(row) {
  var first = str(row[FIRST]);
  if (first) return first;
  return str(row[NAME]).split(/\s+/)[0] || '';
}

/**
 * On a tab that used to keep one Name column, moves each surname into the new
 * Last name column. Splits at the first space, so a two-word surname stays
 * whole; a one-word entry is left alone, and a row where a last name has
 * already been typed is never touched.
 */
function splitFullNames(sheetName) {
  var firstColumn = columnOf(sheetName, FIRST);
  var lastColumn = columnOf(sheetName, LAST);
  if (!firstColumn || !lastColumn) return 0;

  var sheet = sheetFor(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var firsts = sheet.getRange(2, firstColumn, lastRow - 1, 1).getValues();
  var lasts = sheet.getRange(2, lastColumn, lastRow - 1, 1).getValues();
  var split = 0;

  for (var i = 0; i < firsts.length; i++) {
    var whole = str(firsts[i][0]);
    if (!whole || str(lasts[i][0])) continue;
    var gap = whole.indexOf(' ');
    if (gap === -1) continue;
    firsts[i][0] = whole.slice(0, gap);
    lasts[i][0] = whole.slice(gap + 1).trim();
    split++;
  }

  if (split) {
    sheet.getRange(2, firstColumn, firsts.length, 1).setValues(firsts);
    sheet.getRange(2, lastColumn, lasts.length, 1).setValues(lasts);
  }
  return split;
}

/**
 * How many columns the table itself occupies: everything from column 1 up to
 * the first empty header cell. The summary blocks live past that gap, which is
 * what keeps them out of the script's way.
 */
function tableWidth(sheet) {
  var last = sheet.getLastColumn();
  if (!last) return 0;
  var headers = sheet.getRange(1, 1, 1, last).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === '') return i;
  }
  return headers.length;
}

/** Finds a column by its header, so reordering columns doesn't misplace writes. */
function columnOf(sheetName, header) {
  var sheet = sheetFor(sheetName);
  var width = tableWidth(sheet);
  if (!width) return 0;
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === header.toLowerCase()) return i + 1;
  }
  return 0;
}

/* --------------------------------------------- recipient group combinations */

/**
 * Every non-empty combination of the groups: each on its own, each pair, each
 * trio, and all of them. Four groups gives 15 options, ordered shortest first.
 * Used only by the send dialog.
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

/** Everyone on the given groups' tabs, one entry per email address. */
function peopleIn(groups) {
  var byEmail = {};
  var people = [];

  groups.forEach(function (group) {
    var spec = GROUP_TABS[group];
    if (!spec) return;

    readTab(spec.tab).forEach(function (r) {
      var email = str(r.Email);
      if (!email) return;

      var key = email.toLowerCase();
      if (byEmail[key]) {
        // Already listed under another group: keep one entry, remember both stamps.
        byEmail[key].stamps.push({ tab: spec.tab, row: r._row, header: SENT });
        return;
      }

      var person = {
        name: personName(r) || 'there',
        first: personFirstName(r) || 'there',
        email: email,
        group: group,
        grade: str(r.Grade),
        subject: str(r.Subject),
        role: str(r.Role),
        children: str(r['Child(ren)']),
        division: str(r.Division),
        borough: str(r.Borough),
        stamps: [{ tab: spec.tab, row: r._row, header: SENT }]
      };
      byEmail[key] = person;
      people.push(person);
    });
  });

  return people;
}

/**
 * The parents listed in the Parent 1 / Parent 2 columns of the Students tab.
 * A different list from the Parents tab: that one is parents who are working,
 * this one is every working student's parents, whether they are helping or not.
 */
function parentsOfStudents() {
  var byEmail = {};
  var people = [];

  readTab(STUDENTS).forEach(function (s) {
    var child = personName(s);
    [['Parent 1 name', 'Parent 1 email'], ['Parent 2 name', 'Parent 2 email']].forEach(function (pair) {
      var email = str(s[pair[1]]);
      if (!email) return;

      var key = email.toLowerCase();
      var stamp = { tab: STUDENTS, row: s._row, header: PARENT_SENT };

      if (byEmail[key]) {
        byEmail[key].stamps.push(stamp);
        if (child && byEmail[key].childList.indexOf(child) === -1) byEmail[key].childList.push(child);
        return;
      }

      var person = {
        name: str(s[pair[0]]) || 'there',
        first: str(s[pair[0]]).split(/\s+/)[0] || 'there',
        email: email,
        group: 'Parent',
        grade: '',
        childList: child ? [child] : [],
        division: '',
        borough: str(s.Borough),
        stamps: [stamp]
      };
      byEmail[key] = person;
      people.push(person);
    });
  });

  people.forEach(function (p) {
    p.children = p.childList.join(', ');
    delete p.childList;
  });
  return people;
}

/* ------------------------------------------------------------ send dialog */

function showSendDialog() {
  var combos = recipientCombinations();
  var counts = combos.map(function (c) { return peopleIn(c.groups).length; });
  var studentParentCount = parentsOfStudents().length;

  var recipientOptions = combos.map(function (c, i) {
    return '<option value="' + esc(c.key) + '">' + esc(c.label) +
      ' — ' + counts[i] + ' ' + (counts[i] === 1 ? 'person' : 'people') + '</option>';
  }).join('');

  var divisionOptions = ['<option value="">Any division</option>']
    .concat(DIVISIONS.map(function (d) {
      return '<option value="' + esc(d) + '">' + esc(d) + '</option>';
    })).join('');

  var wordingOptions = ['<option value="__own__">Each person’s own group wording</option>']
    .concat(readTab(TEMPLATES).map(function (t) {
      return '<option value="' + esc(t.Group) + '">' + esc(t.Group) + ' wording</option>';
    })).join('');

  var html = [
    '<style>',
    'body{font:13px/1.5 Arial,sans-serif;margin:0;padding:14px;color:#202124}',
    'label{display:block;font-weight:bold;margin:12px 0 4px}',
    'select{width:100%;padding:6px;font:13px Arial}',
    '.check{margin-top:10px;font-weight:normal;display:flex;gap:7px;align-items:flex-start}',
    '.btns{margin-top:18px;display:flex;gap:8px;flex-wrap:wrap}',
    'button{padding:8px 12px;font:13px Arial;cursor:pointer}',
    '#status{margin-top:14px;min-height:34px;white-space:pre-wrap}',
    '#out{width:100%;height:100px;font:12px monospace;margin-top:8px;display:none}',
    '.hint{color:#5f6368;font-size:12px;margin-top:4px}',
    '</style>',
    '<label for="groups">Send to</label>',
    '<select id="groups" onchange="syncWording()">', recipientOptions, '</select>',
    '<label class="check"><input type="checkbox" id="kin"><span>Also the parents listed on the <b>Students</b> tab (' +
      studentParentCount + ') — every working student’s parents, not just the ones helping out</span></label>',
    '<label for="division">Parent division</label>',
    '<select id="division">', divisionOptions, '</select>',
    '<div class="hint">Filters the <b>Parents</b> tab only, since that is the tab with a Division column. ' +
      'A parent marked <b>LS/MS</b> is included by both LS and MS.</div>',
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
    '    groups:sel("groups").value, wording:sel("wording").value,',
    '    division:sel("division").value, kin:sel("kin").checked, mode:mode });',
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
    HtmlService.createHtmlOutput(html).setWidth(440).setHeight(640), 'Send Open House emails');
}

/** Called from the dialog. Returns {message, addresses}. */
function runSend(payload) {
  var groups = str(payload.groups).split('|').filter(String);
  var people = groups.length ? peopleIn(groups) : [];

  // "LS/MS" means children in both, so such a parent matches either filter.
  // Anyone with no division set is left in rather than silently dropped.
  var division = str(payload.division);
  if (division) {
    people = people.filter(function (p) {
      if (p.group !== 'Parent' || !p.division) return true;
      return p.division === division || p.division === 'LS/MS' || division === 'LS/MS';
    });
  }

  var kin = parentsOfStudents();

  if (payload.kin) {
    var known = {};
    people.forEach(function (p) { known[p.email.toLowerCase()] = true; });
    kin.forEach(function (p) {
      if (!known[p.email.toLowerCase()]) { known[p.email.toLowerCase()] = true; people.push(p); }
    });
  }

  // A parent on the Parents tab whose Child(ren) cell is still empty: work the
  // children out from the Students tab so {{children}} isn't blank. Running
  // "Fill in children on the Parents tab" makes this permanent in the sheet.
  var childrenByEmail = {};
  kin.forEach(function (p) { childrenByEmail[p.email.toLowerCase()] = p.children; });
  people.forEach(function (p) {
    if (p.group === 'Parent' && !p.children) {
      p.children = childrenByEmail[p.email.toLowerCase()] || '';
    }
  });

  if (!people.length) {
    return { message: 'Nobody to email. Check the ' + (groups.map(function (g) { return GROUP_TABS[g].tab; }).join(' and ') || 'group') + ' tab has names with email addresses.' };
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
    var key = payload.wording === '__own__' ? 'Combined' : payload.wording;
    var template = templateFor(key);
    if (!template) return { message: 'No "' + key + '" row on the Templates tab.' };
    if (!str(template.body)) {
      return { message: 'The "' + key + '" row on the Templates tab has no wording in it yet. ' +
        'Write the email there first, then come back.' };
    }

    var inbox = me();
    if (!inbox) return { message: 'Could not work out your own email address, which the To field of a BCC draft needs. Use "One draft per person" instead.' };

    var greeting = { name: groups.length === 1 ? plural(groups[0]).toLowerCase() : 'everyone' };
    GmailApp.createDraft(inbox, fill(template.subject, greeting, sender), fill(template.body, greeting, sender), {
      bcc: people.map(function (p) { return p.email; }).join(',')
    });
    return { message: 'One DRAFT created, addressed to you with ' + people.length +
      ' BCC recipient(s), using the ' + key + ' wording.\n\n' +
      'Nothing has been sent. Open Gmail, read it, and press Send yourself. ' +
      'Afterwards, run "Update who has been emailed" to fill in the ' +
      '"' + SENT + '" column from what actually left your account.' };
  }

  var missing = {};
  var blank = {};
  var made = 0;
  people.forEach(function (person) {
    var key = payload.wording === '__own__' ? person.group : payload.wording;
    var template = templateFor(key);
    if (!template) { missing[key] = true; return; }
    if (!str(template.body)) { blank[key] = true; return; }

    person.room = rooms[person.name.toLowerCase()] || '';
    GmailApp.createDraft(person.email, fill(template.subject, person, sender), fill(template.body, person, sender));
    made++;
  });

  var empties = Object.keys(blank);
  var gaps = Object.keys(missing);

  if (!made) {
    if (empties.length) {
      return { message: 'Nothing sent: the "' + empties.join('" and "') +
        '" row(s) on the Templates tab have no wording in them yet. Write the email there first.' };
    }
    if (gaps.length) {
      return { message: 'Nothing sent: there is no "' + gaps.join('" or "') +
        '" row on the Templates tab. Run "Set up / repair sheet".' };
    }
  }

  var message = made + ' DRAFT' + (made === 1 ? '' : 'S') + ' created in Gmail.\n\n' +
    'Nothing has been sent. Open Drafts, read them, and press Send yourself. ' +
    'Afterwards, run "Update who has been emailed" to fill in the "' + SENT +
    '" column from what actually left your account.';
  if (empties.length) {
    message += '\n\nSkipped anyone needing the "' + empties.join('" or "') +
      '" wording -- that row on the Templates tab is still blank.';
  }
  if (gaps.length) {
    message += '\n\nSkipped anyone needing a "' + gaps.join('" or "') +
      '" template -- no such row on the Templates tab.';
  }
  return { message: message };
}

/* ------------------------------------------------------ the room picker */

/**
 * Google Sheets validation allows one value per cell, so a cell cannot hold a
 * multi-select dropdown. This dialog does the job instead: pick a room, tick
 * the teachers and students who belong in it, and their names are written into
 * that room's row on the Rooms tab -- the same cells you could type into by
 * hand, so everything that reads the Rooms tab keeps working.
 */
function showRoomDialog() {
  var data = roomDialogData();
  if (!data.rooms.length) {
    alert('The Rooms tab has no rooms yet. Add a room name in the Room column first.');
    return;
  }
  if (!data.students.length && !data.teachers.length) {
    alert('There is nobody to assign yet. Add people to the Students and Teachers tabs first.');
    return;
  }

  // JSON inlined into the page; < is escaped so it can never end the script.
  var payload = JSON.stringify(data).replace(/</g, '\\u003c');

  var html = [
    '<style>',
    'body{font:13px/1.5 Arial,sans-serif;margin:0;padding:14px;color:#202124}',
    'label.top{display:block;font-weight:bold;margin:0 0 4px}',
    'select,input[type=search]{width:100%;padding:6px;font:13px Arial;box-sizing:border-box}',
    '.cols{display:flex;gap:14px;margin-top:14px;align-items:flex-start}',
    '.col{flex:1 1 0;min-width:0}',
    '.col h3{font:bold 13px Arial;margin:0 0 6px}',
    '.list{border:1px solid #dadce0;border-radius:6px;height:230px;overflow:auto;padding:6px}',
    '.row{display:flex;gap:6px;align-items:flex-start;padding:3px 2px}',
    '.row label{min-width:0;word-break:break-word}',
    '.meta{color:#5f6368}',
    '.warn{color:#b26500}',
    '.count{color:#5f6368;font-size:12px;margin-top:5px}',
    '.btns{margin-top:16px;display:flex;gap:8px;flex-wrap:wrap;align-items:center}',
    'button{padding:8px 14px;font:13px Arial;cursor:pointer}',
    '#status{margin-top:12px;min-height:32px;white-space:pre-wrap}',
    '</style>',
    '<label class="top" for="room">Room</label>',
    '<select id="room"></select>',
    '<div class="cols">',
    '<div class="col"><h3>Teachers</h3><div class="list" id="tlist"></div><div class="count" id="tcount"></div></div>',
    '<div class="col"><h3>Students</h3>',
    '<input type="search" id="find" placeholder="Filter students" autocomplete="off">',
    '<div class="list" id="slist" style="margin-top:6px"></div><div class="count" id="scount"></div></div>',
    '</div>',
    '<div class="btns">',
    '<button id="save" type="button">Save to this room</button>',
    '<button id="clear" type="button">Untick all</button>',
    '</div>',
    '<div id="status"></div>',
    '<script>',
    'var DATA = ', payload, ';',
    'var picked = {};',  // room -> {teachers:{}, students:{}}
    'function sel(id){return document.getElementById(id);}',
    'function roomNow(){return sel("room").value;}',
    // Start from what each room already holds, so nothing is lost on save.
    'DATA.rooms.forEach(function(r){',
    '  var t={},s={};',
    '  r.teachers.forEach(function(n){t[n]=true;});',
    '  r.students.forEach(function(n){s[n]=true;});',
    '  picked[r.room]={teachers:t,students:s};',
    '});',
    'DATA.rooms.forEach(function(r){',
    '  var o=document.createElement("option");',
    '  o.value=r.room;',
    '  o.textContent=r.room+(r.location?" \u2014 "+r.location:"");',
    '  sel("room").appendChild(o);',
    '});',
    // Where is this person ticked, other than the room on screen?
    'function elsewhere(kind,name){',
    '  var hits=[];',
    '  Object.keys(picked).forEach(function(room){',
    '    if(room!==roomNow() && picked[room][kind][name]) hits.push(room);',
    '  });',
    '  return hits;',
    '}',
    'function draw(kind,people,box,counter,filter){',
    '  box.textContent="";',
    '  var shown=0;',
    '  people.forEach(function(p){',
    '    if(filter && p.name.toLowerCase().indexOf(filter)===-1 &&',
    '       (p.detail||"").toLowerCase().indexOf(filter)===-1) return;',
    '    shown++;',
    '    var row=document.createElement("div"); row.className="row";',
    '    var cb=document.createElement("input");',
    '    cb.type="checkbox"; cb.id=kind+"-"+shown;',
    '    cb.checked=!!picked[roomNow()][kind][p.name];',
    '    cb.addEventListener("change",function(){',
    '      if(cb.checked) picked[roomNow()][kind][p.name]=true;',
    '      else delete picked[roomNow()][kind][p.name];',
    '      render();',
    '    });',
    '    var lab=document.createElement("label");',
    '    lab.htmlFor=cb.id;',
    '    lab.appendChild(document.createTextNode(p.name));',
    '    if(p.detail){',
    '      var m=document.createElement("span"); m.className="meta";',
    '      m.textContent=" \u00b7 "+p.detail; lab.appendChild(m);',
    '    }',
    '    var other=elsewhere(kind,p.name);',
    '    if(other.length){',
    '      var w=document.createElement("span"); w.className="warn";',
    '      w.textContent=" \u00b7 also in "+other.join(", "); lab.appendChild(w);',
    '    }',
    '    row.appendChild(cb); row.appendChild(lab); box.appendChild(row);',
    '  });',
    '  var n=Object.keys(picked[roomNow()][kind]).length;',
    '  counter.textContent=n+" ticked"+(shown<people.length?" \u00b7 "+shown+" of "+people.length+" shown":"");',
    '}',
    'function render(){',
    '  draw("teachers",DATA.teachers,sel("tlist"),sel("tcount"),"");',
    '  draw("students",DATA.students,sel("slist"),sel("scount"),sel("find").value.trim().toLowerCase());',
    '}',
    'sel("room").addEventListener("change",function(){ sel("status").textContent=""; render(); });',
    'sel("find").addEventListener("input",render);',
    'sel("clear").addEventListener("click",function(){',
    '  picked[roomNow()]={teachers:{},students:{}}; render();',
    '});',
    'sel("save").addEventListener("click",function(){',
    '  var room=roomNow();',
    '  sel("status").textContent="Saving\u2026";',
    '  sel("save").disabled=true;',
    '  google.script.run.withSuccessHandler(function(message){',
    '    sel("save").disabled=false; sel("status").textContent=message;',
    '  }).withFailureHandler(function(e){',
    '    sel("save").disabled=false; sel("status").textContent="Error: "+e.message;',
    '  }).saveRoomAssignment({ room: room,',
    '    teachers: Object.keys(picked[room].teachers),',
    '    students: Object.keys(picked[room].students) });',
    '});',
    'render();',
    '<\/script>'
  ].join('');

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(620).setHeight(560), 'Assign people to a room');
}

/** Rooms, plus everyone who could be put in one. */
function roomDialogData() {
  return {
    rooms: readTab(ROOMS).filter(function (r) { return str(r.Room); }).map(function (r) {
      return {
        room: str(r.Room),
        location: str(r.Location),
        teachers: namesIn(r.Teachers),
        students: namesIn(r.Students)
      };
    }),
    teachers: readTab(GROUP_TABS.Teacher.tab).filter(function (r) { return personName(r); })
      .map(function (r) {
        return {
          name: personName(r),
          detail: [str(r.Role), str(r.Division), str(r.Grade), str(r.Subject)]
            .filter(String).join(' \u00b7 ')
        };
      }),
    students: readTab(STUDENTS).filter(function (r) { return personName(r); })
      .map(function (r) {
        return { name: personName(r), detail: str(r.Grade) ? 'Grade ' + str(r.Grade) : '' };
      })
  };
}

/** Called from the dialog: writes one room's people into its row. */
function saveRoomAssignment(payload) {
  var room = str(payload.room);
  if (!room) return 'No room given.';

  var match = null;
  readTab(ROOMS).forEach(function (r) {
    if (!match && str(r.Room).toLowerCase() === room.toLowerCase()) match = r;
  });
  if (!match) return 'Room "' + room + '" is no longer on the Rooms tab.';

  var sheet = sheetFor(ROOMS);
  var teacherColumn = columnOf(ROOMS, 'Teachers');
  var studentColumn = columnOf(ROOMS, 'Students');
  if (!teacherColumn || !studentColumn) {
    return 'The Rooms tab is missing its Teachers or Students column. Run "Set up / repair sheet".';
  }

  var teachers = (payload.teachers || []).map(str).filter(String);
  var students = (payload.students || []).map(str).filter(String);
  sheet.getRange(match._row, teacherColumn).setValue(teachers.join(', '));
  sheet.getRange(match._row, studentColumn).setValue(students.join(', '));

  return 'Room ' + room + ' saved: ' +
    teachers.length + ' teacher' + (teachers.length === 1 ? '' : 's') + ' and ' +
    students.length + ' student' + (students.length === 1 ? '' : 's') + '.' +
    '\n\nPick another room to carry on, or close this window.';
}

/* ------------------------------------------------------- room assignments */

/** Lowercased person name -> the room they are in, for the {{room}} placeholder. */
function roomsByPerson() {
  var map = {};
  readTab(ROOMS).forEach(function (r) {
    var room = str(r.Room);
    if (!room) return;
    namesIn(r.Teachers).concat(namesIn(r.Students)).forEach(function (n) {
      var key = n.toLowerCase();
      map[key] = map[key] ? map[key] + ', ' + room : room;
    });
  });
  return map;
}

/** Splits a "Pat Chen, Alex Rivera" cell into names. Commas, semicolons or newlines. */
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
    var roomName = str(room.Room);
    if (!roomName) return;

    var teachers = namesIn(room.Teachers);
    var students = namesIn(room.Students);

    teachers.forEach(function (teacher) {
      var address = emails[teacher.toLowerCase()];
      if (!address) { unknown.push(teacher + ' (room ' + roomName + ')'); return; }

      var body = [
        'Hi ' + givenName(teacher) + ',',
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
      '\n\nAdd them to a group tab or the Team tab and run this again.';
  }
  if (!made && !unknown.length) message = 'No teachers are listed in the Teachers column of the Rooms tab.';
  alert(message);
}

/** Lowercased name -> email, across every group tab plus Team. */
function emailLookup() {
  var map = {};
  var tabs = GROUPS.map(function (g) { return GROUP_TABS[g].tab; }).concat([TEAM]);
  tabs.forEach(function (name) {
    readTab(name).forEach(function (r) {
      var person = personName(r);
      var email = str(r.Email);
      if (person && email && !map[person.toLowerCase()]) map[person.toLowerCase()] = email;
    });
  });
  // Parents named only on the Students tab still count as findable.
  readTab(STUDENTS).forEach(function (s) {
    [['Parent 1 name', 'Parent 1 email'], ['Parent 2 name', 'Parent 2 email']].forEach(function (pair) {
      var person = str(s[pair[0]]);
      var email = str(s[pair[1]]);
      if (person && email && !map[person.toLowerCase()]) map[person.toLowerCase()] = email;
    });
  });
  return map;
}

/* ------------------------------------------------- who has actually been sent to */

/**
 * The "''' + SENT + '''" column is a claim about what left your account, so
 * nothing fills it in except this, which reads your Sent mail and takes the
 * answer from there. Creating a draft deliberately stamps nothing: a draft is
 * not a sent email, and the column would be lying until you pressed Send.
 *
 * It finds your sent messages by the subject lines on the Templates tab, so
 * editing a subject after sending will hide those messages from it.
 */
function updateSentColumn() {
  var ui = SpreadsheetApp.getUi();
  var inbox = me();

  var subjects = [];
  TEMPLATE_KEYS.forEach(function (key) {
    var template = templateFor(key);
    var subject = template ? str(template.subject) : '';
    if (subject && subjects.indexOf(subject) === -1) subjects.push(subject);
  });

  if (!subjects.length) {
    alert('No row on the Templates tab has a subject line yet, so there is nothing to look for in your Sent mail.');
    return;
  }

  var confirm = ui.alert('Read your Sent mail?',
    'This searches your own Sent mail for these subject lines:\n\n  ' +
    subjects.join('\n  ') + '\n\nand fills in the "' + SENT +
    '" column for everyone it finds. It reads only your sent messages and changes nothing in Gmail.',
    ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  // address -> the most recent time you sent to it
  var sentTo = {};
  var found = 0;
  subjects.forEach(function (subject) {
    var query = 'in:sent subject:"' + subject.replace(/"/g, '') + '"';
    GmailApp.search(query, 0, 100).forEach(function (thread) {
      thread.getMessages().forEach(function (message) {
        // A thread holds replies too; only messages you sent count.
        if (inbox && message.getFrom().toLowerCase().indexOf(inbox.toLowerCase()) === -1) return;
        var when = message.getDate();
        var line = [message.getTo(), message.getCc(), message.getBcc()].join(',');
        (line.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || []).forEach(function (address) {
          var key = address.toLowerCase();
          if (!sentTo[key] || sentTo[key] < when) { sentTo[key] = when; }
        });
        found++;
      });
    });
  });

  if (!found) {
    alert('Found no sent messages with those subject lines.\n\n' +
      'If you have not pressed Send yet, that is the answer: the drafts are still sitting in Gmail. ' +
      'Nothing in this sheet sends an email to a parent, student, musician or teacher by itself.');
    return;
  }

  var stamped = 0;
  var wrong = [];

  GROUPS.forEach(function (group) {
    var spec = GROUP_TABS[group];
    var column = columnOf(spec.tab, SENT);
    if (!column) return;
    var sheet = sheetFor(spec.tab);

    readTab(spec.tab).forEach(function (r) {
      var email = str(r.Email).toLowerCase();
      if (!email) return;

      if (sentTo[email]) {
        sheet.getRange(r._row, column).setValue(sentTo[email]);
        stamped++;
      } else if (str(r[SENT])) {
        wrong.push({ tab: spec.tab, row: r._row, column: column, who: personName(r) });
      }
    });
  });

  // The Students tab also tracks whether the parents were written to.
  var parentColumn = columnOf(STUDENTS, PARENT_SENT);
  if (parentColumn) {
    var students = sheetFor(STUDENTS);
    readTab(STUDENTS).forEach(function (r) {
      var when = null;
      ['Parent 1 email', 'Parent 2 email'].forEach(function (header) {
        var hit = sentTo[str(r[header]).toLowerCase()];
        if (hit && (!when || when < hit)) when = hit;
      });
      if (when) {
        students.getRange(r._row, parentColumn).setValue(when);
        stamped++;
      } else if (str(r[PARENT_SENT])) {
        wrong.push({ tab: STUDENTS, row: r._row, column: parentColumn, who: personName(r) + "'s parents" });
      }
    });
  }

  var message = 'Read ' + found + ' sent message(s) and filled in ' + stamped + ' row(s).';

  if (wrong.length) {
    var list = wrong.slice(0, 15).map(function (w) {
      return '  ' + w.tab + ' row ' + w.row + ': ' + w.who;
    }).join('\n');
    var more = wrong.length > 15 ? '\n  ...and ' + (wrong.length - 15) + ' more' : '';

    var clear = ui.alert('Clear ' + wrong.length + ' stamp(s) that do not match your Sent mail?',
      'These rows say an email went out, but nothing matching was found in your Sent mail:\n\n' +
      list + more + '\n\nClear them? Say no if you sent those some other way.',
      ui.ButtonSet.YES_NO);

    if (clear === ui.Button.YES) {
      wrong.forEach(function (w) { sheetFor(w.tab).getRange(w.row, w.column).clearContent(); });
      message += '\n\nCleared ' + wrong.length + ' stamp(s) with no matching sent message.';
    } else {
      message += '\n\nLeft ' + wrong.length + ' unmatched stamp(s) alone.';
    }
  }

  alert(message);
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
    if (task.owner) (byOwner[task.owner] = byOwner[task.owner] || []).push(task);
    else unassigned.push(task);
  });

  var sent = 0;
  var unknown = [];

  Object.keys(byOwner).forEach(function (owner) {
    var address = emails[owner.toLowerCase()];
    if (!address) { unknown.push(owner); return; }

    var mine = byOwner[owner];
    var body = ['Hi ' + givenName(owner) + ',', '', 'Open House tasks that still need you:', '']
      .concat(mine.map(function (t) { return ' - ' + describe(t, today); }))
      .concat(['', 'The full list lives on the Tasks tab of the Open House sheet.', '', 'Thank you,', sender])
      .join('\n');

    var options = {};
    if (inbox && inbox.toLowerCase() !== address.toLowerCase()) options.cc = inbox;

    GmailApp.sendEmail(address,
      'Open House: ' + mine.length + ' task' + (mine.length === 1 ? '' : 's') + ' for you', body, options);

    mine.forEach(function (t) { stampReminded(t.row); });
    sent++;
  });

  if (inbox) {
    var lines = ['Open House task summary - ' +
      Utilities.formatDate(today, Session.getScriptTimeZone(), 'EEEE d MMMM'), ''];
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
      lines.push('No email address on the Team or group tabs for: ' + unknown.join(', '));
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
      task: str(r.Task),
      details: str(r.Details),
      owner: str(r.Owner),
      due: r['Due date'] instanceof Date ? r['Due date'] : null,
      status: str(r.Status)
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
  if (!quiet) alert(removed ? 'Daily reminders are off.' : 'Daily reminders were not on.');
}

/* ------------------------------------------------------------ the checker */

function checkForProblems() {
  var problems = [];
  var emails = emailLookup();
  var seen = {};

  // Every group tab: names, addresses, duplicates.
  GROUPS.forEach(function (group) {
    var spec = GROUP_TABS[group];
    readTab(spec.tab).forEach(function (r) {
      var at = spec.tab + ' row ' + r._row + ': ';
      var email = str(r.Email);
      if (!personName(r)) problems.push(at + 'no name.');
      else if (columnOf(spec.tab, FIRST) && !str(r[LAST])) {
        problems.push(at + '"' + personName(r) + '" has no last name.');
      }
      if (!email) problems.push(at + 'no email address.');
      else if (!looksLikeEmail(email)) problems.push(at + '"' + email + '" does not look like an email address.');
      else {
        var key = email.toLowerCase();
        if (seen[key] && seen[key] !== spec.tab) {
          // Being in two groups is fine and deliberate; just say so once.
          problems.push(at + email + ' is also on the ' + seen[key] + ' tab. They will be emailed once, not twice.');
        }
        seen[key] = spec.tab;
      }
    });
  });

  // Rooms are matched to people by name, so a repeated name is ambiguous.
  GROUPS.forEach(function (group) {
    var spec = GROUP_TABS[group];
    var byName = {};
    readTab(spec.tab).forEach(function (r) {
      var name = personName(r);
      if (!name) return;
      var key = name.toLowerCase();
      if (byName[key]) {
        problems.push(spec.tab + ' row ' + r._row + ': "' + name + '" is also on row ' +
          byName[key] + '. Room assignments go by name, so add a middle initial to tell them apart.');
      } else {
        byName[key] = r._row;
      }
    });
  });

  // Students: grade, and the parent columns.
  readTab(STUDENTS).forEach(function (s) {
    var at = STUDENTS + ' row ' + s._row + ': ';
    if (!str(s.Grade)) problems.push(at + 'no grade.');

    var any = false;
    [['Parent 1 name', 'Parent 1 email'], ['Parent 2 name', 'Parent 2 email']].forEach(function (pair, i) {
      var name = str(s[pair[0]]);
      var email = str(s[pair[1]]);
      if (name || email) any = true;
      if (name && !email) problems.push(at + 'parent ' + (i + 1) + ' "' + name + '" has no email address.');
      if (email && !name) problems.push(at + 'parent ' + (i + 1) + ' has an email address but no name.');
      if (email && !looksLikeEmail(email)) problems.push(at + 'parent ' + (i + 1) + ': "' + email + '" does not look like an email address.');
    });
    if (!any) problems.push(at + 'no parent details at all.');
  });

  // Parents: does each one actually match a student, and do the addresses agree?
  var studentParentEmails = {};
  readTab(STUDENTS).forEach(function (s) {
    [['Parent 1 name', 'Parent 1 email'], ['Parent 2 name', 'Parent 2 email']].forEach(function (pair) {
      var name = str(s[pair[0]]);
      var email = str(s[pair[1]]);
      if (name) studentParentEmails[name.toLowerCase()] = email;
    });
  });
  readTab(PARENTS).forEach(function (p) {
    var at = PARENTS + ' row ' + p._row + ': ';
    var name = personName(p);
    if (!name) return;
    if (!(name.toLowerCase() in studentParentEmails)) {
      problems.push(at + '"' + name + '" is not named as a parent on the Students tab. Fine if their child is not working, worth a look otherwise.');
    } else {
      var onStudents = studentParentEmails[name.toLowerCase()];
      if (onStudents && str(p.Email) && onStudents.toLowerCase() !== str(p.Email).toLowerCase()) {
        problems.push(at + '"' + name + '" has a different email here (' + str(p.Email) + ') than on the Students tab (' + onStudents + ').');
      }
    }
    if (!str(p['Child(ren)'])) {
      problems.push(at + 'no child listed. Run "Fill in children on the Parents tab".');
    }
    var division = str(p.Division);
    if (!division) problems.push(at + 'no division, so the LS / MS filter will always include them.');
    else if (DIVISIONS.indexOf(division) === -1) {
      problems.push(at + 'division "' + division + '" is not one of ' + DIVISIONS.join(', ') + '.');
    }
  });

  // Borough, where it has been filled in: a pasted value can sidestep the dropdown.
  [STUDENTS, PARENTS].forEach(function (name) {
    readTab(name).forEach(function (r) {
      var borough = str(r.Borough);
      if (borough && BOROUGHS.indexOf(borough) === -1) {
        problems.push(name + ' row ' + r._row + ': borough "' + borough +
          '" is not one of the seven options. Pick it from the dropdown.');
      }
    });
  });

  // A teacher's division, same way. Blank is fine here: the LS / MS send
  // filter reads the Parents tab, not this one.
  readTab(GROUP_TABS.Teacher.tab).forEach(function (r) {
    var division = str(r.Division);
    if (division && DIVISIONS.indexOf(division) === -1) {
      problems.push(GROUP_TABS.Teacher.tab + ' row ' + r._row + ': division "' +
        division + '" is not one of ' + DIVISIONS.join(', ') + '.');
    }
  });

  // Team
  readTab(TEAM).forEach(function (r) {
    var at = TEAM + ' row ' + r._row + ': ';
    if (!personName(r)) problems.push(at + 'no name.');
    var email = str(r.Email);
    if (!email) problems.push(at + 'no email address, so this person cannot be reminded.');
    else if (!looksLikeEmail(email)) problems.push(at + '"' + email + '" does not look like an email address.');
  });

  // Tasks
  readTab(TASKS).forEach(function (r) {
    var at = TASKS + ' row ' + r._row + ': ';
    var owner = str(r.Owner);
    var status = str(r.Status);
    if (!str(r.Task)) problems.push(at + 'no task described.');
    if (!owner) problems.push(at + 'nobody is responsible yet.');
    else if (!emails[owner.toLowerCase()]) problems.push(at + 'no email address anywhere for "' + owner + '".');
    if (status && STATUSES.indexOf(status) === -1) problems.push(at + 'status "' + status + '" is not one of ' + STATUSES.join(', ') + '.');
    if (r['Due date'] && !(r['Due date'] instanceof Date)) problems.push(at + 'the due date is text, not a date. Retype it as a date.');
  });

  // Rooms, and anyone booked into two at once.
  var placed = {};
  var roomNames = {};
  readTab(ROOMS).forEach(function (r) {
    var at = ROOMS + ' row ' + r._row + ': ';
    var room = str(r.Room);
    if (!room) { problems.push(at + 'no room name.'); return; }
    if (roomNames[room.toLowerCase()]) {
      problems.push(at + 'room "' + room + '" is also on row ' + roomNames[room.toLowerCase()] +
        '. The room picker can only write to the first one.');
    } else {
      roomNames[room.toLowerCase()] = r._row;
    }
    if (!namesIn(r.Teachers).length) problems.push(at + 'room ' + room + ' has no teacher.');

    namesIn(r.Teachers).concat(namesIn(r.Students)).forEach(function (name) {
      var key = name.toLowerCase();
      if (!emails[key]) problems.push(at + '"' + name + '" is not on any group tab or the Team tab.');
      if (placed[key] && placed[key] !== room) problems.push(at + '"' + name + '" is also in room ' + placed[key] + '.');
      else placed[key] = room;
    });
  });

  // Panel: a panellist should be a student who is actually working.
  var gradeByStudent = {};
  readTab(STUDENTS).forEach(function (r) {
    var who = personName(r);
    if (who) gradeByStudent[who.toLowerCase()] = str(r.Grade);
  });
  readTab(PANEL).forEach(function (r) {
    var at = PANEL + ' row ' + r._row + ': ';
    var who = personName(r);
    if (!who) { problems.push(at + 'no name.'); return; }

    if (!(who.toLowerCase() in gradeByStudent)) {
      problems.push(at + '"' + who + '" is not on the ' + STUDENTS + ' tab.');
      return;
    }
    var onStudents = gradeByStudent[who.toLowerCase()];
    var here = str(r.Grade);
    if (here && onStudents && here !== onStudents) {
      problems.push(at + '"' + who + '" is grade ' + here + ' here but grade ' +
        onStudents + ' on the ' + STUDENTS + ' tab.');
    }
  });

  // Templates
  var keys = {};
  readTab(TEMPLATES).forEach(function (r) {
    var key = str(r.Group);
    keys[key] = true;
    if (key && !str(r.Body)) {
      problems.push(TEMPLATES + ' row ' + r._row + ': the "' + key +
        '" email has no wording yet, so nothing can be sent to them.');
    }
  });
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
    if (str(r.Group).toLowerCase() === String(key).toLowerCase()) {
      match = { subject: str(r.Subject) || 'Open House', body: String(r.Body || '') };
    }
  });
  return match;
}

function fill(text, person, sender) {
  return String(text)
    .replace(/\{\{\s*firstname\s*\}\}/gi, person.first || person.name || '')
    .replace(/\{\{\s*name\s*\}\}/gi, person.name || '')
    .replace(/\{\{\s*email\s*\}\}/gi, person.email || '')
    .replace(/\{\{\s*grade\s*\}\}/gi, person.grade || '')
    .replace(/\{\{\s*children\s*\}\}/gi, person.children || 'your child')
    .replace(/\{\{\s*division\s*\}\}/gi, person.division || '')
    .replace(/\{\{\s*borough\s*\}\}/gi, person.borough || '')
    .replace(/\{\{\s*subject\s*\}\}/gi, person.subject || '')
    .replace(/\{\{\s*role\s*\}\}/gi, person.role || '')
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

function stampReminded(row) {
  var column = columnOf(TASKS, 'Last reminded');
  if (column) sheetFor(TASKS).getRange(row, column).setValue(new Date());
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
