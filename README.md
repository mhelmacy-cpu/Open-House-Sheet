# Open House Needs Sheet

A Google Sheet that holds everyone and everything you need for the open house —
parents, students, musicians, teachers, your team's to-do list, and which
classrooms they're all in — and turns that into email.

Everything is in one file: **`apps-script/Code.gs`**. Paste it in once and the
script builds all five tabs for you.

## Setup (about two minutes)

1. Go to [sheets.new](https://sheets.new) to make a new Google Sheet. Name it
   something like *Open House*.
2. In the menu bar: **Extensions → Apps Script**. A code editor opens in a new tab.
3. Delete the few lines already in the editor (`function myFunction() {}`).
4. Paste in the entire contents of `apps-script/Code.gs`.
5. Click the **Save** icon (or Ctrl/Cmd+S).
6. Go back to your sheet's tab and **reload the page**. An **Open House** menu
   appears to the right of *Help*.
7. Click **Open House → Set up / repair sheet**.

Google will ask for permission the first time you run a command — it needs to
read the sheet and create and send mail. Click through
*Review permissions → your account → Advanced → Go to (project, unsafe) → Allow*.
That "unsafe" warning is Google's standard wording for a script you wrote
yourself rather than installed from their store.

> **Re-run "Set up / repair sheet" whenever you add someone to the Team tab.**
> That's what refreshes the owner dropdown on the Tasks tab. It never overwrites
> rows you've already filled in.

### Already have names in one column?

Setup splits them for you. The old **Name** column on Students and Musicians
becomes **First name**, a **Last name** column is added beside it, and each
surname moves across — at the *first* space, so a two-word surname like
`Van Der Berg` stays whole. A one-word entry is left alone, and a row where
you've already typed a last name is never touched. Setup reports how many it
split, because a two-word *given* name goes the other way: `Mary Jo Smith`
becomes `Mary` / `Jo Smith`. Worth a glance down those two columns afterwards.
Re-running setup won't split anything twice.

### Coming from the earlier single-tab version?

If your sheet still has one **Roster** tab with a *Group* column, setup moves
those rows onto the new Students, Parents, Musicians and Teachers tabs for you,
matching on the group name, and skipping anyone already listed. It renames the
old tab *Roster (old)* rather than deleting it, so you can check nothing was
lost before you remove it. Students that move across will need their grade and
parent details filled in, since the old tab had nowhere to keep them.

## Drafts vs. sent — the one thing to know

The script deliberately does two different things:

| | What happens |
|---|---|
| **Open House emails** (parents, students, musicians, teachers) and **room assignments** | Become **Gmail drafts**. You read each one and press Send yourself. |
| **Task reminders** to your own team | **Send immediately** — that's the point of a reminder, and the optional daily run happens while you're away from the computer. |

## The seven tabs

Each kind of person gets their own tab, so each can carry the columns it
actually needs.

### Students — the master sheet

| First name | Last name | Email | Grade | Borough | Parent 1 name | Parent 2 name | Parent 1 email | Parent 2 email | Notes | Email confirmation sent | Confirmed attending | Parent confirmation sent |
|---|---|---|---|---|---|---|---|---|---|---|---|---|

Every student who is working. The name is split in two so the tab sorts by
surname; the parent name columns stay single. The two stamped columns track different things:
whether the **student** was emailed, and whether their **parents** were. Both
fill in automatically.

Names on the Parents, Teachers and Team tabs stay in a single **Name** column.
Everything that matches people by name — room assignments, task owners, the
Rooms tab — reads either shape, so the two kinds of tab mix freely.

### Parents — the parents who are working

| Name | Email | Child(ren) | Division | Borough | Notes | Email confirmation sent | Confirmed attending |
|---|---|---|---|---|---|---|---|

*Division* is a dropdown: **LS**, **MS**, or **LS/MS** for a parent with
children in both. *Email confirmation sent* stamps itself when a draft is
generated; *Confirmed attending* is a Yes / No / Maybe dropdown for you to fill
in as replies come back.

*Borough* is the same dropdown as on the Students tab — see below.

This is a deliberately different list from the parent columns on the Students
tab:

- **Parents tab** = parents who are *working* the open house
- **Parent 1 / Parent 2 on Students** = every working student's parents, whether
  they're helping or not

So a parent helping out appears on both, and a parent who is just a contact for
their child appears only on Students. The send dialog can reach either list, or
both at once.

The *Child(ren)* column is written for you, in the form
`Maren (9th), Kristen (4th)` — given name plus grade, one entry per child.
**This happens by itself**: Google re-runs the fill after any edit to a name,
grade or parent name on either tab, so adding a student updates their parents
without you doing anything. **Fill in children on the Parents tab** does the
same on demand if you want to force it.

It matches on email first, then name, so two parents with the same name don't
get confused. It writes plain text rather than a formula, so it survives column
edits and you can still type into it by hand — anything it can't match is left
alone and reported. That `(9th)` is also what the grade count reads, so the
format matters.

#### Borough

On both Students and Parents, a dropdown of where the family travels in from:

| | |
|---|---|
| `M - Manhattan` | `X - The Bronx` |
| `B - Brooklyn` | `J - New Jersey` |
| `Q - Queens` | `LI - Long Island` |
| | `CT - Connecticut` |

Reference only — nothing in the script filters or emails by it, so leaving it
blank breaks nothing. *Check for problems* flags a value that isn't one of the
seven, which can only happen if something was pasted in past the dropdown.

### Musicians

| First name | Last name | Email | Notes | Email confirmation sent | Confirmed attending |
|---|---|---|---|---|---|

### Teachers

| Name | Role | Division | Grade | Subject | Notes | Email | Email confirmation sent |
|---|---|---|---|---|---|---|---|

*Role* is free text — you didn't name a fixed set, so it isn't a dropdown.
*Division* is the same LS / MS / LS/MS dropdown as on the Parents tab.

If you already had the combined *Grade / subject* column, setup keeps whatever
was in it under **Grade** and adds an empty **Subject** beside it — it can't
tell which half of `Grade 4` or `Music` was which, so move anything that's
really a subject across by hand.

Teachers are the only tab without *Confirmed attending* — everyone else has it,
and on Students it's what the borough count keys off.

### Team — the people running the event

| Name | Email | Role |
|---|---|---|

Separate from the four groups above, because the people organising aren't
necessarily on any of them. This tab feeds the owner dropdown on Tasks.

### Tasks — the to-do list

| Task | Details | Owner | Due date | Status | Last reminded |
|---|---|---|---|---|---|

*Owner* is a dropdown of everyone on the Team tab. *Status* is a dropdown: Not
started, In progress, Blocked, Done. Anything not marked **Done** counts as open
and gets reminded about. *Last reminded* is stamped automatically.

### Rooms — the classrooms in use

| Room | Location | Activity | Teachers | Students | Notes |
|---|---|---|---|---|---|

Use **Assign people to a room…** rather than typing into this tab — see below.
You still can type, though: several names in the Teachers or Students cell
separated by commas, semicolons or line breaks, and `Pat Chen, Alex Rivera` all
works. Names should match a group tab or the Team tab so the script can find
email addresses; *Check for problems* tells you when one doesn't.

### The summaries

Two live count blocks, each to the right of its tab's data starting at
**column T**, past one blank column. They're built from spreadsheet formulas,
not written-in numbers, so **they recount the moment you type** — no command to
run, nothing to refresh.

**On Parents — grades represented.** Every grade from K to 12th with the number
of parents who have a child in it, read out of the `(9th)` in the *Child(ren)*
column. A parent with children in two grades counts in both. The zeros are the
useful part: they show which grades you have nobody from. Above the table, a
single figure for how many grades are covered at all.

**On Parents — parents by borough.** A count per borough, plus how many parents
have no borough filled in yet, plus the total.

**On Students — confirmed by borough.** The same borough breakdown, but a
student is only counted **once their *Confirmed attending* says Yes**. Said No
or not answered, and they don't appear in any borough row. Underneath: how many
confirmed, how many aren't coming, how many haven't answered, and the total.

Everything from that blank column rightwards belongs to the summary and is
rewritten whenever setup runs, so don't keep your own notes over there. The rest
of the script ignores it: each tab's table is taken to end at the first empty
header cell.

### Templates — the wording

One row per group, plus a **Combined** row used when one email goes to more than
one group at a time. Placeholders that get filled in:

- `{{name}}` — that person's full name
- `{{firstname}}` — their given name alone, which is what the built-in wording
  greets them with, so an email opens "Hi Sam," rather than "Hi Sam Doe,"
- `{{email}}` — their email address
- `{{grade}}` — a student's grade, or a teacher's grade/subject
- `{{children}}` — a parent's child or children
- `{{division}}` — a parent's LS / MS / LS-MS division
- `{{borough}}` — where they travel in from
- `{{subject}}` — a teacher's subject
- `{{room}}` — the room they're assigned to on the Rooms tab (or `TBC`)
- `{{sender}}` — your name (the script asks once and remembers it)

## The commands

### Send emails…

Opens a dialog with two dropdowns.

**Send to** lists every possible combination of the four groups — all 15 of them,
shortest first, each with a live headcount:

```
Parents                                     Parents + Students + Musicians
Students                                    Parents + Students + Teachers
Musicians                                   Parents + Musicians + Teachers
Teachers                                    Students + Musicians + Teachers
Parents + Students                          Everyone (all four)
Parents + Musicians
Parents + Teachers
Students + Musicians
Students + Teachers
Musicians + Teachers
```

Anyone on two tabs — a student who is also a musician, say — is counted and
emailed once, not twice.

Under it, a checkbox: **also the parents listed on the Students tab**. That's
a separate list from the Parents tab — see above — so it's a checkbox rather
than more dropdown entries. Tick it to reach every working student's parents on
top of whatever combination you picked. Anyone already in the selection isn't
added twice.

**Parent division** narrows to LS or MS. It applies to the **Parents** tab only,
since that's the tab with a Division column, and a parent marked **LS/MS** is
included by both — that's the point of the third option. Anyone with no division
set is left in rather than quietly dropped.

**Wording** decides which template to use. It defaults to *each person's own
group wording* — so a Parents + Musicians send gives every parent the parent
email and every musician the musician email, in one go. Or pick a single
template to use for everybody.

Then three buttons:

- **One draft per person** — personalized, one draft each
- **One BCC draft** — a single draft addressed to you with everyone BCC'd, so
  nobody sees anyone else's address. Being one message to many people, it always
  uses one shared wording (the **Combined** row if you left the default).
- **Just show addresses** — the addresses as a comma-separated list to copy

### Fill in children on the Parents tab

Looks each parent up against the Students tab and writes their children's names
into the *Child(ren)* column. See the Parents tab above.

### Restore the built-in wording…

*Set up / repair sheet* only ever **adds** missing template rows. That's what
stops it overwriting wording you've edited — and it's also why a change to the
built-in wording won't appear on a Templates tab you already have.

This is the command that pulls it in. It asks which one — type `Parent`, or
`all` for every row — and overwrites only what you name, leaving your other
rows untouched.

### Assign people to a room…

Google Sheets allows one value per cell, so a cell can't hold a multi-select
dropdown. This dialog does that job instead: pick a room, tick the teachers and
students who belong in it, and their names are written into that room's row on
the Rooms tab — the same cells you could type into by hand, so everything that
reads the Rooms tab keeps working.

It opens with each room's current people already ticked, shows each student's
grade next to their name, has a filter box for long lists, and flags anyone
already ticked in a different room, so you spot a double-booking as you make it
rather than at *Check for problems*.

### Email room assignments to teachers

Makes one draft per teacher per room, listing their room, location, activity,
who they're sharing with, and every student assigned to them. Tells you if it
couldn't find an email address for someone.

### Send task reminders now

Emails everyone with an open task a single digest of *their* tasks — not one
email per task — with overdue items flagged `OVERDUE` and today's flagged
`due TODAY`. You're CC'd on each, and you also get one summary covering
everything, including tasks nobody owns yet. Asks for confirmation first,
because these send for real.

### Turn on / off daily task reminders

Asks what hour you want them, then has Google run the reminders every day on its
own. Works whether or not the sheet is open. Turn it off from the same menu.

### Check for problems

Run this before you send anything. It flags:

- missing names, missing or malformed email addresses
- a borough that isn't one of the seven (only possible if pasted in)
- two people with the same name on one tab (room assignments go by name, so a
  repeat is ambiguous), and two rooms with the same name
- a parent with no division, or one that isn't LS / MS / LS-MS
- anyone on two group tabs (not an error — it just tells you they'll get one email)
- students with no grade, or no parent details at all
- a parent with a name but no email address, or an email address but no name
- **a parent whose email on the Parents tab disagrees with the Students tab**
- a parent on the Parents tab who isn't named as a parent of any working student
- parents with no child filled in yet
- statuses that aren't one of the allowed values
- tasks with no owner, or an owner with no email address anywhere
- due dates that are text rather than real dates
- rooms with no teacher
- people on the Rooms tab who aren't on any group tab or the Team tab
- **anyone booked into two rooms at once**
- missing template rows

## Good to know

- Gmail caps how many messages you can *send* per day (roughly 100 on a free
  account, 400–1,500 on Workspace). Creating drafts doesn't count against the
  cap; sending does. One BCC draft is a single message no matter how many
  recipients, so it's the safe choice for a long list.
- Re-running setup on a sheet you've been filling in is safe. It reads every
  tab by its header row rather than by column position, adds new columns in
  place, renames changed ones, and **puts the columns back into the order
  listed above** — each column's contents travelling with it. So if you
  rearrange columns by hand, the next setup run will undo it; to change the
  order for good, edit that tab's `headers` list near the top of the script.
  Columns you add yourself are left alone, after the ones the script knows
  about.
- The header row of every tab is LREI red, set by `HEADER_FILL` at the top of
  the script. Change that one value and re-run *Set up / repair sheet* to
  recolour every header on every tab.
- To add a group beyond the four, add it to the `GROUPS` list at the top of the
  script and add a matching row on the Templates tab, then re-run
  *Set up / repair sheet*. The dropdown rebuilds itself — five groups would give
  31 combinations, six would give 63.
