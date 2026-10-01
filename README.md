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

## Drafts vs. sent — the one thing to know

The script deliberately does two different things:

| | What happens |
|---|---|
| **Open House emails** (parents, students, musicians, teachers) and **room assignments** | Become **Gmail drafts**. You read each one and press Send yourself. |
| **Task reminders** to your own team | **Send immediately** — that's the point of a reminder, and the optional daily run happens while you're away from the computer. |

## The five tabs

**Roster** — everyone you might email.

| Name | Email | Group | Notes | Emailed? |
|---|---|---|---|---|

*Group* is a dropdown: Parent, Student, Musician, Teacher. *Emailed?* is stamped
with the date automatically whenever a draft is generated for that person.

**Team** — the people running the event, who can own tasks.

| Name | Email | Role |
|---|---|---|

**Tasks** — your to-do list.

| Task | Details | Owner | Due date | Status | Last reminded |
|---|---|---|---|---|---|

*Owner* is a dropdown of everyone on the Team tab. *Status* is a dropdown: Not
started, In progress, Blocked, Done. Anything not marked **Done** counts as open
and gets reminded about. *Last reminded* is stamped automatically.

**Rooms** — the classrooms in use.

| Room | Location | Activity | Teachers | Students | Notes |
|---|---|---|---|---|---|

Put several names in the Teachers or Students cell separated by commas, semicolons
or line breaks — `Pat Chen, Alex Rivera` all works. Names should match the Roster
or Team tab so the script can find email addresses; *Check for problems* tells you
when one doesn't.

**Templates** — the wording, one row per group plus a **Combined** row used when
one email goes to more than one group at a time.

Four placeholders get filled in:

- `{{name}}` — that person's name
- `{{email}}` — their email address
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

Anyone in two groups is counted and emailed once, not twice.

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

- missing names, missing or malformed email addresses, duplicate addresses
- groups and statuses that aren't one of the allowed values
- tasks with no owner, or an owner with no email address anywhere
- due dates that are text rather than real dates
- rooms with no teacher
- people on the Rooms tab who aren't on the Roster or Team tab
- **anyone booked into two rooms at once**
- missing template rows

## Good to know

- Gmail caps how many messages you can *send* per day (roughly 100 on a free
  account, 400–1,500 on Workspace). Creating drafts doesn't count against the
  cap; sending does. One BCC draft is a single message no matter how many
  recipients, so it's the safe choice for a long list.
- Reordering or renaming columns on a tab won't break anything — the script
  reads each tab by its header row, not by column position.
- To add a group beyond the four, add it to the `GROUPS` list at the top of the
  script and add a matching row on the Templates tab, then re-run
  *Set up / repair sheet*. The dropdown rebuilds itself — five groups would give
  31 combinations, six would give 63.
