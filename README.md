# Open House Needs Sheet

A Google Sheet that holds everyone you need for the open house — parents,
students, musicians, teachers — and turns that list into Gmail drafts.

Everything is in one file: **`apps-script/Code.gs`**. Paste it in once and the
script builds the sheet for you.

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
read the sheet and create Gmail drafts. Click through
*Review permissions → your account → Advanced → Go to (project, unsafe) → Allow*.
That "unsafe" warning is Google's standard wording for a script you wrote
yourself rather than installed from their store.

## What you get

Two tabs:

**Roster** — one row per person.

| Name | Email | Group | Notes | Emailed? |
|---|---|---|---|---|
| Jane Doe | jane.doe@example.com | Parent | | |

*Group* is a dropdown: Parent, Student, Musician, or Teacher. That one column is
what replaces having a separate section for each kind of person — sort or filter
by it and you have your parent list, your musician list, and so on.
*Emailed?* is stamped with the date automatically when you generate a draft.

**Templates** — one row per group, holding the subject line and body for that
group's email. Edit the wording here, in the sheet. You never have to touch the
code again.

Three placeholders get filled in when a draft is built:

- `{{name}}` — that person's name from the Roster
- `{{email}}` — their email address
- `{{sender}}` — your name (the script asks once and remembers it)

## The commands

All under the **Open House** menu:

- **Set up / repair sheet** — builds or repairs both tabs. Safe to re-run; it
  won't overwrite rows you've already filled in.
- **Create a draft for each person...** — asks which group, then makes one
  personalized draft per person in it. Good for parents.
- **Create one group draft (BCC)...** — asks which group, then makes a single
  draft with everyone BCC'd, so nobody sees anyone else's address.
- **Copy email addresses...** — asks which group and shows the addresses as a
  comma-separated list you can copy into any mail app.
- **Check for problems** — flags missing names, missing or malformed email
  addresses, duplicates, and bad group names before you send anything.

**Nothing ever sends on its own.** Every command stops at your Gmail Drafts
folder so you can read each message and press Send yourself.

## Good to know

- Gmail caps how many messages you can send per day (roughly 100 on a free
  account, 400–1,500 on a Workspace account). Creating drafts doesn't count
  against it; sending does. The BCC command is one message no matter how many
  recipients, so it's the safe choice for a large list.
- To add a group beyond the four, add it to the `GROUPS` list at the top of the
  script, add a matching row on the Templates tab, then re-run
  *Set up / repair sheet* so the dropdown picks it up.
