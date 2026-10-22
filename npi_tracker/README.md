# NPI tracker: table-driven steps

A replacement for the `UpdateDateV2` macro in `NPI_tracker.xlsm`. The process steps (PFD, TM, PAWI, ...)
now live in a **Steps** table instead of being written into the code by column number. Adding, retiring or
changing a step no longer needs a code change.

| File | What it is |
|---|---|
| `NPITracker.bas` | The macros you use: `Update_Date`, `CheckLayout`, `AddStep`, `RebuildStepFormatting` |
| `NPISetup.bas` | One-time setup (`SetupStepsTable`). Delete it after it has run. |

## How it works

- **Steps sheet, table `tblSteps`.** One row per step:

  | Column | Meaning |
  |---|---|
  | Code | Short unique code, e.g. `PAWI`. Also the key in the Emails log. |
  | Step name | Name used in the emails, and the column title when `AddStep` inserts the step |
  | Section | PROCESS, PACKAGING, Supervisor, EHS, Maintenance (used by `AddStep` for the title in row 7) |
  | Anchor | `PB` (PI/PB sample build date) or `SOP` |
  | Offset | Working days from the anchor date. Negative = before, e.g. `-5` = one week before. |
  | Auto Ready (Pkg/RP) | `Y` = set to Ready on Packaging and Raw part rows |
  | Active | `N` = the macro ignores this step |
  | Owner, Email, CC | Who gets the reminders. Replaces Items_PIC. |
  | Due rule, Past due | Calculated for you; don't type in these columns |

- **Row 1 of NPI_Components.** Each step's code sits above its Due date column, replacing the old column
  numbers. The macro finds columns by these codes, so columns can be inserted or moved freely.
- **Overall status.** PB status covers the PB-anchored steps. SOP status covers every active step. This is
  the same split as before.
- **Settings table `tblSettings`** (Steps sheet): Run password, SCM plan file, SCM plan sheet, Tracker
  link, Send emails. These used to be written in the code.

## Switch-over (once)

1. **Save a copy** of `NPI_tracker.xlsm`.
2. Press Alt+F11, then use File > Import File to import `NPITracker.bas` and `NPISetup.bas`.
3. Open the old module `UpdateDateV2` and keep it open. Setup asks you to paste three values from it:
   `SCM_PLAN_PATH`, the link after `Click <a href=`, and `PASSWORD`.
4. Back in Excel, press Alt+F8 and run `SetupStepsTable`. It:
   - checks that NPI_Components still has the expected layout (and stops if not)
   - adds the Steps sheet, copying Owner / Email / CC from Items_PIC (DFMEA's contacts go to RPSV)
   - replaces `L1:BA1` with the step codes
   - points the Update button at `NPITracker.Update_Date`
   - rebuilds the status colours and dropdowns
5. Check Owner / Email / CC on the Steps sheet, then run `CheckLayout`.
6. **Trial run:** set Send emails = `N`, run `Update_Date`, and spot-check due dates and statuses.
   Projects whose build dates did not change should come out the same as before. Then set Send emails back to `Y`.
7. Delete the modules `Old`, `OldV2`, `Module1`, `UpdateDateV2` and `NPISetup`. Save.

## Day to day

- **Add a step:** add a row to `tblSteps`, then run `AddStep`. It:
  - inserts the two columns after the step listed above it in the table
  - copies that step's formatting
  - writes the code, title and "Due date / Status"
  - places the title under its section
  - updates the colours and dropdowns

  Then run `Update_Date`.
- **Retire a step:** set Active = `N`. The macro stops touching its columns, and you can delete them
  whenever you like.
- **Change a rule or a person:** edit Anchor, Offset, Owner, Email or CC. The next `Update_Date` uses it.
- **Something looks wrong:** run `CheckLayout`. It lists every mismatch between the table and the sheet,
  for example:
  - a code missing from row 1
  - a code in row 1 that's not in the table
  - a code that isn't above a "Due date" column
  - a step with no email

  `Update_Date` runs the same check first and stops with that list instead of writing wrong data.

## Changes in behaviour compared with UpdateDateV2

- **DFMEA is now RPSV**, "Receive Packing Specs from Vendor". Past-due emails used to call it "DFMEA".
  - Reminders already logged under the old code don't count for the new one. A step due on the next working
    day at the switch-over may get one extra reminder.
- **Emails are grouped by recipient (To + CC), not by owner name.** Before, two people with the same team
  name (for example ME/PE Team: PFD/TM/PFMEA go to one person, TSE/MHO/TSA/PRWI to another) got one email,
  sent to whichever of them came first.
- **Columns in the "NPI Tracker Reminder" email are fixed.** The rows were Item / Project # / Project Name
  under the headings Project # / Project Name / Item.
- **Only the columns the macro calculates are written back:** M, O, and each step's Due date and Status.
  Before, the macro rewrote the whole A:BC block.
- **Status colours are rebuilt** as 3 rules across all status columns, replacing about 100 fragmented ones.
  - The status dropdown is on every status column and no longer on some Due date columns.
  - The old dark-red due-date rule compared against `L1`, which held the number 12, so it never fired. It
    has been removed.
- **Items_PIC is no longer read** for contacts. Keep it for the notes or delete it.

These modules were checked for syntax in LibreOffice's VBA mode and checked against the old macro's
step definitions, but not run in Excel. Do the trial run in step 6 before relying on them.
