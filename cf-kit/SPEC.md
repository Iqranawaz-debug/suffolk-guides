# cf-kit layout

The generator (`build_cfkit.py`) is the source of truth; this page describes it for reviewers. Change both together.

- One header row per tab, frozen, in `blue-deep` with white text.
- 60 blank, pre-formatted rows per tab.
- Gray cells calculate; people shouldn't type in them.
- Dates are `yyyy-mm-dd` and validated as dates after 1950.
- Dropdowns are plain lists so they work in both Excel and Google Sheets.

## Calculated columns

| Tab | Column | Formula, in words |
|---|---|---|
| Residences | L, Days since previous address | This row's From minus the previous row's To, minus one. Red above 30. |
| Employment | K, Needs explanation | "Yes" when the reason for leaving is Terminated or Asked to resign. Red when Yes. |
| Employment | M, Days since previous job | Same as Residences. Red above 30. |
| Documents | E, Status | Received if there's a received date; Overdue if requested more than 30 days ago and not received; otherwise Requested. Amber when overdue. |
| Law School App Check | D | Red when Amend. |

Rows must be entered oldest first for the gap columns to work.

## Dashboard

Rows 2 to 14: addresses entered; address gaps over 30 days; jobs entered; employment gaps over 30 days; jobs needing an explanation; schools; legal matters; unresolved financial items; references who agreed; law school answers to amend; documents outstanding; last updated (latest Change Log date); next review (last updated plus 120 days). Counts that need attention turn red or amber above zero.
