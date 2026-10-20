# Character & fitness kit

A spreadsheet for keeping the records a bar application asks for, so you're not rebuilding years of addresses, jobs and tickets from memory in 3L year.

**Download:** [`dist/cf-kit.xlsx`](dist/cf-kit.xlsx) (opens in Excel; upload it to Google Drive to use it in Sheets).

## What's in it

| Tab | What it does |
|---|---|
| Start Here | How to use it, and a reminder that your file stays with you |
| Dashboard | Counts of entries, gaps and open items, with red and amber flags, and your next review date |
| Residences | Every address with dates; flags any gap over 30 days between one address and the next |
| Employment | Every job, with unemployed stretches as their own rows; flags gaps, and flags "terminated" or "asked to resign" as needing an explanation |
| Education | Every school since high school, with any discipline or probation |
| Legal Matters | Tickets (paid ones too), arrests, charges, civil cases, restraining orders, with court, docket number and outcome |
| Financial | Collections, defaults, judgments, bankruptcy, tax issues, student loans, and when each was resolved |
| Licenses | Driver's and professional licenses (record only the last four digits of any number) |
| References | Who you'll ask, and whether they agreed |
| Law School App Check | Your law school application answers next to what your records show now, marked match or amend |
| Documents | What to request and from where; requests over 30 days old turn amber |
| Change Log | Each update, which also sets the Dashboard's last-updated date |

The full layout, column by column, is in [SPEC.md](SPEC.md).

## Why the law school application check matters

A bar application that discloses something your law school application didn't is a common character and fitness problem. If anything on that tab says **Amend**, talk to the Dean of Students about amending your law school application before you file for the bar.

## Privacy

The kit is a blank template. Nothing you type reaches LITSA or this repository. Keep your copy private and backed up, and don't commit a filled-in copy anywhere.

## State notes

`states/` holds notes on what each jurisdiction asks, starting with Massachusetts and New York. They're being researched; until they're marked reviewed, rely on the actual application.

## Rebuild the spreadsheet

```
pip install -r requirements.txt
python3 cf-kit/build_cfkit.py
```

Then open `cf-kit/dist/cf-kit.xlsx` in Excel and confirm it opens without a repair prompt before committing.

*Not legal advice. Always answer the questions on the actual application, completely and truthfully.*
