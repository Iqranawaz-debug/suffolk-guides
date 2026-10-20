# Contributing to suffolk-guides

The org-wide guide is at [suffolklitsa/.github](https://github.com/suffolklitsa/.github/blob/main/CONTRIBUTING.md). This page covers what's specific here.

## Writing a guide

1. Comment on the guide's issue to claim it.
2. Copy `guides/_TEMPLATE.md` to `guides/<short-name>.md` (lowercase, hyphens: `booking-a-room.md`).
3. **Do the process yourself** while you write. Use the exact words the screens use. Take screenshots if they help, but crop out your name, ID and anything personal.
4. Link the official page for every step under `sources`.
5. Leave `status: draft`. Open a pull request.
6. A second person follows your steps on the real system. If they work, they change `status` to `reviewed` and update `last_checked`.

### Style

- Second person, present tense: "Click **Reserve a space**."
- One action per numbered step. Say what the person should see when it worked.
- Name the office to contact when something goes wrong, with its official email or page.
- No opinions about offices or policies; just the steps.
- Never include anyone's personal information, passwords, or internal-only links.

## Keeping guides current

When a step changes, fix it and update `last_checked`. If you can't fix it, open an issue with what you saw.

## Code

- The guide checker and feed builder use only the Python standard library.
- After changing `data/` or `scripts/build_feeds.py`, run `python3 scripts/build_feeds.py` and commit the feeds.
- After changing `cf-kit/build_cfkit.py`, rebuild the spreadsheet and open it in Excel to confirm it opens without a repair prompt before committing. Update `cf-kit/SPEC.md` to match.

## Refreshing the calendar (maintainers)

```
python3 scripts/import_calendar.py --source /srv/dev/suffolk-chatbot/data/current_sources
python3 scripts/build_feeds.py
```
