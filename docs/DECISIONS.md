# Decisions

**2026-10-20. Guides are Markdown with front matter** (status, owner, last_checked, sources), checked in CI. Draft until a second person walks through every step; reviewed guides go stale after 180 days.

**2026-10-20. Calendar data comes from the Suffolk assistant's verified sources**, not a second scraper. Feeds use stable event IDs so subscribers update in place.

**2026-10-20. cf-kit is a generated spreadsheet.** The generator is reviewed like code; the layout is described in cf-kit/SPEC.md; colors come from the design system. Only features that work in both Excel and Google Sheets.

**2026-10-19. cf-kit lives here,** not in its own repo; it's mostly a template and a guide.
