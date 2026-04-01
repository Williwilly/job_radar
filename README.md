# Job Search Automation (RSS → Excel)

- Built to solve a personal problem — tracking remote product analyst and data roles across multiple job boards without manual searching daily.
Runs on a schedule via Windows Task Scheduler and outputs to Excel with application status tracking.

Pulls remote jobs from RSS feeds, filters by keywords, and writes to `jobs.xlsx` with:
- **New Jobs** (this run only)
- **Seen Jobs** (history)
- **Status** dropdown: Not Applied / Applied / Interview / Rejected

## Run
```bash
python job_search_automation_excel_status.py
