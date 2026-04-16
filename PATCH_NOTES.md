# Patch Notes

## Release 1.28.2
Date: 2026-04-16
Type: Patch

### Version Control
- Previous version: 1.28.1
- Current version: 1.28.2
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
This patch improves Google Sheets synchronization so employee record corrections are reflected quickly in the app.

### What Changed
1. Removed the extra in-process memoization layer on sheet reads to prevent stale data persistence.
2. Added a manual sidebar action: **Sync Data from Google Sheets Now**.
3. Added scheduled auto-resync behavior based on a configurable interval.
4. Added new optional secret: `app.data_sync_minutes` (default is 5 minutes).
5. Centralized session data reload logic for consistent refresh behavior.

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/usr/local/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/usr/local/bin/python test_setup.py`
3. Runtime smoke test:
   - `streamlit run streamlit_app.py --server.headless true --server.port 8501`

## Release 1.28.1
Date: 2026-04-16
Type: Patch

### Version Control
- Previous version: 1.28.0
- Current version: 1.28.1
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
This patch adds a manager bulk reminder workflow for self-evaluation completion, fixes a self-evaluation display regression, and validates app startup and integration checks.

### What Changed
1. Added bulk reminder button for managers in the Submit Scorecard tab.
2. Added filtering logic to identify employees who:
   - are under the logged-in manager,
   - still need a scorecard submitted,
   - and have not submitted a self-evaluation yet.
3. Added one-click send flow that emails all incomplete employees in that filtered set.
4. Added aggregated feedback after send (sent count, failed count, and failed recipients list).
5. Added app link fallback guidance when email delivery fails.
6. Fixed compact self-evaluation rendering bug by correcting `question.aget(...)` to `question.get(...)`.

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/usr/local/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/usr/local/bin/python test_setup.py`
3. Runtime smoke test:
   - `streamlit run streamlit_app.py --server.headless true --server.port 8501`

### Deployment Notes
- App code is ready for deployment.
- If Streamlit Cloud is connected to this repository branch, deployment triggers automatically after pushing this commit.
