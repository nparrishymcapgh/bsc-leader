# Patch Notes

## Release 1.3.8
Date: 2026-04-22
Type: Patch

### Version Control
- Previous version: 1.3.7
- Current version: 1.3.8
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Resolved an issue where multiple manager evaluation rows could accumulate in the queue for the same employee, with stale drafts persisting after submission or subsequent draft saves.

### What Changed
1. Added `delete_all_manager_drafts_for_employee` — deletes every draft row matching a manager+employee pair in a single sheet pass (reverse-index safe).
2. Added `scrape_duplicate_manager_drafts` — on manager page load, removes all but the latest draft per manager+employee pair, recovering from any previously accumulated duplicates.
3. Updated `submit_manager_scorecard` to accept an optional `delete_all_drafts` callback; when provided it bulk-removes all drafts instead of only the single `selected_draft`.
4. Updated the "Save as Draft" flow to call `delete_all_manager_drafts_for_employee` both when updating an existing draft (to clear other orphaned drafts) and when creating a new one (to prevent accumulation).
5. Submission call-site now passes `delete_all_manager_drafts_for_employee` as `delete_all_drafts`.
6. Added two new unit tests covering the `delete_all_drafts` code path (both with and without a `selected_draft`).

### Files Updated
- streamlit_app.py
- response_submission.py
- test_response_submission.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/usr/local/bin/python -m py_compile streamlit_app.py response_submission.py`
2. Regression + new unit tests (5 total, all passing):
   - `/usr/local/bin/python -m unittest test_response_submission -v`
3. Environment, sheets, SMTP, and app-url integration check:
   - `/usr/local/bin/python test_setup.py`

Date: 2026-04-21
Type: Patch

### Version Control
- Previous version: 1.3.6
- Current version: 1.3.7
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Updated manager draft submission so approval submissions no longer leave the underlying draft response behind.

### What Changed
1. Added a dedicated submission helper for manager scorecards.
2. Changed draft submission flow to create a fresh approval response instead of converting the draft row in place.
3. Removed the saved draft immediately after successful submission.
4. Added rollback handling so a new approval response is deleted if draft cleanup fails.
5. Added regression tests covering draft removal and rollback behavior.

### Files Updated
- streamlit_app.py
- response_submission.py
- test_response_submission.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/workspaces/bsc-leader/.venv/bin/python -m py_compile streamlit_app.py response_submission.py test_response_submission.py`
2. Focused regression tests:
   - `/workspaces/bsc-leader/.venv/bin/python -m unittest test_response_submission.py`
3. Environment, sheets, SMTP, and app-url integration check:
   - `/workspaces/bsc-leader/.venv/bin/python test_setup.py`

## Release 1.3.6
Date: 2026-04-20
Type: Patch

### Version Control
- Previous version: 1.3.5
- Current version: 1.3.6
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Added manager comment output to the final approved balanced scorecard PDF.

### What Changed
1. Updated PDF generation to include a new `Manager Comments` section in the scorecard export.
2. Manager comments now print from the response `comments` field.
3. Multi-line manager comments preserve line breaks in the PDF.
4. If no comments are present, the PDF shows `No manager comments provided.`

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/workspaces/bsc-leader/.venv/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/workspaces/bsc-leader/.venv/bin/python test_setup.py`

## Release 1.3.5
Date: 2026-04-20
Type: Patch

### Version Control
- Previous version: 1.3.4
- Current version: 1.3.5
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Hardened self-evaluation reminder sender handling so emails use manager email as the SMTP sender context instead of silently falling back.

### What Changed
1. Updated email sending to set both:
   - the `From` header, and
   - the SMTP envelope sender (`from_addr`)
   using manager sender context when provided.
2. Added `Reply-To` to manager sender email when available.
3. Enforced manager sender requirement for self-evaluation reminder emails so these messages fail instead of falling back to default sender.

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/workspaces/bsc-leader/.venv/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/workspaces/bsc-leader/.venv/bin/python test_setup.py`

## Release 1.3.4
Date: 2026-04-20
Type: Patch

### Version Control
- Previous version: 1.3.3
- Current version: 1.3.4
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Updated outbound email sender behavior to use manager email context and enhanced self-evaluation reminder content.

### What Changed
1. Added sender resolution logic so outbound email headers prefer `manager_email` when available.
2. Updated scorecard workflow emails to pass manager email context as the sender where applicable.
3. Updated self-evaluation reminder email flow to pass the logged-in manager email as the sender.
4. Added the requested competencies resource link to self-evaluation reminders:
   - Text: `Learn more about YMCA Leadership Competencies here!`
   - URL: `https://drive.google.com/file/d/1ZboHZAlHWBv-2eqPiTEaBtqygg-9qRya/view?usp=sharing`

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/workspaces/bsc-leader/.venv/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/workspaces/bsc-leader/.venv/bin/python test_setup.py`

## Release 1.3.3
Date: 2026-04-17
Type: Patch

### Version Control
- Previous version: 1.3.2
- Current version: 1.3.3
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Updated employee self-evaluation behavior so submitted responses remain editable until the manager scorecard enters approval, then lock automatically.

### What Changed
1. Replaced employee self-evaluation "delete and start over" flow with in-place edit and update.
2. Added manager-status lock logic for employee self-evaluations:
   - Editable when there is no manager scorecard for the employee.
   - Editable when the latest manager scorecard status is `Draft`.
   - Read-only once the latest manager scorecard status is any non-`Draft` status (approval/in-progress/completed).
3. Added status-aware lock messaging on the employee dashboard.
4. Preserved editability after manager rejection reset by relying on the manager scorecard reset/removal behavior.

### Files Updated
- streamlit_app.py
- README.md
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/workspaces/bsc-leader/.venv/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/workspaces/bsc-leader/.venv/bin/python test_setup.py`

---

## Release 1.3.2
Date: 2026-04-16
Type: Patch

### Version Control
- Previous version: 1.3.1
- Current version: 1.3.2
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Improved PDF layout so all content wraps cleanly within a letter-size page with 1-inch margins.

### What Changed
1. Set explicit 1-inch margins (top, bottom, left, right) on the PDF document.
2. Computed usable column widths proportionally from the available page width.
3. Replaced plain string cells with `Paragraph`-wrapped cells so long question and response text wraps within column boundaries instead of overflowing.
4. Reduced base font size to 8pt with 10pt leading for a compact, readable layout.
5. Added bold header row font for all tables.
6. Section/question/answer columns: 15% / 50% / 35%.
7. Approval table columns: 20% / 15% / 65%.

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

---

## Release 1.3.1
Date: 2026-04-16
Type: Hotfix

### Version Control
- Previous version: 1.3
- Current version: 1.3.1
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
Fixed a regression where executives were incorrectly routed to the Employee Dashboard instead of the Executive Dashboard after login.

### What Changed
1. Corrected duplicate `elif` block structure in the main routing chain.
2. Executive dashboard (`elif user_role == 'executive':`) is now properly placed in the dashboard routing chain so executives land on the correct dashboard.

### Files Updated
- streamlit_app.py
- PATCH_NOTES.md

---

## Release 1.3
Date: 2026-04-16
Type: Feature Release

### Version Control
- Previous version: 1.28.2
- Current version: 1.3
- Repository: nparrishymcapgh/bsc-leader
- Branch: main

### Summary
This release introduces manager draft support, executive login and branch dashboards, approved-scorecard PDF export, and confirmation prompts before all mass-email actions.

### What Changed
1. Added manager draft capability in scorecard submission flow.
2. Added Save as Draft button that stores the current scorecard state per employee without entering approval workflow.
3. Added draft restoration so managers can return to an employee and continue from saved answers/comments.
4. Added executive login section using the Executives sheet (`executive_email` + `password`).
5. Added executive dashboard view of submitted scorecards and statuses scoped to executive branch assignments.
6. Added executive mass-email workflow to notify managers missing employee scorecards, including employee lists.
7. Added executive admin actions for `nparrish@ymcapgh.org`:
   - Email all executive passwords.
   - Email all managers everywhere who are missing reviews.
8. Added approved-only PDF export for managers and executives by employee response.
9. PDF output now includes:
   - Balanced scorecard responses,
   - Employee self-evaluation,
   - Approval decisions and timestamps.
10. Added mandatory Yes/No confirmation before all mass-email operations.

### Files Updated
- streamlit_app.py
- requirements.txt
- PATCH_NOTES.md

### Testing and Debugging Completed
1. Python syntax compile check:
   - `/usr/local/bin/python -m py_compile streamlit_app.py`
2. Environment, sheets, SMTP, and app-url integration check:
   - `/usr/local/bin/python test_setup.py`

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
