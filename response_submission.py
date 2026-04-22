def submit_manager_scorecard(
    response_entry,
    selected_draft,
    append_response,
    delete_response,
    delete_all_drafts=None,
):
    append_response(response_entry)

    manager_email = response_entry.get("manager_email", "")
    employee_id = response_entry.get("employee_id", "")

    # Prefer bulk-delete to remove ALL drafts for this manager+employee pair.
    if delete_all_drafts is not None:
        deleted = delete_all_drafts(manager_email, employee_id)
        # If no drafts existed at all that's still success.
        return True, response_entry, ""

    # Fallback: delete only the explicitly selected draft (legacy path).
    if not selected_draft:
        return True, response_entry, ""

    if delete_response(selected_draft["response_id"]):
        return True, response_entry, ""

    rollback_deleted = delete_response(response_entry["response_id"])
    if rollback_deleted:
        return False, None, "Unable to clear the saved draft, so the submission was rolled back. Please try again."

    return False, None, "Unable to clear the saved draft. Please confirm whether the submission was created before retrying."