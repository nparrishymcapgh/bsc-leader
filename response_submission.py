def submit_manager_scorecard(
    response_entry,
    selected_draft,
    append_response,
    delete_response,
):
    append_response(response_entry)

    if not selected_draft:
        return True, response_entry, ""

    if delete_response(selected_draft["response_id"]):
        return True, response_entry, ""

    rollback_deleted = delete_response(response_entry["response_id"])
    if rollback_deleted:
        return False, None, "Unable to clear the saved draft, so the submission was rolled back. Please try again."

    return False, None, "Unable to clear the saved draft. Please confirm whether the submission was created before retrying."