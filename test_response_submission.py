import unittest

from response_submission import submit_manager_scorecard


class SubmitManagerScorecardTests(unittest.TestCase):
    def test_submitted_draft_is_deleted_after_append(self):
        appended = []
        deleted = []
        response_entry = {"response_id": "submitted-1", "status": "Pending Employee"}
        selected_draft = {"response_id": "draft-1", "status": "Draft"}

        def append_response(row):
            appended.append(row["response_id"])

        def delete_response(response_id):
            deleted.append(response_id)
            return True

        submitted, created_response, error = submit_manager_scorecard(
            response_entry,
            selected_draft,
            append_response,
            delete_response,
        )

        self.assertTrue(submitted)
        self.assertEqual(created_response, response_entry)
        self.assertEqual(error, "")
        self.assertEqual(appended, ["submitted-1"])
        self.assertEqual(deleted, ["draft-1"])

    def test_failed_draft_delete_rolls_back_new_submission(self):
        appended = []
        deleted = []
        response_entry = {"response_id": "submitted-2", "status": "Pending Employee"}
        selected_draft = {"response_id": "draft-2", "status": "Draft"}

        def append_response(row):
            appended.append(row["response_id"])

        def delete_response(response_id):
            deleted.append(response_id)
            return response_id == "submitted-2"

        submitted, created_response, error = submit_manager_scorecard(
            response_entry,
            selected_draft,
            append_response,
            delete_response,
        )

        self.assertFalse(submitted)
        self.assertIsNone(created_response)
        self.assertIn("rolled back", error)
        self.assertEqual(appended, ["submitted-2"])
        self.assertEqual(deleted, ["draft-2", "submitted-2"])

    def test_submission_without_draft_does_not_delete_anything(self):
        appended = []
        deleted = []
        response_entry = {"response_id": "submitted-3", "status": "Pending Employee"}

        def append_response(row):
            appended.append(row["response_id"])

        def delete_response(response_id):
            deleted.append(response_id)
            return True

        submitted, created_response, error = submit_manager_scorecard(
            response_entry,
            None,
            append_response,
            delete_response,
        )

        self.assertTrue(submitted)
        self.assertEqual(created_response, response_entry)
        self.assertEqual(error, "")
        self.assertEqual(appended, ["submitted-3"])
        self.assertEqual(deleted, [])


    def test_delete_all_drafts_called_when_provided(self):
        """When delete_all_drafts is supplied it is used instead of the legacy single-delete path."""
        appended = []
        all_drafts_deleted = []
        response_entry = {"response_id": "sub-4", "manager_email": "mgr@test.com", "employee_id": "42", "status": "Pending Employee"}
        selected_draft = {"response_id": "draft-4", "status": "Draft"}

        def append_response(row):
            appended.append(row["response_id"])

        def delete_response(response_id):
            raise AssertionError("delete_response should not be called when delete_all_drafts is provided")

        def delete_all_drafts(manager_email, employee_id):
            all_drafts_deleted.append((manager_email, employee_id))
            return 2  # simulating 2 drafts removed

        submitted, created_response, error = submit_manager_scorecard(
            response_entry,
            selected_draft,
            append_response,
            delete_response,
            delete_all_drafts=delete_all_drafts,
        )

        self.assertTrue(submitted)
        self.assertEqual(error, "")
        self.assertEqual(appended, ["sub-4"])
        self.assertEqual(all_drafts_deleted, [("mgr@test.com", "42")])

    def test_delete_all_drafts_clears_drafts_without_selected_draft(self):
        """delete_all_drafts works even when no selected_draft was identified."""
        appended = []
        all_drafts_deleted = []
        response_entry = {"response_id": "sub-5", "manager_email": "mgr2@test.com", "employee_id": "99", "status": "Pending Employee"}

        def append_response(row):
            appended.append(row["response_id"])

        def delete_response(response_id):
            raise AssertionError("delete_response should not be called")

        def delete_all_drafts(manager_email, employee_id):
            all_drafts_deleted.append((manager_email, employee_id))
            return 0  # no drafts existed

        submitted, created_response, error = submit_manager_scorecard(
            response_entry,
            None,
            append_response,
            delete_response,
            delete_all_drafts=delete_all_drafts,
        )

        self.assertTrue(submitted)
        self.assertEqual(error, "")
        self.assertEqual(all_drafts_deleted, [("mgr2@test.com", "99")])


if __name__ == "__main__":
    unittest.main()