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


if __name__ == "__main__":
    unittest.main()