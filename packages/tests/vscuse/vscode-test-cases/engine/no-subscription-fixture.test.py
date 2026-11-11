import base64
import contextlib
import io
import json
import os
from pathlib import Path
import unittest
from unittest.mock import patch
import urllib.error


class NoSubscriptionFixtureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        plan_path = Path(__file__).parent.parent / "plans" / "feature-sign-in-no-subscription.json"
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        step = next(step for step in plan["steps"] if step["step_id"].startswith("step_verifyNoAzureSubscriptions_"))
        cls.script = step["parameters"]["sample"].split("python3 - <<'PY'\n", 1)[1].rsplit("\nPY\n```", 1)[0]

    def token(self, account="fixture@example.test"):
        claims = base64.urlsafe_b64encode(json.dumps({"upn": account}).encode()).decode().rstrip("=")
        return {"access_token": "header." + claims + ".signature"}

    def run_fixture(self, responses, environment=None):
        requests = []
        pending = iter(responses)

        def open_request(request, timeout):
            requests.append(request)
            self.assertEqual(timeout, 60)
            value = next(pending)
            if isinstance(value, Exception):
                raise value
            response = io.BytesIO(value if isinstance(value, bytes) else json.dumps(value).encode())
            response.status = 200
            return response

        output = io.StringIO()
        errors = io.StringIO()
        status = 0
        credentials = {"AZURE_NO_SUB_ACCOUNT_NAME": "fixture@example.test", "M365_ACCOUNT_PASSWORD": "fake-private-password"}
        if environment is not None:
            credentials = environment
        with patch.dict(os.environ, credentials, clear=True), patch("urllib.request.urlopen", side_effect=open_request), contextlib.redirect_stdout(output), contextlib.redirect_stderr(errors):
            try:
                exec(compile(self.script, "generated-fixture", "exec"), {})
            except SystemExit as error:
                status = error.code
        combined = output.getvalue() + errors.getvalue()
        self.assertNotIn("fake-private-password", combined)
        self.assertNotIn("header.", combined)
        self.assertNotIn("fixture@example.test", combined)
        return status, combined, requests

    def test_VCB_198_empty_subscriptions_succeed_for_same_user(self):
        status, output, requests = self.run_fixture([self.token(), {"value": []}])
        self.assertEqual(status, 0)
        self.assertIn("VSCUSE_NO_AZURE_SUBSCRIPTIONS_VERIFIED", output)
        self.assertEqual(requests[0].full_url, "https://login.microsoftonline.com/example.test/oauth2/v2.0/token")
        self.assertIn(b"username=fixture%40example.test", requests[0].data)
        self.assertEqual(requests[1].full_url, "https://management.azure.com/subscriptions?api-version=2022-12-01")

    def test_VCB_198_empty_pages_are_all_verified(self):
        continuation = "https://management.azure.com/subscriptions?api-version=2022-12-01&skiptoken=next"
        status, _, requests = self.run_fixture([self.token(), {"value": [], "nextLink": continuation}, {"value": []}])
        self.assertEqual(status, 0)
        self.assertEqual(len(requests), 3)

    def test_VCB_198_nonempty_page_never_passes(self):
        continuation = "https://management.azure.com/subscriptions?skiptoken=next"
        for response_list in [[self.token(), {"value": [{"subscriptionId": "fake"}]}], [self.token(), {"value": [], "nextLink": continuation}, {"value": [{"subscriptionId": "fake"}]}]]:
            with self.subTest(response_list=response_list):
                status, output, _ = self.run_fixture(response_list)
                self.assertEqual(status, 1)
                self.assertNotIn("VSCUSE_NO_AZURE_SUBSCRIPTIONS_VERIFIED", output)

    def test_VCB_198_missing_credentials_do_not_request_tokens(self):
        status, _, requests = self.run_fixture([], {})
        self.assertEqual(status, 1)
        self.assertEqual(requests, [])

    def test_VCB_198_wrong_identity_stops_before_ARM(self):
        status, _, requests = self.run_fixture([self.token("another@example.test")])
        self.assertEqual(status, 1)
        self.assertEqual(len(requests), 1)

    def test_VCB_198_http_auth_failure_is_sanitized(self):
        error = urllib.error.HTTPError("https://login.microsoftonline.com/", 400, "bad", {}, io.BytesIO(json.dumps({"error_codes": [50076], "error_description": "fake-private-password"}).encode()))
        status, output, requests = self.run_fixture([error])
        self.assertEqual(status, 1)
        self.assertIn("50076", output)
        self.assertEqual(len(requests), 1)

    def test_VCB_198_malformed_responses_and_missing_token_fail(self):
        for response_list in [[{}], [b"not JSON fake-private-password"], [self.token(), {}], [self.token(), {"value": None}], [self.token(), []], [self.token(), {"value": [], "nextLink": 1}]]:
            with self.subTest(response_list=response_list):
                status, output, _ = self.run_fixture(response_list)
                self.assertEqual(status, 1)
                self.assertNotIn("VSCUSE_NO_AZURE_SUBSCRIPTIONS_VERIFIED", output)

    def test_VCB_198_network_and_ARM_errors_fail(self):
        for error in [urllib.error.URLError("fake-private-password"), urllib.error.HTTPError("https://management.azure.com/", 403, "Forbidden", {}, io.BytesIO(b"fake-private-password"))]:
            with self.subTest(error=type(error).__name__):
                status, _, _ = self.run_fixture([self.token(), error])
                self.assertEqual(status, 1)

    def test_VCB_198_unsafe_or_repeated_continuation_is_rejected(self):
        for continuation in ["https://example.test/subscriptions", "http://management.azure.com/subscriptions", "https://management.azure.com/other", "https://management.azure.com/subscriptions?api-version=2022-12-01"]:
            with self.subTest(continuation=continuation):
                status, _, requests = self.run_fixture([self.token(), {"value": [], "nextLink": continuation}])
                self.assertEqual(status, 1)
                self.assertEqual(len(requests), 2)


if __name__ == "__main__":
    unittest.main()