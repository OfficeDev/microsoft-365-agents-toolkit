import copy
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import vscuse_aad_validation as validation

ROOT = Path(__file__).resolve().parents[3]


class ValidationTests(unittest.TestCase):
    def setUp(self):
        self.env = {
            "GITHUB_REF": validation.VALIDATION_REF,
            "GITHUB_EVENT_NAME": "workflow_dispatch",
            "VSCUSE_VERSION": "v1.0.3rc1",
            "VSCUSE_SHA256": validation.RELEASE_SHA256,
            "TEST_PLAN": validation.SAFE_PLAN,
            "IMAGE_TAG": validation.IMAGE_TAG,
            "MAX_RETRIES": "1",
            "SCHEDULE_TRIGGER": "false",
            "EMAIL_RECEIVER": "",
            "TEMPLATE_VERSION": "",
        }
        self.config = yaml.safe_load(
            (ROOT / "packages/tests/vscuse/vscode-test-cases/config.yaml").read_text()
        )

    def test_explicit_isolated_inputs_pass(self):
        validation.validate_inputs(self.env)

    def test_unsafe_inputs_fail_closed(self):
        for name, values in {
            "GITHUB_REF": ["refs/heads/dev", "refs/heads/release/6.12", ""],
            "GITHUB_EVENT_NAME": ["schedule", "push"],
            "VSCUSE_VERSION": ["", "latest", "v1.0.3", "v1.0.3rc2", "v1.0.3rc1;echo unsafe", "v1.0.3rc1\n"],
            "VSCUSE_SHA256": ["", "a" * 63, "a" * 64, "g" * 64, validation.RELEASE_SHA256 + "\n"],
            "TEST_PLAN": ["", "DA_Provision", validation.SAFE_PLAN + ",second"],
            "IMAGE_TAG": ["latest", "", "unverified-image"],
            "MAX_RETRIES": ["7", "0", ""],
            "SCHEDULE_TRIGGER": ["true", ""],
            "EMAIL_RECEIVER": ["recipient@example.test"],
            "TEMPLATE_VERSION": ["local"],
        }.items():
            for value in values:
                with self.subTest(name=name, value=value):
                    with self.assertRaises(ValueError):
                        validation.validate_inputs({**self.env, name: value})

    def test_config_preserves_application_contract_without_host_identity(self):
        original = copy.deepcopy(self.config)
        prepared = validation.prepare_config(self.config)
        self.assertEqual(self.config, original)
        self.assertNotIn("api_key", prepared["azure_openai"])
        self.assertNotIn("use_azure_credential", prepared["azure_openai"])
        self.assertNotIn("api_key", prepared["azure_vision"])
        self.assertIs(prepared["azure_service_principal"]["enabled"], False)
        for key in ("client_id", "client_secret", "tenant_id"):
            self.assertEqual(prepared["azure_service_principal"][key], "")
        expected = {
            **original["docker"]["environment"],
            "AZURE_OPENAI_API_KEY": "${VSCUSE_APP_AZURE_OPENAI_API_KEY}",
        }
        self.assertEqual(prepared["docker"]["environment"], expected)
        for name in (
            "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID", "AZURE_CONFIG_DIR",
            "ACTIONS_ID_TOKEN_REQUEST_TOKEN", "ACTIONS_ID_TOKEN_REQUEST_URL", "VSCUSE_AAD_CREDENTIAL",
        ):
            self.assertNotIn(name, prepared["docker"]["environment"])
        self.assertEqual(prepared["m365"], original["m365"])
        self.assertEqual(validation.prepare_config(prepared), prepared)
        self.assertNotIn("literal-secret", yaml.safe_dump(prepared))

    def test_container_identity_injection_rejected(self):
        self.config["azure_service_principal"]["enabled"] = True
        with self.assertRaises(ValueError):
            validation.prepare_config(self.config)
        self.config["azure_service_principal"]["enabled"] = False
        self.config["docker"]["environment"]["AZURE_CLIENT_SECRET"] = "${VSCUSE_AAD_CLIENT_SECRET}"
        with self.assertRaises(ValueError):
            validation.prepare_config(self.config)

    def test_run_configures_host_auth_and_preserves_application_alias(self):
        prepared = validation.prepare_config(self.config)
        with patch.dict(os.environ, {"AZURE_OPENAI_API_KEY": "sample-only"}, clear=True):
            with (
                patch("pathlib.Path.read_text", return_value=yaml.safe_dump(prepared)),
                patch("vscuse_aad_preflight.verify_version") as version,
                patch("vscuse_aad_preflight.configure_aad_environment") as configure,
                patch("subprocess.run") as run,
            ):
                run.return_value.returncode = 19
                result = validation.run_plan("config.aad.yaml", f"{validation.SAFE_PLAN}.json", "v1.0.3rc1")
                self.assertEqual(result, 19)
                version.assert_called_once_with("v1.0.3rc1")
                configure.assert_called_once_with()
                self.assertEqual(os.environ["VSCUSE_APP_AZURE_OPENAI_API_KEY"], "sample-only")
                self.assertEqual(run.call_args.args[0][0:2], ["vscuse", "execute"])

    def test_workflow_isolation_and_unchanged_daily_defaults(self):
        workflows = ROOT / ".github/workflows"
        common = yaml.safe_load((workflows / "ui-test-vscuse-common.yml").read_text(encoding="utf-8"))
        template = yaml.safe_load((workflows / "ui-test-vscuse-template.yml").read_text(encoding="utf-8"))
        # YAML 1.1's boolean interpretation of "on" is irrelevant to the other nodes.
        template_inputs = template[True]["workflow_dispatch"]["inputs"]
        common_inputs = common[True]["workflow_call"]["inputs"]
        for inputs in (template_inputs, common_inputs):
            self.assertEqual(inputs["vscuse_version"]["default"], "latest")
            self.assertEqual(inputs["image_tag"]["default"], "latest")
            self.assertEqual(inputs["max_retries"]["default"], 7)
        self.assertEqual(
            template["jobs"]["run"]["with"]["vscuse_version"],
            "${{ github.event.inputs.vscuse_version }}",
        )
        self.assertEqual(
            template["jobs"]["run"]["with"]["vscuse_sha256"],
            "${{ github.event.inputs.vscuse_sha256 }}",
        )
        discovery = common["jobs"]["discover-test-plans"]["steps"]
        self.assertEqual(discovery[1]["name"], "Guard isolated AAD validation inputs")
        steps = {step["name"]: step for step in common["jobs"]["main"]["steps"]}
        branch_only = f"github.ref == '{validation.VALIDATION_REF}'"
        non_validation = f"github.ref != '{validation.VALIDATION_REF}'"
        for name in (
            "Login to Azure for isolated AAD validation",
            "Verify installed prerelease and AAD data-plane access",
            "Verify validation image digest",
            "Run isolated AAD validation plan",
        ):
            self.assertIn(branch_only, steps[name]["if"])
        for name in (
            "Verify Azure Configuration", "Run test", "Login to Azure",
            "\U0001f310Upload files to Azure Blob Storage",
        ):
            self.assertIn(non_validation, steps[name]["if"])
        for name in ("rerun", "report"):
            self.assertIn(non_validation, common["jobs"][name]["if"])
        self.assertEqual(
            steps["Encrypt and zip test report"]["if"],
            "${{ always() && (" + non_validation
            + " || hashFiles('packages/tests/vscuse/vscode-test-cases/test_report/index.html') != '') }}",
        )
        self.assertIn(non_validation, template["jobs"]["trigger-ui-test-features"]["if"])
        names = list(steps)
        self.assertLess(
            names.index("Login to Azure for isolated AAD validation"),
            names.index("Verify installed prerelease and AAD data-plane access"),
        )
        self.assertLess(
            names.index("Verify installed prerelease and AAD data-plane access"),
            names.index("Set m365 account randomly"),
        )
        host_env = common["jobs"]["main"]["env"]
        self.assertNotIn("AZURE_CLIENT_ID", host_env)
        self.assertNotIn("AZURE_CLIENT_SECRET", host_env)
        login = steps["Login to Azure for isolated AAD validation"]
        self.assertEqual(login["uses"], "azure/login@v2")
        self.assertEqual(login["with"], {
            "client-id": "${{ secrets.DEVOPS_CLIENT_ID }}",
            "tenant-id": "${{ secrets.DEVOPS_TENANT_ID }}",
            "subscription-id": "${{ secrets.DEVOPS_SUB_ID }}",
        })
        self.assertEqual(login["env"], {"AZURE_CORE_OUTPUT": "none"})
        self.assertEqual(common["jobs"]["main"]["permissions"]["id-token"], "write")
        self.assertEqual(template["jobs"]["run"]["permissions"]["id-token"], "write")
        self.assertEqual(common["jobs"]["main"]["environment"], "engineering")
        self.assertNotIn("VSCUSE_AAD_CREDENTIAL", common_inputs)
        self.assertNotIn("VSCUSE_AAD_CREDENTIAL", template_inputs)
        self.assertEqual(steps["Login to Azure"]["with"], {
            "client-id": "${{secrets.DEVOPS_CLIENT_ID}}",
            "tenant-id": "${{secrets.DEVOPS_TENANT_ID}}",
            "subscription-id": "${{secrets.DEVOPS_SUB_ID}}",
            "enable-AzPSSession": True,
        })
        for name in ("Verify installed prerelease and AAD data-plane access", "Run isolated AAD validation plan"):
            self.assertEqual(
                steps[name]["env"]["VSCUSE_AAD_CREDENTIAL"], "AzureCliCredential",
            )
            for key in ("CLIENT_ID", "CLIENT_SECRET", "TENANT_ID"):
                self.assertNotIn(f"VSCUSE_AAD_{key}", steps[name]["env"])
                self.assertNotIn(f"AZURE_{key}", steps[name]["env"])


if __name__ == "__main__":
    unittest.main()
