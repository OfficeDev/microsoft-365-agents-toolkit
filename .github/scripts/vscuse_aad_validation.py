"""Safety gates and process isolation for the temporary AAD validation branch."""

import argparse
import copy
import os
from pathlib import Path
import re
import subprocess
import sys


VALIDATION_REF = "refs/heads/qfai/vscuse-aad-prerelease-validation"
RELEASE_TAG = "v1.0.3rc1"
SAFE_PLAN = "Feature_Validate_copilot_manifest"
# Successful template run 38201441001, image pull in job 114655614516.
IMAGE_TAG = "qinzhouxu-rai_20261012033454"
IMAGE_DIGEST = "sha256:194a00acd2e5b235afd75d3d3e125fd4acbb08cdc2ae8460cd9ab2c4a70ad7d3"


def validate_inputs(env):
    if env.get("GITHUB_REF") != VALIDATION_REF:
        raise ValueError("AAD validation requires the isolated branch")
    if env.get("GITHUB_EVENT_NAME") != "workflow_dispatch":
        raise ValueError("AAD validation requires manual dispatch")
    if env.get("VSCUSE_VERSION") != RELEASE_TAG:
        raise ValueError("Explicit vscuse_version=v1.0.3rc1 is required; latest is forbidden")
    if re.fullmatch(r"[a-f0-9]{64}", env.get("VSCUSE_SHA256", "")) is None:
        raise ValueError("The release owner's expected wheel SHA256 is required")
    if env.get("TEST_PLAN") != SAFE_PLAN:
        raise ValueError("Only the approved local-only DA plan is allowed")
    if env.get("IMAGE_TAG") != IMAGE_TAG:
        raise ValueError("The provenance-pinned product image is required")
    if env.get("MAX_RETRIES") != "1":
        raise ValueError("AAD validation requires max_retries=1")
    if env.get("SCHEDULE_TRIGGER") != "false":
        raise ValueError("AAD validation cannot trigger scheduled workflows")
    if env.get("EMAIL_RECEIVER") or env.get("TEMPLATE_VERSION"):
        raise ValueError("Email and template overrides must be empty")


def prepare_config(config):
    prepared = copy.deepcopy(config)
    for section in ("azure_openai", "azure_vision"):
        prepared[section].pop("api_key", None)
    prepared["azure_openai"].pop("use_azure_credential", None)

    principal = prepared["azure_service_principal"]
    if principal.get("enabled") is not False:
        raise ValueError("Toolkit service-principal authentication must remain disabled")
    for key in ("client_id", "client_secret", "tenant_id"):
        principal[key] = ""

    environment = prepared["docker"]["environment"]
    if any(name in environment for name in (
        "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET", "AZURE_TENANT_ID",
        "VSCUSE_AAD_CLIENT_ID", "VSCUSE_AAD_CLIENT_SECRET", "VSCUSE_AAD_TENANT_ID",
    )):
        raise ValueError("Host AAD credentials must not be passed to the application container")
    # The old runner forwarded this application key automatically. The RC does not.
    environment["AZURE_OPENAI_API_KEY"] = "${VSCUSE_APP_AZURE_OPENAI_API_KEY}"
    return prepared


def run_plan(config_path, plan_path, version):
    import yaml
    from vscuse_aad_preflight import configure_aad_environment, verify_version

    verify_version(version)
    config = yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))
    if config != prepare_config(config):
        raise ValueError("Run config has not been isolated for host-only AAD")
    if Path(plan_path).name != f"{SAFE_PLAN}.json":
        raise ValueError("Only the approved local-only DA plan is allowed")
    os.environ["VSCUSE_APP_AZURE_OPENAI_API_KEY"] = os.environ.get("AZURE_OPENAI_API_KEY", "")
    configure_aad_environment()
    return subprocess.run([
        "vscuse", "execute", "--config-file", config_path,
        "--groups-dir", "groups", plan_path,
    ], check=False).returncode


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("guard")
    prepare = subparsers.add_parser("prepare")
    prepare.add_argument("source")
    prepare.add_argument("destination")
    run = subparsers.add_parser("run")
    run.add_argument("--config", required=True)
    run.add_argument("--plan", required=True)
    run.add_argument("--version", required=True)
    subparsers.add_parser("image-digest")
    args = parser.parse_args()
    if args.command == "guard":
        validate_inputs(os.environ)
        print("AAD validation inputs: PASS")
    elif args.command == "prepare":
        import yaml

        config = yaml.safe_load(Path(args.source).read_text(encoding="utf-8"))
        Path(args.destination).write_text(
            yaml.safe_dump(prepare_config(config), sort_keys=False), encoding="utf-8"
        )
    elif args.command == "image-digest":
        print(IMAGE_DIGEST)
    else:
        return run_plan(args.config, args.plan, args.version)
    return 0


if __name__ == "__main__":
    sys.exit(main())
