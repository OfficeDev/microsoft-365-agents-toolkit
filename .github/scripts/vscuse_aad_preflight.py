#!/usr/bin/env python3
"""Opt-in AAD smoke for an already installed, explicitly pinned vscuse prerelease."""

import argparse
import asyncio
from contextlib import AsyncExitStack, redirect_stderr, redirect_stdout
from importlib import metadata
import io
import logging
import os
from pathlib import Path
import re
import sys


EXPECTED_TEXT = "AAD OK"
REQUEST_TIMEOUT = 30
KEY_VARIABLES = ("AZURE_OPENAI_API_KEY", "AZURE_VISION_KEY", "OPENAI_API_KEY")
CREDENTIAL_NAMES = ("CLIENT_ID", "CLIENT_SECRET", "TENANT_ID")


class PreflightError(RuntimeError):
    """Carries only a fixed, non-sensitive failure category."""


class SafeArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        # argparse normally echoes arbitrary argument values.
        raise PreflightError("arguments")


def verify_version(tag: str) -> str:
    """Verify the installed prerelease without importing the candidate runtime."""
    expected = tag.removeprefix("v")
    if not re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+(?:a|b|rc)[0-9]+", expected):
        raise PreflightError("prerelease_required")
    installed = metadata.version("vscuse")
    if installed != expected:
        raise PreflightError("version_mismatch")
    return installed


def configure_aad_environment() -> dict[str, str]:
    """Configure process-local AAD credentials, preserving the sample-app key alias."""
    for name in KEY_VARIABLES:
        os.environ.pop(name, None)
    if "AZURE_OPENAI_AD_TOKEN" in os.environ:
        raise PreflightError("static_token")
    credentials = {}
    for name in CREDENTIAL_NAMES:
        value = os.environ.get(f"VSCUSE_AAD_{name}", "")
        if not value.strip() or "${" in value:
            raise PreflightError("credentials")
        credentials[f"AZURE_{name}"] = value
    os.environ.update(credentials)
    os.environ.update({
        "AZURE_TOKEN_CREDENTIALS": "EnvironmentCredential",
        "PYTHON_DOTENV_DISABLED": "1",
        "OPENAI_AGENTS_DISABLE_TRACING": "1",
        "AZURE_TRACING_ENABLED": "false",
        "OTEL_SDK_DISABLED": "true",
    })
    return credentials


async def run_smoke(config_path):
    credentials = configure_aad_environment()
    # The candidate package's __init__ loads .env with override=True. Recheck
    # even when an older python-dotenv does not honor PYTHON_DOTENV_DISABLED.
    import vscuse_v2  # noqa: F401

    if configure_aad_environment() != credentials:
        raise PreflightError("credentials_changed")

    from agents import set_tracing_disabled

    set_tracing_disabled(True)
    from vscuse_v2.config import (
        get_azure_openai_config,
        get_azure_vision_config,
        get_config,
        reset_config,
    )

    # The candidate silently uses defaults for a missing explicit config file.
    path = Path(config_path).expanduser()
    with path.open("rb"):
        pass
    reset_config()
    get_config(path)
    azure_config = get_azure_openai_config()
    vision_config = get_azure_vision_config()
    for value in (
        azure_config.endpoint,
        vision_config.endpoint,
        azure_config.api_version,
        azure_config.model,
    ):
        if not isinstance(value, str) or not value.strip() or "${" in value:
            raise PreflightError("configuration")
    model = azure_config.model
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,127}", model):
        raise PreflightError("model")

    # Importing vscuse_v2.agent before core triggers the candidate's circular imports.
    import vscuse_v2.core  # noqa: F401

    from azure.ai.vision.imageanalysis.models import VisualFeatures
    from PIL import Image, ImageDraw, ImageFont
    from vscuse_v2.core.ocr_agent import OcrAgent
    from vscuse_v2.utils.azure_client import AzureOpenAIClientFactory

    async with AsyncExitStack() as cleanup:
        client = await AzureOpenAIClientFactory.create_client()
        cleanup.push_async_callback(client.close)
        client.max_retries = 0
        response = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "Reply with OK."}],
            max_completion_tokens=16,
            timeout=REQUEST_TIMEOUT,
        )
        if response is None or not response.choices:
            raise PreflightError("openai_result")
        content = response.choices[0].message.content
        if not isinstance(content, str) or not content.strip():
            raise PreflightError("openai_result")

        ocr = OcrAgent()

        async def close_ocr():
            try:
                if ocr._client is not None:
                    await ocr._client.close()
            finally:
                if ocr._vision_client is not None:
                    ocr._vision_client.close()

        # Register before initialization so partial initialization is cleaned up.
        cleanup.push_async_callback(close_ocr)
        await ocr._initialize_client()
        if ocr._client is None or ocr._vision_client is None:
            raise PreflightError("ocr_initialization")
        with Image.new("RGB", (480, 160), "white") as image, io.BytesIO() as data:
            ImageDraw.Draw(image).text(
                (24, 32), EXPECTED_TEXT, fill="black",
                font=ImageFont.load_default(size=64),
            )
            image.save(data, format="PNG")
            result = ocr._vision_client.analyze(
                image_data=data.getvalue(),
                visual_features=[VisualFeatures.READ],
                connection_timeout=REQUEST_TIMEOUT,
                read_timeout=REQUEST_TIMEOUT,
                retry_total=0,
                logging_enable=False,
                tracing_enable=False,
            )
        if result is None or not result.read or not result.read.blocks:
            raise PreflightError("vision_result")
        text = " ".join(
            line.text for block in result.read.blocks for line in block.lines
        )
        if not re.search(r"\bAAD\s+OK\b", text.upper()):
            raise PreflightError("vision_result")
    return model


def main(argv=None):
    parser = SafeArgumentParser(description=__doc__, allow_abbrev=False)
    parser.add_argument("--config", help="vscuse YAML config (required for live checks)")
    parser.add_argument("--version", required=True, help="exact prerelease tag, e.g. v1.0.3rc1")
    parser.add_argument("--check-version", action="store_true", help="check metadata only; no live calls")
    previous_log_level = logging.root.manager.disable
    failure = None
    try:
        args = parser.parse_args(argv)
        if not args.check_version and not args.config:
            raise PreflightError("arguments")
        logging.disable(sys.maxsize)
        # Imports, SDK diagnostics, response bodies and exception messages must
        # never reach the host log, including during SDK cleanup.
        with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            try:
                version = verify_version(args.version)
                model = None if args.check_version else asyncio.run(run_smoke(args.config))
            except Exception as error:
                category = str(error) if isinstance(error, PreflightError) else "runtime"
                status = getattr(error, "status_code", None)
                if status is None:
                    status = getattr(getattr(error, "response", None), "status_code", None)
                failure = f"check={category} status=failed type={type(error).__name__}"
                if type(status) is int and 100 <= status <= 599:
                    failure += f" http_status={status}"
    except PreflightError:
        failure = "check=arguments status=failed type=PreflightError"
    finally:
        logging.disable(previous_log_level)
    if failure:
        print(failure, file=sys.stderr)
        return 1
    if args.check_version:
        print(f"vscuse={version} check=version status=ok")
    else:
        for check in ("openai", "vision"):
            print(f"vscuse={version} model={model} check={check} status=ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
