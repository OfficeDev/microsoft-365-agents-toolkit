import builtins
from contextlib import contextmanager, redirect_stderr, redirect_stdout
import importlib.util
import io
import logging
import os
from pathlib import Path
import sys
from types import ModuleType, SimpleNamespace
import unittest
from unittest.mock import AsyncMock, MagicMock, Mock, call, mock_open, patch
import warnings


SCRIPT = Path(__file__).parents[1] / "vscuse_aad_preflight.py"
VERSION = "1.2.3rc4"
CONFIG = "offline-smoke.yaml"
ALIASES = {
    "VSCUSE_AAD_CLIENT_ID": "synthetic-client",
    "VSCUSE_AAD_CLIENT_SECRET": "synthetic-secret",
    "VSCUSE_AAD_TENANT_ID": "synthetic-tenant",
}
API_KEYS = ("AZURE_OPENAI_API_KEY", "AZURE_VISION_KEY", "OPENAI_API_KEY")
SDK_ROOTS = {"vscuse_v2", "agents", "azure", "PIL", "openai", "dotenv"}
REAL_IMPORT = builtins.__import__
SENSITIVE = (
    "synthetic-secret synthetic-token synthetic-account "
    "https://private-endpoint.invalid raw-exception Traceback"
)
PNG_BYTES = b"synthetic-image-bytes"


def offline_import(name, globals=None, locals=None, fromlist=(), level=0):
    if name.split(".")[0] in SDK_ROOTS:
        raise AssertionError("Unmocked SDK import")
    return REAL_IMPORT(name, globals, locals, fromlist, level)


SPEC = importlib.util.spec_from_file_location("vscuse_aad_preflight", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader
with patch("builtins.__import__", offline_import):
    SPEC.loader.exec_module(MODULE)


def completion(content="OK"):
    return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))])


def vision_result(*lines):
    return SimpleNamespace(read=SimpleNamespace(blocks=[
        SimpleNamespace(lines=[SimpleNamespace(text=text) for text in lines])
    ]))


class FakeHttpError(RuntimeError):
    pass


class Candidate:
    """Only in-memory doubles; unknown SDK imports fail instead of reaching disk."""

    def __init__(self):
        self.imports = []
        self.on_root_import = None
        self.azure_config = SimpleNamespace(
            endpoint="https://synthetic-openai.invalid",
            api_version="2025-01-01-preview",
            model="synthetic-deployment",
        )
        self.vision_config = SimpleNamespace(endpoint="https://synthetic-vision.invalid")
        self.reset_config = Mock()
        self.get_config = Mock()
        self.create_completion = AsyncMock(return_value=completion())
        self.client = SimpleNamespace(
            max_retries=9,
            chat=SimpleNamespace(completions=SimpleNamespace(create=self.create_completion)),
            close=AsyncMock(),
        )
        self.factory = SimpleNamespace(create_client=AsyncMock(return_value=self.client))
        self.ocr_client = SimpleNamespace(close=AsyncMock())
        self.vision = SimpleNamespace(
            analyze=Mock(return_value=vision_result("AAD OK")), close=Mock()
        )
        self.ocr = SimpleNamespace(_client=None, _vision_client=None)

        async def initialize():
            self.ocr._client = self.ocr_client
            self.ocr._vision_client = self.vision

        self.ocr._initialize_client = AsyncMock(side_effect=initialize)
        self.ocr_constructor = Mock(return_value=self.ocr)
        self.set_tracing_disabled = Mock()
        self.read_feature = object()
        self.image = MagicMock()
        self.image.__enter__.return_value = self.image
        self.image.save.side_effect = lambda stream, **kwargs: stream.write(PNG_BYTES)
        self.image_api = SimpleNamespace(new=Mock(return_value=self.image))
        self.draw = Mock()
        self.draw_api = SimpleNamespace(Draw=Mock(return_value=self.draw))
        self.font = object()
        self.font_api = SimpleNamespace(load_default=Mock(return_value=self.font))
        attributes = {
            "vscuse_v2": {},
            "vscuse_v2.config": {
                "reset_config": self.reset_config,
                "get_config": self.get_config,
                "get_azure_openai_config": Mock(return_value=self.azure_config),
                "get_azure_vision_config": Mock(return_value=self.vision_config),
            },
            "vscuse_v2.core.ocr_agent": {"OcrAgent": self.ocr_constructor},
            "vscuse_v2.utils.azure_client": {"AzureOpenAIClientFactory": self.factory},
            "agents": {"set_tracing_disabled": self.set_tracing_disabled},
            "azure.ai.vision.imageanalysis.models": {
                "VisualFeatures": SimpleNamespace(READ=self.read_feature),
            },
            "PIL": {
                "Image": self.image_api, "ImageDraw": self.draw_api, "ImageFont": self.font_api,
            },
        }
        self.modules = {}
        for name, values in attributes.items():
            parts = name.split(".")
            for length in range(1, len(parts) + 1):
                qualified = ".".join(parts[:length])
                if qualified not in self.modules:
                    module = ModuleType(qualified)
                    module.__path__ = []
                    self.modules[qualified] = module
                    if length > 1:
                        setattr(self.modules[".".join(parts[:length - 1])], parts[length - 1], module)
            self.modules[name].__dict__.update(values)
        self.events = Mock()
        for name, method in (
            ("reset_config", self.reset_config),
            ("get_config", self.get_config),
            ("create_client", self.factory.create_client),
            ("complete", self.create_completion),
            ("ocr", self.ocr_constructor),
            ("initialize_ocr", self.ocr._initialize_client),
            ("analyze", self.vision.analyze),
            ("close_ocr", self.ocr_client.close),
            ("close_vision", self.vision.close),
            ("close_client", self.client.close),
        ):
            self.events.attach_mock(method, name)

    def import_module(self, name, globals=None, locals=None, fromlist=(), level=0):
        if name.split(".")[0] not in SDK_ROOTS:
            return offline_import(name, globals, locals, fromlist, level)
        if name not in self.modules:
            raise AssertionError("Unmocked SDK import")
        self.imports.append((name, dict(os.environ)))
        if name == "vscuse_v2" and self.on_root_import is not None:
            self.on_root_import()
        return self.modules[name if fromlist else name.split(".")[0]]


@contextmanager
def candidate(environment=None):
    stub = Candidate()
    with (
        patch.dict(os.environ, ALIASES if environment is None else environment, clear=True),
        patch.dict(sys.modules, stub.modules),
        patch("builtins.__import__", stub.import_module),
        patch.object(MODULE.metadata, "version", return_value=VERSION) as installed_version,
        patch.object(MODULE.Path, "open", mock_open()) as open_config,
    ):
        stub.installed_version = installed_version
        stub.open_config = open_config
        yield stub


def emit_diagnostics():
    print(SENSITIVE)
    print(SENSITIVE, file=sys.stderr)
    logging.getLogger("synthetic-candidate").error(SENSITIVE)
    warnings.warn(SENSITIVE, UserWarning)


class AadPreflightTests(unittest.TestCase):
    def invoke(self, argv=None):
        stdout, stderr = io.StringIO(), io.StringIO()
        previous = logging.root.manager.disable
        with redirect_stdout(stdout), redirect_stderr(stderr):
            code = MODULE.main(
                ["--version", f"v{VERSION}", "--config", CONFIG] if argv is None else argv
            )
        self.assertEqual(logging.root.manager.disable, previous)
        return code, stdout.getvalue(), stderr.getvalue()

    def assert_failure(self, category, error_type="PreflightError", status=None, argv=None,
                       phase="credential_setup", passed=(), codes=(), api_details=()):
        message = f"phase={phase} check={category} status=failed type={error_type}"
        if status is not None:
            message += f" http_status={status}"
        if codes:
            message += " aadsts_codes=" + ",".join(codes)
        if api_details:
            message += " " + " ".join(api_details)
        successes = "".join(
            f"vscuse={VERSION} model=synthetic-deployment check={check} status=ok\n"
            for check in passed
        )
        self.assertEqual(self.invoke(argv), (1, successes, message + "\n"))

    def assert_success(self, model="synthetic-deployment"):
        expected = "".join(
            f"vscuse={VERSION} model={model} check={check} status=ok\n"
            for check in ("openai", "vision")
        )
        self.assertEqual(self.invoke(), (0, expected, ""))

    def test_exact_prerelease_versions_with_and_without_v(self):
        for version in ("1.2.3a1", "1.2.3b2", VERSION):
            for prefix in ("", "v"):
                with self.subTest(tag=prefix + version), candidate() as stub:
                    stub.installed_version.return_value = version
                    self.assertEqual(MODULE.verify_version(prefix + version), version)
                    stub.installed_version.assert_called_once_with("vscuse")
                    self.assertEqual(stub.imports, [])

    def test_non_prerelease_and_symbolic_tags_are_rejected(self):
        for tag in (
            "", "stable", "latest", "dev", "1.2.3", "v1.2.3", "1.2.3.dev1", "v",
            "V1.2.3rc4", "vv1.2.3rc4", "1.2.3RC4", "1.2.3rc", "1.2.3-rc4",
            " 1.2.3rc4", "1.2.3rc4\n", "1.2.3rc4+local",
        ):
            with self.subTest(tag=tag), candidate() as stub:
                self.assert_failure(
                    "prerelease_required", argv=["--version", tag, "--check-version"]
                )
                stub.installed_version.assert_not_called()
                self.assertEqual(stub.imports, [])

    def test_wrong_installed_version_fails_before_candidate_import(self):
        with candidate(environment={}) as stub:
            stub.installed_version.return_value = "1.2.3rc3"
            self.assert_failure("version_mismatch")
            stub.installed_version.assert_called_once_with("vscuse")
            self.assertEqual(stub.imports, [])
            stub.open_config.assert_not_called()

    def test_missing_distribution_is_a_safe_failure(self):
        with candidate(environment={}) as stub:
            stub.installed_version.side_effect = MODULE.metadata.PackageNotFoundError(SENSITIVE)
            self.assert_failure("runtime", "PackageNotFoundError")
            self.assertEqual(stub.imports, [])
            stub.open_config.assert_not_called()

    def test_check_version_needs_no_candidate_credentials_or_config(self):
        for environment in ({}, {"AZURE_OPENAI_AD_TOKEN": "", "OPENAI_API_KEY": "synthetic-key"}):
            with self.subTest(environment=environment), candidate(environment) as stub:
                stub.open_config.side_effect = AssertionError("Config must not be opened")
                self.assertEqual(
                    self.invoke(["--version", VERSION, "--check-version"]),
                    (0, f"vscuse={VERSION} check=version status=ok\n", ""),
                )
                stub.installed_version.assert_called_once_with("vscuse")
                self.assertEqual(stub.imports, [])
                stub.open_config.assert_not_called()
                stub.factory.create_client.assert_not_awaited()
                self.assertEqual(dict(os.environ), environment)

    def test_static_token_is_blocked_even_when_empty(self):
        for token in ("", "synthetic-token"):
            environment = {**ALIASES, **dict.fromkeys(API_KEYS, "synthetic-key"),
                           "AZURE_OPENAI_AD_TOKEN": token}
            with self.subTest(token=token), candidate(environment) as stub:
                self.assert_failure("static_token")
                self.assertEqual(stub.imports, [])
                self.assertTrue(all(key not in os.environ for key in API_KEYS))

    def test_missing_blank_or_unresolved_alias_does_not_use_ambient_credentials(self):
        ambient = {f"AZURE_{name}": "ambient-value"
                   for name in ("CLIENT_ID", "CLIENT_SECRET", "TENANT_ID")}
        for alias in ALIASES:
            for value in (None, "", " \t", "${UNRESOLVED}"):
                environment = {**ALIASES, **ambient}
                if value is None:
                    environment.pop(alias)
                else:
                    environment[alias] = value
                with self.subTest(alias=alias, value=value), candidate(environment) as stub:
                    self.assert_failure("credentials")
                    self.assertEqual(stub.imports, [])
                    stub.factory.create_client.assert_not_awaited()

    def test_environment_maps_aliases_restricts_credentials_and_disables_tracing(self):
        environment = {
            **ALIASES, **dict.fromkeys(API_KEYS, "synthetic-key"),
            "AZURE_CLIENT_ID": "ambient-client", "AZURE_CLIENT_SECRET": "ambient-secret",
            "AZURE_TENANT_ID": "ambient-tenant", "AZURE_TOKEN_CREDENTIALS": "AzureCliCredential",
        }
        with candidate(environment):
            expected = {name.removeprefix("VSCUSE_AAD_"): value for name, value in ALIASES.items()}
            expected = {f"AZURE_{name}": value for name, value in expected.items()}
            self.assertEqual(MODULE.configure_aad_environment(), expected)
            for name, value in {
                **expected,
                "AZURE_TOKEN_CREDENTIALS": "EnvironmentCredential",
                "PYTHON_DOTENV_DISABLED": "1",
                "OPENAI_AGENTS_DISABLE_TRACING": "1",
                "AZURE_TRACING_ENABLED": "false",
                "OTEL_SDK_DISABLED": "true",
            }.items():
                self.assertEqual(os.environ[name], value)
            self.assertTrue(all(key not in os.environ for key in API_KEYS))

    def test_reusable_environment_helper_preserves_sample_app_key_without_imports(self):
        environment = {
            **ALIASES, **dict.fromkeys(API_KEYS, "synthetic-service-key"),
            "VSCUSE_APP_AZURE_OPENAI_API_KEY": "synthetic-app-key",
        }
        with candidate(environment) as stub:
            first = MODULE.configure_aad_environment()
            self.assertEqual(MODULE.configure_aad_environment(), first)
            self.assertEqual(os.environ["VSCUSE_APP_AZURE_OPENAI_API_KEY"], "synthetic-app-key")
            self.assertTrue(all(key not in os.environ for key in API_KEYS))
            self.assertEqual(stub.imports, [])
            stub.installed_version.assert_not_called()
            self.assert_success()
            self.assertEqual(os.environ["VSCUSE_APP_AZURE_OPENAI_API_KEY"], "synthetic-app-key")

    def test_smoke_uses_candidate_factory_ocr_initialization_and_synthetic_image(self):
        with candidate({**ALIASES, **dict.fromkeys(API_KEYS, "synthetic-key")}) as stub:
            async def respond(**kwargs):
                self.assertEqual(stub.client.max_retries, 0)
                return completion()

            stub.create_completion.side_effect = respond
            self.assert_success()
            stub.installed_version.assert_called_once_with("vscuse")
            stub.open_config.assert_called_once_with("rb")
            stub.set_tracing_disabled.assert_called_once_with(True)
            request = {
                "model": "synthetic-deployment",
                "messages": [{"role": "user", "content": "Reply with OK."}],
                "max_completion_tokens": 16, "timeout": 30,
            }
            vision_request = {
                "image_data": PNG_BYTES, "visual_features": [stub.read_feature],
                "connection_timeout": 30, "read_timeout": 30, "retry_total": 0,
                "logging_enable": False, "tracing_enable": False,
            }
            self.assertEqual(stub.events.mock_calls, [
                call.reset_config(), call.get_config(Path(CONFIG)),
                call.create_client(), call.complete(**request),
                call.ocr(), call.initialize_ocr(), call.analyze(**vision_request),
                call.close_ocr(), call.close_vision(), call.close_client(),
            ])
            stub.factory.create_client.assert_awaited_once_with()
            stub.create_completion.assert_awaited_once_with(**request)
            stub.ocr._initialize_client.assert_awaited_once_with()
            stub.client.close.assert_awaited_once_with()
            stub.ocr_client.close.assert_awaited_once_with()
            stub.image_api.new.assert_called_once_with("RGB", (480, 160), "white")
            stub.draw_api.Draw.assert_called_once_with(stub.image)
            stub.font_api.load_default.assert_called_once_with(size=64)
            stub.draw.text.assert_called_once_with(
                (24, 32), "AAD OK", fill="black", font=stub.font
            )
            self.assertEqual(stub.image.save.call_args.kwargs, {"format": "PNG"})
            stub.image.__exit__.assert_called_once()
            self.assertEqual({name for name, _ in stub.imports}, {
                "vscuse_v2", "agents", "vscuse_v2.config", "vscuse_v2.core",
                "azure.ai.vision.imageanalysis.models", "PIL",
                "vscuse_v2.core.ocr_agent", "vscuse_v2.utils.azure_client",
            })
            for name, environment in stub.imports:
                with self.subTest(import_name=name):
                    self.assertTrue(all(key not in environment for key in API_KEYS))
                    self.assertEqual(environment["AZURE_TOKEN_CREDENTIALS"], "EnvironmentCredential")
                    self.assertEqual(environment["AZURE_CLIENT_SECRET"], ALIASES["VSCUSE_AAD_CLIENT_SECRET"])

    def test_core_is_initialized_before_importing_candidate_agents(self):
        with candidate() as stub:
            self.assert_success()
            imports = [name for name, _ in stub.imports]
            self.assertIn("vscuse_v2.core", imports)
            core_index = imports.index("vscuse_v2.core")
            agent_imports = [
                index for index, name in enumerate(imports)
                if name == "vscuse_v2.core.ocr_agent" or name.startswith("vscuse_v2.agent")
            ]
            self.assertTrue(agent_imports)
            self.assertTrue(all(core_index < index for index in agent_imports))

    def test_model_is_resolved_from_candidate_config_not_hardcoded(self):
        for model in ("another-deployment", "Custom.Model_2", "a" * 128):
            with self.subTest(model=model), candidate() as stub:
                stub.azure_config.model = model
                self.assert_success(model)
                self.assertEqual(stub.create_completion.call_args.kwargs["model"], model)

    def test_invalid_config_values_fail_after_reset_and_explicit_load(self):
        for target, field in (
            ("azure_config", "endpoint"), ("vision_config", "endpoint"),
            ("azure_config", "api_version"), ("azure_config", "model"),
        ):
            for value in (None, 42, "", " \t", "${UNRESOLVED}"):
                with self.subTest(target=target, field=field, value=value), candidate() as stub:
                    setattr(getattr(stub, target), field, value)
                    self.assert_failure("configuration")
                    self.assertEqual(stub.events.mock_calls, [
                        call.reset_config(), call.get_config(Path(CONFIG)),
                    ])
                    stub.factory.create_client.assert_not_awaited()

    def test_unsafe_model_names_fail_without_echoing_them(self):
        for model in ("a" * 129, "model/name", "-model", "model\nsecret", SENSITIVE):
            with self.subTest(model=model), candidate() as stub:
                stub.azure_config.model = model
                self.assert_failure("model")
                stub.factory.create_client.assert_not_awaited()

    def test_missing_explicit_config_cannot_fall_back_to_defaults(self):
        with candidate() as stub:
            stub.open_config.side_effect = FileNotFoundError(SENSITIVE)
            self.assert_failure("runtime", "FileNotFoundError")
            stub.open_config.assert_called_once_with("rb")
            stub.reset_config.assert_not_called()
            stub.get_config.assert_not_called()
            stub.factory.create_client.assert_not_awaited()

    def test_live_mode_requires_an_explicit_config_argument(self):
        with candidate() as stub:
            self.assert_failure("arguments", argv=["--version", VERSION])
            stub.installed_version.assert_not_called()
            self.assertEqual(stub.imports, [])

    def test_openai_missing_or_blank_results_fail_and_close_client(self):
        responses = [None, SimpleNamespace(choices=[])]
        responses.extend(completion(value) for value in (None, "", " \t\n", 123))
        for response in responses:
            with self.subTest(response=response), candidate() as stub:
                stub.create_completion.return_value = response
                self.assert_failure("openai_result", phase="openai")
                stub.client.close.assert_awaited_once_with()
                stub.ocr_constructor.assert_not_called()
                stub.vision.analyze.assert_not_called()

    def test_vision_missing_or_wrong_text_fails_and_closes_all_clients(self):
        results = [
            None, SimpleNamespace(read=None), SimpleNamespace(read=SimpleNamespace(blocks=[])),
            vision_result(), vision_result(""), vision_result("AAD NO"),
            vision_result("NOTAAD OK"), vision_result("AAD OKAY"),
        ]
        for result in results:
            with self.subTest(result=result), candidate() as stub:
                stub.vision.analyze.return_value = result
                self.assert_failure("vision_result", phase="vision", passed=("openai",))
                stub.client.close.assert_awaited_once_with()
                stub.ocr_client.close.assert_awaited_once_with()
                stub.vision.close.assert_called_once_with()

    def test_vision_recognizes_case_insensitive_text_across_lines(self):
        with candidate() as stub:
            stub.vision.analyze.return_value = vision_result("aad", "ok")
            self.assert_success()

    def test_factory_and_service_call_failures_close_acquired_clients(self):
        for stage in ("factory", "openai", "ocr_constructor", "vision"):
            with self.subTest(stage=stage), candidate() as stub:
                method = {
                    "factory": stub.factory.create_client, "openai": stub.create_completion,
                    "ocr_constructor": stub.ocr_constructor, "vision": stub.vision.analyze,
                }[stage]
                method.side_effect = FakeHttpError(SENSITIVE)
                self.assert_failure(
                    "runtime", "FakeHttpError",
                    phase="openai" if stage in ("factory", "openai") else "vision",
                    passed=() if stage in ("factory", "openai") else ("openai",),
                )
                self.assertEqual(stub.client.close.await_count, int(stage != "factory"))
                self.assertEqual(stub.ocr_client.close.await_count, int(stage == "vision"))
                self.assertEqual(stub.vision.close.call_count, int(stage == "vision"))

    def test_partial_or_failed_ocr_initialization_closes_available_clients(self):
        for has_openai, has_vision in ((False, False), (True, False), (False, True), (True, True)):
            for raises in (False, True):
                if has_openai and has_vision and not raises:
                    continue
                with self.subTest(openai=has_openai, vision=has_vision, raises=raises), candidate() as stub:
                    async def initialize():
                        stub.ocr._client = stub.ocr_client if has_openai else None
                        stub.ocr._vision_client = stub.vision if has_vision else None
                        if raises:
                            raise FakeHttpError(SENSITIVE)

                    stub.ocr._initialize_client.side_effect = initialize
                    if raises:
                        self.assert_failure("runtime", "FakeHttpError", phase="vision", passed=("openai",))
                    else:
                        self.assert_failure("ocr_initialization", phase="vision", passed=("openai",))
                    stub.client.close.assert_awaited_once_with()
                    self.assertEqual(stub.ocr_client.close.await_count, int(has_openai))
                    self.assertEqual(stub.vision.close.call_count, int(has_vision))
                    stub.vision.analyze.assert_not_called()

    def test_cleanup_failure_fails_gate_but_still_closes_other_clients(self):
        for target in ("client", "ocr_client", "vision"):
            with self.subTest(target=target), candidate() as stub:
                getattr(stub, target).close.side_effect = FakeHttpError(SENSITIVE)
                self.assert_failure(
                    "runtime", "FakeHttpError", phase="cleanup", passed=("openai", "vision")
                )
                stub.client.close.assert_awaited_once_with()
                stub.ocr_client.close.assert_awaited_once_with()
                stub.vision.close.assert_called_once_with()

    def test_candidate_root_import_cannot_change_aliases(self):
        for alias in ALIASES:
            with self.subTest(alias=alias), candidate() as stub:
                stub.on_root_import = lambda: os.environ.update({alias: "changed-value"})
                self.assert_failure("credentials_changed")
                self.assertEqual([name for name, _ in stub.imports], ["vscuse_v2"])
                stub.open_config.assert_not_called()

    def test_candidate_root_import_cannot_introduce_a_static_token(self):
        for token in ("", "synthetic-token"):
            with self.subTest(token=token), candidate() as stub:
                stub.on_root_import = lambda: os.environ.update({"AZURE_OPENAI_AD_TOKEN": token})
                self.assert_failure("static_token")
                self.assertEqual([name for name, _ in stub.imports], ["vscuse_v2"])
                stub.factory.create_client.assert_not_awaited()

    def test_candidate_root_import_keys_are_removed_and_credentials_are_remapped(self):
        with candidate() as stub:
            stub.on_root_import = lambda: os.environ.update({
                **dict.fromkeys(API_KEYS, "dotenv-key"),
                "AZURE_CLIENT_ID": "dotenv-client", "AZURE_CLIENT_SECRET": "dotenv-secret",
                "AZURE_TENANT_ID": "dotenv-tenant", "AZURE_TOKEN_CREDENTIALS": "AzureCliCredential",
            })
            self.assert_success()
            for _, environment in stub.imports:
                self.assertTrue(all(key not in environment for key in API_KEYS))
                self.assertEqual(environment["AZURE_TOKEN_CREDENTIALS"], "EnvironmentCredential")
                for alias, value in ALIASES.items():
                    self.assertEqual(environment[alias.replace("VSCUSE_AAD_", "AZURE_")], value)

    def test_failure_hides_candidate_prints_logs_warnings_and_raw_exceptions(self):
        previous = logging.root.manager.disable
        self.addCleanup(logging.disable, previous)
        logging.disable(logging.INFO)
        for stage in ("import", "openai", "initialize", "vision", "cleanup"):
            with self.subTest(stage=stage), candidate() as stub, warnings.catch_warnings():
                warnings.simplefilter("always")
                error = FakeHttpError(SENSITIVE)
                error.status_code = 401

                def fail(*args, **kwargs):
                    emit_diagnostics()
                    raise error

                stub.on_root_import = emit_diagnostics
                if stage == "import":
                    stub.on_root_import = fail
                else:
                    method = {
                        "openai": stub.create_completion, "initialize": stub.ocr._initialize_client,
                        "vision": stub.vision.analyze, "cleanup": stub.client.close,
                    }[stage]
                    method.side_effect = fail
                phase = {
                    "import": "credential_setup", "openai": "openai",
                    "initialize": "vision", "vision": "vision", "cleanup": "cleanup",
                }[stage]
                passed = {
                    "import": (), "openai": (), "initialize": ("openai",),
                    "vision": ("openai",), "cleanup": ("openai", "vision"),
                }[stage]
                self.assert_failure("runtime", "FakeHttpError", status=401, phase=phase, passed=passed)
                self.assertEqual(logging.root.manager.disable, logging.INFO)

    def test_success_hides_candidate_diagnostics_including_cleanup(self):
        with candidate() as stub, warnings.catch_warnings():
            warnings.simplefilter("always")
            stub.on_root_import = emit_diagnostics
            stub.client.close.side_effect = emit_diagnostics
            self.assert_success()

    def test_only_valid_integer_http_status_is_reported(self):
        for status in (400, 599, None, True, 99, 600, SENSITIVE):
            for nested in (False, True):
                with self.subTest(status=status, nested=nested), candidate() as stub:
                    error = FakeHttpError(SENSITIVE)
                    if nested:
                        error.response = SimpleNamespace(status_code=status)
                    else:
                        error.status_code = status
                    stub.create_completion.side_effect = error
                    expected = status if type(status) is int and 100 <= status <= 599 else None
                    self.assert_failure("runtime", "FakeHttpError", status=expected, phase="openai")
                    stub.client.close.assert_awaited_once_with()

    def test_only_numeric_aadsts_codes_escape_secret_bearing_authentication_errors(self):
        for code in ("7000215", "7000222", "700016"):
            for source in ("args", "message"):
                with self.subTest(code=code, source=source), candidate() as stub:
                    message = f"{SENSITIVE}; AADSTS{code}: invalid credentials; {SENSITIVE}"
                    error = FakeHttpError(message if source == "args" else SENSITIVE)
                    if source == "message":
                        error.message = message
                    stub.create_completion.side_effect = error
                    self.assert_failure("runtime", "FakeHttpError", phase="openai", codes=(code,))

    def test_aadsts_codes_from_cyclic_exception_chains_are_deduplicated(self):
        with candidate() as stub:
            error = FakeHttpError(SENSITIVE)
            cause = RuntimeError(f"{SENSITIVE} AADSTS7000215: hidden")
            context = ValueError(f"AADSTS7000222: {SENSITIVE}")
            inner = RuntimeError(f"{SENSITIVE} AADSTS700016: hidden AADSTS7000215")
            error.__cause__ = cause
            error.__context__ = context
            cause.inner_exception = inner
            inner.__cause__ = error
            stub.create_completion.side_effect = error
            self.assert_failure(
                "runtime", "FakeHttpError", phase="openai",
                codes=("700016", "7000215", "7000222"),
            )

    def test_malformed_aadsts_tokens_and_response_bodies_are_not_logged(self):
        with candidate() as stub:
            error = FakeHttpError(
                f"{SENSITIVE} AADSTSsecret AADSTS7000215secret AADSTS1234567890123456 "
                "secretAADSTS700016 AADSTS7000222_secret AADSTS\u0667\u0660\u0660\u0660\u0661\u0666"
            )
            error.response = SimpleNamespace(
                status_code=401, text=f"{SENSITIVE} AADSTS7000215", headers={"secret": SENSITIVE}
            )
            stub.create_completion.side_effect = error
            self.assert_failure("runtime", "FakeHttpError", status=401, phase="openai")

    def test_vision_authentication_failure_preserves_successful_openai_evidence(self):
        with candidate() as stub:
            error = FakeHttpError(SENSITIVE)
            error.__cause__ = RuntimeError(f"AADSTS7000222: {SENSITIVE}")
            stub.vision.analyze.side_effect = error
            self.assert_failure(
                "runtime", "FakeHttpError", phase="vision", passed=("openai",), codes=("7000222",),
            )
            stub.create_completion.assert_awaited_once()
            stub.vision.analyze.assert_called_once()

    def test_openai_error_reports_only_known_structured_code_and_param(self):
        for code, param in (
            ("unsupported_parameter", "max_completion_tokens"),
            ("unsupported_value", "max_tokens"),
            ("OperationNotSupported", "model"),
            ("InvalidApiVersionParameter", "api-version"),
        ):
            with self.subTest(code=code, param=param), candidate() as stub:
                error = FakeHttpError(SENSITIVE)
                error.status_code = 400
                error.code, error.param = code, param
                error.body = {"message": SENSITIVE, "headers": SENSITIVE}
                stub.create_completion.side_effect = error
                self.assert_failure(
                    "runtime", "FakeHttpError", status=400, phase="openai",
                    api_details=(f"api_code={code}", f"api_param={param}"),
                )

    def test_unsafe_structured_identifiers_are_replaced_with_fixed_markers(self):
        for value in (
            SENSITIVE, "secret-looking-identifier", "a" * 1000, 123, True, [],
            "12345678-1234-1234-1234-123456789012", "max_tokens\nsecret",
            "max_tokens.secret", "unsupported_parameter/secret", "\uff2d\uff21\uff38\uff3f\uff34\uff2f\uff2b\uff25\uff2e\uff33",
        ):
            with self.subTest(value=value), candidate() as stub:
                error = FakeHttpError(SENSITIVE)
                error.status_code = 400
                error.code = error.param = value
                stub.create_completion.side_effect = error
                self.assert_failure(
                    "runtime", "FakeHttpError", status=400, phase="openai",
                    api_details=("api_code=unrecognized", "api_param=unrecognized"),
                )

    def test_generic_openai_error_emits_only_fixed_request_constraint_categories(self):
        messages = (
            ("Unsupported parameter: 'max_completion_tokens'.", "unsupported_max_completion_tokens"),
            ("'max_tokens' is not supported with this model.", "unsupported_max_tokens"),
            ("The chatCompletion operation does not work with the specified model.", "unsupported_model_operation"),
            ("API version 2025-01-01-preview is not supported.", "unsupported_api_version"),
            ("Invalid max_tokens: integer below minimum value.", "max_token_range"),
        )
        for text, expected in messages:
            for code in (None, "invalid_request_error", "BadRequest"):
                for source in ("args", "message"):
                    with self.subTest(text=text, code=code, source=source), candidate() as stub:
                        message = f"{SENSITIVE} {text} {SENSITIVE}"
                        error = FakeHttpError(message if source == "args" else SENSITIVE)
                        if source == "message":
                            error.message = message
                        error.status_code = 400
                        error.code = code
                        stub.create_completion.side_effect = error
                        details = (() if code is None else (f"api_code={code}",))
                        self.assert_failure(
                            "runtime", "FakeHttpError", status=400, phase="openai",
                            api_details=details + (f"request_constraint={expected}",),
                        )

    def test_request_constraint_matching_does_not_guess_from_unrelated_text_or_body(self):
        for message in (
            SENSITIVE, "max_tokens is supported", "max_tokens_secret is not supported",
            "secret_max_completion_tokens is not supported", "This model works with the operation",
            "Unsupported parameter: 'max_tokens.secret'.", "secret.max_tokens is not supported",
        ):
            with self.subTest(message=message), candidate() as stub:
                error = FakeHttpError(message)
                error.status_code = 400
                error.body = {"code": "unsupported_parameter", "param": "max_tokens",
                              "message": f"max_tokens is not supported {SENSITIVE}"}
                error.response = SimpleNamespace(
                    headers={"secret": SENSITIVE}, text=f"max_tokens is not supported {SENSITIVE}"
                )
                stub.create_completion.side_effect = error
                self.assert_failure("runtime", "FakeHttpError", status=400, phase="openai")

    def test_specific_code_does_not_trigger_message_inference(self):
        with candidate() as stub:
            error = FakeHttpError(f"max_tokens is not supported {SENSITIVE}")
            error.status_code = 400
            error.code = "unsupported_parameter"
            stub.create_completion.side_effect = error
            self.assert_failure(
                "runtime", "FakeHttpError", status=400, phase="openai",
                api_details=("api_code=unsupported_parameter",),
            )

    def test_openai_details_are_not_applied_to_other_phases_or_statuses(self):
        for phase, status in (("vision", 400), ("openai", 401)):
            with self.subTest(phase=phase, status=status), candidate() as stub:
                error = FakeHttpError(f"max_tokens is not supported {SENSITIVE}")
                error.status_code = status
                error.code, error.param = "unsupported_parameter", "max_tokens"
                if phase == "vision":
                    stub.vision.analyze.side_effect = error
                else:
                    stub.create_completion.side_effect = error
                self.assert_failure(
                    "runtime", "FakeHttpError", status=status, phase=phase,
                    passed=("openai",) if phase == "vision" else (),
                )

    def test_argument_errors_never_echo_secret_arguments(self):
        arguments = [
            ["--config", SENSITIVE],
            ["--version", VERSION, "--config", CONFIG, "--unknown", SENSITIVE],
            ["--version", VERSION, "--config", SENSITIVE, "--check-version=" + SENSITIVE],
            ["--vers", SENSITIVE, "--check-version"],
            ["--version", VERSION, "--check-version", SENSITIVE],
        ]
        for argv in arguments:
            with self.subTest(argv=argv), candidate() as stub:
                self.assert_failure("arguments", argv=argv)
                stub.installed_version.assert_not_called()
                self.assertEqual(stub.imports, [])
                stub.open_config.assert_not_called()


if __name__ == "__main__":
    unittest.main()
