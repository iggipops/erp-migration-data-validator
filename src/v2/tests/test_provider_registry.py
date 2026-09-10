from registry.provider_registry import ProviderRegistry


class FakeProviderModule:
    PROVIDER_NAME = "fakeprovider"
    REQUIRED_SETTINGS = ["ai_endpoint"]


class FakeProviderModuleNoRequiredSettings:
    PROVIDER_NAME = "bareprovider"


def test_register_builtins():
    reg = ProviderRegistry()
    reg.register_builtins()
    assert reg.has("anthropic")


def test_case_insensitive_lookup():
    reg = ProviderRegistry()
    reg.register(FakeProviderModule)
    assert reg.has("fakeprovider")
    assert reg.has("FAKEPROVIDER")
    assert reg.get("FakeProvider") is FakeProviderModule


def test_unregistered_provider_returns_none():
    reg = ProviderRegistry()
    assert reg.get("nope") is None
    assert reg.has("nope") is False


def test_required_settings_tracked():
    reg = ProviderRegistry()
    reg.register(FakeProviderModule)
    assert reg.get_required_settings("fakeprovider") == ["ai_endpoint"]


def test_required_settings_defaults_to_empty_list_when_not_declared():
    reg = ProviderRegistry()
    reg.register(FakeProviderModuleNoRequiredSettings)
    assert reg.get_required_settings("bareprovider") == []


def test_required_settings_for_unregistered_provider_is_empty_list():
    reg = ProviderRegistry()
    assert reg.get_required_settings("nope") == []


def test_list_names():
    reg = ProviderRegistry()
    reg.register_builtins()
    assert reg.list_names() == ["anthropic"]


def test_duplicate_registration_keeps_first_and_does_not_raise():
    reg = ProviderRegistry()
    reg.register(FakeProviderModule)
    reg.register(FakeProviderModule)
    assert reg.list_names() == ["fakeprovider"]
