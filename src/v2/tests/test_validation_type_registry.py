from validators.registry import ValidationTypeRegistry


def test_register_builtins():
    reg = ValidationTypeRegistry()
    reg.register_builtins()
    for t in ("EMPTY", "NULL", "DUPLICATE", "DUPLICATE2", "FORMAT",
              "REFERENCE", "CUSTOM", "AI", "DATE", "NUMERIC", "INTEGER"):
        assert reg.has(t), f"{t} not registered"


def test_get_validate_fn_returns_callable():
    reg = ValidationTypeRegistry()
    reg.register_builtins()
    fn = reg.get_validate_fn("EMPTY")
    assert callable(fn)


def test_unknown_type_not_registered():
    reg = ValidationTypeRegistry()
    reg.register_builtins()
    assert reg.has("BOGUS") is False
    assert reg.get_validate_fn("BOGUS") is None


def test_list_types():
    reg = ValidationTypeRegistry()
    reg.register_builtins()
    types = reg.list_types()
    assert "EMPTY" in types
    assert "NUMERIC" in types
