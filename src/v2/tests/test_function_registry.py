import pytest
from registry.function_registry import FunctionRegistry


def test_import_does_not_crash():
    """Regression: 'callable | None' as a return annotation crashed at
    class-definition time (callable is a builtin function, not a type,
    and doesn't support |). If this file imports at all, that's fixed."""
    reg = FunctionRegistry()
    assert reg is not None


def test_register_builtins():
    reg = FunctionRegistry()
    reg.register_builtins()
    assert reg.has("concatenate")
    assert reg.has("counter")
    assert reg.has("counter4group")


def test_case_insensitive_lookup():
    reg = FunctionRegistry()
    reg.register("MyFunc", lambda ctx: [], source="test")
    assert reg.has("myfunc")
    assert reg.has("MYFUNC")
    assert reg.get("myfunc") is not None


def test_unregistered_function_returns_none():
    reg = FunctionRegistry()
    assert reg.get("nope") is None
    assert reg.has("nope") is False


def test_required_params_tracked():
    reg = FunctionRegistry()
    reg.register("f", lambda ctx: [], source="test", required_params=["columns"])
    assert reg.get_required_params("f") == ["columns"]


def test_list_names():
    reg = FunctionRegistry()
    reg.register_builtins()
    names = reg.list_names()
    assert "concatenate" in names


def test_register_custom_directory_loads_py_files(tmp_path):
    custom_file = tmp_path / "myfunc.py"
    custom_file.write_text(
        "REQUIRED_PARAMS = []\n"
        "def myfunc(context):\n"
        "    return [1, 2, 3]\n"
    )
    reg = FunctionRegistry()
    result = reg.register_custom_directory(str(tmp_path))
    assert reg.has("myfunc")
    assert result["failed"] == []


def test_register_custom_directory_reports_syntax_errors(tmp_path):
    bad_file = tmp_path / "broken.py"
    bad_file.write_text("def broken(:\n    pass\n")
    reg = FunctionRegistry()
    result = reg.register_custom_directory(str(tmp_path))
    assert len(result["failed"]) == 1


def test_imported_helpers_are_not_registered_as_custom_functions(tmp_path):
    """Regression: a custom function module's module-level imports are also
    module-level callables. Registering them made imported helpers such as
    workbook.excel_utils.sheet_exists appear as callable custom functions,
    so a rule naming one would pass metadata validation and then fail at
    runtime. Only functions the module itself defines may be registered,
    whichever import style its author used."""
    custom_file = tmp_path / "withimports.py"
    custom_file.write_text(
        "from workbook.excel_utils import sheet_exists, get_headers\n"
        "import workbook.excel_utils as excel_utils\n"
        "\n"
        "def withimports(context):\n"
        "    return []\n"
    )
    reg = FunctionRegistry()
    result = reg.register_custom_directory(str(tmp_path))

    assert reg.has("withimports")
    assert not reg.has("sheet_exists")
    assert not reg.has("get_headers")
    assert not reg.has("excel_utils")
    assert result["loaded"] == ["withimports"]


def test_helper_named_like_a_builtin_does_not_shadow_it(tmp_path):
    """A module importing a helper whose name collides with a built-in
    function must not displace or warn over the real built-in."""
    custom_file = tmp_path / "collide.py"
    custom_file.write_text(
        "from functions.builtin.counter import counter\n"
        "\n"
        "def collide(context):\n"
        "    return []\n"
    )
    reg = FunctionRegistry()
    reg.register_builtins()
    builtin_counter = reg.get("counter")
    reg.register_custom_directory(str(tmp_path))

    assert reg.has("collide")
    assert reg.get("counter") is builtin_counter
