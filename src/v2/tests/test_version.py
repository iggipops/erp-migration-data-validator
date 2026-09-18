import re

from __version__ import __version__


def test_version_matches_scheme():
    # vN_L_rK_M per docs/versioning_scheme.md
    assert re.fullmatch(r"v\d+_\d+_r\d+_\d+", __version__)


def test_entry_script_imports_same_version():
    import erp_migration_data_validator
    assert erp_migration_data_validator.__version__ == __version__
