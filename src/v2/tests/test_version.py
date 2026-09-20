import re
from pathlib import Path

from __version__ import __version__


def test_version_matches_scheme():
    # vN_L_rK_M per docs/versioning_scheme.md
    assert re.fullmatch(r"v\d+_\d+_r\d+_\d+", __version__)


def test_entry_script_imports_same_version():
    import erp_migration_data_validator
    assert erp_migration_data_validator.__version__ == __version__


def test_changelog_top_entry_matches_version():
    """
    Guards against the __version__.py/CHANGELOG.md drift that produced the
    orphaned v2_9_r0_3 changelog entry: a version bumped in one file but
    never in the other. The topmost dated "## [vN_L_rK_M] - date" heading
    in CHANGELOG.md (an "[Unreleased]" heading above it, if present, is not
    dated and is skipped) must always name the current __version__.
    """
    changelog = next(
        p for p in Path(__file__).resolve().parents if (p / "CHANGELOG.md").exists()
    ) / "CHANGELOG.md"
    text = changelog.read_text(encoding="utf-8")

    match = re.search(r"^## \[(v\d+_\d+_r\d+_\d+)\] - \d{4}-\d{2}-\d{2}", text, re.MULTILINE)
    assert match, f"No dated '## [vN_L_rK_M] - date' heading found in {changelog}"
    assert match.group(1) == __version__
