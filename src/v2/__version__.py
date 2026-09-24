"""
Single source of truth for the codebase version.

Format is vN_L_rK_M per docs/versioning_scheme.md — N_L_rK identifies the
Functional Specification version this codebase implements; M counts
code-only change rounds against that same FS version (resets to 0
whenever N, L, or K changes).
"""

__version__ = "v2_9_r2_0"
