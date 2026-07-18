"""Deprecated embedded model orchestration.

This optional namespace exists only for short-term compatibility. New code must
use the external xlsliberator-swe orchestration repository and pass generated
target-native artifacts to the deterministic XLSLiberator primitives.
"""

from warnings import warn

warn(
    "xlsliberator.legacy_agent is deprecated; use xlsliberator-swe for model "
    "orchestration and pass target-native artifacts to deterministic primitives",
    DeprecationWarning,
    stacklevel=2,
)
