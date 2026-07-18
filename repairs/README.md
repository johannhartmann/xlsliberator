# Generic repair records

Each directory contains a machine-readable record for one reusable repair.
Records bind classification, minimization, regression, patch, exact
fail-before/pass-after execution, affected-corpus evidence, skill updates,
independent review, and upstream review.

The checked-in TDF-172479 record points to the existing Docker-only LibreOffice
source-build evidence. A record is not proof by itself: its artifact hashes and
cross-file identities are validated in CI.
