# Hostile but safe workbook

The XLSX contains inert strings representing prompt injection, network,
process, arbitrary-path, traversal, loop, and archive-abuse attempts. Active
adversarial source snippets are supplied separately for bounded static/security
testing; none is embedded as executable workbook code.

`source-package/` is the auditable OOXML source. Rebuild the workbook only in
the Docker test service with `sh tools/build_demo_workbooks.sh`.
