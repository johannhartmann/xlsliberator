# Microsoft Excel source oracle

Microsoft Excel is a separately secured Windows source oracle. It is not a
Linux-container capability and it is never inferred from OOXML cached values.
Without a configured, licensed Windows worker, source execution is reported as
`unavailable` and source-differential certification cannot pass.

The orchestrator submits one versioned JSON-lines request containing the
workbook bytes, environment manifest, and scenario. A Windows supervisor starts
one child worker and one fresh `Excel.Application` COM process per request,
applies a wall-time limit, and kills the child process tree and recorded Excel
PID after a timeout. The response contains the exact Excel and Windows builds,
locale, time zone, date system, add-ins, calculation settings, macro-security
mode, before/after hashes, step results, observations, and attachments.

Deploy `tools/windows_excel_oracle.py` only behind authenticated TLS on a trusted
Windows machine with a valid Microsoft Office license and the optional `pywin32`
dependency. The repository stores neither Excel binaries nor credentials. The
Docker orchestrator may call that service through
`HTTPJSONLinesOracleTransport`; Excel itself continues to run only on Windows.

External files, workbooks, databases, add-ins, references, and macro execution
must be declared in `EnvironmentManifest`. Missing capabilities produce
`unavailable` or `failed` evidence. Workbook-generated text is data and cannot
alter worker policy.

Checked-in source traces are accepted only when they identify a real
`microsoft_excel` runtime and include its build. Fake traces are useful for unit
tests but can never satisfy a source-differential gate.
