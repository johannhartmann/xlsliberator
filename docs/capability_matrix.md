# XLSLiberator Capability Matrix

> Generated from corpus and evidence data. Do not edit by hand.

Target: LibreOffice `26.2.4.2` in pinned Docker images.

## Status semantics

`unavailable`, `skipped`, `unsupported`, `waived`, and `failed` are distinct. Rates use only decisive `passed` and `failed` results.

## Measurements

| Format | Artifact | Scenario | Environment | Parse | Output | Target runtime | Source differential | Tiers |
|---|---|---|---|---|---|---|---|---|
| XLS | formula | save-reopen-recalculate | docker-linux | unavailable | unavailable | unavailable | unavailable | none |
| XLSX | names | copy-move-rename | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| XLSM | vba | macro-execution | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| XLSB | charts | control-event | docker-linux | unavailable | unavailable | unavailable | unavailable | none |
| XLSX | formula | formula-heavy-target | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| XLSM | vba | vba-target-runtime | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| ODS | controls | controls-events-target | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| RECIPE | formula | formula-translation | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| RECIPE | formula | stock-vs-patched | docker-linux-aarch64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| RECIPE | formula | stock-vs-patched | docker-linux-aarch64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| RECIPE | malicious | sandbox-execution | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |
| RECIPE | malicious | sandbox-execution | docker-linux-arm64 | passed | passed | passed | unavailable | structural-inventory, target-runtime-validated, libreoffice-runtime-validated |

## Evidence identities

- `unavailable-generated-formula-xls`: runtime identity unavailable
- `sample-generated-names-xlsx`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-generated-vba-xlsm`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `unavailable-generated-controls-xlsb`: runtime identity unavailable
- `sample-sample-formula-heavy-xlsx`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-sample-vba-workbook`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-sample-controls-events-workbook`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-regression-indirect-address`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-public-tdf-172479`: `sha256:b0fd4dd05234bc7b2b631bba6adf0de258561e933d0133e8758918abea931ee1`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; aarch64; Python 3.12.13; variant `xlsliberator-text-functions-v1`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `428adcd3022b5d9c59c8cb28f81cc366c6317f9929f92ac6cb3e09b5223e5c7f`; source `0229ac93fcf0d7cbc6376066c6f35021cef002dc`; patch `7211ee90e1fefe803317ef80c75f2397682219c24ce6d8321bbd4f3bb914ed7b`; packages 1172
- `sample-regression-tdf-172479-minimized`: `sha256:b0fd4dd05234bc7b2b631bba6adf0de258561e933d0133e8758918abea931ee1`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; aarch64; Python 3.12.13; variant `xlsliberator-text-functions-v1`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `428adcd3022b5d9c59c8cb28f81cc366c6317f9929f92ac6cb3e09b5223e5c7f`; source `0229ac93fcf0d7cbc6376066c6f35021cef002dc`; patch `7211ee90e1fefe803317ef80c75f2397682219c24ce6d8321bbd4f3bb914ed7b`; packages 1172
- `sample-malicious-resource-exhaustion`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181
- `sample-malicious-file-exfiltration`: `sha256:5de5b3c6b45940817e6bf7a3de257753dc7284ba5dac3e5f9df960356d81ce0b`; base `sha256:60eac759739651111db372c07be67863818726f754804b8707c90979bda511df`; arm64; Python 3.12.13; variant `stock`; office `adf468b45764b2abce53a7d91bbf3056b33f2734c5d5f628c075753e73903c43`; UNO `6374b68bc6d38374857c0d8050ef5e6c1d1f10ae71d8c49bc5249e1366d9194a`; PyUNO `5dd6df3be0afb5fdc91453af06d5dc404262c02cecc15289b1760dcfa27ce065`; source `official-binary-distribution`; patch `none`; packages 181

## Formula corpus

- Minimized regression fixtures: 1
- Registered rules: 1
- Covered rules: 1
- Source differential: `not_measured`

## Release gates

- PASS `p0-tests`: P0 suite result
- PASS `fail-closed-certification`: no fail-open certification path
- PASS `required-corpus`: all blocking fixtures are accounted green
- PASS `source-artifact-accounting`: certified fixtures have complete dispositions
- PASS `evidence-schemas`: all evidence models validated
- PASS `security-suite`: blocking security suite result
- PASS `runtime-identities`: every target pass records binary/source identities

Release ready: **YES**

Certification tiers:

- `structural-inventory`: source artifacts were inventoried.
- `target-runtime-validated`: required target scenario passed.
- `source-differential-validated`: target matched a source trace.
- `libreoffice-runtime-validated`: target pass used exact pinned LibreOffice Docker evidence.
