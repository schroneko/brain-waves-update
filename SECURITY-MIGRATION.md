# Electron downloader security migration

Use the current @electron/get 5.1.0 download API on Node 22.22.3+, which replaces the legacy Got/cacheable-request chain. This removes http-cache-semantics rather than retaining a security backport. Remove patch-package, whose own glob dependency introduced the newly reported braces alert. Electron application behavior is unchanged. Validate with npm run test:security, unsigned packaging and actual Electron/Flask startup.

## Independent validation, 2026-10-25

A fresh Python3.11.15 environment installed both requirements files. The development pins now include Black26.10.1 and setuptools84.2.1 for newly observed advisories. Runtime and development dependency audits report zero vulnerabilities. Two new stdlib unittests exercise the actual Flask upload, synthetic EEG analysis, generated Word document with 15 images, and attachment download. These pass after the final development install.

The Electron graph additionally updates compatible fast-uri and Undici transitive patches. Clean npm ci, existing security tests and npm audit pass (zero vulnerabilities). Electron42.11.18's existing Mac x64 target packaged successfully as an unsigned local app directory with publication explicitly disabled. The GUI and real hardware/data were not exercised; signing, notarization and Windows packaging remain unverified.

The repository has historical Heroku deployments and an active webhook. Current external auto-deployment settings are not fully exposed by these reads, so applying the branch may publish externally. Hosted application CI at the final head is still required separately; skipped Dependabot jobs and GitGuardian success are not application-test evidence.
