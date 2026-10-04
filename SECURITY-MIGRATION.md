# Electron downloader security migration

Use the current @electron/get 5.1.0 download API on Node 22.22.3+, which replaces the legacy Got/cacheable-request chain. This removes http-cache-semantics rather than retaining a security backport. Remove patch-package, whose own glob dependency introduced the newly reported braces alert. Electron application behavior is unchanged. Validate with npm run test:security, unsigned packaging and actual Electron/Flask startup.
