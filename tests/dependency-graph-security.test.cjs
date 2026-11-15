const test=require("node:test");const assert=require("node:assert/strict");const lock=require("../package-lock.json");test("Electron download dependency graph removes vulnerable cache and braces",()=>{for(const name of ["http-cache-semantics","braces"]){assert.equal(Object.keys(lock.packages).some(p=>p.endsWith("node_modules/"+name)),false,name)}});

test("minimatch versions affected by globparts amplification use 11.0.1 or newer", () => {
  const entries = Object.entries(lock.packages).filter(([path]) =>
    path.endsWith("node_modules/minimatch")
  );
  assert.ok(entries.length > 0);
  for (const [path, { version }] of entries) {
    const [major, minor, patch] = version.split(".").map(Number);
    // GHSA-g3hv-75wg-m249 affects >=7.0.0 <11.0.1. Older callers
    // retain their patched legacy versions because the callable API changed.
    assert.ok(major < 7 || major > 11 ||
      (major === 11 && (minor > 0 || patch >= 1)), `${path}: ${version}`);
  }
  assert.equal(lock.packages["node_modules/minimatch"].version, "11.0.1");
});
