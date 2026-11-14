#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
EXT_DIR="${ROOT_DIR}/extension"
DIST_DIR="${ROOT_DIR}/dist"
CHROME_DIR="${DIST_DIR}/chrome"
FIREFOX_DIR="${DIST_DIR}/firefox"

VERSION="$(node -e "const fs=require('fs'); console.log(JSON.parse(fs.readFileSync('${EXT_DIR}/manifest.json', 'utf8')).version);")"
if [[ -z "${VERSION}" ]]; then
  echo "Failed to determine extension version from manifest.json" >&2
  exit 1
fi

CHROME_ZIP="${DIST_DIR}/transcriptonic-chrome-v${VERSION}.zip"
FIREFOX_ZIP="${DIST_DIR}/transcriptonic-firefox-v${VERSION}.zip"

rm -rf "${CHROME_DIR}" "${FIREFOX_DIR}" "${CHROME_ZIP}" "${FIREFOX_ZIP}"
mkdir -p "${CHROME_DIR}" "${FIREFOX_DIR}"

# Copy all extension assets except source manifest template
copy_assets() {
  local target_dir="$1"
  for item in "${EXT_DIR}"/*; do
    local base
    base="$(basename "${item}")"
    if [[ "${base}" == "manifest.json" ]]; then
      continue
    fi
    cp -a "${item}" "${target_dir}/"
  done
}

copy_assets "${CHROME_DIR}"
copy_assets "${FIREFOX_DIR}"

# Generate clean Chrome/Edge/Opera MV3 manifest (enable Chrome side_panel & strip Firefox-only keys)
node -e '
const fs = require("fs");
const src = process.argv[1];
const dst = process.argv[2];
const manifest = JSON.parse(fs.readFileSync(src, "utf8"));
manifest.side_panel = { default_path: "side-panel/index.html" };
manifest.permissions = Array.isArray(manifest.permissions) ? manifest.permissions : [];
if (!manifest.permissions.includes("sidePanel")) {
  manifest.permissions.push("sidePanel");
}
manifest.background = {
  service_worker: "background-script/index.js",
  type: "module"
};
delete manifest.sidebar_action;
delete manifest.browser_specific_settings;
fs.writeFileSync(dst, JSON.stringify(manifest, null, 2) + "\n", "utf8");
' "${EXT_DIR}/manifest.json" "${CHROME_DIR}/manifest.json"

# Install Firefox MV3 manifest
cp "${EXT_DIR}/manifest.json" "${FIREFOX_DIR}/manifest.json"

# Create ZIP packages (use zip if installed, otherwise Python zipfile)
package_dir() {
  local src_dir="$1"
  local out_zip="$2"
  if command -v zip >/dev/null 2>&1; then
    (cd "${src_dir}" && zip -r -q "${out_zip}" .)
  else
    python3 - "${src_dir}" "${out_zip}" <<'PY'
import os
import sys
import zipfile

src_dir = sys.argv[1]
out_zip = sys.argv[2]

with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(src_dir):
        dirs.sort()
        files.sort()
        for fname in files:
            full_path = os.path.join(root, fname)
            rel_path = os.path.relpath(full_path, src_dir).replace(os.sep, "/")
            zf.write(full_path, rel_path)
PY
  fi
}

package_dir "${CHROME_DIR}" "${CHROME_ZIP}"
package_dir "${FIREFOX_DIR}" "${FIREFOX_ZIP}"

echo "Built Chrome MV3 directory: ${CHROME_DIR}"
echo "Built Firefox MV3 directory: ${FIREFOX_DIR}"
echo "Packaged Chrome zip: ${CHROME_ZIP}"
echo "Packaged Firefox zip: ${FIREFOX_ZIP}"
