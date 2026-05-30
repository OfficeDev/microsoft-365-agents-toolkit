# Vulnerability fix TODO

This file was generated automatically because the CD vulnerability scan
found an issue that could not be patched mechanically. Please review and
replace this file with the actual fix, then re-open this PR.

- Ecosystem: `npm`
- Scan target: `templates/vsc`
- File: `templates/vsc/common/declarative-agent-meta-os-new-project/package.json.tpl`
- Package: `serialize-javascript`
- Current version: ``
- Fixed version: `14.0.0`
- Severity: `high`
- Advisory: https://github.com/advisories/GHSA-5c6j-r48x-rmvq
- Title: Serialize JavaScript is Vulnerable to RCE via RegExp.flags and Date.prototype.toISOString()
