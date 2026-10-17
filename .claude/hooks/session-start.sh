#!/bin/bash
# SessionStart hook for cloud sessions (claude.ai/code, claude --cloud).
# Exits at once outside the cloud. In the cloud it tells Claude what this Linux
# VM can and cannot run, since the kit itself needs Windows and Microsoft Access.
# Whatever this script prints is added to Claude's context at session start.

if [ "$CLAUDE_CODE_REMOTE" != "true" ]; then
  exit 0
fi

cd "$CLAUDE_PROJECT_DIR" || exit 0

if command -v pwsh >/dev/null 2>&1; then
  pwsh_line="pwsh $(pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>/dev/null) is installed."
else
  pwsh_line="pwsh is NOT installed: add .claude/cloud/setup.sh to this cloud environment's setup script."
fi

cat <<EOF
Cloud session notes for access-modern-ui (Linux VM, no Microsoft Access):
- $pwsh_line
- Can run here:
  - pwsh -NoProfile -File tools/review-mcp/FormSnapshot.Tests.ps1   (form snapshot diff checks, no Access needed)
  - python3 tools/make_themes.py <app_src_folder>                  (palette generator and contrast check)
  - tools/VbaLineNumbers.ps1, dot-sourced in pwsh                   (line-numbering of .bas/.vba text)
- Cannot run here: anything that drives Access over COM, such as examples/build.ps1,
  examples/demo/build.ps1, tests/native-smoke.ps1, tools/AccessWatch.ps1 and the
  Access Review MCP (tools/review-mcp/AccessReview*.ps1). Say which native checks
  still need a run on Windows instead of claiming they passed.
- .bas, .vba and .ps1 files are CRLF (.gitattributes). Keep CRLF line endings when editing them.
EOF
exit 0
