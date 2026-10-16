#!/bin/bash
# Setup script for the access-modern-ui cloud environment.
# Paste this file into the environment's "Setup script" field at claude.ai/code
# (environment selector > settings icon). It runs as root on Ubuntu 24.04 before
# Claude starts, and the result is cached for about seven days.
# Needs the "Trusted" network level (packages.microsoft.com is on its allowlist).

# PowerShell 7, for the portable scripts: FormSnapshot.Tests.ps1 and VbaLineNumbers.ps1
if ! command -v pwsh >/dev/null 2>&1; then
  wget -q https://packages.microsoft.com/config/ubuntu/24.04/packages-microsoft-prod.deb -O /tmp/packages-microsoft-prod.deb \
    && dpkg -i /tmp/packages-microsoft-prod.deb \
    && apt-get update \
    && apt-get install -y powershell \
    || echo "PowerShell install failed; the session starts without pwsh"
  rm -f /tmp/packages-microsoft-prod.deb
fi

exit 0
