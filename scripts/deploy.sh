#!/usr/bin/env bash
#
# Deploy a generated LIC-DSF package to its target GitHub repository.
#
# Usage: ./scripts/deploy.sh <template-date>
#   e.g.: ./scripts/deploy.sh 2025-08-12
#
# Requires:
#   - DEPLOY_TOKEN env var (GitHub PAT with contents:write on target repos)
#   - Generated package at dist/lic-dsf-<template-date>/
#
set -euo pipefail

TEMPLATE_DATE="${1:?Usage: deploy.sh <template-date>}"
REPO_ORG="Teal-Insights"
REPO_NAME="lic-dsf-${TEMPLATE_DATE}"
PACKAGE_DIR="lic_dsf_${TEMPLATE_DATE//-/_}"

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST_DIR="${PROJECT_ROOT}/dist/lic-dsf-${TEMPLATE_DATE}"
WORK_DIR="$(mktemp -d)"

trap 'rm -rf "$WORK_DIR"' EXIT

# Validate inputs
if [ ! -d "$DIST_DIR/$PACKAGE_DIR" ]; then
    echo "ERROR: Generated package not found at $DIST_DIR/$PACKAGE_DIR"
    echo "Run the export pipeline first: uv run python -m src.lic_dsf_export --template $TEMPLATE_DATE"
    exit 1
fi

if [ -z "${DEPLOY_TOKEN:-}" ]; then
    echo "ERROR: DEPLOY_TOKEN environment variable is not set"
    exit 1
fi

REMOTE_URL="https://x-access-token:${DEPLOY_TOKEN}@github.com/${REPO_ORG}/${REPO_NAME}.git"

echo "==> Deploying lic-dsf-${TEMPLATE_DATE} to ${REPO_ORG}/${REPO_NAME}"

# Clone (or init if empty) the target repo
if git ls-remote "$REMOTE_URL" HEAD &>/dev/null; then
    echo "==> Cloning target repo (shallow)..."
    git clone --depth 1 "$REMOTE_URL" "$WORK_DIR/repo" 2>/dev/null || {
        # Empty repo — initialize fresh
        echo "==> Target repo is empty, initializing..."
        git init "$WORK_DIR/repo"
        git -C "$WORK_DIR/repo" remote add origin "$REMOTE_URL"
        git -C "$WORK_DIR/repo" checkout -b main
    }
else
    echo "ERROR: Cannot reach ${REPO_ORG}/${REPO_NAME}"
    exit 1
fi

cd "$WORK_DIR/repo"

# Configure git for CI
git config user.name "github-actions[bot]"
git config user.email "41898282+github-actions[bot]@users.noreply.github.com"

# Set up Git LFS for large files
echo "==> Configuring Git LFS..."
git lfs install --local 2>/dev/null || true
if [ ! -f .gitattributes ]; then
    cat > .gitattributes <<'ATTR'
# Track large generated Python modules with Git LFS
*.py filter=lfs diff=lfs merge=lfs -text
ATTR
    git add .gitattributes
fi

# Clear existing package content (so deletions are tracked)
rm -rf "$PACKAGE_DIR" pyproject.toml README.md

# Copy generated package
echo "==> Syncing generated package..."
cp -r "$DIST_DIR/$PACKAGE_DIR" "$PACKAGE_DIR"

# Generate pyproject.toml from template
VERSION="1.0.0"
if [ -f pyproject.toml ]; then
    # Preserve existing version if present
    EXISTING_VERSION=$(python3 -c "
import re, sys
text = open('pyproject.toml').read()
m = re.search(r'version\s*=\s*\"([^\"]+)\"', text)
print(m.group(1) if m else '1.0.0')
" 2>/dev/null || echo "1.0.0")
    VERSION="$EXISTING_VERSION"
fi
sed -e "s/{{TEMPLATE_DATE}}/${TEMPLATE_DATE}/g" \
    -e "s/{{VERSION}}/${VERSION}/g" \
    "$SCRIPT_DIR/pyproject.toml.template" > pyproject.toml

# Generate a minimal README
cat > README.md <<README
# lic-dsf-${TEMPLATE_DATE}

Python library duplicating the ${TEMPLATE_DATE} version of the IMF's LIC DSF
(Low-Income Country Debt Sustainability Framework) template.

This package is **auto-generated** by
[lic-dsf-programmatic-extraction](https://github.com/${REPO_ORG}/lic-dsf-programmatic-extraction).

## Installation

\`\`\`bash
pip install lic-dsf-${TEMPLATE_DATE}
\`\`\`

## Usage

\`\`\`python
from ${PACKAGE_DIR} import make_context, compute_all

ctx = make_context()
results = compute_all(ctx)
\`\`\`
README

# Stage all changes
git add -A

# Check if there are actual changes to commit
if git diff --cached --quiet 2>/dev/null; then
    echo "==> No changes detected, skipping commit."
    exit 0
fi

echo "==> Committing and pushing..."
git commit -m "Update generated package from lic-dsf-programmatic-extraction

Source: https://github.com/${REPO_ORG}/lic-dsf-programmatic-extraction
Template: ${TEMPLATE_DATE}"

git push -u origin main

echo "==> Done! Deployed to https://github.com/${REPO_ORG}/${REPO_NAME}"
