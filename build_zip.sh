#!/usr/bin/env bash
# build_zip.sh
# T.A.N.G.S - One-command project ZIP builder
#
# Usage:  bash build_zip.sh
# Output: tangs_project.zip   (ready to upload to GitHub or submit)
#
# Run from the project root directory (same folder as app.py).
# ─────────────────────────────────────────────────────────────────────

set -e   # Exit immediately on any error

PROJECT_DIR="tangs_project"
ZIP_NAME="tangs_project.zip"

echo "═══════════════════════════════════════════════════"
echo "  T.A.N.G.S — Building project ZIP..."
echo "═══════════════════════════════════════════════════"

# ── Clean previous build ──────────────────────────────────────────────
if [ -d "$PROJECT_DIR" ]; then
    echo "→ Removing previous build directory: $PROJECT_DIR/"
    rm -rf "$PROJECT_DIR"
fi
if [ -f "$ZIP_NAME" ]; then
    echo "→ Removing previous ZIP: $ZIP_NAME"
    rm -f "$ZIP_NAME"
fi

# ── Create directory structure ────────────────────────────────────────
echo "→ Creating directory structure..."
mkdir -p "$PROJECT_DIR/.streamlit"

# ── Copy Python modules ───────────────────────────────────────────────
echo "→ Copying Python modules..."
for f in \
    models.py \
    pdf_extractor.py \
    academic_module.py \
    timetable_module.py \
    financial_module.py \
    query_router.py \
    output_generator.py \
    app.py
do
    if [ -f "$f" ]; then
        cp "$f" "$PROJECT_DIR/"
        echo "   ✓ $f"
    else
        echo "   ✗ WARNING: $f not found — skipping"
    fi
done

# ── Copy configuration and support files ─────────────────────────────
echo "→ Copying configuration files..."
for f in requirements.txt packages.txt README.md; do
    if [ -f "$f" ]; then
        cp "$f" "$PROJECT_DIR/"
        echo "   ✓ $f"
    else
        echo "   ✗ WARNING: $f not found — skipping"
    fi
done

# ── Copy Streamlit config ─────────────────────────────────────────────
if [ -f ".streamlit/config.toml" ]; then
    cp ".streamlit/config.toml" "$PROJECT_DIR/.streamlit/"
    echo "   ✓ .streamlit/config.toml"
else
    echo "   ✗ WARNING: .streamlit/config.toml not found — skipping"
fi

# ── Count files copied ────────────────────────────────────────────────
FILE_COUNT=$(find "$PROJECT_DIR" -type f | wc -l | tr -d ' ')
echo ""
echo "→ $FILE_COUNT file(s) staged in $PROJECT_DIR/"

# ── Create ZIP archive ────────────────────────────────────────────────
echo "→ Creating $ZIP_NAME..."
zip -r "$ZIP_NAME" "$PROJECT_DIR/" -x "*.pyc" -x "*/__pycache__/*" -x "*/.DS_Store"

# ── Clean up staging directory ────────────────────────────────────────
echo "→ Cleaning up staging directory..."
rm -rf "$PROJECT_DIR"

# ── Report ────────────────────────────────────────────────────────────
ZIP_SIZE=$(du -sh "$ZIP_NAME" | cut -f1)
echo ""
echo "═══════════════════════════════════════════════════"
echo "  ✅  Done!  →  $ZIP_NAME  ($ZIP_SIZE)"
echo "═══════════════════════════════════════════════════"
echo ""
echo "Next steps:"
echo "  1. Unzip locally:       unzip $ZIP_NAME"
echo "  2. Push to GitHub:      git init && git add . && git commit -m 'T.A.N.G.S' && git push"
echo "  3. Deploy free:         share.streamlit.io → New app → select repo → app.py"
echo "  4. Add secret on Cloud: GROQ_API_KEY = \"gsk_your_key_here\""
echo ""
echo "Get a free Groq API key at: https://console.groq.com"
