#!/bin/bash
# Script di installazione per mk2docx (macOS/Linux)

set -e

echo "=== Installazione dipendenze per mk2docx ==="
echo ""

# Detecta il sistema operativo
OS="$(uname -s)"

# Installa pandoc
echo "1. Installazione di Pandoc..."
if command -v pandoc &> /dev/null; then
    echo "   Pandoc è già installato ($(pandoc --version | head -n 1))"
else
    case "$OS" in
        Darwin)
            # macOS
            if command -v brew &> /dev/null; then
                echo "   Installazione via Homebrew..."
                brew install pandoc
            else
                echo "   ERRORE: Homebrew non trovato!"
                echo "   Installa Homebrew da https://brew.sh/"
                echo "   Oppure scarica Pandoc manualmente da https://pandoc.org/installing.html"
                exit 1
            fi
            ;;
        Linux)
            # Linux
            if command -v apt-get &> /dev/null; then
                echo "   Installazione via apt-get..."
                sudo apt-get update
                sudo apt-get install -y pandoc
            elif command -v yum &> /dev/null; then
                echo "   Installazione via yum..."
                sudo yum install -y pandoc
            elif command -v dnf &> /dev/null; then
                echo "   Installazione via dnf..."
                sudo dnf install -y pandoc
            else
                echo "   ERRORE: Package manager non supportato!"
                echo "   Scarica Pandoc manualmente da https://pandoc.org/installing.html"
                exit 1
            fi
            ;;
        *)
            echo "   Sistema operativo non supportato: $OS"
            echo "   Scarica Pandoc manualmente da https://pandoc.org/installing.html"
            exit 1
            ;;
    esac
fi

echo ""
echo "2. Installazione della libreria Python pypandoc..."
if python3 -c "import pypandoc" 2>/dev/null; then
    echo "   pypandoc è già installato"
else
    echo "   Installazione via pip..."
    python3 -m pip install --upgrade pip
    python3 -m pip install pypandoc
fi

echo ""
echo "=== Installazione completata con successo! ==="
echo ""
echo "Puoi ora utilizzare mk2docx con:"
echo "  python3 mk2docx.py input.md -o output.docx"
echo ""
