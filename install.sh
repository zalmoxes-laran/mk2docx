#!/bin/bash
# Installazione di mk2docx (macOS/Linux)
#   1. installa Pandoc (via Homebrew su macOS, via apt/yum/dnf su Linux)
#   2. crea l'ambiente virtuale .venv nella cartella del progetto e vi installa pypandoc
#   3. (solo macOS) installa l'azione rapida del Finder "Convert md -> docx"
#
# Uso:  ./install.sh              installazione completa
#       ./install.sh --no-quick-action   salta il punto 3

set -e

DIR="$(cd "$(dirname "$0")" && pwd)"
OS="$(uname -s)"
QUICK_ACTION=1
[ "$1" = "--no-quick-action" ] && QUICK_ACTION=0

# Su macOS con Apple Silicon Homebrew sta in /opt/homebrew, su Intel in /usr/local
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

echo "=== Installazione di mk2docx ==="
echo "Cartella del progetto: $DIR"
echo ""

# --- 1. Pandoc -------------------------------------------------------------
echo "1. Pandoc..."
if command -v pandoc &> /dev/null; then
    echo "   già installato ($(pandoc --version | head -n 1))"
else
    case "$OS" in
        Darwin)
            if command -v brew &> /dev/null; then
                echo "   installazione via Homebrew..."
                brew install pandoc
            else
                echo "   ERRORE: Homebrew non trovato."
                echo "   Installalo con:"
                echo '   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
                echo "   poi rilancia ./install.sh (vedi INSTALL.md)."
                exit 1
            fi
            ;;
        Linux)
            if command -v apt-get &> /dev/null; then
                sudo apt-get update && sudo apt-get install -y pandoc python3-venv
            elif command -v dnf &> /dev/null; then
                sudo dnf install -y pandoc
            elif command -v yum &> /dev/null; then
                sudo yum install -y pandoc
            else
                echo "   ERRORE: package manager non supportato. Scarica Pandoc da https://pandoc.org/installing.html"
                exit 1
            fi
            ;;
        *)
            echo "   Sistema non supportato: $OS"
            exit 1
            ;;
    esac
fi

# --- 2. Ambiente virtuale ----------------------------------------------------
echo ""
echo "2. Ambiente virtuale Python (.venv)..."
if ! command -v python3 &> /dev/null; then
    echo "   ERRORE: python3 non trovato. Su macOS: brew install python"
    exit 1
fi
if [ ! -x "$DIR/.venv/bin/python" ]; then
    python3 -m venv "$DIR/.venv"
    echo "   creato con $(python3 --version)"
else
    echo "   già presente"
fi
"$DIR/.venv/bin/python" -m pip install --quiet --upgrade pip
"$DIR/.venv/bin/python" -m pip install --quiet --upgrade pypandoc
echo "   pypandoc installato in .venv"

# --- 3. Azione rapida del Finder (macOS) -----------------------------------
if [ "$OS" = "Darwin" ] && [ "$QUICK_ACTION" = "1" ]; then
    echo ""
    echo "3. Azione rapida del Finder..."
    NAME="Convert md -> docx.workflow"
    SRC="$DIR/macos/$NAME"
    DST="$HOME/Library/Services/$NAME"
    if [ ! -d "$SRC" ]; then
        echo "   modello non trovato in macos/, salto (vedi INSTALL.md per crearla a mano)"
    else
        mkdir -p "$HOME/Library/Services"
        rm -rf "$DST"
        cp -R "$SRC" "$DST"
        WF="$DST/Contents/document.wflow"
        CMD="$(plutil -extract actions.0.action.ActionParameters.COMMAND_STRING raw -o - "$WF")"
        CMD="${CMD//__MK2DOCX_DIR__/$DIR}"
        plutil -replace actions.0.action.ActionParameters.COMMAND_STRING -string "$CMD" "$WF"
        /System/Library/CoreServices/pbs -flush  >/dev/null 2>&1 || true
        /System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true
        echo "   installata in ~/Library/Services"
        echo "   (se non compare subito: killall Finder, oppure Azioni rapide > Personalizza...)"
    fi
fi

echo ""
echo "=== Installazione completata ==="
echo ""
echo "Da terminale:"
echo "  $DIR/.venv/bin/python $DIR/mk2docx.py file.md -o file.docx"
[ "$OS" = "Darwin" ] && echo "Dal Finder: clic destro su un .md > Azioni rapide > Convert md -> docx"
echo ""
