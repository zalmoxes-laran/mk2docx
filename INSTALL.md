# Installare mk2docx su macOS

Guida passo passo per installare mk2docx su un Mac nuovo o appena formattato, fino ad avere l'azione rapida del Finder che converte un `.md` in `.docx` nella stessa cartella.

Serve circa un quarto d'ora, quasi tutto di attesa durante l'installazione di Homebrew. Tutti i comandi vanno dati nel Terminale (Applicazioni › Utility › Terminale).

---

## 1. Strumenti da riga di comando di Apple

Contengono `git` e il compilatore che Homebrew richiede.

```bash
xcode-select --install
```

Si apre una finestra: clic su **Installa** e aspetta la fine. Se ti risponde che sono già installati, vai avanti.

## 2. Homebrew

Homebrew è il gestore di pacchetti che userai per installare Python e Pandoc.

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Ti chiederà la password del Mac. Alla fine l'installer stampa due righe da eseguire per aggiungere `brew` al PATH. Su un Mac Apple Silicon sono queste:

```bash
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"
```

Per controllare: `brew --version` deve rispondere con un numero di versione.

## 3. Python e Pandoc

```bash
brew install python pandoc
```

Pandoc è il motore che fa la conversione vera e propria. Lo script Python lo usa tramite la libreria `pypandoc`. Anche se questo passo lo salti, `install.sh` installa comunque Pandoc da sé. Python invece conviene prenderlo da Homebrew, perché quello di sistema è vecchio.

Per controllare: `python3 --version` e `pandoc --version`.

## 4. Scaricare il codice

Mettilo in `~/Documents/GitHub/`:

```bash
mkdir -p ~/Documents/GitHub && cd ~/Documents/GitHub
git clone https://github.com/zalmoxes-laran/mk2docx.git
cd mk2docx
```

Con GitHub Desktop va bene lo stesso (File › Clone repository). Conta solo che la cartella finisca dove vuoi tenerla, perché l'azione rapida punta a quel percorso.

## 5. Installare con `install.sh`

```bash
./install.sh
```

Lo script fa tre cose, nell'ordine:

1. controlla Pandoc e lo installa via Homebrew se manca;
2. crea l'ambiente virtuale `.venv` dentro la cartella del progetto e ci installa `pypandoc` (così il Python di sistema non viene toccato);
3. installa l'azione rapida **Convert md -> docx** in `~/Library/Services`. Il percorso della cartella viene rilevato e scritto nell'azione in automatico, quindi non c'è niente da configurare a mano.

Se ti serve solo la riga di comando, senza azione rapida, usa `./install.sh --no-quick-action`.

Se risponde `Permission denied`, dai prima `chmod +x install.sh`.

La `.venv` è esclusa da git (`.gitignore`), quindi va creata di nuovo su ogni macchina. Basta rilanciare `./install.sh`, che si può eseguire quante volte vuoi.

## 6. Usare l'azione rapida

Nel Finder fai clic destro su un file `.md`, poi **Azioni rapide** › **Convert md -> docx**. Il `.docx` viene creato accanto al file, con lo stesso nome, e una notifica ti avvisa quando è pronto. Puoi anche selezionare più `.md` insieme.

Se la voce non compare:

- riavvia il Finder con `killall Finder`;
- clic destro › **Azioni rapide** › **Personalizza…** e attiva *Convert md -> docx*;
- oppure vai in Impostazioni di Sistema › Generali › Elementi login ed estensioni › Finder.

La prima volta che la usi su una cartella di OneDrive, iCloud o Documenti, macOS può chiederti il permesso di accedervi: concedilo.

Attenzione: se nella cartella c'è già un `.docx` con lo stesso nome, viene sovrascritto.

## 7. Usarlo da terminale

```bash
~/Documents/GitHub/mk2docx/.venv/bin/python ~/Documents/GitHub/mk2docx/mk2docx.py documento.md -o documento.docx
```

Opzioni:

- `-o, --output` imposta il percorso del `.docx` (di default è lo stesso nome del `.md`);
- `--refdoc modello.docx` usa un file Word come modello di stili (titoli, font, spaziature).

Per comodità puoi aggiungere un alias in `~/.zshrc`:

```bash
alias mk2docx='~/Documents/GitHub/mk2docx/.venv/bin/python ~/Documents/GitHub/mk2docx/mk2docx.py'
```

---

## Solo in caso di problemi: creare l'azione rapida a mano

Normalmente non serve, perché `install.sh` fa già tutto.

1. Apri **Automator** › Nuovo documento › **Azione rapida**.
2. In alto: *Il flusso di lavoro riceve* **file o cartelle** *in* **Finder**. Questo è il punto decisivo: se resta su "testo", l'azione non compare nel Finder.
3. Aggiungi l'azione **Esegui script shell**, con Shell `/bin/zsh` e *Passa input* **come argomenti** (non "a stdin").
4. Incolla lo script e correggi `MK2DOCX_DIR` se la cartella è altrove:

```zsh
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

MK2DOCX_DIR="$HOME/Documents/GitHub/mk2docx"
SCRIPT="$MK2DOCX_DIR/mk2docx.py"
PY="$MK2DOCX_DIR/.venv/bin/python"
[ -x "$PY" ] || PY="$(command -v python3)"
LOG="$HOME/mk2docx_errors.log"

for f in "$@"; do
  case "$f" in
    *.md|*.markdown|*.MD)
      out="${f%.*}.docx"
      cd "$(dirname "$f")"
      if "$PY" "$SCRIPT" "$f" -o "$out" 2>>"$LOG"; then
        osascript -e "display notification \"$(basename "$out")\" with title \"DOCX creato\""
      else
        osascript -e "display notification \"Controlla ~/mk2docx_errors.log\" with title \"mk2docx: errore\""
      fi
      ;;
  esac
done
```

5. Salva con Cmd+S e un nome a scelta. Il file va in `~/Library/Services`.

Nota tecnica: il modello dell'azione sta in `macos/Convert md -> docx.workflow` e contiene il segnaposto `__MK2DOCX_DIR__`, così va bene su qualunque Mac. `install.sh` lo sostituisce con il percorso reale nella copia che installa. Il file nel repository non va toccato.

## Se qualcosa non funziona

- **Notifica "mk2docx: errore"**: leggi `~/mk2docx_errors.log`. Di solito manca Pandoc (`brew install pandoc`) oppure la `.venv` (rilancia `./install.sh`).
- **Dopo aver spostato la cartella del progetto**: rilancia `./install.sh`, che reinstalla l'azione rapida con il nuovo percorso.
- **Dopo un aggiornamento di Python con Homebrew**, se la `.venv` smette di funzionare: `rm -rf .venv && ./install.sh`.
- **Rimuovere l'azione rapida**: cancella `~/Library/Services/Convert md -> docx.workflow`.
