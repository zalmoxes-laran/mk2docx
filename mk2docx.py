# no nobis

#!/usr/bin/env python3
import argparse
import os
import sys
from pathlib import Path

def convert_with_pypandoc(src: Path, dst: Path, reference_docx: Path | None):
    try:
        import pypandoc
    except ImportError:
        return False, "pypandoc non installato"
    extra_args = ["--from=markdown", "--to=docx"]
    if reference_docx and reference_docx.exists():
        extra_args += [f"--reference-doc={reference_docx}"]
    try:
        pypandoc.convert_file(str(src), "docx", outputfile=str(dst), extra_args=extra_args)
        return True, None
    except OSError as e:
        # Tipicamente indica che pandoc binario non è installato
        return False, f"pandoc non trovato o errore: {e}"
    except Exception as e:
        return False, str(e)

def convert_with_subprocess(src: Path, dst: Path, reference_docx: Path | None):
    import subprocess, shlex
    cmd = ["pandoc", str(src), "-o", str(dst)]
    if reference_docx and reference_docx.exists():
        cmd += ["--reference-doc", str(reference_docx)]
    try:
        subprocess.run(cmd, check=True)
        return True, None
    except FileNotFoundError:
        return False, "pandoc non è installato nel PATH"
    except subprocess.CalledProcessError as e:
        return False, f"pandoc ha restituito errore: {e}"

def main():
    parser = argparse.ArgumentParser(
        description="Converti un file Markdown in DOCX (Word) da riga di comando."
    )
    parser.add_argument("input_md", type=Path, help="Percorso del file .md in ingresso")
    parser.add_argument(
        "-o", "--output", type=Path, default=None,
        help="Percorso del file .docx in uscita (default: stesso nome con estensione .docx)"
    )
    parser.add_argument(
        "--refdoc", type=Path, default=None,
        help="File DOCX di riferimento per lo stile (opzionale)"
    )
    args = parser.parse_args()

    src = args.input_md
    if not src.exists():
        print(f"Errore: file di input non trovato: {src}", file=sys.stderr)
        sys.exit(1)

    dst = args.output or src.with_suffix(".docx")
    # Crea cartella di output, se manca
    dst.parent.mkdir(parents=True, exist_ok=True)

    # 1) pypandoc (preferito)
    ok, err = convert_with_pypandoc(src, dst, args.refdoc)
    if ok:
        print(f"Creato: {dst}")
        return

    # 2) subprocess su pandoc
    ok, err2 = convert_with_subprocess(src, dst, args.refdoc)
    if ok:
        print(f"Creato: {dst}")
        return

    # Se entrambe falliscono, mostra motivi
    print("Conversione fallita.", file=sys.stderr)
    if err:
        print(f"- Tentativo pypandoc: {err}", file=sys.stderr)
    if err2:
        print(f"- Tentativo pandoc (subprocess): {err2}", file=sys.stderr)
    print("Suggerimento: installa pandoc (vedi istruzioni) oppure assicurati che sia nel PATH.", file=sys.stderr)
    sys.exit(2)

if __name__ == "__main__":
    main()
