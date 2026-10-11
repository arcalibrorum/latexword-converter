# LatexWord Converter

This tool converts an Overleaf or LaTeX project `.zip` package into an editable Microsoft Word `.docx` file.

The generated Word file uses manuscript-style formatting and keeps text, equations, figures, and captions editable whenever possible.

## Requirements

- Python 3
- Pandoc

On systems that use environment modules, the tool can load a configured pandoc module automatically.

On a personal computer, install pandoc first:

- macOS: `brew install pandoc`
- Linux: install pandoc with the system package manager
- Windows: install pandoc from the official installer, then use PowerShell or WSL

## Quick Start

Unzip this package:

```bash
unzip latex2reviewdocx_release.zip
cd latex2reviewdocx_release
```

Run the converter:

```bash
./latex2reviewdocx /path/to/overleaf_package.zip /path/to/output.docx
```

Example:

```bash
./latex2reviewdocx ~/paper/overleaf_project.zip ~/paper/paper_review.docx
```

The first path is the Overleaf `.zip` package.

The second path is where the Word file should be saved.

The output path must include the `.docx` file name, not only a directory.

## Optional Setup

If you want to run `latex2reviewdocx` from any folder, run this once:

```bash
./install_for_me.sh
```

Then open a new terminal or log in again, and check:

```bash
which latex2reviewdocx
```

After that, you can run:

```bash
latex2reviewdocx ~/paper/overleaf_project.zip ~/paper/paper_review.docx
```

## If The Main TeX File Is Not Detected

Most projects work automatically. If the program asks for the main TeX file, run:

```bash
./latex2reviewdocx ~/paper/overleaf_project.zip ~/paper/paper_review.docx --main-tex main.tex
```

Replace `main.tex` with the main file name used by the Overleaf project.

## Remote Conversion

The Python entry point can also synchronize an extracted LaTeX project to a remote machine, run pandoc there, and download the result:

```bash
python3 latex2reviewdocx.py remote \
  --project-dir ~/paper/project \
  --main-tex main.tex \
  --output ~/paper/paper_review.docx \
  --remote user@cluster.example.edu \
  --remote-project-dir /remote/work/path
```

Both the account and remote path must be supplied explicitly.

## Notes

The output Word file is intended for editing and review. It may not exactly match the PDF layout, and complex custom LaTeX commands may require manual cleanup in Word.
