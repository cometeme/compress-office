# Compress Office 

English | [中文](README_ZH_CN.md)

A simple tool that losslessly compresses images in docx/pptx/xlsx documents.

![screenshot](screenshot/example.gif)

## Installing

Clone this repo first:

```
git clone https://github.com/cometeme/compress-office.git
```

Then we need:

1. **uv**: https://docs.astral.sh/uv/ (Python package manager)
2. **optipng**: http://optipng.sourceforge.net/ (PNG lossless compression)
3. **zopfli**: https://github.com/google/zopfli (PNG recompressor, stronger than optipng)
4. **pngcrush**: https://pmt.sourceforge.io/pngcrush/ (PNG optimizer)
5. **jpegoptim**: https://github.com/tjko/jpegoptim (JPEG lossless optimization)
6. **jpeg**: https://ijg.org/ (provides jpegtran for JPEG Huffman table optimization)
7. **gifsicle**: https://www.lcdf.org/gifsicle/ (GIF lossless optimization)
8. **trash**: https://github.com/ali-rantakari/trash
9. **fd** (Optional, speeds up file search): https://github.com/sharkdp/fd

If you have Homebrew, run these commands to install dependencies:

```
brew install uv optipng zopfli pngcrush jpegoptim jpeg gifsicle trash
```

Then install Python dependencies with uv:

```
uv sync
```

## Usage

```
uv run python compress_office.py [OPTIONS] PATH [PATH ...]
```

### Options

| Option | Description |
|--------|-------------|
| `-w N`, `--workers N` | Number of parallel workers (default: auto) |
| `--no-parallel` | Disable parallel compression |
| `-h`, `--help` | Show help message |

### Examples

```bash
# Compress all Office files in a directory
uv run python compress_office.py ~/Documents

# Compress specific files
uv run python compress_office.py report.docx slides.pptx

# Use 2 parallel workers
uv run python compress_office.py -w 2 ~/Documents

# Sequential (no parallelism)
uv run python compress_office.py --no-parallel ~/Documents
```

When the program runs for the first time, it will create a file called `process_history.csv`, which records the path of the compressed file and its modify time. When running again, if program finds that the file has not changed (file path is in the history and its modify time is same as recorded), then the program will skip the file without re-compressing, because doing so is not meaningful. If you really need to recompress a file, just delete its record from csv file.

After the compression is complete, the original document will be moved to recycle bin. **Please check your documents before emptying the recycle bin.**

## Use `fd` to speed up searching (optional)

The program uses python's built-in `glob` for file traversal by default, but it is very slow when there are many files, so the program supports the use of [fd](https://github.com/sharkdp/fd) to speed up the search.

If you have Homebrew, enter `brew install fd` in the console to install `fd`.

## FAQ

### Q1: Can I use this tool on other platforms?

Yes. This tool uses standard command-line tools (optipng, jpegoptim, gifsicle, trash) that are available on most platforms. To use on non-macOS systems:

- Modify `cache_folder` in `office.py` to a platform-appropriate temp directory
- Install the equivalent image compression tools for your platform
- Replace `trash` with an alternative (e.g., `gio trash` on Linux)

### Q2: How does this tool compress images? Will it corrupt my documents?

The docx/pptx/xlsx document is essentially a zip compressed package. This tool extracts the document to a cache directory, compresses all PNG/JPEG/GIF images losslessly using optipng, jpegoptim, and gifsicle, then repacks everything back.

Therefore, **in theory**, using this tool will not corrupt files when compressing, but in order to prevent unpredictable bugs, it is recommended to backup files before compressing them. At the same time, this tool will move the original document to the recycle bin after compression, and if you find a problem, you can restore the original file from the recycle bin.

All compression is lossless -- image quality is never reduced.

## Acknowledgments

Inspired by [ImageOptim](https://github.com/ImageOptim/ImageOptim).

Built with these great open source tools:

- [optipng](http://optipng.sourceforge.net/) — PNG optimizer
- [zopflipng](https://github.com/google/zopfli) — PNG recompressor
- [pngcrush](https://pmt.sourceforge.io/pngcrush/) — PNG optimizer
- [jpegoptim](https://github.com/tjko/jpegoptim) — JPEG optimizer
- [jpegtran](https://ijg.org/) — JPEG transformation
- [gifsicle](https://www.lcdf.org/gifsicle/) — GIF optimizer
- [rich](https://github.com/Textualize/rich) — Terminal UI
- [uv](https://docs.astral.sh/uv/) — Python package manager
- [vhs](https://github.com/charmbracelet/vhs) — Terminal recording (demo)
