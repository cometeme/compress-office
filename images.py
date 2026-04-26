import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from os.path import getsize, join, splitext
from subprocess import DEVNULL, run

from rich.console import Console as RichConsole
from rich.progress import BarColumn, Progress, TaskProgressColumn, TextColumn
from rich.theme import Theme

_plain_console = RichConsole(
    theme=Theme({"bar.complete": "green", "bar.finished": "green", "bar.pulse": ""})
)


def _check_tool(name: str) -> bool:
    return run(["which", name], stdout=DEVNULL, stderr=DEVNULL).returncode == 0


def _run_step(args: list[str], file_path: str) -> bool:
    try:
        run(
            [file_path if a == "{file}" else a for a in args],
            stdout=DEVNULL,
            stderr=DEVNULL,
            check=True,
        )
        return True
    except Exception:
        return False


def _build_chain() -> dict[str, list[list[str]]]:
    chain: dict[str, list[list[str]]] = {}

    png: list[list[str]] = []
    if _check_tool("optipng"):
        png.append(["optipng", "-o2", "-quiet", "-preserve", "{file}"])
        png.append(["optipng", "-o7", "-quiet", "-preserve", "{file}"])
    if _check_tool("zopflipng"):
        png.append(["zopflipng", "-y", "{file}", "{file}"])
    if _check_tool("pngcrush"):
        png.append(["pngcrush", "-q", "-ow", "-rem", "alla", "{file}"])
    if png:
        chain[".png"] = png

    jpg: list[list[str]] = []
    if _check_tool("jpegoptim"):
        jpg.append(["jpegoptim", "--strip-all", "-q", "{file}"])
    if _check_tool("jpegtran"):
        jpg.append(
            ["jpegtran", "-optimize", "-copy", "none", "-outfile", "{file}", "{file}"]
        )
    if jpg:
        chain[".jpg"] = jpg
        chain[".jpeg"] = jpg

    gif: list[list[str]] = []
    if _check_tool("gifsicle"):
        gif.append(["gifsicle", "--batch", "-O3", "{file}"])
    if gif:
        chain[".gif"] = gif

    return chain


_COMPRESSION_CHAINS = _build_chain()
_IMAGE_EXTS = set(_COMPRESSION_CHAINS.keys())
_WORKERS = min(4, os.cpu_count() or 4)


def _format_size(size_in_byte: int) -> str:
    units = ("B", "KB", "MB", "GB")
    for unit in units:
        if size_in_byte < 1024:
            return f"{round(size_in_byte, 2)}{unit}"
        size_in_byte /= 1024
    return f"{round(size_in_byte, 2)}TB"


def compress_images(folder: str, workers: int | None = None) -> tuple[int, int, int]:
    images: list[str] = []
    for root, _dirs, files in os.walk(folder):
        for f in files:
            ext = splitext(f)[1].lower()
            if ext in _IMAGE_EXTS:
                images.append(join(root, f))

    if not images:
        return (0, 0, 0)

    if workers is None:
        workers = _WORKERS

    with Progress(
        TextColumn("{task.description}"),
        BarColumn(),
        TaskProgressColumn(text_format="{task.percentage:>3.0f}%"),
        TextColumn("{task.fields[result]}"),
        console=_plain_console,
    ) as progress:
        tasks: dict[str, int] = {}
        for fp in images:
            name = os.path.basename(fp)
            ext = splitext(fp)[1].lower()
            steps = len(_COMPRESSION_CHAINS.get(ext, []))
            tasks[fp] = progress.add_task(
                f"  {name}",
                total=max(steps, 1),
                start=True,
                result="",
            )

        results: list[tuple[int, int]] = []

        def _compress_one(fp: str) -> None:
            task_id = tasks[fp]
            name = os.path.basename(fp)
            ext = splitext(fp)[1].lower()
            before = getsize(fp)

            steps = _COMPRESSION_CHAINS.get(ext, [])
            for step in steps:
                progress.update(task_id, description=f"  {name}")
                _run_step(step, fp)
                progress.advance(task_id)

            after = getsize(fp)

            if after < before:
                pct = round(100 * (before - after) / before, 1)
                result = f"{_format_size(before)} -> {_format_size(after)} (-{pct}%)"
            else:
                result = f"{_format_size(before)} (already optimized)"

            progress.update(
                task_id,
                completed=len(steps),
                description=f"  {name}",
                result=result,
            )
            results.append((before, after))

        if workers == 1:
            for fp in images:
                _compress_one(fp)
        else:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {executor.submit(_compress_one, fp): fp for fp in images}
                for future in as_completed(futures):
                    future.result()

        total_before = sum(r[0] for r in results)
        total_after = sum(r[1] for r in results)

    return (len(images), total_before, total_after)
