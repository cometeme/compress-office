from os import mkdir
from os.path import abspath, basename, exists, expanduser, getsize, join
from shutil import move, rmtree
from subprocess import DEVNULL, run
from typing import Tuple

from rich.console import Console

from images import compress_images

console = Console()

cache_folder = expanduser("~/Library/Caches/compress-office/")


def convert_size(size_in_byte: int) -> str:
    units = ("B", "KB", "MB", "GB")
    for unit in units:
        if size_in_byte < 1024:
            return f"{round(size_in_byte, 2)}{unit}"
        size_in_byte /= 1024
    return f"{round(size_in_byte, 2)}TB"


def compress(file_path: str, workers: int | None = None) -> Tuple:
    file_path = abspath(file_path)
    file_name = basename(file_path)
    before_size = getsize(file_path)

    if exists(cache_folder):
        rmtree(cache_folder)

    mkdir(cache_folder)

    with console.status("[bold green]Unzipping Document..."):
        run(
            ["unzip", file_path, "-d", cache_folder],
            stdout=DEVNULL,
        ).check_returncode()

    new_file_path = join(cache_folder, file_name)
    img_count, img_before, img_after = compress_images(cache_folder, workers=workers)

    with console.status("[bold green]Packing Document..."):
        run(
            ["zip", file_name, "-r", "."],
            stdout=DEVNULL,
            cwd=cache_folder,
        ).check_returncode()
    after_size = getsize(new_file_path)
    if after_size >= before_size:
        print("File size unchanged")
        return (before_size, before_size)

    run(
        ["trash", file_path],
    ).check_returncode()
    move(new_file_path, file_path)

    if exists(cache_folder):
        rmtree(cache_folder)

    print(
        f"Successfully compressed {convert_size(before_size)} -> {convert_size(after_size)}"
    )

    return (before_size, after_size)
