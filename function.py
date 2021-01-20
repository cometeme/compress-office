from os import mkdir
from os.path import abspath, basename, exists, expanduser, getsize, join
from shlex import quote
from shutil import move, rmtree
from subprocess import DEVNULL, run
from typing import Tuple

from rich.console import Console

console = Console()

# Modify these to use on other platform
cache_folder = expanduser("~/Library/Caches/compress-office/")
compress_command = (
    "/Applications/ImageOptim.app/Contents/MacOS/ImageOptim ~/Library/Caches/compress-office/*"
)


def convert_size(size_in_byte: int) -> str:
    units = ("B", "KB", "MB", "GB")
    for unit in units:
        if size_in_byte < 1024:
            return f"{round(size_in_byte, 2)}{unit}"
        size_in_byte /= 1024
    return f"{round(size_in_byte, 2)}TB"


def compress(file_path: str) -> Tuple:
    file_path = abspath(file_path)
    file_name = basename(file_path)
    before_size = getsize(file_path)

    # Delete old files
    if exists(cache_folder):
        rmtree(cache_folder)

    # Create cache folder
    mkdir(cache_folder)

    # Unzip file
    with console.status("[bold green]Unzipping Document..."):
        run(
            f"unzip {quote(file_path)} -d {quote(cache_folder)}",
            shell=True,
            stdout=DEVNULL,
        ).check_returncode()

    # Compress
    new_file_path = join(cache_folder, file_name)
    with console.status("[bold green]Compressing Images..."):
        run(
            compress_command,
            shell=True,
            stdout=DEVNULL,
            stderr=DEVNULL,
        ).check_returncode()

    # Rezip file
    with console.status("[bold green]Packing Document..."):
        run(
            f"cd {quote(cache_folder)} && zip {quote(file_name)} -r *",
            shell=True,
            stdout=DEVNULL,
        ).check_returncode()
    after_size = getsize(new_file_path)

    if after_size >= before_size:
        print("File size unchanged")
        return (before_size, before_size)

    # Move new file to origin place
    run(
        f"trash {quote(file_path)}",
        shell=True,
    ).check_returncode()
    move(new_file_path, file_path)

    # Delete caches
    if exists(cache_folder):
        rmtree(cache_folder)

    print(f"Successfully compressed {convert_size(before_size)} -> {convert_size(after_size)}")

    return (before_size, after_size)
