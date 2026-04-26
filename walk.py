from glob import glob
from os.path import abspath, exists, isfile, join
from shlex import quote
from subprocess import DEVNULL, PIPE, run

from rich.console import Console

console = Console(highlighter=None)

use_fd = run("fd --version", shell=True, stdout=DEVNULL, stderr=DEVNULL).returncode == 0


def find_files(paths: list[str], exts: tuple[str, ...]) -> list[str]:
    result = []

    for path in paths:
        if not exists(path):
            console.print(f"Error: {path} not found.", style="bold red")
            exit(-1)
        path = abspath(path)

        if isfile(path):
            if path.split(".")[-1] not in exts:
                console.print(
                    f"Error: {path} is not a supported file.", style="bold red"
                )
                exit(-1)
            result.append(path)
        else:
            if use_fd:
                command = f"fd -ip . {quote(path)}"
                for ext in exts:
                    command += f" -e {ext}"
                res = run(command, shell=True, stdout=PIPE)
                res.check_returncode()
                for file_path in res.stdout.decode().splitlines():
                    result.append(file_path)
            else:
                for ext in exts:
                    for file_path in glob(join(path, "**", f"*.{ext}"), recursive=True):
                        result.append(file_path)

    return result
