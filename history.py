import csv
from os.path import abspath, exists
import os

class history:
    def __init__(self, path: str) -> None:
        self.path = path
        self.history = dict()  # <file_path: modify_time>
        if exists(self.path):
            with open(self.path, "r") as f:
                reader = csv.reader(f)
                for row in reader:
                    self.history[row[0]] = int(row[1])

    def __del__(self) -> None:
        with open(self.path, "w") as f:
            writer = csv.writer(f)
            for file in self.history.items():
                writer.writerow(file)

    def add_history(self, file_path: str) -> None:
        file_path = abspath(file_path)
        mtime = os.stat(file_path).st_mtime_ns

        # Update history
        self.history[file_path] = mtime

    def file_in_history(self, file_path: str) -> bool:
        file_path = abspath(file_path)
        # Search from history
        if file_path not in self.history:
            return False

        # Get file's modify_time
        mtime = os.stat(file_path).st_mtime_ns

        return mtime == self.history[file_path]

    def clean_up(self) -> None:
        clean_up_list = []
        for file_path in self.history:
            if not exists(file_path):
                clean_up_list.append(file_path)

        for file_path in clean_up_list:
            self.history.pop(file_path)
            # print(f"Clean up {file_path}")
