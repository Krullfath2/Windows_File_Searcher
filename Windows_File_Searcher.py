import platform
import shutil
import pathlib
from datetime import datetime
import time
import subprocess
import ctypes


class FileSearcher:
    def __init__(self):
        # self.minimize_console()
        self.load_settings()

    '''
    def minimize_console(self):
        kernel32 = ctypes.WinDLL("kernel32")
        user32 = ctypes.WinDLL("user32")
        hWnd = kernel32.GetConsoleWindow()
        if hWnd:
            user32.ShowWindowAsync(hWnd, 6)
            user32.UpdateWindow(hWnd)
    '''
    def load_settings(self):
        self.settings_file_path = pathlib.Path.cwd() / "File_searcher_settings.txt"
        default_settings = """\
search_keywords:"example_keyword_one","example_keyword_2"
allowed_file_extensions:".txt",".png"
max_file_size:"1GB"

# Explanation on how to modify these settings without causing issues:
# To modify keywords, add or remove them within the double quotes separated by commas.
# To modify allowed file extensions, add or remove them within the double quotes separated by commas.
# To modify the maximum file size, change the value (e.g., "500MB", "100KB") or type "disabled" to disable the size check.
"""

        try:
            with open(self.settings_file_path, "r") as settings_file:
                settings = settings_file.readlines()
                if not all(setting.startswith(expected) for setting, expected in zip(settings, ['search_keywords:', 'allowed_file_extensions:', 'max_file_size:'])):
                    raise SyntaxError("Settings file contains syntax issues")
        except FileNotFoundError:
            with open(self.settings_file_path, "w") as new_settings_file:
                new_settings_file.write(default_settings)
        except SyntaxError as e:
            print(f"Error in settings file syntax: {e}. Replacing with default settings...")
            with open(self.settings_file_path, "w") as new_settings_file:
                new_settings_file.write(default_settings)
            exit()

        with open(self.settings_file_path, "r") as settings_file:
            settings = settings_file.readlines()
            print("Settings read successfully")

        self.search_keywords = settings[0].split(':')[1].strip().strip('"').split('","')
        self.allowed_file_extensions = settings[1].split(':')[1].strip()
        self.max_file_size = settings[2].split(':')[1].strip().strip('"')

    def get_recent_folders(self, common_folders):
        print("Getting recent folders...")
        recent_folders = []
        recent_path = pathlib.Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows" / "Recent"
        if recent_path.is_dir():
            recent_folders = [file for file in recent_path.iterdir() if file.is_file() and not any(str(common_folder) in str(file.parents[0]) for common_folder in common_folders)]
        return recent_folders

    def get_pinned_folders(self, common_folders):
        print("Getting pinned folders...")
        pinned_folders = []
        pinned_path = pathlib.Path.home() / "AppData" / "Roaming" / "Microsoft" / "Internet Explorer" / "Quick Launch" / "User Pinned" / "TaskBar"
        if pinned_path.is_dir():
            pinned_folders = [file for file in pinned_path.iterdir() if file.is_file() and not any(str(common_folder) in str(file.parents[0]) for common_folder in common_folders)]
        return pinned_folders

    def process_file(self, file, keywords, found_files_keyword, log_file_path):
        duplicated_count = 0

        file_name_lower = file.name.lower()
        file_extension = file.suffix.lower()

        if self.max_file_size != "disabled":
            try:
                max_size_bytes = self.convert_to_bytes(self.max_file_size)
                if file.stat().st_size > max_size_bytes:
                    return 0
            except ValueError as e:
                print(f"Invalid max file size: {e}")
                return 0

        if not self.is_valid_extension(file_extension, self.allowed_file_extensions):
            return 0

        for keyword in keywords:
            if keyword.lower() in file_name_lower:
                try:
                    shutil.copy2(file, found_files_keyword)
                    print(f"[{keyword.lower()}] Successfully duplicated with keyword: {file.name}")
                    with open(log_file_path, "a") as log_file:
                        log_file.write(f"[{keyword.lower()}] Successfully duplicated with keyword: {file.name}\n")
                    duplicated_count += 1
                except Exception as e:
                    print(f"Error duplicating {file.name}: {e}")
                    with open(log_file_path, "a") as log_file:
                        log_file.write(f"Error duplicating {file.name}: {e}\n")
                break

        return duplicated_count

    def convert_to_bytes(self, size_str):
        suffixes = {'KB': 1024, 'MB': 1024 ** 2, 'GB': 1024 ** 3}
        if size_str.endswith(tuple(suffixes.keys())):
            return int(size_str[:-2]) * suffixes[size_str[-2:]]
        else:
            raise ValueError("Invalid size format")

    def is_valid_extension(self, file_extension, allowed_extensions):
        allowed_extensions = [ext.strip().strip('"') for ext in allowed_extensions.split(",")]
        if allowed_extensions == ['']:
            return True
        return file_extension in allowed_extensions

    def run_search(self):
        if platform.system() != "Windows":
            print("This program is only supported on Windows.")
            exit()

        start_time = time.time()

        current_dir = pathlib.Path.cwd()
        results_folder_name = f"Results_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"
        results_folder = current_dir / results_folder_name
        found_files_keyword = results_folder / "Found_Files_Keyword"
        duplicate_logs_folder = results_folder / "Duplicate_Logs"

        try:
            results_folder.mkdir(parents=True, exist_ok=True)
            found_files_keyword.mkdir(parents=True, exist_ok=True)
            duplicate_logs_folder.mkdir(parents=True, exist_ok=True)
            subprocess.run(['attrib', '+h', str(results_folder)])
        except FileExistsError:
            pass

        common_folders = [pathlib.Path.home() / folder for folder in ["Desktop", "Documents", "Downloads", "Pictures", "Videos"]]
        search_folders = [
            *common_folders,
            *self.get_recent_folders(common_folders),
            *self.get_pinned_folders(common_folders)
        ]

        log_file_path = duplicate_logs_folder / "duplicate_log.txt"
        with open(log_file_path, "w") as log_file:
            duplicated_counts = 0
            for folder in search_folders:
                for file in folder.glob("**/*"):
                    if file.is_file():
                        duplicated_counts += self.process_file(file, self.search_keywords, found_files_keyword, log_file_path)

        total_duplicated_files = duplicated_counts

        print(f"Duplicated {total_duplicated_files} files successfully.")
        end_time = time.time()
        duration = end_time - start_time
        print(f"Created Results in {duration} seconds")


if __name__ == "__main__":
    searcher = FileSearcher()
    searcher.run_search()
