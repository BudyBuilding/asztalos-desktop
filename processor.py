import zipfile
import shutil
import os
import time

from logger import log


def wait_for_download(path, retries=10):

    last_size = -1

    for _ in range(retries):

        if not os.path.exists(path):
            return False

        size = os.path.getsize(path)

        if size == last_size:
            return True

        last_size = size
        time.sleep(1)

    return False


def process_zip(path, config):

    log(f"Processing ZIP: {path}")

    if not wait_for_download(path):
        log("File still downloading")
        return

    temp_folder = path + "_tmp"

    if os.path.exists(temp_folder):
        shutil.rmtree(temp_folder)

    os.makedirs(temp_folder)

    with zipfile.ZipFile(path, 'r') as zip_ref:
        zip_ref.extractall(temp_folder)

    folders = [
        f for f in os.listdir(temp_folder)
        if os.path.isdir(os.path.join(temp_folder, f))
    ]

    if len(folders) == 0:
        log("ZIP is empty")
        return

    # feltételezzük hogy van egy root mappa
    root_folder = os.path.join(temp_folder, folders[0])

    subfolders = [
        f for f in os.listdir(root_folder)
        if os.path.isdir(os.path.join(root_folder, f))
    ]

    folder_list = None
    folder_program = None

    for f in subfolders:

        name = f.lower()

        if "lista" in name:
            folder_list = os.path.join(root_folder, f)

        if "program" in name:
            folder_program = os.path.join(root_folder, f)

    if not folder_list or not folder_program:
        log("Missing required folders: program / lista")
        return

    target_list = config["target_list"]
    target_program = config["target_program"]

    os.makedirs(target_list, exist_ok=True)
    os.makedirs(target_program, exist_ok=True)

    for file in os.listdir(folder_list):

        src = os.path.join(folder_list, file)
        dst = os.path.join(target_list, file)

        shutil.move(src, dst)

    for file in os.listdir(folder_program):

        src = os.path.join(folder_program, file)
        dst = os.path.join(target_program, file)

        shutil.move(src, dst)

    archive = config["archive"]
    os.makedirs(archive, exist_ok=True)

    archive_path = os.path.join(archive, os.path.basename(path))

    shutil.copy2(path, archive_path)

    log(f"ZIP archived: {archive_path}")

    os.remove(path)

    log(f"Original ZIP deleted: {path}")
    
    shutil.rmtree(temp_folder)

    log(f"ZIP processed successfully: {path}")