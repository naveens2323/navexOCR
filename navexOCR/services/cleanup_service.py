import os
import shutil


def cleanup_folder(folder):

    if os.path.exists(folder):
        shutil.rmtree(folder, ignore_errors=True)