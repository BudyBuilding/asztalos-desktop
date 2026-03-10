import os
from datetime import datetime


def get_log_path():
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "log.txt")


def log(message):

    log_file = get_log_path()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(log_file, "a", encoding="utf8") as f:
        f.write(f"{timestamp}  {message}\n")

    print(message)