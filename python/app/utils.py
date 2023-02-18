import os


def join_from_cwd(path: str):
    return os.path.join(os.getcwd(), path)


def join(x, y):
    return os.path.join(x, y)
