import os


def get_desktop_path():
    """获取Windows桌面路径"""
    return os.path.join(os.path.expanduser("~"), 'Desktop')


if __name__ == "__main__":
    print(get_desktop_path())
