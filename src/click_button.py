import os
import logging
from datetime import datetime
from pywinauto.application import Application
from pywinauto.findbestmatch import MatchError

from tools import get_desktop_path


logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename='Debug.log',
                    filemode='w')

# ApogeeConnect.exe文件路径，请在此修改
EXE_PATH = "E:\Program Files (x86)\Apogee Connect\dist\ApogeeConnect.exe"
# 日志文件存放路径，为None时存放在 桌面\ApogeeConnectLogo\ 文件夹下
# LOG_FILE_PATH = "E:\\test_test\\test"
LOG_FILE_PATH = None


def click_data_logging():
    # Apogee Connect程序未打开用start
    # app = Application(backend="uia").Start(cmd_line=EXE_PATH)
    # Apogee Connect程序已打开用connect
    app = Application(backend='win32' ).connect(path=EXE_PATH)
    app['Apogee Connect']['0.0Button'].click()
    app['Apogee Connect']['Setup'].click()
    file_name = '{}.csv'.format(datetime.now().strftime("%Y_%m_%d_%H_%M_%S"))

    if LOG_FILE_PATH is None:
        log_file_path = os.path.join(get_desktop_path(), 'ApogeeConnectLogo')
    else:
        log_file_path = LOG_FILE_PATH

    file_path = os.path.join(log_file_path, file_name)
    logging.info("file_path:{}, log_file_path:{}".format(file_path, log_file_path))

    # 如果日志问价夹不存在则创建
    if not os.path.exists(log_file_path):
        os.makedirs(log_file_path)

    # 新建日志文件
    if not os.path.exists(file_path):
        with open(file_path, 'w'):
            pass

    window = app.Dialog
    # 设置编辑框文字
    combobox_edit = window["Edit"]
    combobox_edit.set_edit_text(file_path)

    button_save = window.Save
    button_save.click()
    # Start
    app['Apogee Connect']['Button5'].click()
    # 重写日志文件弹框，点击确认
    window["确定"].click()

    # 输出 Button5 按键的文字信息   Start/Stop
    logging.info(app['Apogee Connect']['Button5'].texts())


if __name__ == "__main__":
    logging.info(os.getcwd())
    os.chdir(os.path.dirname(EXE_PATH))
    logging.info(os.getcwd())

    # 循环 失败则重新开始，直至成功
    while True:
        try:
            click_data_logging()
            break
        except MatchError as e:
            # print("MatchError!!!")
            logging.error(e)
            continue
