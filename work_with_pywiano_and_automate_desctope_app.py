import pyautogui as pg
import time
from pywinauto import application
from mutagen.mp3 import MP3
import os
import shutil
from pathlib import Path


input_folder_path = Path("C:\\IzotopProcess\\Input")
# input_folder_path = Path(r"C:\test")
output_folder_path = Path("C:\\IzotopProcess\\Output")
# output_folder_path = Path(r"C:\test2")
preset_time_list = {'module1': 27, 'module2': 52, 'module3': 28, 'module4': 36, 'module5': 27, 'module6': 30,
                    'module7': 52,
                    'module8': 30, 'module9': 7}


def exploring_files(folder_files):
    num_tries = 0
    while True:
        if any(File.endswith(".mp3") for File in os.listdir(folder_files)):
            num_tries = 0

            file_name = os.listdir(folder_files)[0]
            parse_file_name(file_name)

            file_path = os.path.join(input_folder_path, file_name)
            delete_file(file_path)

    num_tries += 1
    print("Exploring for mp3 files in '%s' folder, num tries: %s" % (folder_files, num_tries))
    time.sleep(5)


def parse_file_name(file_name):
    print(file_name)

    # input file path
    file_path = os.path.join(input_folder_path, file_name)

    # get module name and number
    module_number = file_name[file_name.rfind("_") + 1:-4]
    module_name = file_name[:file_name.rfind(module_number) - 1]
    module_name = module_name[:module_name.rfind("_")]
    print("module_name: ", module_name)
    print("module_number: ", module_number)
    print("file_path: ", file_path)
    print("output_dir: ", output_folder_path)

    flag = True
    if module_number != "all":
        flag = False
        if module_number.isdigit() and int(module_number) in range(1, 10):
            flag = True

    if flag:
        if module_number.lower() == "all":
            module_number = 10

        izotop_work(int(module_number), file_path, file_name)
    else:
        print("Error")
        with open("C:\\IzotopProcess\\Output\\error.txt", 'tw', encoding='utf-8') as f:
            f.write("Plase input file with example: (file_1.mp3 or file_all.mp3). Your file name: _%s_"%file_name)
        shutil.move(file_path,input_folder_path)




def izotop_work(preset_number, file_path, file_name):
    if preset_number in range(1, 11):
        print("Start logic with module %s" % preset_number)

        if preset_number == 10:
            for pr in range(1, 10):
                logic(file_path, int(pr), file_name)
        else:
            logic(file_path, preset_number, file_name)
    else:
        print("Unsuported preset")
        delete_file(file_path)


def delete_file(file_path):
    print("Delete input file: %s" % file_path)
    if os.path.isfile(file_path):
        os.remove(file_path)
    with open("C:\\IzotopProcess\\Output\\done.txt", 'tw', encoding='utf-8') as f:
        pass


def select_preset(preset_number):
    if preset_number == 1:
        pg.click(1006, 332)
    elif preset_number == 2:
        pg.click(1024, 357)
    elif preset_number == 3:
        pg.click(1017, 384)
    elif preset_number == 4:
        pg.click(1016, 409)
    elif preset_number == 5:
        pg.click(1009, 438)
    elif preset_number == 6:
        pg.click(1015, 464)
    elif preset_number == 7:
        pg.click(1009, 486)
    elif preset_number == 8:
        pg.click(1016, 513)
    elif preset_number == 9:
        pg.click(1002, 527)


def calculate_time(preset_number, file_path):
    audio = MP3(file_path)
    print("Input file: %s" % file_path)

    audio_length = audio.info.length
    print("Audio file length: %d" % audio_length)

    # calculate origin render time according module_list + 20 %
    origin_render = (preset_time_list["module%d" % preset_number] / 60) * audio_length
    render_time = origin_render + (origin_render * 0.2)
    print("Aproximately render time: %f seconds" % render_time)

    return render_time


def logic(input_file_path, preset_number, file_name):
    # start application
    app = application.Application()
    app.start("C:\\Program Files (x86)\iZotope\\RX 7 Audio Editor\\win64\\iZotope RX 7 Audio Editor.exe")
    time.sleep(2)
    pg.click(1129, 607)  # click if izotop crash
    time.sleep(1)
    app["iZotope RX 7 Advanced Audio Editor"].wait('ready')

    app['iZotope RX 7 Advanced Audio Editor'].menu_select("File->Open")
    time.sleep(1)

    # app['Select audio files'].print_control_identifiers()

    app['Select audio files']["'&Имя файла:Edit'"].set_text(input_file_path)
    time.sleep(1)
    app['Select audio file']['&ОткрытьButton'].click()
    time.sleep(1)

    app[u'iZ_RX7_v7010_Win32Window.140B9A8B0'].menu_item(u'&Window->&Module Chain\tC').select()
    # app['Module Chain'].print_control_identifiers()
    pg.click(1183, 267)  # select menu
    time.sleep(1)

    select_preset(preset_number)  # select module
    time.sleep(1)

    pg.click(1265, 618)  # click on Button "render"
    time.sleep(calculate_time(preset_number, input_file_path))

    pg.click(1318, 227)  # close window "module chain"
    print("close window module chain")
    time.sleep(2)
    app[u'iZ_RX7_v7010_Win32Window.140B9A8B0'].wait('ready')
    app[u'iZ_RX7_v7010_Win32Window.140B9A8B0'].menu_select(u'&File->Save As...\tCtrl+Shift+S')

    print("save file ")
    print("ready to click button OK in save as windows")
    pg.click(1066, 730)  # click button OK in save as windows
    # pg.click(1151, 684)  # click button OK in save as windows

    print("click button OK in save as windows")
    time.sleep(2)

    print("ready to rename  save audio file")
    time.sleep(2)

    saved_file = os.path.join(output_folder_path, file_name.partition('.')[0] + '_' + str(preset_number))
    # print(input_file_path.partition('.')[0] + '_' + str(preset_number))
    print("Saved_files %s" % str(saved_file))

    app["Save As"]["'&Имя файла:Edit'"].set_text(saved_file)
    app["Save As"]['Со&хранитьButton'].click()

    # app['iZotope RX 7 Advanced Audio Editor'].menu_select("File->Open")
    time.sleep(2)
    # pg.click(753,204)#close window with render file

    app[u'iZ_RX7_v7010_Win32Window.D24A10'].menu_item("u'&File->Close All Files\tCtrl+Shift+W'").select()
    # app[u'iZotope RX 7 Advanced Audio Editor'].menu_item(u'&File->E&xit').click()
    time.sleep(1)

    # print(input_file_path + saved_file + '.mp3')
    # shutil.move(output_folder_path + saved_file + '.mp3',
    #       r"C:\IzotopProcess\Output")
    time.sleep(3)
    print("ready to close program")

    app.kill()

    time.sleep(1)
    print("closed Izotop")
    # app['Module Chain']'''


if __name__ == "__main__":
    exploring_files(input_folder_path)
    # logic(sys.argv[1], sys.argv[2])

    # search_file()

# print(logic())
