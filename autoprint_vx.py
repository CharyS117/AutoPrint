import os, shutil, win32api, win32print, time
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import *

fileFold = '~/'
printedFold = '~/printed'


def print_doc(filename):
    open(filename, "r")
    win32api.ShellExecute(0, "print", filename,
                          '/d:"%s"' % win32print.GetDefaultPrinter(), ".", 0)


def print_pdf(filename):
    os.system(r'start acrobat /P /h "%s"' % filename)


def findFile(fold):
    files = os.listdir(fold)
    fileList = []
    for fileName in files:
        file = [fileName, os.path.join(fold, fileName)]
        if os.path.isdir(file[1]):
            fileList = fileList + findFile(file[1])
        else:
            fileList.append(file)
    return fileList


def cutSuffix(filename):
    doc = ['.docx', '.doc']
    if '.pdf' in filename:
        i = filename[::-1].find('fdp.')
        if i == 0:
            return 1, filename
        else:
            return 1, filename[:-i]
    for suffix in doc:
        if suffix in filename:
            j = filename[::-1].find(suffix[::-1])
            if j == 0:
                return 2, filename
            else:
                return 2, filename[:-j]
    return 0


class handler(FileSystemEventHandler):

    def on_created(self, event):
        cuted = cutSuffix(event.src_path)
        if cuted:
            filetype = cuted[0]
            newpath = cuted[1]
            while True:
                if os.path.exists(newpath):
                    break
                time.sleep(1)
            filename = os.path.split(newpath)[-1]
            Time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ctime = time.strftime("%Y%m%d-%H%M%S",
                                  time.localtime(os.stat(newpath).st_ctime))
            ctimeFold = os.path.join(printedFold, ctime)
            if not os.path.exists(ctimeFold):
                os.mkdir(ctimeFold)
            filePath = os.path.join(ctimeFold, filename)
            shutil.copy(newpath, filePath)
            print('%s is printed %s' % (filename, Time))
            # pushNotification(filename)
            if filetype == 1:
                print_pdf(filePath)
            if filetype == 2:
                print_doc(filePath)


def autoprint_vx():
    observer = Observer()
    observer.schedule(handler(), fileFold, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()


def main():
    try:
        print("Print from wechat start")
        autoprint_vx()
    except Exception:
        print(Exception.__context__)
        print('Restart in 60 seconds')
        time.sleep(60)
        main()
    finally:
        print('Sucssess')


if __name__ == '__main__':
    main()