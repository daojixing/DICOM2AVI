

import SimpleITK as sitk
import os
import cv2
import pydicom
import numpy as np
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import configparser
import sys


def get_cf(key):
    cf = configparser.ConfigParser()
    cur_path = os.path.dirname(os.path.realpath(__file__))
    config_path = os.path.join(cur_path, "dicom.conf")
    #print(config_path)
    cf.read(config_path)
    value = cf.get("dicom", key)
    return value
def set_cf(key,value):
    cf = configparser.ConfigParser()
    cur_path = os.path.dirname(os.path.realpath(__file__))
    config_path = os.path.join(cur_path, "dicom.conf")
    cf.read(config_path)
    cf.set('dicom',key,value)
    with open(config_path, "w+") as f:
        cf.write(f)
    source_dir = get_cf('source_dir')
    dest_dir = get_cf('dest_dir')
def loadFile(filename):
    ds = sitk.ReadImage(filename)
    img_array = sitk.GetArrayFromImage(ds)
    frame_num, width, height = img_array.shape
    return img_array, frame_num, width, height


def loadFileInformation(filename):
    global  information
    information = {}
    ds = pydicom.read_file(filename)
    information['PatientID'] = ds.PatientID
    information['PatientName'] = ds.PatientName
    # information['PatientBirthDate'] = ds.PatientBirthDate
    information['PatientSex'] = ds.PatientSex
    information['StudyID'] = ds.StudyID
    information['StudyDate'] = ds.StudyDate
    information['StudyTime'] = ds.StudyTime
    information['InstitutionName'] = ds.InstitutionName
    information['Manufacturer'] = ds.Manufacturer

    information['InstanceNumber'] = ds.InstanceNumber
    angle1 = ds.PositionerPrimaryAngle
    if angle1 >= 0:
        information['PositionerPrimaryAngle'] = 'LAO ' + str(abs(angle1))
    else:
        information['PositionerPrimaryAngle'] = 'RAO ' + str(abs(angle1))
    # information['PositionerPrimaryAngle']=ds.PositionerPrimaryAngle
    angle2 = ds.PositionerSecondaryAngle
    if angle2 >= 0:
        information['PositionerSecondaryAngle'] = 'CRA ' + str(abs(angle2))
    else:
        information['PositionerSecondaryAngle'] = 'CAU ' + str(abs(angle2))

    try:
        information['NumberOfFrames'] = ds.NumberOfFrames
        information['RecommendedDisplayFrameRate'] = ds.RecommendedDisplayFrameRate
    except:
        pass
    information['ImageType'] = ds.ImageType
    return information


def autoEqualize(img_array):
    img_array_list = []
    for img in img_array:
        img_array_list.append(cv2.equalizeHist(img))
    img_array_equalized = np.array(img_array_list)
    return img_array_equalized




def limitedEqualize(img_array, limit=4.0,write_img=False):
    img_array_list = []
    for img in img_array:
        img = cv2.putText(img, str(information['StudyDate']), (10, 20), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                          (255, 255, 255), 2)  # 视频添加文字
        img = cv2.putText(img, str(information['PositionerPrimaryAngle']), (10, 50), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                          (255, 255, 255), 2)
        img = cv2.putText(img, str(information['PositionerSecondaryAngle']), (10, 80), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                          (255, 255, 255), 2)
        # 视频添加文字
        #优化对比度
        #if RB_origin.getboolean(RB_origin)
        if check_var.get():
            if write_img:
                dirs = dest_dir + "/" + str(information['PatientName']).replace('^', '').replace(' ', '_') + '_' + \
                       information['StudyDate']
                cv2.imwrite(
                    dirs + '/' + str(information['InstanceNumber']) + '_' +get_origin()+ str(information['PatientName']).replace('^',
                                                                                                                    '').replace(
                        ' ', '_') + '.jpg', img)
        else:
            clahe = cv2.createCLAHE(clipLimit=limit,
                                    tileGridSize=(8, 8))  # CLAHE (Contrast Limited Adaptive Histogram Equalization)
            img = clahe.apply(img)
            if write_img:
                dirs = dest_dir + "/" + str(information['PatientName']).replace('^', '').replace(' ', '_') + '_' + \
                       information['StudyDate']
                cv2.imwrite(
                    dirs + '/' + str(information['InstanceNumber']) + '_' +get_origin()+ str(information['PatientName']).replace('^',
                                                                                                                    '').replace(
                        ' ', '_') + '.jpg', img)
        img_array_list.append(img)


    img_array_limited_equalized = np.array(img_array_list)
    cv2.destroyAllWindows()
    return img_array_limited_equalized


def writeVideo(img_array, directory):
    frame_num, width, height = img_array.shape
    #print(information['InstanceNumber'])
    #filename_output=''
    #print(information)
    if information['ImageType'].__contains__('SECONDARY'):
        filename_output = directory + '/0' + str(information['InstanceNumber']) + '_' + str(
            information['PatientName']).replace('^', '').replace(' ', '_') + '.avi'
        print(filename_output)
    else:
        filename_output = directory + '/' + str(information['InstanceNumber']) + '_'+ get_origin()+ str(
            information['PatientName']).replace('^', '').replace(' ', '_') + '.avi'
    #filename_output = directory + '/' + str(information['InstanceNumber'])+'_'+str(information['PatientName']).replace('^', '').replace(' ', '_') + '.avi'
    fourcc = cv2.VideoWriter_fourcc(*'MJPG')
    video = cv2.VideoWriter(filename_output, fourcc, 15, (width, height))
    # Above is for Mac OSX use only./////////////////////////////////////////////////////////////

    #video = cv2.VideoWriter(filename_output, -1, 15, (width, height))  # Initialize Video File

    for img in img_array:
        #print(img_array.size)
        img_rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
        video.write(img_rgb)  # Write video file frame by frame
    video.release()
## GUI Helper Function
def loadFileButton():
    global img_array, frame_num, width, height,  img_array,go_on
    try:
        img_array, frame_num, width, height = loadFile(filename)
        #print("DICOM File Loaded", "DICOM file successfully loaded!")
    except:
        print("File Loading Failed", "Sorry, file loading failed! Please check the file format.")
        go_on=0
def convertVideoButton():
    global  clipLimit,filename,go_on
    filepath=source_dir
    pathDir = os.listdir(filepath)
    clipLimit = scale.get()
    for allDir in pathDir:
        go_on = 1
        child = os.path.join('%s/%s' % (filepath, allDir))
        filename=child
        #print(filename)
        try:
            loadFileInformation(filename)
            dirs = dest_dir + "/" + str(information['PatientName']).replace('^', '').replace(' ', '_') + '_' + \
                   information['StudyDate']  #目标文件为姓名+手术日期
            if not os.path.exists(dirs):  #判断目标文件夹是否存在
                os.makedirs(dirs)
            loadFileButton()
            if go_on == 1:
                #img_array=autoEqualize(img_array)
                if information['ImageType'].__contains__('SECONDARY'): #如果是静态图片
                    limitedEqualize(img_array, clipLimit, True)
                else:
                    img_arrays = limitedEqualize(img_array, clipLimit)
                    writeVideo(img_arrays, dirs)
        except:
            pass
    messagebox.showinfo('提示','完事了')
def config():
    global source_dir,dest_dir
    cur_path = os.path.dirname(os.path.realpath(__file__))
    config_path = os.path.join(cur_path, "dicom.conf")
    if os.path.exists(config_path):
        source_dir=get_cf('source_dir')
        dest_dir=get_cf('dest_dir')
    else:
        config = configparser.ConfigParser()
        config['dicom'] = {'dest_dir': sys.path[0], 'source_dir': sys.path[0]}

        with open(config_path, 'w') as configfile:
            config.write(configfile)
        source_dir=get_cf('source_dir')
        dest_dir=get_cf('dest_dir')
def SlectSourceFloder():
    global  source_dir
    directory = filedialog.askdirectory(initialdir=source_dir)
    if directory=='':
        return
    source_dir=directory
    set_cf("source_dir",directory)
    text_sources.delete('1.0', tk.END)
    text_sources.insert('1.0',directory)
def SlectDestFloder():
    global dest_dir
    directory = filedialog.askdirectory(initialdir=dest_dir)
    if directory=='':
        return
    dest_dir=directory
    set_cf("dest_dir",directory)
    text_dest.delete('1.0', tk.END)
    text_dest.insert('1.0',directory)
def check():
    if check_var.get():
        scale['state'] = 'disabled'
    else:
        scale['state'] = 'normal'
def get_origin():
    if check_var.get():
        return 'origin_'
    else:
        return ''
## Main Stream

# Main Frame////////////////////////////////////////////////////////////////////////////////////////
root = tk.Tk()

w = 250  # width for the Tk root
h = 180  # height for the Tk root

# get screen width and height
ws = root.winfo_screenwidth()  # width of the screen
hs = root.winfo_screenheight()  # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws / 2) - (w / 2)
y = (hs / 2) - (h / 2)

# set the dimensions of the screen
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
#root.geometry()
# root.attributes('-fullscreen', True)
root.title('DICOM TO AVI')
#img = PhotoImage(file='Heart.ico')
#root.iconphoto(True, PhotoImage(file='Heart.ico'))
#root.iconbitmap('Heart.ico')

isLoad = 0
clipLimit = 3.0
filename = ''

#button_browse = ttk.Button(root, text='浏览', width=20, command=browseFileButton)
button_browse = ttk.Button(root, text='...',width=3, command=SlectSourceFloder)
button_browse.place(x=200, y=8)

button_dest = ttk.Button(root, text='...', width=3, command=SlectDestFloder)
button_dest.place(x=200, y=48)

button_convert = ttk.Button(root, text='转换', width=20, command=convertVideoButton)
button_convert.place(x=40, y=130)


#button_close = ttk.Button(root, width=20, text='退出', command=root.destroy)
#button_close.place(x=40, y=170)

label_s_t=tk.Label(root, text='优化程度', font=('tahoma', 10))
label_s_t.place(x=10,y=90)
scale = tk.Scale(root, from_=0, to=10, orient=tk.HORIZONTAL,label='')
scale.set(3)                                                                     
scale.place(x=80,y=75)

text_sources=tk.Text(root, width=18,height=1, font=('tahoma', 9), bd=2)
text_sources.place(x=70, y=10)
label_sources=tk.Label(root, text='来源目录', font=('tahoma', 10))
label_sources.place(x=10,y=10)
text_dest=tk.Text(root, width=18,height=1, font=('tahoma', 9), bd=2)
text_dest.place(x=70, y=50)
label_dest=tk.Label(root, text='目标目录', font=('tahoma', 10))
label_dest.place(x=10,y=50)

check_var=tk.IntVar()
RB_origin=tk.Checkbutton(root,text='原图',variabl=check_var,command=check)
RB_origin.place(x=190,y=95)
RB_origin.deselect()
cv2.destroyAllWindows()

config()
text_dest.delete('1.0', tk.END)
text_dest.insert('1.0', dest_dir)
text_sources.delete('1.0', tk.END)
text_sources.insert('1.0', source_dir)
root.mainloop()

### !!! Make sure to downgrade setuptools to 19.2. If this does get the frozen binary with PyInstaller !!!!
# Just hit this myself. Can confirm that downgrading to setuptools 19.2 fixes the issue for me.

### To install the SimpleITK package with conda run:
'''
```powershell
conda install --channel https://conda.anaconda.org/SimpleITK SimpleITK
```
'''