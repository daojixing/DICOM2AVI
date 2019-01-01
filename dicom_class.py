
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

class MyApp(object):
    def __init__(self, parent):
        self.rootframe = tk.Frame(parent)
        self.rootframe.pack()
        self.setupUI()
        self.config()
    def setupUI(self):
        self.button_browse = ttk.Button(root, text='...', width=3, command=self.SlectSourceFloder)
        self.button_browse.place(x=200, y=8)

        self.button_dest = ttk.Button(root, text='...', width=3, command=self.SlectDestFloder)
        self.button_dest.place(x=200, y=48)

        self.button_convert = ttk.Button(root, text='转换', width=20, command=self.convertVideoButton)
        self.button_convert.place(x=40, y=130)

        self.scale = tk.Scale(root, from_=0, to=10, orient=tk.HORIZONTAL, label='')
        self.scale.set(3)
        self.scale.place(x=80, y=75)

        self.text_sources = tk.Text(root, width=18, height=1, font=('tahoma', 9), bd=2)
        self.text_sources.place(x=70, y=10)
        self.label_sources = tk.Label(root, text='来源目录', font=('tahoma', 10))
        self.label_sources.place(x=10, y=10)
        self.text_dest = tk.Text(root, width=18, height=1, font=('tahoma', 9), bd=2)
        self.text_dest.place(x=70, y=50)
        self.label_dest = tk.Label(root, text='目标目录', font=('tahoma', 10))
        self.label_dest.place(x=10, y=50)

        self.check_CLAHE = tk.IntVar()
        self.RB_origin = tk.Checkbutton(root, text='CLAHE', variabl=self.check_CLAHE, command='')
        self.RB_origin.place(x=10, y=90)
        self.RB_origin.select()
      # 均衡
        self.check_equalizeHist = tk.IntVar()
        self.equalizeHist_check = tk.Checkbutton(root, text='均衡', variabl=self.check_equalizeHist, command='')
        self.equalizeHist_check.place(x=180, y=90)
        self.equalizeHist_check.select()

    def get_cf(self,key):
        cf = configparser.ConfigParser()
        cur_path = os.path.dirname(os.path.realpath(__file__))
        config_path = os.path.join(cur_path, "dicom.conf")
        # print(config_path)
        cf.read(config_path)
        value = cf.get("dicom", key)
        return value

    def set_cf(self,key, value):
        cf = configparser.ConfigParser()
        cur_path = os.path.dirname(os.path.realpath(__file__))
        config_path = os.path.join(cur_path, "dicom.conf")
        cf.read(config_path)
        cf.set('dicom', key, value)
        with open(config_path, "w+") as f:
            cf.write(f)

    def loadFile(self,filename):
        ds = sitk.ReadImage(filename)
        img_array = sitk.GetArrayFromImage(ds)
        frame_num, width, height = img_array.shape
        return img_array, frame_num, width, height

    def loadFileInformation(self,filename):
        self.information = {}
        ds = pydicom.read_file(filename)
        self.information['PatientID'] = ds.PatientID
        self.information['PatientName'] = ds.PatientName
        # self.information['PatientBirthDate'] = ds.PatientBirthDate
        self.information['PatientSex'] = ds.PatientSex
        self.information['StudyID'] = ds.StudyID
        self.information['StudyDate'] = ds.StudyDate
        self.information['StudyTime'] = ds.StudyTime
        self.information['InstitutionName'] = ds.InstitutionName
        self.information['Manufacturer'] = ds.Manufacturer

        self.information['InstanceNumber'] = ds.InstanceNumber
        angle1 = ds.PositionerPrimaryAngle
        if angle1 >= 0:
            self.information['PositionerPrimaryAngle'] = 'LAO ' + str(abs(angle1))
        else:
            self.information['PositionerPrimaryAngle'] = 'RAO ' + str(abs(angle1))
        # self.information['PositionerPrimaryAngle']=ds.PositionerPrimaryAngle
        angle2 = ds.PositionerSecondaryAngle
        if angle2 >= 0:
            self.information['PositionerSecondaryAngle'] = 'CRA ' + str(abs(angle2))
        else:
            self.information['PositionerSecondaryAngle'] = 'CAU ' + str(abs(angle2))

        try:
            self.information['NumberOfFrames'] = ds.NumberOfFrames
            self.information['RecommendedDisplayFrameRate'] = ds.RecommendedDisplayFrameRate
        except:
            pass
        self.information['ImageType'] = ds.ImageType
        #return self.self.information

    def autoEqualize(self,img_array_a):
        img_array_list = []
        for img in img_array_a:
            img_array_list.append(cv2.equalizeHist(img))
        img_array_equalized = np.array(img_array_list)
        return img_array_equalized

    def writeVideo(self,img_array, directory):
        frame_num, width, height = img_array.shape
        if self.information['ImageType'].__contains__('SECONDARY'):  # 如果是静态图片
            for img in img_array:
                cv2.imwrite(self.get_jpgname(), img)
            return
        filename_output = self.get_filename(directory)
        fourcc = cv2.VideoWriter_fourcc(*'MJPG')
        video = cv2.VideoWriter(filename_output, fourcc, 15, (width, height))
        for img in img_array:
            img_rgb = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            video.write(img_rgb)  # Write video file frame by frame
        video.release()



    def convertVideoButton(self):
        filepath = self.source_dir
        pathDir = os.listdir(filepath)
        for allDir in pathDir:

            filename = os.path.join('%s/%s' % (filepath, allDir))
            # try:
            self.loadFileInformation(filename)
            dirs = self.get_dir()
            #self.loadFileButton(filename)
            img_array, frame_num, width, height = self.loadFile(filename)
            img_arrays = self.core(img_array)
            self.writeVideo(img_arrays, dirs)
            # except e:
            #     print(e)

        messagebox.showinfo('提示', '完事了')

    def core(self,img_array_c):
        if self.check_equalizeHist.get():
            #print("均衡")
            img_array_c = self.autoEqualize(img_array_c)
        if self.check_CLAHE.get():
            img_array_c = self.limitedEqualize(img_array_c)
        return img_array_c

    def limitedEqualize(self,img_array):
        img_array_list = []
        for img in img_array:
            clahe = cv2.createCLAHE(clipLimit=self.scale.get(),
                                    tileGridSize=(8, 8))  # CLAHE (Contrast Limited Adaptive Histogram Equalization)
            img_array_list.append(clahe.apply(img))

        img_array_limited_equalized = np.array(img_array_list)
        return img_array_limited_equalized

    def config(self):
        cur_path = os.path.dirname(os.path.realpath(__file__))
        config_path = os.path.join(cur_path, "dicom.conf")
        if os.path.exists(config_path):
            self.source_dir = self.get_cf('source_dir')
            self.dest_dir = self.get_cf('dest_dir')
        else:
            config =configparser.ConfigParser()
            config['dicom'] = {'dest_dir': sys.path[0], 'source_dir': sys.path[0]}

            with open(config_path, 'w') as configfile:
                config.write(configfile)
            self.source_dir =  self.get_cf('source_dir')
            self.dest_dir =  self.get_cf('dest_dir')
        self.text_dest.delete('1.0', tk.END)
        self.text_dest.insert('1.0', self.dest_dir)
        self.text_sources.delete('1.0', tk.END)
        self.text_sources.insert('1.0', self.source_dir)


    def SlectSourceFloder(self):

        directory = filedialog.askdirectory(initialdir=self.source_dir)
        if directory == '':
            return
        self.source_dir = directory
        self.set_cf("source_dir", self.source_dir)
        self.text_sources.delete('1.0', tk.END)
        self.text_sources.insert('1.0', self.source_dir)

    def SlectDestFloder(self):

        directory = filedialog.askdirectory(initialdir=self.dest_dir)
        if directory == '':
            return
        self.dest_dir = directory
        self.set_cf("dest_dir", self.dest_dir)
        self.text_dest.delete('1.0', tk.END)
        self.text_dest.insert('1.0', self.dest_dir)

    def check(self):
        if self.check_CLAHE.get():
            scale['state'] = 'normal'
        else:
            scale['state'] = 'disable'

    def get_origin(self):
        str = ''
        if self.check_CLAHE.get():
            str = str + '_CLAHE'
        if self.check_equalizeHist.get():
            str = str + '_HIST'
        return str

    def get_dir(self):
        dirs = self.dest_dir + "/" + str(self.information['PatientName']).replace('^', '').replace(' ', '_') + '_' + \
               self.information['StudyDate']  # 目标文件为姓名+手术日期
        if not os.path.exists(dirs):  # 判断目标文件夹是否存在
            os.makedirs(dirs)
        return dirs

    def get_filename(self,directory):
        if self.information['ImageType'].__contains__('SECONDARY'):
            filename_output = directory + '/0' + str(self.information['InstanceNumber']) + self.get_origin() + '.avi'
        else:
            filename_output = directory + '/' + str(self.information['InstanceNumber']) + self.get_origin() + '.avi'
        return filename_output

    def add_text(self,imgs):
        imgs = cv2.putText(imgs, str(self.information['StudyDate']), (10, 20), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                           (255, 255, 255), 2)  # 视频添加文字
        imgs = cv2.putText(imgs, str(self.information['PositionerPrimaryAngle']), (10, 50), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                           (255, 255, 255), 2)
        imgs = cv2.putText(imgs, str(self.information['PositionerSecondaryAngle']), (10, 80), cv2.FONT_HERSHEY_COMPLEX, 0.5,
                           (255, 255, 255), 2)
        return imgs

    def creat_clahe(self,imgs, limit):  # 优化对比第一道工序
        clahe = cv2.createCLAHE(clipLimit=limit,
                                tileGridSize=(8, 8))  # CLAHE (Contrast Limited Adaptive Histogram Equalization)
        imgs = clahe.apply(imgs)
        return imgs

    def get_jpgname(self):
        jpg = self.get_dir() + '/' + str(self.information['InstanceNumber']) + self.get_origin() + '.jpg'
        return jpg


if __name__ == '__main__':
  root = tk.Tk()
  w = 250
  h = 180
  ws = root.winfo_screenwidth()
  hs = root.winfo_screenheight()
  x = (ws / 2) - (w / 2)
  y = (hs / 2) - (h / 2)
  root.geometry('%dx%d+%d+%d' % (w, h, x, y))
  root.title('DICOM TO AVI')
  MyApp(root) # 注意这句
  root.mainloop()
