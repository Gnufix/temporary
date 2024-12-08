# -*- coding: utf-8 -*-

import os
import re
import sys
import time
#import tkinter as tk
import pyautogui
import pandas as pd
import win32gui
import win32con
import ctypes
import subprocess
import customtkinter
#import CTkMessagebox

#import traceback
import inspect

#pygame import時のログ抑制
os.environ["PYGAME_HIDE_SUPPORT_Prompt"] = "hide"

import pygame.mixer 
import pyperclip
import threading
import keyboard
#import signal

from tkinter import messagebox
from ctypes import *
from PIL import Image,ImageTk

global label_font,label_config,label_restart,label_end,label_set
global sleepsec
global root
global quality
global serviceKeywordList
global securityKeywordList
global exclusionKeywordList
global tmpExcKeywordlist
global tmpOnOffList
global yesnoMsgflg
global afterid
global label
global text2
global alarm_old
global exitflg
global outputflg
global handle
global keywordlist,keywordlist2,keywordlist3,keywordlist4
global buttonlist,buttonlist2,buttonlist3,buttonlist4
global rowcount,rowcount2,rowcount3,rowcount4
global modserviceFlg,modsecurityFlg,modexclusionFlg,closeflg
global frame1
global setfontsize
global button_plus,button_minus,buttonSmall,buttonMiddle,buttonLarge,buttonNext,buttonExit,buttonConfig

serviceKeywordList=[]
securityKeywordList=[]
exclusionKeywordList=[]

keywordlist = []
keywordlist2 = []
keywordlist3 = []
keywordlist4 = []

buttonlist = []
buttonlist2 = []
buttonlist3 = []
buttonlist4 = []

tmpOnOffList = []
tmpExcKeywordlist = []

rowcount=0
rowcount2=0
rowcount3=0
rowcount4=0
afterid = 0

sleepsec = 10
setfontsize = 18

FONT_TYPE = "meiryo"

filename = "temp.txt"
pyautogui.FAILSAFE = False

#qualty = 0.77
qualty = 0.95
hwnd = 0

with open(filename,'w',encoding='utf-8') as f:
    f.truncate(0)

try:
    df = pd.read_excel(r'KeywordList.xlsx',sheet_name=[0,1,2],header=None,index_col=None)
except FileNotFoundError:
    messagebox.showerror("エラーダイアログ","keywordlist.xlsxが見つかりません。")
    sys.exit()

for i in range(df[0].shape[0]):

    if type(df[0].iloc[i,0])==str:
        serviceKeywordList.append(df[0].iloc[i,0])

for i in range(df[1].shape[0]):
    if type(df[1].iloc[i,0])==str:
        securityKeywordList.append(df[1].iloc[i,0])

for i in range(df[2].shape[0]):

    if type(df[2].iloc[i,0])==str:
        exclusionKeywordList.append(df[2].iloc[i,0])

pygame.mixer.init()
handle = 0

class Form4(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        global  exclusionKeywordList,rowcount4,modexclusionFlg,button_plus

        super().__init__(master, **kwargs)

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)

        modexclusionFlg = False

        #self.grid_columnconfigure(0,weight=1)
        #self.grid_rowconfigure(0,weight=1)

        # テキストボックスを表示する
        self.textbox = customtkinter.CTkEntry(master=self, placeholder_text="検索から除外したいキーワードを入力してください", width=450, font=self.fonts)
        self.textbox.grid(row=0, column=0, padx=20)
        #self.textbox.focus()

        # ボタンを表示する
        self.button = customtkinter.CTkButton(master=self, text="",width=10,fg_color="transparent",image=button_plus,command=self.button_function4)
        self.button.grid(row=0, column=1)

        # エラーメッセージ表示場所を確保
        self.labelErr= customtkinter.CTkLabel(master=self,text_color="red",text="")
        self.labelErr.grid(row=1, column=0)

        rowcount4=0
        keywordlist4.clear()
        buttonlist4.clear()

        for i in range(len(exclusionKeywordList)):
            self.setdata4(i)

    def setdata4(self,i):    
        global keywordlist4,buttonlist4,exclusionKeywordList,rowcount4,button_minus

        buttonlist4.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event4(i)))
        buttonlist4[i].grid(row=i+2,column=0,padx=20,pady=10,sticky="w")

        keywordlist4.insert(i,customtkinter.CTkLabel(master=self, text=exclusionKeywordList[i]))
        keywordlist4[i].grid(row=i+2,column=0,padx=70,sticky="w")
        rowcount4 = rowcount4 + 1

    def button_function4(self):
        global keywordlist4,buttonlist4,exclusionKeywordList,rowcount4,modexclusionFlg,button_minus

        if self.textbox.get() == "":
            self.labelErr.configure(text="キーワード入力欄が空欄です。")
            return

        elif self.textbox.get() in exclusionKeywordList:
            self.labelErr.configure(text="既に登録されています。")
            return            
        else:
            self.labelErr.configure(text="")

        i = len(exclusionKeywordList)
        rowcount4 = rowcount4 + 1

        buttonlist4.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event4(i)))
        buttonlist4[i].grid(row=rowcount4+1,column=0,padx=20,pady=10,sticky="w")

        keywordlist4.insert(i,customtkinter.CTkLabel(master=self, text=self.textbox.get()))
        keywordlist4[i].grid(row=rowcount4+1,column=0,padx=70,sticky="w")

        exclusionKeywordList.insert(i,self.textbox.get())
        
        self.textbox.delete(first_index=0,last_index="end")
        modexclusionFlg = True

    def button_event4(self,setnum):
        global keywordlist4,buttonlist4,exclusionKeywordList,modexclusionFlg

        #print(setnum)
        keywordlist4[setnum].destroy()
        buttonlist4[setnum].destroy()
        exclusionKeywordList[setnum] = "NoDataNoDataNoDataNoDataNoData"
        modexclusionFlg = True

class Form3(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        global  securityKeywordList,rowcount3,modsecurityFlg,button_plus

        super().__init__(master, **kwargs)

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)

        modsecurityFlg = False

        #self.grid_columnconfigure(0,weight=1)
        #self.grid_rowconfigure(0,weight=1)

        # テキストボックスを表示する
        self.textbox = customtkinter.CTkEntry(master=self, placeholder_text="検索したいキーワードを入力してください。", width=450, font=self.fonts)
        self.textbox.grid(row=0, column=0, padx=20)
        #self.textbox.focus()

        # ボタンを表示する
        self.button = customtkinter.CTkButton(master=self, text="",width=10,fg_color="transparent",image=button_plus,command=self.button_function3)
        self.button.grid(row=0, column=1)

        # エラーメッセージ表示場所を確保
        self.labelErr= customtkinter.CTkLabel(master=self,text_color="red",text="")
        self.labelErr.grid(row=1, column=0)

        rowcount3=0
        keywordlist3.clear()
        buttonlist3.clear()

        for i in range(len(securityKeywordList)):
            self.setdata3(i)

    def setdata3(self,i):    
        global keywordlist3,buttonlist3,securityKeywordList,rowcount3,button_minus

        buttonlist3.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event3(i)))
        buttonlist3[i].grid(row=i+2,column=0,padx=20,pady=10,sticky="w")

        keywordlist3.insert(i,customtkinter.CTkLabel(master=self, text=securityKeywordList[i]))
        keywordlist3[i].grid(row=i+2,column=0,padx=70,sticky="w")
        rowcount3 = rowcount3 + 1

    def button_function3(self):
        global keywordlist3,buttonlist3,securityKeywordList,rowcount3,modsecurityFlg,button_minus

        if self.textbox.get() == "":
            self.labelErr.configure(text="キーワード入力欄が空欄です。")
            return

        elif self.textbox.get() in securityKeywordList:
            self.labelErr.configure(text="既に登録されています。")
            return            
        else:
            self.labelErr.configure(text="")

        i = len(securityKeywordList)
        rowcount3 = rowcount3 + 1

        buttonlist3.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event3(i)))
        buttonlist3[i].grid(row=rowcount3+1,column=0,padx=20,pady=10,sticky="w")

        keywordlist3.insert(i,customtkinter.CTkLabel(master=self, text=self.textbox.get()))
        keywordlist3[i].grid(row=rowcount3+1,column=0,padx=70,sticky="w")

        securityKeywordList.insert(i,self.textbox.get())
        
        self.textbox.delete(first_index=0,last_index="end")
        modsecurityFlg = True

    def button_event3(self,setnum):
        global keywordlist3,buttonlist3,securityKeywordList,modsecurityFlg

        #print(setnum)
        keywordlist3[setnum].destroy()
        buttonlist3[setnum].destroy()
        securityKeywordList[setnum] = "NoDataNoDataNoDataNoDataNoData"
        modsecurityFlg = True

class Form2(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        global  serviceKeywordList,rowcount2,modserviceFlg,button_plus

        super().__init__(master, **kwargs)

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)

        modserviceFlg = False

        #self.grid_columnconfigure(0,weight=1)
        #self.grid_rowconfigure(0,weight=1)

        # テキストボックスを表示する
        self.textbox = customtkinter.CTkEntry(master=self, placeholder_text="検索したいキーワードを入力してください。", width=450, font=self.fonts)
        self.textbox.grid(row=0, column=0, padx=20)
        #self.textbox.focus()

        # ボタンを表示する
        self.button = customtkinter.CTkButton(master=self, text="",width=10,fg_color="transparent",image=button_plus,command=self.button_function2)
        self.button.grid(row=0, column=1)

        # エラーメッセージ表示場所を確保
        self.labelErr= customtkinter.CTkLabel(master=self,text_color="red",text="")
        self.labelErr.grid(row=1, column=0)

        rowcount2=0
        keywordlist2.clear()
        buttonlist2.clear()

        for i in range(len(serviceKeywordList)):
            self.setdata2(i)

    def setdata2(self,i):    
        global keywordlist2,buttonlist2,serviceKeywordList,rowcount2,button_minus

        buttonlist2.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event2(i)))
        buttonlist2[i].grid(row=i+2,column=0,padx=20,pady=10,sticky="w")

        keywordlist2.insert(i,customtkinter.CTkLabel(master=self, text=serviceKeywordList[i]))
        keywordlist2[i].grid(row=i+2,column=0,padx=70,sticky="w")
        rowcount2 = rowcount2 + 1

    def button_function2(self):
        global keywordlist2,buttonlist2,serviceKeywordList,rowcount2,modserviceFlg,button_minus

        if self.textbox.get() == "":
            self.labelErr.configure(text="キーワード入力欄が空欄です。")
            return

        elif self.textbox.get() in serviceKeywordList:
            self.labelErr.configure(text="既に登録されています。")
            return            
        else:
            self.labelErr.configure(text="")

        i = len(serviceKeywordList)
        rowcount2 = rowcount2 + 1

        buttonlist2.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event2(i)))
        buttonlist2[i].grid(row=rowcount2+1,column=0,padx=20,pady=10,sticky="w")

        keywordlist2.insert(i,customtkinter.CTkLabel(master=self, text=self.textbox.get()))
        keywordlist2[i].grid(row=rowcount2+1,column=0,padx=70,sticky="w")

        serviceKeywordList.insert(i,self.textbox.get())
        
        self.textbox.delete(first_index=0,last_index="end")
        modserviceFlg = True

    def button_event2(self,setnum):
        global keywordlist2,buttonlist2,serviceKeywordList,modserviceFlg

        #print(setnum)
        keywordlist2[setnum].destroy()
        buttonlist2[setnum].destroy()
        serviceKeywordList[setnum] = "NoDataNoDataNoDataNoDataNoData"
        modserviceFlg = True

class Form1(customtkinter.CTkScrollableFrame):
    def __init__(self, master, **kwargs):
        global  tmpExcKeywordlist,rowcount,button_plus

        super().__init__(master, **kwargs)

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)

        #self.grid_columnconfigure(0,weight=1)
        #self.grid_rowconfigure(0,weight=1)

        # テキストボックスを表示する
        self.textbox = customtkinter.CTkEntry(master=self, placeholder_text="一時的に無視したいキーワードを入力してください。", width=450, font=self.fonts)
        self.textbox.grid(row=0, column=0, padx=20)
        #self.textbox.focus()

        # ボタンを表示する
        #self.button = customtkinter.CTkButton(master=self, text="add", width=10,height=8,command=self.button_function, font=self.fonts)
        self.button = customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_plus,command=self.button_function)
        self.button.grid(row=0, column=1)

        # エラーメッセージ表示場所を確保
        self.labelErr= customtkinter.CTkLabel(master=self,text_color="red",text="")
        self.labelErr.grid(row=1, column=0)

        rowcount=0
        keywordlist.clear()
        buttonlist.clear()

        for i in range(len(tmpExcKeywordlist)):
            self.setdata(i)

    def setdata(self,i):    
        global keywordlist,buttonlist,tmpExcKeywordlist,tmpOnOffList,rowcount,button_minus

        buttonlist.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event(i)))
        buttonlist[i].grid(row=i+2,column=0,padx=20,pady=10,sticky="w")

        if tmpOnOffList[i] == "off":
            self.switch_var = customtkinter.StringVar(value="off")
        else:
            self.switch_var = customtkinter.StringVar(value="on")

        keywordlist.insert(i,customtkinter.CTkSwitch(master=self, text=tmpExcKeywordlist[i], command=lambda:self.switch_event(tmpExcKeywordlist[i],i),variable=self.switch_var, onvalue="on", offvalue="off"))
        keywordlist[i].grid(row=i+2,column=0,padx=70,sticky="w")
        rowcount = rowcount + 1

    def button_function(self):
        global keywordlist,buttonlist,tmpExcKeywordlist,tmpOnOffList,rowcount,button_minus

        if self.textbox.get() == "":
            self.labelErr.configure(text="キーワード入力欄が空欄です。")
            return

        elif self.textbox.get() in tmpExcKeywordlist:
            self.labelErr.configure(text="既に登録されています。")
            return            
        else:
            self.labelErr.configure(text="")

        i = len(tmpExcKeywordlist)
        rowcount = rowcount + 1

        self.switch_var = customtkinter.StringVar(value="on")

        buttonlist.insert(i,customtkinter.CTkButton(master=self,text="",width=10,fg_color="transparent",image=button_minus,command=lambda:self.button_event(i)))
        #buttonlist[i].grid(row=rowcount,column=0,padx=20,pady=10,sticky="w")
        buttonlist[i].grid(row=rowcount+1,column=0,padx=20,pady=10,sticky="w")

        keyword=self.textbox.get()

        keywordlist.insert(i,customtkinter.CTkSwitch(master=self, text=self.textbox.get(), command=lambda:self.switch_event(keyword,i),variable=self.switch_var, onvalue="on", offvalue="off"))
        #keywordlist[i].grid(row=rowcount,column=0,padx=70,sticky="w")
        keywordlist[i].grid(row=rowcount+1,column=0,padx=70,sticky="w")

        tmpExcKeywordlist.insert(i,self.textbox.get())
        tmpOnOffList.insert(i,"on")

        self.textbox.delete(first_index=0,last_index="end")

    def button_event(self,setnum):
        global keywordlist,buttonlist,tmpExcKeywordlist,tmpOnOffList

        #print(setnum)
        keywordlist[setnum].destroy()
        buttonlist[setnum].destroy()
        #del keywordlist[setnum]
        #del buttonlist[setnum]
        tmpExcKeywordlist[setnum] = "NoData"
        tmpOnOffList[setnum] = "NoData"

    def switch_event(self,keyword,num):
        global tmpOnOffList,tmpExcKeywordlist

        #print(str(num) + "," + keyword)

        i = tmpExcKeywordlist.index(keyword)

        if tmpOnOffList[i] == "on":
            tmpOnOffList[i]="off"
        else:
            tmpOnOffList[i]="on"

#class App(customtkinter.CTk):
class App(customtkinter.CTkToplevel):
    def __init__(self):
        global  button_plus,button_minus
        super().__init__()

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)

        self.attributes("-topmost",True)

        self.protocol("WM_DELETE_WINDOW",lambda:self.quit_me(self))

        # フォームサイズ設定
        self.geometry("640x690")
        self.title("アラームキーワード設定ツール")
        self.resizable(0,0)

        button_plus = customtkinter.CTkImage(Image.open("buttonimage//plus.png"),size=(20,20))
        button_minus = customtkinter.CTkImage(Image.open("buttonimage//minus.png"),size=(20,20))
        button_saveexit = customtkinter.CTkImage(Image.open("buttonimage//saveexit.png"),size=(30,30))

        customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
        customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

        self.tabview = customtkinter.CTkTabview(self)
        self.tabview.pack(padx=20,pady=40)

        self.tabview.add("一時無視")
        self.tabview.add("サービス制御")
        self.tabview.add("法人セキュリティ")
        self.tabview.add("検索除外")
        self.tabview.set("一時無視")

        self.my_frame = Form1(master=self.tabview.tab("一時無視"),width=550,height=520)
        self.my_frame.grid(row=0,column=0,padx=20,pady=10,sticky="nsew")

        self.my_frame2 = Form2(master=self.tabview.tab("サービス制御"),width=550,height=520)
        self.my_frame2.grid(row=0,column=0,padx=20,pady=10,sticky="nsew")

        self.my_frame3 = Form3(master=self.tabview.tab("法人セキュリティ"),width=550,height=520)
        self.my_frame3.grid(row=0,column=0,padx=20,pady=10,sticky="nsew")

        self.my_frame4 = Form4(master=self.tabview.tab("検索除外"),width=550,height=520)
        self.my_frame4.grid(row=0,column=0,padx=20,pady=10,sticky="nsew")

        self.label_2 = customtkinter.CTkLabel(master=self,text="画面サイズ")
        self.label_2.place(x=485, y=10)

        self.buttonSmall = customtkinter.CTkButton(master=self, text="小", width=10,height=8,command=lambda:self.modsize_function("小"), font=self.fonts)
        self.buttonSmall.place(x=555, y=10)
        
        self.buttonMiddle = customtkinter.CTkButton(master=self, text="中", width=10,height=8,command=lambda:self.modsize_function("中"), font=self.fonts)
        self.buttonMiddle.place(x=575, y=10)

        self.buttonLarge = customtkinter.CTkButton(master=self, text="大", width=10,height=8,command=lambda:self.modsize_function("大"), font=self.fonts)
        self.buttonLarge.place(x=595, y=10)

        self.label_3 = customtkinter.CTkLabel(master=self,text="設定を保存して終了",font=self.fonts)
        self.label_3.place(x=443, y=650)

        self.buttonClose = customtkinter.CTkButton(master=self, text="",width=10,fg_color="transparent",image=button_saveexit,command=lambda:self.close_function(self))
        #self.buttonClose = customtkinter.CTkButton(master=self, text="save&close", width=10,height=8,command=lambda:self.close_function(self), font=self.fonts)
        self.buttonClose.place(x=582, y=645)


    def modsize_function(self,text):

        if text == "小":
            customtkinter.set_widget_scaling(1.0)
            customtkinter.set_window_scaling(1.0)
        elif text == "中":
            customtkinter.set_widget_scaling(1.25)
            customtkinter.set_window_scaling(1.25)
        else:
            customtkinter.set_widget_scaling(1.5)
            customtkinter.set_window_scaling(1.5)
            
    def close_function(self,root_window):
        global keywordlist,buttonlist,tmpExcKeywordlist,tmpOnOffList
        global exclusionKeywordList,serviceKeywordList,securityKeywordList
        global closeflg,modsecurityFlg,modserviceFlg,modexclusionFlg

        tmplist1 = [a for a in tmpExcKeywordlist if a != "NoData"]
        tmplist2 = [b for b in tmpOnOffList if b != "NoData"]
        tmplist3 = [c for c in serviceKeywordList if c != "NoDataNoDataNoDataNoDataNoData"]
        tmplist4 = [d for d in securityKeywordList if d != "NoDataNoDataNoDataNoDataNoData"]
        tmplist5 = [e for e in exclusionKeywordList if e != "NoDataNoDataNoDataNoDataNoData"]

        tmpExcKeywordlist = tmplist1
        tmpOnOffList = tmplist2
        serviceKeywordList = tmplist3
        securityKeywordList = tmplist4
        exclusionKeywordList = tmplist5

        for i in range(len(tmpExcKeywordlist)):
            keywordlist[i].destroy
            buttonlist[i].destroy

        if modexclusionFlg == True or modsecurityFlg == True or modserviceFlg == True:

            dfser = pd.DataFrame(serviceKeywordList)
            dfsec = pd.DataFrame(securityKeywordList)
            dfexc = pd.DataFrame(exclusionKeywordList)
            
            with pd.ExcelWriter(r'KeywordList.xlsx') as writer:
                dfser.to_excel(writer,sheet_name="サービス制御",index=False,header=False)
                dfsec.to_excel(writer,sheet_name="法人セキュリティ",index=False,header=False)
                dfexc.to_excel(writer,sheet_name="除外",index=False,header=False)

        customtkinter.set_widget_scaling(1.0)
        customtkinter.set_window_scaling(1.0)

        closeflg = True

        # root_window.label_2.place_forget()
        # root_window.label_3.place_forget()
        # root_window.buttonSmall.place_forget()
        # root_window.buttonMiddle.place_forget()
        # root_window.buttonLarge.place_forget()
        # root_window.buttonClose.place_forget()
        # root_window.my_frame.grid_forget()
        # root_window.my_frame2.grid_forget()
        # root_window.my_frame3.grid_forget()
        # root_window.my_frame4.grid_forget()
        # root_window.tabview.delete("一時無視")
        # root_window.tabview.delete("サービス制御")
        # root_window.tabview.delete("法人セキュリティ")
        # root_window.tabview.delete("検索除外")
        # root_window.tabview.pack_forget

        root_window.withdraw()

        #root_window.quit()
        #root_window.destroy()
        #sys.exit()

    def quit_me(self,root_window):
        global exitflg

        customtkinter.set_widget_scaling(1.0)
        customtkinter.set_window_scaling(1.0)

        exitflg = True

        #root_window.quit()
        #root_window.destroy()
        #sys.exit()

#ファイル実行時に立ち上がるコマンドプロンプト画面をPC画面の右上に表示
def enumHandler(hwnd,lParam):
    global handle
    
    scr_w,scr_h = pyautogui.size()
    
    try:
        appname = "exe"
        width = 600
        length = 100
        #xpos = int(scr_w/1.3333) - length
        #ypos = int(scr_h/1.3333) - width 
        xpos = scr_w - width
        #ypos = scr_h - length 
        ypos = scr_h - length - 50

        if win32gui.IsWindowVisible(hwnd):
            if appname in win32gui.GetWindowText(hwnd):         
                win32gui.MoveWindow(hwnd,xpos,ypos,width,length,True)
                handle = hwnd

    except:
        raise

win32gui.EnumWindows(enumHandler,None)

def imageclick(sImageFileName,iSleepCount,x,y,clickcount):
    global qualty,exitflg

    for i in range(iSleepCount):
        
        try:
            #confidence(曖昧検索)指定時は画像ファイルパスに日本語が含まれているとエラーになる
            Result = pyautogui.locateCenterOnScreen("image\\" + sImageFileName,grayscale=True,confidence=qualty)
            break

        except pyautogui.ImageNotFoundException:
            time.sleep(1)

    #画面認識失敗時にエラー
    if i == iSleepCount-1:
        if sImageFileName == "01_ClickSyoudakuSousa.png":
            messagebox.showerror("エラーダイアログ","承諾操作開始ボタンのクリックに失敗しました。")
        elif sImageFileName == "03_ClickSyoudakuSousaStop.png":
            messagebox.showerror("エラーダイアログ","承諾操作停止ボタンのクリックに失敗しました。")
        else:
            messagebox.showerror("エラーダイアログ","画像(" + sImageFileName + ")マッチングに失敗しました。")
            
        exitflg = True
        sys.exit()

    pyautogui.click(Result.x + x,Result.y + y,clicks = clickcount,interval=3)
    return True

def Init():
    global alarm_old,root,afterid,yesnoMsgflg,label

    #再帰関数起動回数上限値取得
    #print(sys.getrecursionlimit())

    failureStop = win32gui.FindWindow(None,"ProactnesII NM-発生中障害一覧-IPOPS(MTB)-承諾中 - Internet Explorer")
    
    if failureStop != 0:
        win32gui.SetWindowPos(failureStop,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        
        # 承諾操作停止ボタンクリック
        imageclick("00_ClickSyoudakuSousaTyuushi.png",5,0,0,1)

        win32gui.SetWindowPos(failureStop,win32con.HWND_NOTOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    alarm_old = ""
    yesnoMsgflg = False
    
    #メインウィンドウ作成
    root = customtkinter.CTk()

    #bindid = root.bind('<Escape>',on_esc)

    #ツール画面サイズと表示座標設定
    root.geometry("860x50+0+0")

    #タイトルバー非表示(終了はalt+f4)
    root.overrideredirect(True)

    #ウィンドウサイズ変更を制限
    root.resizable(False,False)
    
    #背景の透明度設定
    root.attributes("-alpha",0.9)

    label = customtkinter.CTkLabel(root,text="【IPOPS監視中･･･(10秒毎チェック)】※ESCキーでツール終了",width=860,height=50,fg_color=("yellow","yellow"),text_color=("black","black"),font=("meiryo",30))
    label.pack()

    root.attributes("-topmost",True)
    
    afterid = root.after(sleepsec*1000,alarmCheck)

    #ツールウィンドウの表示
    root.mainloop()

def alarmCheck():
    global serviceKeywordList,securityKeywordList,exclusionKeywordList
    global quality,root,yesnoMsgflg,afterid,alarm_old,exitflg,outputflg,handle,tmpExcKeywordlist,tmpOnOffList,sleepsec
    global frame1,text2,buttonSmall,buttonMiddle,buttonLarge,buttonNext,buttonExit,buttonConfig,setfontsize
    global label,label_font,label_config,label_restart,label_end,label_set

    #depth = len(inspect.stack())
    #print(f"stack-depth: {depth}")
    #print(inspect.stack())
    #print()

    if yesnoMsgflg == True:

        #root.withdraw()

        yesnoMsgflg = False       

        frame1.grid_forget()
        text2.grid_forget()
        label_font.place_forget()
        label_config.place_forget()
        label_restart.place_forget()
        label_end.place_forget()
        buttonSmall.place_forget()
        buttonMiddle.place_forget()
        buttonLarge.place_forget()
        buttonNext.place_forget()
        buttonExit.place_forget()
        buttonConfig.place_forget()
        label_set.place_forget()

        #root.destroy()

        #root = customtkinter.CTk()
        root.geometry("860x50+0+0")

        root.overrideredirect(True)

        #ウィンドウサイズ変更を制限
        root.resizable(False,False)
    
        #背景の透明度設定
        root.attributes("-alpha",0.9)

        #label = customtkinter.CTkLabel(root,text="【IPOPS監視中･･･(10秒毎チェック)】※ESCキーでツール終了",width=860,height=50,fg_color=("yellow","yellow"),text_color=("black","black"),font=("meiryo",30))
        #label.pack()

        root.attributes("-topmost",True)
        #root.deiconify()

        label.configure(text="【IPOPS監視中･･･(10秒毎チェック)】※ESCキーでツール終了",width=860,height=50,fg_color=("yellow","yellow"))
        label.place(x=0, y=0)
        #label.pack()

        #ウィンドウサイズ変更を制限
        #root.resizable(False,False)

        #GUIフォームの透明度設定
        #root.attributes("-alpha",0.9)

        #最前面表示
        #root.attributes("-topmost",True)

        afterid = root.after(sleepsec*1000,alarmCheck)
  
    else:
        # IPOPS MTの操作画面を最前面 & 位置調整
        failure = win32gui.FindWindow(None,"ProactnesII NM-発生中障害一覧-IPOPS(MTB) - Internet Explorer")

        if failure == 0:
            messagebox.showerror("エラーダイアログ","「ProactnesII NM-発生中障害一覧-IPOPS(MTB) - Internet Explorer」画面が見つかりません。")
            exitflg == True
            sys.exit()

        win32gui.SetWindowPos(failure,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

        # 承諾操作開始ボタンクリック
        imageclick("01_ClickSyoudakuSousa.png",5,0,0,1)

        #time.sleep(1)

        pyperclip.copy("")

        for i in range(3):
            # 障害箇所欄の下をクリック
            imageclick("02_ClickUnderSyougaiKasyo.png",5,0,30,1)
            #time.sleep(1)

            # MTBの全アラームをクリップボードへコピー 
            pyautogui.keyDown('ctrl')
            #time.sleep(1)
            pyautogui.press('a')
            #time.sleep(1)
            pyautogui.press('c')
            #time.sleep(1)
            almdata = pyperclip.paste()
            pyautogui.keyUp('ctrl') 

            if len(almdata) > 0:
                break
        
        if len(almdata) == 0:
            messagebox.showerror("エラーダイアログ","アラーム取得失敗")
            exitflg == True
            sys.exit()

        # 承諾操作停止ボタンクリック
        imageclick("03_ClickSyoudakuSousaStop.png",5,0,0,1)       

        win32gui.SetWindowPos(failure,win32con.HWND_NOTOPMOST,0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

        #メモリにコピーした全アラームを1行単位でリスト分割
        almlist = re.split("\r\n",almdata)

        serviceList = ""
        securityList = ""

        hitflg = False
        
        for iRpCnt in range(len(almlist)):
            skipflg = False

            # 除外キーワードリスト検索
            for iRpCnt2 in range(len(exclusionKeywordList)):
                if (exclusionKeywordList[iRpCnt2] in almlist[iRpCnt] )== True:
                    skipflg = True
                    break

            if skipflg == True:
                continue
            
            #一時除外キーワードリスト検索
            for iRpCnt2 in range(len(tmpOnOffList)):

                if tmpOnOffList[iRpCnt2]=="on":
                
                    if (tmpExcKeywordlist[iRpCnt2] in almlist[iRpCnt] )== True:
                        skipflg = True
                        break

            if skipflg == True:
                continue

            date_type = re.compile(r"""(
                (\d{4})         # First 4 digits number
                (\D)            # Something other than numbers
                (\d{1,2})       # 1 or 2 digits number
                (\D)            # Something other than numbers
                (\d{1,2})       # 1 or 2 digits number
                (-)
                (\d{2})         # 2 digits number
                (:)
                (\d{2})         # 2 digits number
                (:)
                (\d{2})         # 2 digits number
                )""",re.VERBOSE)

            # サービス制御キーワードリスト検索
            for iRpCnt2 in range(len(serviceKeywordList)):
                if (serviceKeywordList[iRpCnt2] in almlist[iRpCnt]) == True:
                    #worklist = str(re.findall("\(MTB\).*","    ".join((almlist[iRpCnt].replace("\t","    ").split()))))
                    worklist = str(re.findall(r"\(MTB\).*",almlist[iRpCnt].replace("\t","    ")))
                    worklist = worklist.lstrip("['").rstrip("']")
                
                    hit_date = date_type.search(worklist)
                    start,end = hit_date.span()
                    worklist = worklist[0:end]

                    serviceList = serviceList + worklist + "\n"
                    hitflg = True
                    break

            # 法人セキュリティキーワードリスト検索
            for iRpCnt2 in range(len(securityKeywordList)):
                if (securityKeywordList[iRpCnt2] in almlist[iRpCnt]) == True:
                    #worklist = str(re.findall("\(MTB\).*"," ".join((almlist[iRpCnt].replace("\t"," ").split()))))
                    worklist = str(re.findall(r"\(MTB\).*",almlist[iRpCnt].replace("\t","    ")))
                    worklist = worklist.lstrip("['").rstrip("']")
                    
                    hit_date = date_type.search(worklist)
                    start,end = hit_date.span()
                    worklist = worklist[0:end]

                    securityList = securityList + worklist + "\n"
                    hitflg = True
                    break

        stringList = "【アラーム発生_サービス制御】\n" + serviceList + "\n" + "【アラーム発生_法人セキュリティ】\n" + securityList
        stringList = stringList.replace("[","")

        if hitflg == False or alarm_old == stringList:
            #root.deiconify() 
            afterid = root.after(sleepsec*1000,alarmCheck)

        else:
            alarm_old = stringList

            ctypes.windll.user32.SetForegroundWindow(handle)
            #print("hit:\n"+stringList)

            with open(filename,'a',encoding='utf-8') as f:
                print("hit:\n"+stringList,file=f)
        
            outputflg = True

            #監視メッセージ非表示
            #root.withdraw()
            #label.pack_forget()
            label.configure(text="",fg_color="transparent",width=1,height=1)
            label.place(x=20, y=850)

            #ツール画面サイズと表示座標設定
            root.geometry(f"1800x920+0+0")
            
            #タイトルバー非表示(終了はalt+f4)
            root.overrideredirect(True)
            
            root.attributes("-alpha",1.0)

            #ウィンドウサイズ変更を制限
            root.resizable(False,False)

            customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
            customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

            button_restart = customtkinter.CTkImage(Image.open("buttonimage//restart.png"),size=(30,30))
            button_end = customtkinter.CTkImage(Image.open("buttonimage//end.png"),size=(30,30))
            button_config = customtkinter.CTkImage(Image.open("buttonimage//haguruma.png"),size=(30,30))

            label_font = customtkinter.CTkLabel(master=root,text="フォントサイズ")
            label_font.place(x=1627, y=10)

            buttonSmall = customtkinter.CTkButton(master=root, text="小", width=10,height=8,command=lambda:modfsize_function("小"), font=("meiryo",15))
            buttonSmall.place(x=1707, y=10)
        
            buttonMiddle = customtkinter.CTkButton(master=root, text="中", width=10,height=8,command=lambda:modfsize_function("中"), font=("meiryo",15))
            buttonMiddle.place(x=1737, y=10)
        
            buttonLarge = customtkinter.CTkButton(master=root, text="大", width=10,height=8,command=lambda:modfsize_function("大"), font=("meiryo",15))
            buttonLarge.place(x=1767, y=10)

            label_set = customtkinter.CTkLabel(master=root,text="ctrl+c,vでデータコピー可   設定ボタンで一時無視/検索/検索除外キーワード設定可",text_color=("green2"),font=("meiryo",22))
            label_set.place(x=10, y=873)

            label_config = customtkinter.CTkLabel(master=root,text="設定",font=("meiryo",22))
            label_config.place(x=1400, y=873)

            buttonConfig = customtkinter.CTkButton(master=root, text="",width=20,fg_color="transparent",image=button_config,command=modconfig)
            #buttonConfig = customtkinter.CTkButton(master=root, text="     設定     ", width=30,height=8,command=modconfig, font=("meiryo",23))
            buttonConfig.place(x=1445, y=870)

            label_restart = customtkinter.CTkLabel(master=root,text="監視再開",font=("meiryo",22))
            label_restart.place(x=1527, y=873)

            buttonNext = customtkinter.CTkButton(master=root, text="",width=20,fg_color="transparent",image=button_restart,command=nextloop)
            #buttonNext = customtkinter.CTkButton(master=root, text="   監視再開   ", width=30,height=8,command=nextloop, font=("meiryo",23))
            buttonNext.place(x=1620, y=870)

            label_end = customtkinter.CTkLabel(master=root,text="終了",font=("meiryo",22))
            label_end.place(x=1705, y=873)

            buttonExit = customtkinter.CTkButton(master=root, text="",width=20,fg_color="transparent",image=button_end,command=lambda:loopexit(root))
            #buttonExit = customtkinter.CTkButton(master=root, text="     終了     ", width=30,height=8,command=lambda:loopexit(root), font=("meiryo",23))
            buttonExit.place(x=1750, y=870)

            frame1 = customtkinter.CTkScrollableFrame(master=root,width=1770,height=800)
  
            frame1.grid(row=0,column=0,padx=0,pady=50)

            text2 = customtkinter.CTkTextbox(master=frame1,width=1770,height=800,font=("meiryo",setfontsize))
            text2.insert(0., stringList)

            text2.grid(row=0,column=0)

            root.attributes("-topmost",True)
            yesnoMsgflg = True

            pygame.mixer.music.load(r'C:\Windows\Media\Ring08.wav')
            pygame.mixer.music.play(loops=1,start=0.0) 

            root.after_cancel(afterid)

            #label.pack()

            #root.deiconify()

            #root.after(sleepsec*1000,alarmCheck)

def modfsize_function(text):
    global text2,setfontsize

    if text == "小":
        text2.configure(font=("meiryo",14))
        setfontsize = 14
    elif text == "中":
        text2.configure(font=("meiryo",18))
        setfontsize = 18
    else:
        text2.configure(font=("meiryo",22))
        setfontsize = 22

def modconfig():
    global root,buttonNext,buttonExit,buttonConfig

    root.attributes("-topmost",False)

    buttonNext.configure(state="disabled")
    buttonExit.configure(state="disabled")
    buttonConfig.configure(state="disabled")
    
    app = App()
    #app.mainloop()

def nextloop():
    alarmCheck()

def loopexit(root_window):
    global exitflg

    exitflg = True
    #root_window.quit()
    #root_window.destroy()

    #sys.exit()

class CustomThread(threading.Thread):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._run = self.run
        self.run = self.set_id_and_run

    def set_id_and_run(self):
        self.id = threading.get_native_id()
        self._run()

    def get_id(self):
        return self.id
        
    def raise_exception(self):

        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(
            ctypes.c_long(self.get_id()), 
            ctypes.py_object(SystemExit)
        )

        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(
                ctypes.c_long(self.get_id()), 
                0
            )
            print('Failure in raising exception')

if __name__ == '__main__':
    outputflg = False
    exitflg = False
    closeflg = False

    thread = CustomThread(target=Init,daemon=True)
    thread.start()
    while True:
        if keyboard.is_pressed('esc') == True or exitflg == True:
            break
        elif closeflg == True:
            buttonNext.configure(state="normal")
            buttonExit.configure(state="normal")
            buttonConfig.configure(state="normal")
            closeflg=False
        
        time.sleep(0.1)
 
    if outputflg == True:
        subprocess.run(['explorer',r'temp.txt'])
    
    # raise_exceptionを呼び出すことでスレッドが終了
    thread.raise_exception()

    # 既に終了しているので処理を待機しないはず
    thread.join(timeout = 5)