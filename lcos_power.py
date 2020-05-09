# -*- coding: UTF-8 -*-
#csvファイルのプロット・解析プログラム

import wx
import os
import csv 
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.font_manager import FontProperties
import matplotlib.ticker as ptick
from statistics import mean, median, variance,stdev
#from scipy import interpolate
import winsound
from ctypes import *
import ctypes.util  
from numpy.ctypeslib import ndpointer 

# import numpy.random.common #exe化時のエラー回避用
# import numpy.random.bounded_integers
# import numpy.random.entropy
import visa
import pyautogui
import time
import threading
import cv2
import math
import datetime

cdir = os.path.dirname(os.path.abspath(__file__))
os.chdir(cdir)
dll = cdll.LoadLibrary(r"SubDisplay.dll")
#dll.SubDisp_TermWin()




rm = visa.ResourceManager()
lis =rm.list_resources()

length = 1440 #横軸（ガウス関数方面の長さ）
height = 1050 
base = 770    #基準（０）になるＬＣＯＳの深さ
row = [0]*length #ガウス関数を反映させるlist
step =9 #lcosの繰り返しステップ
wlength = '785'
a = 330 #基準と最大値の間の変化数（ガウス関数の最大値）
flag = 0 #ガウス関数の波形を保存するタイミング
shot = [(i,0)for i in range(length)] #ガウス関数保存用
cc = 0
dd = 0

mat = [0,0,0]*(length*height)



def MakeWindow(mat):
    #doubleのポインタのポインタ型を用意
    _DOUBLE_PP = ndpointer(dtype=np.uint8, ndim=1, flags='C')
    #関数の引数の型を指定(ctypes)　
    dll.SubDisp_OutputWin.argtypes = [ c_uint32, c_uint32, _DOUBLE_PP]
    #関数が返す値の型を指定(今回は返り値なし)
    dll.SubDisp_OutputWin.restype = None

    arr = (c_uint8 * (length*height*3))(*mat)
    
    #int型もctypeのc_int型へ変換して渡す
    # row = ctypes.c_uint32(length)
    # col = ctypes.c_uint32(height)
    n = np.ctypeslib.as_array(arr)


    dll.SubDisp_OutputWin(length,height,n)


def MakeColumn(r): #ガウス関数方向の値に対応したLCOSグレーティング方向（縦）の作成

    B = [int((r)/step*(j+1)) for j in range(step)]
    A = B*math.ceil(int(height/step)+1)
    A = A[:height]
    T = [A]*length

    return T

MakeWindow(mat)

class TestThread(threading.Thread):
 
    def run(self):
        global start,stop,d,ti,r,cou
        wx.MessageBox('このウインドウをアクティブ化してカーソルをクリックする位置に置いた後、エンターを押してください','セット')
        r = start
        c = 1
        while 1:
            
            img = np.zeros((height, length, 3), np.uint8) #画像の種

            row2 = MakeColumn(r)
            for h in range(len(row2)): #1024段階の値を説明書通り色に変換
                for w in range(len(row2[h])):
                    m = row2[h][w]
                    
               
                    img[w, h] = [(m & 0b1111)<<4, ((m >>4)& 0b111)<<5, ((m>>7) & 0b111)<<5] 
            
            # セカンドディスプレイにフルスクリーンで表示　(メインがフルHDの時)
            img2 = img.tolist()
            mat = img.flatten()
            MakeWindow(mat)
            time.sleep(1)
            for c in range(cou): #データを複数回取るため
                time.sleep(ti)
                pyautogui.click()
                text_entry.SetLabel(str(r)+'のデータを取得しました('+str(c+1)+'個目)') 

            r = r+d
            c = c+1
            
            if r >= stop:
                img = np.zeros((height, length, 3), np.uint8) #画像の種

                row2 = MakeColumn(stop)
                for h in range(len(row2)): #1024段階の値を説明書通り色に変換
                    for w in range(len(row2[h])):
                        m = row2[h][w]
                        
                   
                        img[w, h] = [(m & 0b1111)<<4, ((m >>4)& 0b111)<<5, ((m>>7) & 0b111)<<5] 
                
                # セカンドディスプレイにフルスクリーンで表示　(メインがフルHDの時)
                img2 = img.tolist()
                mat = img.flatten()
                MakeWindow(mat)
                time.sleep(1)
                for c in range(cou): #データを複数回取るため
                    time.sleep(ti)
                    pyautogui.click()
                    text_entry.SetLabel(str(stop)+'のデータを取得しました('+str(c+1)+'個目)')                   
                break



        text_entry.SetLabel('データの取得が完了しました') 
        winsound.Beep(2500, 1000)
        
class TestThread2(threading.Thread):
 
    def run(self):
        power = []
        global wlength
        try:
            myinst=rm.open_resource('USB0::0x1313::0x8072::P2009036::INSTR')
        except:
            wx.MessageBox(' パワーメータが接続されていない可能性があります','error')
            return 0
        myinst.read_termination = '\n'
        myinst.write_termination = '\n'
        myinst.write("AVERage 10 ")
        myinst.write("CORRection:COLLect:ZERO")
        myinst.write("CORRection:WAVelength "+wlength)
        text_entry.SetLabel('波長:'+wlength)
        time.sleep(3)
        
        
        global start,stop,d,ti,r
        r = start
        c = 1
        while 1:
            
            img = np.zeros((height, length, 3), np.uint8) #画像の種

            row2 = MakeColumn(r)
            for h in range(len(row2)): #1024段階の値を説明書通り色に変換
                for w in range(len(row2[h])):
                    m = row2[h][w]
                    
               
                    img[w, h] = [(m & 0b1111)<<4, ((m >>4)& 0b111)<<5, ((m>>7) & 0b111)<<5] 
            
            # セカンドディスプレイにフルスクリーンで表示　(メインがフルHDの時)
            img2 = img.tolist()
            mat = img.flatten()
            MakeWindow(mat)
            

            myinst.write("INITiate")
            time.sleep(ti)
            
            myinst.write("MEASure")
            power.append(myinst.query("FETCh?"))
            text_entry.SetLabel(str(r)+'のデータを取得しました')
            time.sleep(0.5)
            r = r+d
            c = c+1

            if r >= stop:
                img = np.zeros((height, length, 3), np.uint8) #画像の種

                row2 = MakeColumn(stop)
                for h in range(len(row2)): #1024段階の値を説明書通り色に変換
                    for w in range(len(row2[h])):
                        m = row2[h][w]
                        
                   
                        img[w, h] = [(m & 0b1111)<<4, ((m >>4)& 0b111)<<5, ((m>>7) & 0b111)<<5] 
                
                # セカンドディスプレイにフルスクリーンで表示　(メインがフルHDの時)
                img2 = img.tolist()
                mat = img.flatten()
                MakeWindow(mat)
                

                
                myinst.write("INITiate")
                time.sleep(ti)
                myinst.write("MEASure")
                power.append(myinst.query("FETCh?"))
                text_entry.SetLabel(str(stop)+'のデータを取得しました')
                time.sleep(0.5)
                break

        dt_now = datetime.datetime.now()
        os.chdir(cdir)
        lpower = [float(s) for s in power]
        nppower = np.array(lpower)
        np.savetxt(str(dt_now.year)+str(dt_now.month)+str(dt_now.day)+str(dt_now.hour)+str(dt_now.minute)+str(dt_now.second)+'.csv',nppower.T,delimiter=',')
        text_entry.SetLabel(str(dt_now.year)+str(dt_now.month)+str(dt_now.day)+str(dt_now.hour)+str(dt_now.minute)+str(dt_now.second)+'.csv を保存しました')

        text_entry.SetLabel('データの取得が完了しました') 
        winsound.Beep(2500, 1000)


class App(wx.Frame):
    """ GUI """
    def __init__(self, parent, id, title):
        def click_button_1(event):    #ボタン１がクリックされた時のイベント
            global start,stop,d,ti,cou,wlength
            start = int(text_1.GetValue())
            stop = int(text_2.GetValue())
            d = int(text_3.GetValue())
            ti = int(text_4.GetValue())
            cou = int(text_5.GetValue())
            wlength = str(text_6.GetValue())
            th = TestThread()
            th.setDaemon(True)
            th.start() 
            
        def click_button_2(event):    #ボタン１がクリックされた時のイベント
        
            global start,stop,d,ti,wlength
            start = int(text_1.GetValue())
            stop = int(text_2.GetValue())
            d = int(text_3.GetValue())
            ti = int(text_4.GetValue())
            wlength = str(text_6.GetValue())
            th = TestThread2()
            th.setDaemon(True)
            th.start() 

            
        
        
        
        wx.Frame.__init__(self, parent, id, title, size=(600, 600), style=wx.DEFAULT_FRAME_STYLE)

        # パネル
        p = wx.Panel(self, wx.ID_ANY)

        label = wx.StaticText(p, wx.ID_ANY, 'パワー変化\n一つのマスクパターンに対して変化時間×回数の時間がかかる\nセカンドディスプレイを1920*1080で125％の倍率で表示', style=wx.SIMPLE_BORDER | wx.TE_CENTER)
        label.SetBackgroundColour("#e0ffff")
        #s_text = wx.StaticText(p,wx.ID_ANY,'ポート名(このままでOK)')
        s_text_t = wx.StaticText(p,wx.ID_ANY,'開始深さ') #固定文
        s_text_x = wx.StaticText(p,wx.ID_ANY,'終了深さ')
        s_text_y = wx.StaticText(p,wx.ID_ANY,'変化ステップ')
        s_text_z = wx.StaticText(p,wx.ID_ANY,'変化時間*(s)')
        s_text_a = wx.StaticText(p,wx.ID_ANY,'クリック回数*')
        s_text_b = wx.StaticText(p,wx.ID_ANY,'パワーメータ波長')
        # text_x = wx.StaticText(p,wx.ID_ANY,'x軸単位')
        # text_y = wx.StaticText(p,wx.ID_ANY,'y軸単位')
        # text_m = wx.StaticText(p,wx.ID_ANY,'マーカーサイズ')
        
        
     

        # テキスト入力ウィジット
        global text_entry
        text_entry = wx.TextCtrl(p, wx.ID_ANY)
        #text = wx.TextCtrl(p, wx.ID_ANY,'USB0::0x1313::0x8072::P2009036::INSTR') 
        text_1 = wx.TextCtrl(p, wx.ID_ANY,'400') 
        text_2 = wx.TextCtrl(p, wx.ID_ANY,'800')
        text_3 = wx.TextCtrl(p, wx.ID_ANY,'10')
        text_4 = wx.TextCtrl(p, wx.ID_ANY,'2')
        text_5 = wx.TextCtrl(p, wx.ID_ANY,'20')
        text_6 = wx.TextCtrl(p, wx.ID_ANY,'785')
        text_entry.Disable()
        
        #スピンロール
        # spin1 = wx.SpinCtrlDouble(p, wx.ID_ANY, value="10", min=2, max=100)
        # spin2 = wx.SpinCtrlDouble(p, wx.ID_ANY, value="10", min=2, max=100)
        # spin3 = wx.SpinCtrlDouble(p, wx.ID_ANY, value="10", min=2, max=100)
        # spin4 = wx.SpinCtrlDouble(p, wx.ID_ANY, value="10", min=2, max=100)
        # spin5 = wx.SpinCtrlDOuble(p, wx.ID_ANY, value="10", min=2, max=100)
        
        
        
        #チェックボックス
        #checkbox_1 = wx.CheckBox(p, wx.ID_ANY, 'x軸とy軸を入れ替える')
        # checkbox_2 = wx.CheckBox(p, wx.ID_ANY,'凡例を付ける')
        # checkbox_3 = wx.CheckBox(p, wx.ID_ANY,'x軸を数字列にする')
        # checkbox_4 = wx.CheckBox(p, wx.ID_ANY,'x軸を指数表記')
        # checkbox_5 = wx.CheckBox(p, wx.ID_ANY,'y軸を指数表記')
        # checkbox_6 = wx.CheckBox(p, wx.ID_ANY,'FWHM')
        
        # checkbox_4.SetValue(True)
        # checkbox_5.SetValue(True)
        
        
        #コンボボックス
        # element_array = ('線と点','線','点（プロットのみ対応）','近似直線（プロットのみ対応）','スプライン補間（プロットのみ対応）','近似直線（エラーバー付き）','点（エラーバー付き）')
        # combobox_1 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array, style = wx.CB_READONLY)
        
        # element_array2 = ('T','G','M','k','1','c','m','u','n','p','f')
        # combobox_2 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array2, style = wx.CB_READONLY)
       
        # combobox_3 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array2, style = wx.CB_READONLY)
        
        # combobox_1.SetSelection(0)
        # combobox_2.SetSelection(4)
        # combobox_3.SetSelection(4)

       
        
        #ボタン
        button_1 = wx.Button(p,wx.ID_ANY,'クリック')
        button_2 = wx.Button(p,wx.ID_ANY,'パワーメータ')
        # button_3 = wx.Button(p,wx.ID_ANY,'ｙ軸成分のデータ')
        # button_4 = wx.Button(p,wx.ID_ANY,'ｙ軸成分のデータ(dBm用)')
        
        button_1.Bind(wx.EVT_BUTTON,click_button_1)
        button_2.Bind(wx.EVT_BUTTON,click_button_2)
        # button_3.Bind(wx.EVT_BUTTON,click_button_3)
        # button_4.Bind(wx.EVT_BUTTON,click_button_4)

        
        #スライダー
        
        # slider = wx.Slider(p, style=wx.SL_AUTOTICKS|wx.SL_LABELS)
        # slider.SetTickFreq(1)
        # slider.SetMin(0)
        # slider.SetMax(10)
        # slider.SetValue(5)

        # レイアウト
        layout = wx.BoxSizer(wx.VERTICAL)
        #sizer1= wx.BoxSizer(wx.HORIZONTAL)
        # sizer2= wx.BoxSizer(wx.HORIZONTAL)
        # sizer3 = wx.BoxSizer(wx.HORIZONTAL)
        # sizer4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer_txy = wx.FlexGridSizer(6,2,(0,0))
        layout.Add(label, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        #layout.Add(sizer1, flag =wx.GROW, border=10)
        # layout.Add(sizer2, flag =wx.GROW, border=10)
        layout.Add(sizer_txy,flag=wx.EXPAND | wx.ALL, border=10)
        # layout.Add(sizer3, flag =wx.GROW, border=10)
        layout.Add(text_entry, flag=wx.EXPAND | wx.ALL, border=10)
        # layout.Add(sizer4)
        layout.Add(sizer)
        
        # sizer1.Add(combobox_1, flag =wx.GROW| wx.LEFT, border=10)
        #sizer1.Add(checkbox_1, flag =wx.GROW| wx.LEFT, border=10)
        # sizer1.Add(checkbox_2, flag =wx.GROW| wx.LEFT, border=10)
        # sizer1.Add(checkbox_6, flag = wx.EXPAND| wx.LEFT, border = 10)
        
        # sizer2.Add(checkbox_3, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        # sizer2.Add(checkbox_4, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        # sizer2.Add(checkbox_5, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        
        #sizer_txy.Add(s_text, flag=wx.EXPAND | wx.ALL, border=10)
        #sizer_txy.Add(text, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_t, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_1, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_x, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_2, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_y, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_3, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_z, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_4, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_a, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_5, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_b, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_6, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.AddGrowableCol(1) #グリッドサイズ変更
        
        # sizer3.Add(text_x, flag =wx.GROW| wx.LEFT, border=10)
        # sizer3.Add(combobox_2, flag =wx.GROW| wx.LEFT, border=10)
        # sizer3.Add(text_y, flag =wx.GROW| wx.LEFT, border=10)
        # sizer3.Add(combobox_3, flag =wx.GROW| wx.LEFT, border=10)
        
        # sizer4.Add(text_m,1, flag = wx.ALL, border=10)
        # sizer4.Add(slider ,4)
        
        sizer.Add(button_1, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        sizer.Add(button_2, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        # sizer.Add(button_3, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        # sizer.Add(button_4, flag =wx.GROW| wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        p.SetSizer(layout)
        
        self.Centre()

        self.Show()
        
        def OnTimer(self, event): #時間経過のイベント
            rm = visa.ResourceManager()
            lis =rm.list_resources()
            ss_text.SetLabel('ポートリスト:'+str(lis))
    
        

app = wx.App()
App(None, -1, 'lcos_power')
app.MainLoop()