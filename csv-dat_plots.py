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
from scipy import interpolate
from scipy.optimize import curve_fit
import codecs

# import numpy.random.common #exe化時のエラー回避用
# import numpy.random.bounded_integers
# import numpy.random.entropy


#fp = FontProperties(fname =r'C:\Users\yamagishi\Desktop\python\font\ipaexg.ttf') #日本語フォントの位置
dirname ='' #ファイル未選択時の判別用の初期化

def nonlinear_fit(x,a,b,c):
    return  a*np.sin(b*(x-c))**2



class ChildFrame(wx.Dialog): #凡例用ダイヤログ
    def __init__(self,parent):
    
        def click_button_1(event): #凡例の取得
            global legend
            legend =self.text_entry.GetValue()
            if legend =='': 
                legend = '-'
            self.Destroy()
        
            
           
        wx.Dialog.__init__(self, parent, -1, '凡例', size=(300, 150), style=wx.DEFAULT_FRAME_STYLE  )
    
        
        self.parent =parent
        # パネル
        
        p = wx.Panel(self, wx.ID_ANY)

        text = wx.StaticText(p,wx.ID_ANY,os.path.basename(file[c])+'の凡例') #固定文
        self.text_entry = wx.TextCtrl(p, wx.ID_ANY)

        
       
        
        #ボタン
        button_1 = wx.Button(p,wx.ID_ANY,'OK')
        
        button_1.Bind(wx.EVT_BUTTON,click_button_1)
        


        # レイアウト
        layout2 = wx.BoxSizer(wx.VERTICAL)
        sizer1= wx.BoxSizer(wx.HORIZONTAL)
        layout2.Add(text, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        layout2.Add(self.text_entry, flag=wx.EXPAND | wx.ALL, border=10)
        layout2.Add(sizer1, flag =wx.GROW, border=10)
        
        sizer1.Add(button_1, flag =wx.GROW, border=10)
        
        p.SetSizer(layout2)
  
        
        


class FileDropTarget(wx.FileDropTarget):
    """ Drag & Drop Class """
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, files):
        global file
        file = files
        self.window.text_entry.SetLabel('{0:}'.format(len(file))+'個のファイル  最初のファイル名'+os.path.basename(file[0])) #ファイル絶対パスの取得
        
        return 0


class App(wx.Frame):
    """ GUI """
    def __init__(self, parent, id, title):
        mk = [ "o", "s","*","x", "v", "^","o", "s","*","x", "v", "^"]
        def click_button_1(event):    #ボタン１がクリックされた時のイベント
            global c #複数ファイルプロット用カウンタ
            c = 0

            cmap = plt.get_cmap("tab10")
            fig, ax=plt.subplots(figsize=(5*1.618,5))
            
            if 'file' in globals():
                pass
            else:
                wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                plt.close(fig)
                return 0
            for fname in file:
                
                dirname = os.path.dirname(fname)   #ファイルの位置とファイル名に分解   
                basename = os.path.basename(fname)
                if dirname =='' or file[0] == 'None':
                    wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                    plt.close(fig)
                    return 0
      
                os.chdir(dirname) #ファイルのある位置に移動
                ext = basename[-3:] #ファイル名の下から3文字の取得
                if not (ext == 'csv' or ext == 'DAT'or ext == 'SPE'): #拡張子がcsvであるかの判定
                    wx.MessageBox('対応していない拡張子のファイルが含まれています','error') 
                    plt.close(fig)
                    return 0
                    
                else:
                    
                    
                    if ext == 'csv':
                        f = open(basename) 
                        reader = csv.reader(f)
                        l =[row for row in reader]
                        
                    elif ext == 'DAT': #datファイルはタブ区切り
                        f = open(basename) 
                        reader = csv.reader(f,delimiter='\t')
                        l =[row for row in reader]
                    elif ext == 'SPE':
                        with codecs.open(basename, 'r', 'utf-8', 'ignore') as f:
                            count = 0
                            l = []
                            for line in f:
                                line = line.strip()
                                line = line.split(',')
                                l.append(line)
                                count = count+1
                                if count == 510:
                                    break
                            l = l[10:]  
                        
                    
                    

                    if len(l)==1 and l[0] ==[]: #ファイルの配列がない場合
                        plt.close(fig)
                        wx.MessageBox('空のファイルが含まれています','error') 
                        return 0
                    elif len(l[0])== 5 or len(l[0])== 2 or len(l[0]) == 3:    
                        title = text_1.GetValue()
                        x_label = text_2.GetValue()
                        y_label = text_3.GetValue()
                        if len(l[0])== 5:
                            l_x = [x[3] for x in l] #リスト列に変換
                            l_y = [y[4] for y in l] 
                            
                        elif len(l[0])== 2:
                            l_y = []
                            for y in l:  #DATファイルの２行目がないところを補てんしてからリスト列に変換
                                if len(y) == 1:
                                    y = y+['None']
                                l_y = l_y+[y[1]]
                                
                            l_x = [x[0] for x in l] #リスト列に変換
                            
                            
                            l_ch = [i for i, x in enumerate(l_x) if x == '!'] #スペアナ用リスト変換
                            if len(l_ch) == 3:
                                pt2 = l_ch[1]
                                pt3 = l_ch[2]
                                del l_x[pt3]
                                del l_y[pt3]
                                del l_x[:pt2+2]
                                del l_y[:pt2+2]
                                if checkbox_3.GetValue():  #x軸の作成
                                    l_x = list(range(len(l_x)))
                        
                        elif len(l[0]) == 3:
                            l_x = [x[0] for x in l] #リスト列に変換
                            l_y = [y[1] for y in l]
                            try:
                                if combobox_1.GetSelection() == 5 or combobox_1.GetSelection() == 6 or combobox_1.GetSelection() == 7:
                                    l_yerr = [y[2] for y in l]
                            
                            except:
                                wx.MessageBox('３列目に標準偏差のデータが必要です','error')
                                plt.close(fig)
                                return 0   
                        
                        try:
                            li_x = [float(s) for s in l_x] #文字列を数字列に変換
                            li_y = [float(s) for s in l_y]
    
                            if combobox_1.GetSelection() == 5 or combobox_1.GetSelection() == 6 or combobox_1.GetSelection() == 7:
                                li_yerr = [float(s) for s in l_yerr]
                            
                        except ValueError:
                            wx.MessageBox('csv,datデータに除去できない文字が含まれています','error')
                            plt.close(fig)
                            return 0
                       
                        if checkbox_2.GetValue(): #凡例用ダイヤログの生成
                            childFrame = ChildFrame(self)
                            childID = childFrame.ShowModal()
                            
                        if 'legend' in globals(): #凡例が存在しない場合
                            pass
                        else:
                            global legend
                            legend ='-'
                            
                        t = combobox_2.GetSelection()
                        t2 = combobox_3.GetSelection()
                        
                        
                        for i in range(len(li_x)): #グラフ単位変換
                            x = li_x[i]
                            if  t == 0: 
                                x=x*10**(-12)
                            elif t == 1:
                                x=x*10**(-9)
                            elif t == 2:
                                x=x*10**(-6)
                            elif t == 3:
                                x=x*10**(-3)
                            elif t == 4:
                                x=x*10**(0)
                            elif t == 5:
                                x=x*10**(2)
                            elif t == 6:
                                x=x*10**(3)
                            elif t == 7:
                                x=x*10**(6)
                            elif t == 8:
                                x=x*10**(9)
                            elif t == 9:
                                x=x*10**(12)
                            li_x[i] = x
                            
                        
                        for i in range(len(li_y)):
                            y = li_y[i]
                            if  t2 == 0: 
                                y=y*10**(-12)
                            elif t2 == 1:
                                y=y*10**(-9)
                            elif t2 == 2:
                                y=y*10**(-6)
                            elif t2 == 3:
                                y=y*10**(-3)
                            elif t2 == 4:
                                y=y*10**(0)
                            elif t2 == 5:
                                y=y*10**(2)
                            elif t2 == 6:
                                y=y*10**(3)
                            elif t2 == 7:
                                y=y*10**(6)
                            elif t2 == 8:
                                y=y*10**(9)
                            elif t2 == 9:
                                y=y*10**(12)
                            
                            li_y[i] = y
                            
                            
                        if checkbox_6.GetValue():
                            M = max(li_y)
                            in_M = li_y.index(M)
                            m = (mean(li_y[:int(len(li_y)/8)])+mean(li_y[int(len(li_y)/8*7):]))/2
                            h = (M-m)/2
                            d_y = [abs(i-m-h) for i in li_y]
                            d_y1 = d_y[:in_M ]
                            d_y2 = d_y[in_M :]
                            m1 = d_y1.index(min(d_y1))
                            m2 = d_y2.index(min(d_y2))
                            list_0 = [m]*len(li_x)
                            list_h = [m+h]*len(li_x)
                            fwhm = abs(li_x[m2+in_M ]-li_x[m1])
                            plt.plot(li_x,list_0,color=cmap(c),label = '0',linestyle=':')
                            plt.plot(li_x,list_h,color=cmap(c),label = '1/2',linestyle=':')
                            wx.MessageBox('ファイル名:'+basename+'\n'+'最大値:　　{0: .3e}'.format(M)+'\n'+'基準値:　　{0: .3e}'.format(m)+'\n'+'半値:　　　{0: .3e}'.format(m+h)+'\n'+'半値全幅:　{0: .3e}'.format(fwhm)+'\n')
                            
                        
                        
                        if title !='':
                            fig.canvas.set_window_title(title)
                        
                        plt.title(title) #plt.title(title,fontproperties = fp) で日本語対応
                        plt.xlabel(x_label)
                        plt.ylabel(y_label)
                        
                        plt.rcParams['font.family'] = 'Times New Roman' #全体のフォントを設定
                        plt.rcParams["mathtext.fontset"] = "stix" 
                        plt.rcParams["font.size"] = 12
                        plt.rcParams['axes.linewidth'] = 1.0# 軸の線幅edge linewidth。囲みの太さ
                        #plt.rcParams['axes.grid'] = True
                        plt.rcParams['xtick.direction'] = 'in'#x軸の目盛線
                        plt.rcParams['ytick.direction'] = 'in'#y軸の目盛線
                        
                        m_size =slider.GetValue()
                        
                        if  combobox_1.GetSelection() == 0: #点と折れ線
                            plt.plot(li_x,li_y,label = legend,marker = mk[c],color=cmap(c),linestyle=':', linewidth = 1.0,markersize=m_size)
                        elif combobox_1.GetSelection() == 1: #折れ線
                            plt.plot(li_x,li_y,label = legend,color=cmap(c))
                        elif combobox_1.GetSelection() == 2:#点
                            plt.plot(li_x,li_y,label = legend,marker = mk[c],color=cmap(c),linestyle='None', linewidth = 1.0,markersize=m_size)
                        elif combobox_1.GetSelection() == 3:#直線近似
                            ab = np.polyfit(li_x,li_y,1)
                            fh = np.poly1d(ab)(li_x)
                            plt.plot(li_x,fh,color=cmap(c),label = legend,linestyle=':')
                            plt.plot(li_x,li_y,marker = mk[c],color=cmap(c),linestyle='None', linewidth = 1.0,markersize=m_size)
                            wx.MessageBox('y={0:.3e}'.format(ab[0])+'x'+'{0:+.3e}'.format(ab[1]),'近似直線')
                        elif combobox_1.GetSelection() == 4: #スプライン補間による曲線近似
                            try:
                                tck,u = interpolate.splprep([li_x,li_y],k=3,s=0) 
                                u = np.linspace(0,1,num=100,endpoint=True) 
                                spline = interpolate.splev(u,tck)
                                plt.plot(spline[0],spline[1],label = legend,color=cmap(c))
                                plt.plot(li_x,li_y,color=cmap(c),marker = mk[c],linestyle='None', linewidth = 1.0,markersize=m_size)
                            except:
                                wx.MessageBox('このグラフのx成分は単調に増加していなければなりません','error')
                                plt.close(fig)
                                return 0
                        
                        elif combobox_1.GetSelection() == 5: #近似直線とエラーバー
                            ab = np.polyfit(li_x,li_y,1)
                            fh = np.poly1d(ab)(li_x)
                            plt.plot(li_x,fh,color=cmap(c),label = legend,linestyle=':')
                            plt.errorbar(li_x, li_y, yerr = li_yerr,fmt='o',capsize = 3, markersize=m_size, color=cmap(c))
                            wx.MessageBox('y={0:.3e}'.format(ab[0])+'x'+'{0:+.3e}'.format(ab[1]),'近似直線')
                        
                        elif combobox_1.GetSelection() == 6: #エラーバーを用いたプロット
                            plt.errorbar(li_x, li_y, yerr = li_yerr,fmt='o',capsize = 3, markersize=m_size, color=cmap(c))
                        
                        elif combobox_1.GetSelection() == 7: #sin^2フィッティング
                            lst = list(range(1,20000,1))
                            param, cov = curve_fit(nonlinear_fit, li_x, li_y,p0=(10000,0.00001,0),maxfev = 1000000)
                            print(param)
                            list_y = []
                            for num in lst:
                                list_y.append(nonlinear_fit(num,param[0],param[1],param[2]))
                            plt.errorbar(li_x, li_y, yerr = li_yerr,fmt='o',capsize = 3, markersize=m_size, color=cmap(c))
                            plt.plot(lst,list_y,label = legend,color=cmap(c))

                         
                                
                        if checkbox_5.GetValue():        
                            ax.yaxis.set_major_formatter(ptick.ScalarFormatter(useMathText=True)) #グラフ軸を指数表記に
                            ax.yaxis.offsetText.set_fontsize(12)
                            ax.ticklabel_format(style='sci',axis='y',scilimits=(0,0))
                        if checkbox_4.GetValue():
                            ax.xaxis.set_major_formatter(ptick.ScalarFormatter(useMathText=True))
                            ax.xaxis.offsetText.set_fontsize(12)
                            ax.ticklabel_format(style='sci',axis='x',scilimits=(0,0))
                        
                        if checkbox_1.GetValue(): #チェックボックス選択時に原点を端に固定する
                            left, right = ax.get_xlim() #グラフの範囲の取得
                            bottom, top = ax.get_ylim()
                            min_x = min(li_x)  #要素の最大値、最小値の取得
                            max_x = max(li_x)
                            min_y = min(li_y)
                            max_y = max(li_y)
                            if min_x >= 0 and max_x >= 0: #グラフの端を原点に設定
                                ax.set_xlim(0, right)
                            elif min_x <= 0 and max_x <= 0:
                                ax.set_xlim(left, 0)
                                
                            if min_y >= 0 and max_y >= 0:
                                ax.set_ylim(0, top)
                            elif min_y <= 0 and max_y <= 0:
                                ax.set_xlim(bottom, 0)       
                                
                        
                    
                    
                    else: #csvファイルが5列でない場合
                        wx.MessageBox('対応できないcsv,datファイルがあります（列数）','error')
                        plt.close(fig)
                        return 0
                c += 1
            if checkbox_2.GetValue(): #凡例の追加
                plt.legend()
            plt.show() 
                    
                    
        def click_button_2(event):    #ボタン2がクリックされた時のイベント
            global c #複数ファイルプロット用カウンタ
            c = 0
            cmap = plt.get_cmap("tab10")
            fig, ax=plt.subplots(figsize=(5*1.618,5))
            if 'file' in globals():
                pass
            else:
                wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                plt.close(fig)
                return 0
            for fname in file:
                
                dirname = os.path.dirname(fname)   #ファイルの位置とファイル名に分解   
                basename = os.path.basename(fname)
                if dirname =='':
                    wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                    plt.close(fig)
                    return 0
      
                os.chdir(dirname) #ファイルのある位置に移動
                ext = basename[-3:] #ファイル名の下から3文字の取得
                if not (ext == 'csv' or ext == 'DAT'): #拡張子がcsvであるかの判定
                    wx.MessageBox('対応していない拡張子のファイルが含まれています','error') 
                    plt.close(fig)
                    return 0
                    
                else:
                    
                    f = open(basename) 
                    if ext == 'csv':
                        reader = csv.reader(f)
                        
                    elif ext == 'DAT': #datファイルはタブ区切り
                        reader = csv.reader(f,delimiter='\t')
                        
                    l =[row for row in reader]

                    if len(l)==1 and l[0] ==[]: #ファイルの配列がない場合
                        wx.MessageBox('空のファイルが含まれています','error') 
                        plt.close(fig)
                        return 0
                    elif len(l[0])== 5 or len(l[0])== 2:    
                        title = text_1.GetValue()
                        x_label = text_2.GetValue()
                        y_label = text_3.GetValue()
                        if len(l[0])== 5:
                            l_x = [x[3] for x in l] #リスト列に変換
                            l_y = [y[4] for y in l] 
                            
                        elif len(l[0])== 2:
                            l_y = []
                            for y in l:  #DATファイルの２行目がないところを補てんしてからリスト列に変換
                                if len(y) == 1:
                                    y = y+['None']
                                l_y = l_y+[y[1]]
                                
                            l_x = [x[0] for x in l] #リスト列に変換
                            
                            
                            l_ch = [i for i, x in enumerate(l_x) if x == '!'] #スペアナ用リスト変換
                            if len(l_ch) == 3:
                                pt2 = l_ch[1]
                                pt3 = l_ch[2]
                                del l_x[pt3]
                                del l_y[pt3]
                                del l_x[:pt2+2]
                                del l_y[:pt2+2]
                                if checkbox_1.GetValue():  #x軸の作成
                                    l_x = list(range(len(l_x)))
                        try:
                            li_x = [float(s) for s in l_x] #文字列を数字列に変換
                            li_y = [float(s) for s in l_y]
                        except ValueError:
                            wx.MessageBox('csv,datデータに除去できない文字が含まれています','error')
                            plt.close(fig)
                            return 0
                       
                        if checkbox_2.GetValue(): #凡例用ダイヤログの生成
                            childFrame = ChildFrame(self)
                            childID = childFrame.ShowModal()
                            
                        if 'legend' in globals(): #凡例が存在しない場合
                            pass
                        else:
                            global legend
                            legend ='-'
                        
                        
                        
                        plt.title(title) #plt.title(title,fontproperties = fp) で日本語対応
                        plt.xlabel(x_label)
                        plt.ylabel(y_label)
                        
                        plt.rcParams['font.family'] = 'Times New Roman' #全体のフォントを設定
                        plt.rcParams["mathtext.fontset"] = "stix" 
                        plt.rcParams["font.size"] = 12
                        plt.rcParams['axes.linewidth'] = 1.0# 軸の線幅edge linewidth。囲みの太さ
                        #plt.rcParams['axes.grid'] = True
                        plt.rcParams['xtick.direction'] = 'in'#x軸の目盛線
                        plt.rcParams['ytick.direction'] = 'in'#y軸の目盛線
                        
                        
                        N = len(li_x) #フーリエ変換
                        h = int(N/2)
                        dt = li_x[h+1]-li_x[h]
                        freq = np.linspace(0, 1.0/dt, N)
                        
                        F = np.fft.fft(li_y)
                        F_abs = np.abs(F)
                        F_abs_amp = F_abs/N*2
                        F_abs_amp[0] = F_abs_amp[0]/2
                        
                        dialog = wx.MessageDialog(None,  '0 Hz(DC成分)を除きますか','グラフ修正', style=wx.YES_NO)
                        result = dialog.ShowModal()
                        if result == wx.ID_YES:  #DC成分除去
                            F_abs_amp = np.delete(F_abs_amp, 0)
                            
                        if combobox_1.GetSelection() == 0: 
                            plt.plot(freq[:int(N/2)+1],F_abs_amp[:int(N/2)+1],label =legend,linestyle=':', linewidth = 1.0, marker='x',markersize=5)
                        else:
                            plt.plot(freq[:int(N/2)+1],F_abs_amp[:int(N/2)+1],label =legend)
                                
                                
                        if checkbox_5.GetValue():        
                            ax.yaxis.set_major_formatter(ptick.ScalarFormatter(useMathText=True)) #グラフ軸を指数表記に
                            ax.yaxis.offsetText.set_fontsize(12)
                            ax.ticklabel_format(style='sci',axis='y',scilimits=(0,0))
                        if checkbox_4.GetValue():
                            ax.xaxis.set_major_formatter(ptick.ScalarFormatter(useMathText=True))
                            ax.xaxis.offsetText.set_fontsize(12)
                            ax.ticklabel_format(style='sci',axis='x',scilimits=(0,0))
                        
                        if checkbox_1.GetValue(): #チェックボックス選択時に原点を端に固定する
                            left, right = ax.get_xlim() #グラフの範囲の取得
                            bottom, top = ax.get_ylim()
                            ax.set_xlim(0, right)
                            ax.set_ylim(0, top)      
                       
                    
                    
                    else: #csvファイルが5列でない場合
                        wx.MessageBox('対応できないcsv,datファイルがあります（列数）','error') 
                        plt.close(fig)
                        return 0
         
            if checkbox_2.GetValue(): #凡例の追加
                plt.legend()        
            plt.show()
                    
           
        def click_button_3(event):    #ボタン3がクリックされた時のイベント
            if 'file' in globals():
                pass
            else:
                wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                return 0
            for fname in file:
                
                dirname = os.path.dirname(fname)   #ファイルの位置とファイル名に分解   
                basename = os.path.basename(fname)
                if dirname =='':
                    wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                    return 0
      
                os.chdir(dirname) #ファイルのある位置に移動
                ext = basename[-3:] #ファイル名の下から3文字の取得
                if not (ext == 'csv' or ext == 'DAT'): #拡張子がcsvであるかの判定
                    wx.MessageBox('対応していない拡張子のファイルが含まれています','error') 
                    return 0
                    
                else:
                    
                    f = open(basename) 
                    if ext == 'csv':
                        reader = csv.reader(f)
                        
                    elif ext == 'DAT': #datファイルはタブ区切り
                        reader = csv.reader(f,delimiter='\t')
                        
                    l =[row for row in reader]

                    if len(l)==1 and l[0] ==[]: #ファイルの配列がない場合
                        wx.MessageBox('空のファイルが含まれています','error') 
                        return 0
                    elif len(l[0])== 5 or len(l[0])== 2:    
                        title = text_1.GetValue()
                        x_label = text_2.GetValue()
                        y_label = text_3.GetValue()
                        if len(l[0])== 5:
                            l_x = [x[3] for x in l] #リスト列に変換
                            l_y = [y[4] for y in l] 
                            
                        elif len(l[0])== 2:
                            l_y = []
                            for y in l:  #DATファイルの２行目がないところを補てんしてからリスト列に変換
                                if len(y) == 1:
                                    y = y+['None']
                                l_y = l_y+[y[1]]
                                
                            l_x = [x[0] for x in l] #リスト列に変換
                            
                            
                            l_ch = [i for i, x in enumerate(l_x) if x == '!'] #スペアナ用リスト変換
                            if len(l_ch) == 3:
                                pt2 = l_ch[1]
                                pt3 = l_ch[2]
                                del l_x[pt3]
                                del l_y[pt3]
                                del l_x[:pt2+2]
                                del l_y[:pt2+2]

                        try:
                            li_x = [float(s) for s in l_x] #文字列を数字列に変換
                            li_y = [float(s) for s in l_y]
                        except ValueError:
                            wx.MessageBox('csv,datデータに除去できない文字が含まれています','error')
                            return 0
                       
                        
                        
                        
                        l = len(li_y)
                        ma = max(li_y)
                        mi = min(li_y)
                        m = mean(li_y)
                        med = median(li_y)
                        var = variance(li_y)
                        std = stdev(li_y)
                        wx.MessageBox('ファイル名:'+basename+'\n''データ数:　  　{0:}'.format(l)+'\n'+'最大値:　　{0: .3e}'.format(ma)+'\n'+'最小値:　　{0: .3e}'.format(mi)+'\n'+'平均:　　　{0: .3e}'.format(m)+'\n'+'中央値:　　{0: .3e}'.format(med)+'\n'+'分散:　　　{0: .3e}'.format(var)+'\n'+'標準偏差:　{0: .3e}'.format(std),'y軸成分のデータ')
 
                       
                    
                    
                    else: #csvファイルが5列でない場合
                        wx.MessageBox('対応できないcsv,datファイルがあります（列数）','error') 
                        return 0
            
            
        def click_button_4(event):    #ボタン3がクリックされた時のイベント
            if 'file' in globals():
                pass
            else:
                wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                return 0
            for fname in file:
                
                dirname = os.path.dirname(fname)   #ファイルの位置とファイル名に分解   
                basename = os.path.basename(fname)
                if dirname =='':
                    wx.MessageBox('ファイルを指定してください','error') #ファイルが指定されていないとき
                    return 0
      
                os.chdir(dirname) #ファイルのある位置に移動
                ext = basename[-3:] #ファイル名の下から3文字の取得
                if not (ext == 'csv' or ext == 'DAT'): #拡張子がcsvであるかの判定
                    wx.MessageBox('対応していない拡張子のファイルが含まれています','error') 

                    return 0
                    
                else:
                    
                    f = open(basename) 
                    if ext == 'csv':
                        reader = csv.reader(f)
                        
                    elif ext == 'DAT': #datファイルはタブ区切り
                        reader = csv.reader(f,delimiter='\t')
                        
                    l =[row for row in reader]

                    if len(l)==1 and l[0] ==[]: #ファイルの配列がない場合
                        wx.MessageBox('空のファイルが含まれています','error') 
                        return 0
                    elif len(l[0])== 5 or len(l[0])== 2:    
                        title = text_1.GetValue()
                        x_label = text_2.GetValue()
                        y_label = text_3.GetValue()
                        if len(l[0])== 5:
                            l_x = [x[3] for x in l] #リスト列に変換
                            l_y = [y[4] for y in l] 
                            
                        elif len(l[0])== 2:
                            l_y = []
                            for y in l:  #DATファイルの２行目がないところを補てんしてからリスト列に変換
                                if len(y) == 1:
                                    y = y+['None']
                                l_y = l_y+[y[1]]
                                
                            l_x = [x[0] for x in l] #リスト列に変換
                            
                            
                            l_ch = [i for i, x in enumerate(l_x) if x == '!'] #スペアナ用リスト変換
                            if len(l_ch) == 3:
                                pt2 = l_ch[1]
                                pt3 = l_ch[2]
                                del l_x[pt3]
                                del l_y[pt3]
                                del l_x[:pt2+2]
                                del l_y[:pt2+2]

                        try:
                            li_x = [float(s) for s in l_x] #文字列を数字列に変換
                            li_y = [float(s) for s in l_y]
                        except ValueError:
                            wx.MessageBox('csv,datデータに除去できない文字が含まれています','error')
               
                            return 0
                       
                        li_y = [10**(0.1*y) for y in li_y]
                        
                        
                        l = len(li_y)
                        ma = max(li_y)
                        mi = min(li_y)
                        m = mean(li_y)
                        med = median(li_y)
                        var = variance(li_y)
                        std = stdev(li_y)
                        wx.MessageBox('ファイル名:'+basename+'\n''データ数:　  　{0:}'.format(l)+'\n'+'最大値:　　{0: .3e}'.format(ma)+'\n'+'最小値:　　{0: .3e}'.format(mi)+'\n'+'平均:　　　{0: .3e}'.format(m)+'\n'+'中央値:　　{0: .3e}'.format(med)+'\n'+'分散:　　　{0: .3e}'.format(var)+'\n'+'標準偏差:　{0: .3e}'.format(std),'y軸成分のデータ')
 
                       
                    
                    
                    else: #csvファイルが5列でない場合
                        wx.MessageBox('対応できないcsv,datファイルがあります（列数）','error') 
                        return 0            
                   
                    
                   
                    
                
            
        
        
        
        wx.Frame.__init__(self, parent, id, title, size=(600, 600), style=wx.DEFAULT_FRAME_STYLE)

        # パネル
        p = wx.Panel(self, wx.ID_ANY)

        label = wx.StaticText(p, wx.ID_ANY, 'csvもしくはdatファイルをドロップしてください\n まとめてドロップすると同じグラフに重ねてプロットできます \n エラーバー付きのプロットは３列の整形したデータのみ有効です（3列目:標準偏差）\n LaTex表記が可能です\n FWHMはピークパルス内に十分なプロットがないと正確に出ません', style=wx.SIMPLE_BORDER | wx.TE_CENTER)
        label.SetBackgroundColour("#e0ffff")
        s_text_t = wx.StaticText(p,wx.ID_ANY,'タイトル') #固定文
        s_text_x = wx.StaticText(p,wx.ID_ANY,'x軸ラベル')
        s_text_y = wx.StaticText(p,wx.ID_ANY,'y軸ラベル')
        text_x = wx.StaticText(p,wx.ID_ANY,'x軸単位')
        text_y = wx.StaticText(p,wx.ID_ANY,'y軸単位')
        text_m = wx.StaticText(p,wx.ID_ANY,'マーカーサイズ')
        
        # 閉じるイベント
        self.Bind(wx.EVT_CLOSE, self.frame_close)
        
        # ドロップ対象の設定
        self.SetDropTarget(FileDropTarget(self))
     

        # テキスト入力ウィジット
        self.text_entry = wx.TextCtrl(p, wx.ID_ANY)
        text_1 = wx.TextCtrl(p, wx.ID_ANY) 
        text_2 = wx.TextCtrl(p, wx.ID_ANY)
        text_3 = wx.TextCtrl(p, wx.ID_ANY)
        self.text_entry.Disable()
        
        #チェックボックス
        checkbox_1 = wx.CheckBox(p, wx.ID_ANY, '原点を端にする')
        checkbox_2 = wx.CheckBox(p, wx.ID_ANY,'凡例を付ける')
        checkbox_3 = wx.CheckBox(p, wx.ID_ANY,'x軸を数字列にする')
        checkbox_4 = wx.CheckBox(p, wx.ID_ANY,'x軸を指数表記')
        checkbox_5 = wx.CheckBox(p, wx.ID_ANY,'y軸を指数表記')
        checkbox_6 = wx.CheckBox(p, wx.ID_ANY,'FWHM')
        
        checkbox_4.SetValue(True)
        checkbox_5.SetValue(True)
        
        
        #コンボボックス
        element_array = ('線と点','線','点（プロットのみ対応）','近似直線（プロットのみ対応）','スプライン補間（プロットのみ対応）','近似直線（エラーバー付き）','点（エラーバー付き）','sin^2フィッティング')
        combobox_1 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array, style = wx.CB_READONLY)
        
        element_array2 = ('T','G','M','k','1','c','m','u','n','p','f')
        combobox_2 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array2, style = wx.CB_READONLY)
       
        combobox_3 = wx.ComboBox(p, wx.ID_ANY , 'プロット表記の選択', choices = element_array2, style = wx.CB_READONLY)
        
        combobox_1.SetSelection(0)
        combobox_2.SetSelection(4)
        combobox_3.SetSelection(4)

       
        
        #ボタン
        button_1 = wx.Button(p,wx.ID_ANY,'プロット')
        button_2 = wx.Button(p,wx.ID_ANY,'フーリエ変換')
        button_3 = wx.Button(p,wx.ID_ANY,'ｙ軸成分のデータ')
        button_4 = wx.Button(p,wx.ID_ANY,'ｙ軸成分のデータ(dBm用)')
        
        button_1.Bind(wx.EVT_BUTTON,click_button_1)
        button_2.Bind(wx.EVT_BUTTON,click_button_2)
        button_3.Bind(wx.EVT_BUTTON,click_button_3)
        button_4.Bind(wx.EVT_BUTTON,click_button_4)
        
        #スライダー
        
        slider = wx.Slider(p, style=wx.SL_AUTOTICKS|wx.SL_LABELS)
        slider.SetTickFreq(1)
        slider.SetMin(0)
        slider.SetMax(10)
        slider.SetValue(5)

        # レイアウト
        layout = wx.BoxSizer(wx.VERTICAL)
        sizer1= wx.BoxSizer(wx.HORIZONTAL)
        sizer2= wx.BoxSizer(wx.HORIZONTAL)
        sizer3 = wx.BoxSizer(wx.HORIZONTAL)
        sizer4 = wx.BoxSizer(wx.HORIZONTAL)
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer_txy = wx.FlexGridSizer(3,2,(0,0))
        layout.Add(label, flag=wx.EXPAND | wx.ALL, border=10, proportion=1)
        layout.Add(sizer1, flag =wx.GROW, border=10)
        layout.Add(sizer2, flag =wx.GROW, border=10)
        layout.Add(sizer_txy,flag=wx.EXPAND | wx.ALL, border=10)
        layout.Add(sizer3, flag =wx.GROW, border=10)
        layout.Add(self.text_entry, flag=wx.EXPAND | wx.ALL, border=10)
        layout.Add(sizer4)
        layout.Add(sizer)
        
        sizer1.Add(combobox_1, flag =wx.GROW| wx.LEFT, border=10)
        sizer1.Add(checkbox_1, flag =wx.GROW| wx.LEFT, border=10)
        sizer1.Add(checkbox_2, flag =wx.GROW| wx.LEFT, border=10)
        sizer1.Add(checkbox_6, flag = wx.EXPAND| wx.LEFT, border = 10)
        
        sizer2.Add(checkbox_3, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        sizer2.Add(checkbox_4, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        sizer2.Add(checkbox_5, flag =wx.GROW| wx.LEFT|wx.TOP, border=10)
        
        sizer_txy.Add(s_text_t, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_1, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_x, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_2, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(s_text_y, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.Add(text_3, flag=wx.EXPAND | wx.ALL, border=10)
        sizer_txy.AddGrowableCol(1) #グリッドサイズ変更
        
        sizer3.Add(text_x, flag =wx.GROW| wx.LEFT, border=10)
        sizer3.Add(combobox_2, flag =wx.GROW| wx.LEFT, border=10)
        sizer3.Add(text_y, flag =wx.GROW| wx.LEFT, border=10)
        sizer3.Add(combobox_3, flag =wx.GROW| wx.LEFT, border=10)
        
        sizer4.Add(text_m,1, flag = wx.ALL, border=10)
        sizer4.Add(slider ,4)
        
        sizer.Add(button_1, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        sizer.Add(button_2, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        sizer.Add(button_3, flag =wx.GROW| wx.LEFT|wx.BOTTOM, border=10)
        sizer.Add(button_4, flag =wx.GROW| wx.LEFT|wx.RIGHT|wx.BOTTOM, border=10)
        p.SetSizer(layout)
        
        self.Centre()

        self.Show()
    def frame_close(self, event):
        """ 閉じたときに発生するイベント """
        
        plt.close('all')
        #event.Skip()
        self.Destroy()   
    
        

app = wx.App()
App(None, -1, 'csv_shape')
app.MainLoop()