from typing import ChainMap
import openpyxl as op
import pandas as pd
import os , sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog

def Exit():
    sys.exit()  

def filedialog_clicked():#iFilePath1
        fTyp = [("エクセルファイル", "*.xlsx")]
        iFile = os.path.abspath(os.path.dirname(__file__))
        #ファイルダイアログを表示
        iFilePath1 = filedialog.askopenfilename(filetype = fTyp, initialdir = iFile)
        entry0.set(iFilePath1)

def TKFILEPATH(name):
    """
    TKinterのfile名とpathをゲット
    """
    #初期化
    colcount = 1
    rowcount = 2
    maxcount = 5001
    sheetcount = 1
    cheak = True

    filePath0 = entry0.get()
    sheetname = name.get()
    if filePath0 == '':
        messagebox.showerror('エラー','エクセルファイルを選択してください')
    
    print(filePath0)


    def GETLIST(maxcount , cheak, sheetcount ,rowcount):
        """
        listに追加していく
        """
        print(sheetname)
        print("5000の上限",maxcount)
        print(sheetcount)
        print("現在のシートの最低",rowcount)

        #pdf用のlist作成
        SKUlist = list()
        ASINlist = list()
        NUMBERlist = list()
        PRICElist = list()
        CONDITIONlist = list()
        PRICElist = list()
        COSTlist = list()
        AKAJIlist = list()
        PRICETRACElist = list()
        LEADTIMElist = list()

        wb =op.load_workbook(filePath0,data_only=True)
        ws = wb.worksheets[0]

        MIN_COL = colcount
        MIN_ROW = rowcount
        MAX_COl = colcount
        MAX_ROW = maxcount

        for row in ws.iter_rows(min_col=MIN_COL, min_row=MIN_ROW, max_col=MAX_COl, max_row=MAX_ROW):
            for cell in row:
                sku = cell.value
                asin = cell.offset (0,1)
                number = cell.offset (0,3)
                price = cell.offset (0,4)
                cost = cell.offset (0,5)
                akaji = cell.offset (0,6)
                condition = cell.offset (0,8)
                pricetrace = cell.offset (0,10)
                leadtime = cell.offset (0,11)

                SKUlist.append(sku)
                ASINlist.append(asin.value)
                NUMBERlist.append(number.value)
                PRICElist.append(price.value)
                COSTlist.append(cost.value)
                AKAJIlist.append(akaji.value)
                CONDITIONlist.append(condition.value)
                PRICETRACElist.append(pricetrace.value)
                LEADTIMElist.append(leadtime.value)

        maxcount += 5000
        rowcount += 5000 
        info = {'SKU':SKUlist,'ASIN':ASINlist,'number':NUMBERlist,'price':PRICElist,'cost':COSTlist,'akaji':AKAJIlist,'condition':CONDITIONlist,'priceTrace':PRICETRACElist,'leadtime':LEADTIMElist}
        df = pd.DataFrame(info)
        df.to_csv( sheetname + str(sheetcount) + '.csv',index=False,encoding='utf-8')
        print("listの要素",len(SKUlist))

        for listindex in SKUlist:
            if listindex is None:
                cheak = False
                break
            else:
                cheak = True
                
        sheetcount+=1

        return maxcount,cheak,sheetcount,rowcount
    
    while True:
        print("while",cheak)
        if cheak == False:
            Exit()
        else:
            maxcount , cheak , sheetcount , rowcount = GETLIST(maxcount , cheak , sheetcount , rowcount)
            continue


root = Tk()
root.title("ファイル選択")
# Frame1の作成
frame1 = ttk.Frame(root, padding=5)
frame1.grid(row=0, column=0, sticky=E)
# 「ファイル参照」ラベルの作成
IFileLabel = ttk.Label(frame1, text="Excelを選択")
IFileLabel.pack(side='top')
# 「ファイル参照」エントリーの作成
entry0 = StringVar()
IFileEntry = ttk.Entry(frame1, textvariable=entry0, width=30)
IFileEntry.pack(fill = 'x',side=LEFT)
# 「ファイル参照」ボタンの作成
IFileButton = ttk.Button(frame1, text="参照", command=filedialog_clicked)
IFileButton.pack(side=LEFT)
##################################################################################################
framestr = ttk.Frame(root, padding=5)
framestr.grid(row=10, column=0, sticky=E)
# ラベルの作成
leLabel = ttk.Label(framestr, text="シートの名前を入力")
leLabel.pack(side='top')
# 「入力」エントリーの作成
#書き込む白い四角部分を作成
name = StringVar()
nameEntry = ttk.Entry(framestr, textvariable=name, width=43)
nameEntry.pack(fill = 'x',side=RIGHT)
###################################################################################################
# Frame3の作成
frame3 = ttk.Frame(root, padding=10)
frame3.grid(row=12,column=0,sticky=W+E)

# 実行ボタンの設置
start_button = ttk.Button(frame3, text="実行", command=lambda:TKFILEPATH(name))
start_button.pack(side=LEFT)

#実行ボタン
close_button = ttk.Button(frame3, text='閉じる', command = lambda:Exit())
close_button.pack(side=RIGHT)

root.mainloop()