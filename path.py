import fitz  # pip install PyMuPDF
import os
from IPython.display import Image
import pandas as pd
from os import path 
from openpyxl import load_workbook 
from pathlib import Path

from shutil import copy
import xlwt
import xlrd

# 批量获取pdf 指定区域内容 写入exls
# mac os 使用python3 运行
#    python3 -m pip install install pandas==1.2.5
#    python3 -m pip install PyMuPDF
#    python3 -m pip install Image
#    python3 -m pip install IPython
#    python3 -m pip install pandas
#    python3 -m pip install openpyxl
#    python3 -m pip install ExcelWriter
#    python3 -m pip install xlwt
#    python3 -m pip install xlrd


# 批量获取的结果，一次写入，追加写入异常
result = []
#老模板数据
def getOldPDF(url):

    pdf_path = url
    if not os.path.exists("imgs"):
        os.mkdir("imgs")

    
    with fitz.open(pdf_path) as pdfDoc:
        for i in range(1):
            page_num = i+1
            print("--------------", page_num, "--------------")
            file = Path(pdf_path).stem
            fileName = file +'.pdf'
            page = pdfDoc[i]
            mat = fitz.Matrix(1.3, 1.3)  
            rect = page.rect

            a_text = getImageData(page,file,0.25*rect.width, 0.18*rect.height,
                              rect.width*0.5, 170)
            # b_text = getImageData(page,file,0.22*rect.width, 0.2*rect.height,
            #                  rect.width*0.5, 200)
            # c_text = getImageData(page,file,0.69*rect.width, 0.1*rect.height,
            #                  rect.width*1, 110)
            # d_text = getImageData(page,file,0.7*rect.width, 0.125*rect.height,
            #                  rect.width*1, 130)

            result.append([a_text,fileName])

#截取图片数据
def getImageData(page,fileName,x,y,width,height):
    mat = fitz.Matrix(1.5, 1.5)  
    clip = fitz.Rect(x, y, width, height)  
    pix = page.getPixmap(matrix=mat.preRotate(0),
                         alpha=False, clip=clip)  
    #pix.writeImage(f"imgs/{fileName}_a.png")
    img1 = pix.getImageData()
    #display(Image(img1))
    a_text = page.getText(clip=clip)
    print(a_text)
    return a_text

#有电子签章
def getNewPDF(url,tag):

    pdf_path = url
    if not os.path.exists("imgs"):
        os.mkdir("imgs")

    with fitz.open(pdf_path) as pdfDoc:
        for i in range(1):
            page_num = i+1
            print("---------new-----", page_num, "--------------")
            file = Path(pdf_path).stem
            fileName = file +'.pdf'
            page = pdfDoc[i]
            mat = fitz.Matrix(1.5, 1.5)  
            rect = page.rect

            #获取借款人姓名
           
            a_text = getImageData(page,file,0.18*rect.width, 0.09*rect.height,
                             rect.width*0.5, 95)
            """
            #获取借款人身份证
            b_text = getImageData(page,file,0.15*rect.width, 0.1*rect.height,
                             rect.width*0.5, 110)
            #获取协议编号
            c_text = getImageData(page,file,0.6*rect.width, 0.04*rect.height,
                             rect.width*1, 55)
            #获取广金服编号
            d_text = getImageData(page,file,0.6*rect.width, 0.06*rect.height,
                             rect.width*1, 80)
            if (tag == 2) :
                 d_text = getImageData(page,file,0.75*rect.width, 0.06*rect.height,
                             rect.width*1, 80)

            #pix = page.getPixmap(matrix=mat.preRotate(0),
             #                    alpha=False, clip=clip)  
            # 储存截取得图片
            # pix.writeImage(f"imgs/{fileName}_d.png")
            #img4 = pix.getImageData()
            # display(Image(img1))
            #d_text = page.getText(clip=clip)
            #tmp = pd.DataFrame(d_text.splitlines(), columns=["a"])
            #tmp["b"] = (tmp.a.str[:2]).astype("category")
            #tmp.b.cat.set_categories(
            #    ['An', 're', 'vi', 'Ma', 'co', 'VC', 'ES'], inplace=True)
            #tmp.sort_values('b', inplace=True)
            #d_text = ''.join(tmp.a.to_list())
            #print(d_text)

            if '编号' in d_text:
                 getNewPDF(url,2)
            else :
                result.append((a_text, b_text,c_text,d_text,fileName))
            """
            result.append([a_text,fileName])
# 监测elxs 是否存在
def createElxs(wb_name):
    try:
        with xlrd.open_workbook(wb_name) as wb:
            sh1 = wb.sheet_by_index(0)
            for sh in wb.sheets():
                for r in range(sh.nrows):
                    # 输出指定行
                    print(sh.row(r))
    except FileNotFoundError:
        # 创建工作簿
        wb = xlwt.Workbook(wb_name)

#搜索文件函数
def scaner_file (url , elxsName):
        # createElxs(elxsName)
        #遍历当前路径下所有文件
        file  = os.listdir(url)
        for f in file:
            #字符串拼接
            real_url = path.join (url , f)
            #打印出来
            print(real_url)
            file = Path(real_url).stem
            print(file)
            if '.pdf' in real_url:
                ll = file[3:11]
                if (int(ll) < 20180127):
                    getOldPDF(real_url)
                else :
                    getNewPDF(real_url,1)
            # break
        print(result)
        """
        # 写入数据到xlsx
        data = pd.DataFrame(result, columns=["姓名","身份证号","协议编号","广金所编号","文件名"])
        writer = pd.ExcelWriter(elxsName,mode='a', engine='openpyxl',if_sheet_exists='new')
        data.to_excel(writer, sheet_name='sheet1')
        writer.save()
        writer.close()
        """
        for ll in result:
            name = "./pdf/" + ll[0].replace('\n', '')
            pdf = ll[1]
            if not os.path.exists(name):
                os.mkdir(name)
            print("---- " + name + " --- " + pdf)
            copy("./all/"+pdf ,name)

#调用自定义函数
scaner_file("./all" , 'result.xlsx') 

