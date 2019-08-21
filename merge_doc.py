# -*- coding: utf-8 -*-：
#    Author : wonder_zz
#    Time   : 2019/8/20  20:41
import os
import sys
import time
from docx import Document
from win32com import client as wc


def ReSaveDoc(path, filename):
    """
    将doc转换docx
    """
    time1 = time.time()
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(path + "\\" + filename)
    filename = filename.replace("doc", "docx")
    doc.SaveAs(path + "\\" + filename, 12, False, "",
               True, "", False, False, False, False)
    doc.Close()
    word.Quit()
    time2 = time.time()
    print("success change file " + filename + " in " + str(time2 - time1))


def ReSaveAllDoc(path):
    """
    保存全部文档
    """
    filelist = []
    dirs = os.listdir()
    for f in dirs:
        filelist.append(str(f))
    for file in filelist:
        if file.find(".doc") != -1:
            ReSaveDoc(path, file)
            print(file)
        else:
            continue


def combine_word_documents(path, files):
    """
    创建合并文件
    """
    merged_document = Document()
    for index, file in enumerate(files):
        sub_doc = Document(path + "\\" + file)

        # Don't add a page break if you've reached the last file.
        if index < len(files) - 1:
            sub_doc.add_page_break()

        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save(path + "\\" + '合并文件.docx')


def MergeDocx(path):
    """
    合并文件
    """
    files = []
    filelist = []
    dirs = os.listdir()
    for f in dirs:
        filelist.append(str(f))
    for file in filelist:
        if (file.find(".docx") != -1) and (file.find("~$") == -1):
            files.append(file)
    combine_word_documents(path, files)


if __name__ == '__main__':
    path = os.getcwd()
    ReSaveAllDoc(path)
    # 转换文件格式
    MergeDocx(path)
