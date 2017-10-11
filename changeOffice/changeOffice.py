# -*- coding: utf-8 -*-
import os
import win32com.client as win32
import win32com
class Change():
    def __init__(self,dataDir):
        self.root=dataDir
    def change_singleDoc(self,targetFile):
        if targetFile.endswith('.doc'):
            print ("converting",targetFile)
            targetFile = os.path.abspath(targetFile)
            newname = targetFile.replace('.doc', '.docx')
            # wrd = win32.Dispatch("Word.Application")
            wrd = win32.gencache.EnsureDispatch('Word.Application')
            wrd.Visible = False
            # wrd.Application.DisplayAlerts = False
            wb = wrd.Documents.Open(targetFile)
            wb.SaveAs(newname, FileFormat=12)
            wb.Close()
            wrd.Quit()
            os.remove(targetFile)
    def change_singleXls(self,targetFile):
        if targetFile.endswith('.xls'):
            print ("converting",targetFile)
            newname = targetFile.replace('.xls', '.xlsx')
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Application.DisplayAlerts = False
            wb = excel.Workbooks.Open(targetFile)
            wb.SaveAs(newname, FileFormat=51)
            excel.Application.Quit()
            os.remove(targetFile)
    def change_signlePpt(self,file):
        if file.endswith('.ppt'):
            print ("converting",file)
            newname = file.replace('.ppt', '.pptx')
            ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
            pres = ppt.Presentations.Open(file,True,False,False)
            pres.SaveAs(newname)
            pres.Close()
            ppt.Quit()
            os.remove(file)
    def change_sigleEt(self,targetFile):
        if targetFile.endswith('.et'):
            print ("converting",targetFile)
            newname = targetFile.replace('.xls', '.xlsx')
            os.rename(targetFile,targetFile.replace('.et','.xls'))
    def get_allPath(self):
        files = []
        for root, dirs, _files in os.walk(self.root):
            for file in _files:
                files.append(os.path.abspath(os.path.join(root, file)))
        return files
    def xls2xlsx(self):
        for f in self.get_allPath():
            self.change_singleXls(f)
    def et2xls(self):
        for f in self.get_allPath():
            self.change_sigleEt(f)
    def ppt2pptx(self):
        for f in self.get_allPath():
            self.change_signlePpt(f)
    def doc2docx(self):
        for f in self.get_allPath():
            self.change_singleDoc(f)










