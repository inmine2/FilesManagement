import docx,re
import win32com.client as wc


# def doSaveAas(oldfile):
#
#     if oldfile=='':
#         pass
#     else:
#         word = wc.Dispatch('Word.Application')
#         print('this')
#         doc = word.Documents.Open(oldfile)  # 目标路径下的文件
#         newfile = oldfile.replace('doc','docx')
#         doc.SaveAs(newfile, 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件
#         doc.Close()
#         word.Quit()
#         return newfile

class ReadDoc:
    def __init__(self,file):
        self.doc=docx.Document(file)
        self.texts=''
        for i in self.doc.paragraphs:
            self.texts = self.texts+i.text+'\n'
        print(self.texts)

        try:
            self.L2 = re.findall("合同号：(.*?)\n",self.texts)[0]
        except:
            self.L2 = ""
        try:
            self.L3 = re.findall("卖方：(.*?)\n",self.texts)[0]
        except:
            self.L3 = ""
        try:
            self.L4 = re.findall("买方：(.*?)\n",self.texts)[0]
        except:
            self.L4 = ""
        try:
            self.L5 = re.findall("品牌：(.*?)\n",self.texts)[0]
        except:
            self.L5 = ""
        try:
            self.L6 = re.findall("数量：(.*?)\n",self.texts)[0]
        except:
            self.L6 = ""
        try:
            self.L7 = re.findall("单价：(.*?)\n",self.texts)[0]
        except:
            self.L7 = ""
        try:
            self.L8 = re.findall("金额：￥(.*?)\n",self.texts)[0]
        except:
            self.L8 = ""
        try:
            self.L9 = re.findall("支付条款：本合同采用(.*?)方式结算",self.texts)[0]
            JIESUANFANGSHI= {'A':"信用证","B":"托收","C":"汇款"}
            self.L9=JIESUANFANGSHI[self.L9]
        except:
            self.L9 = ""

        print(self.L2,self.L2,self.L3,self.L4,self.L5,self.L6,self.L8)


if __name__ == '__main__':
    afile = r'.\文件\亚洲\CON004木星有限公司.docx'
    #doSaveAas(afile)
    adoc = ReadDoc(afile)