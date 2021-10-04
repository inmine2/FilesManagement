import pandas as pd

class NoteBook:
    def __init__(self,NoteBookPath,sheetname):
        self.ZhuanZu=pd.read_excel(NoteBookPath,sheet_name=sheetname,parse_dates=True)
        self.ZhuanZu.fillna('',inplace=True)
        self.index = self.ZhuanZu.columns.tolist()
        self.info=[]
        print(self.ZhuanZu)

    def eachContract(self,lineno):
        try:
            self.info = self.ZhuanZu.iloc[lineno].tolist()
        except:
            self.info=self.index
        print(self.info)
        return self.info

if __name__ == '__main__':
    path = '文件管理.xlsx'
    sheetname = '欧洲'
    note=NoteBook(path,sheetname)
    note.eachContract(1)
    print(note.index)
    print(note.info)