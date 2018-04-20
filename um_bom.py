#coding=utf-8
import xlrd
import xlsxwriter
import os
import re
import multiprocessing

class Bom(object):
    def __init__(self,model):
        self.i=0
        self.model=model
        self.th=[]
        self.rootdir=u'J:\\PIE Process Manual\\新工序文件(CMP)\\工序手冊\\{}\\'.format(model)
        self.get_file(self.rootdir)

    #返回excel文件对象
    def get_file(self,rootdir):
        files=os.listdir(rootdir)
        for fil in files:
            if os.path.splitext(fil)[1]!='.pdf':
                self.th.append(fil)
        return self.th

    #返回excel文件名和sheet表对象
def openxls(rootdir,fil):
    txt=[]
    results=[]
    xls=xlrd.open_workbook(rootdir+fil)
    xls_sheet_names=xls.sheet_names()
    for sheets in xls_sheet_names:
        if not 'ECN' in sheets:
            sht=xls.sheet_by_name(sheets)
            rows=sht.nrows
            for i in range(0,rows):
                if sht.cell(i,9).ctype==1:
                    txt.append(sht.cell(i,9).value)

    txt=''.join(txt).replace(' ','').replace('\n','')
    return fil,txt

def writeExcel(dic,enter):
    f=xlsxwriter.Workbook('{}.xlsx'.format(enter))
    sheet=f.add_worksheet('sh')
    i=0
    for d,v in dic.items():
        sheet.write(i,0,d)
        sheet.write(i,1,v)
        i+=1
    f.close()

if __name__=='__main__':
    multiprocessing.freeze_support()
    re_=re.compile('[A-Z]{3}-[0-9]{3}-[0-9]{3}',re.S)
    v=[]
    i=0
    enter=raw_input('Enter model:>>>\n')
    rootdir=u'J:\\PIE Process Manual\\新工序文件(CMP)\\工序手冊\\{}\\'.format(enter)
    f=Bom(enter)
    t=f.th
    results=[]
    p=multiprocessing.Pool(processes=4)
    for tt in t:
        result=p.apply_async(openxls,(rootdir,tt,))
        results.append(result)
    p.close()
    p.join()
    dic={}
    for x in results:
        try:
            txtvalue=x.get()
            lists=re.findall(re_,txtvalue[1])
            lists.sort()
            key=' \n'.join(lists)
            dic[txtvalue[0]]=key
        except:
            pass
    writeExcel(dic,enter)
    print u'完成........'

'''
    workbook=xlwt.Workbook(encoding='gb18030')
    worksheet=workbook.add_sheet('sheet')
    p=multiprocessing.Pool(processes=4)
    for tt in t:
        p.apply(openxls,(tt,))
    p.close()
    p.join()
    workbook.save('{}.xls'.format(enter))
    for x in t:
        #openxls(x)
        thr=multiprocessing.Process(target=openxls,args=(rootdir,x,))
        thr.start()
        thr.join()
    workbook.save('{}.xls'.format(enter))
        thr=multiprocessing.Process(target=f.openxls,args=(x,))
        threads.append(thr)
    for child in threads:
        child.start()
    child.join()
    f.workbook.save('{}.xls'.format(enter))
'''