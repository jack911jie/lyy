import os
import sys
import re
from docx import Document
from datetime import datetime

class Infos:
    def __init__(self):
        self.dir='e:\\py\\lyy'
        self.fmt_file=os.path.join(self.dir,'info_format.txt')
        with open(self.fmt_file,'r',encoding='utf-8') as fn:
            self.ptns=fn.readlines()

    def check_date(self,fn='e:\\py\\lyy\\通知\\tz01.docx'):

        # ptn=re.compile('\d{2,4}年.*月.*日.*\d+(?:\:|\：)\d?.*在.*(?:召开|举行).*会议(?:\.|\,|，|。)')
        # self.ptns=[ "r'会议时间.*'"]
        # self.ptns=["r'\d{2,4}年.*月.*日.*\d+(?:\:|\：)\d?.*在.*(?:召开|举行).*会议(?:\.|\,|，|。)'"]

        doc=Document(fn)    
        txt=''
        for prgh in doc.paragraphs:
            txt=txt+prgh.text.strip()
        txt=txt.strip().replace(' ','')
        

        for pattern in self.ptns:
        # ?: 是正则括号的非捕获版本。 匹配在括号内的任何正则表达式，但该分组所匹配的子字符串 不能 在执行匹配后被获取或是之后在模式中被引用。
           
            ptn=re.compile(eval(pattern))
            core_txts=ptn.findall(txt)            
            # print('core_txts',core_txts)
            if core_txts:
                core_txt=core_txts[0].strip().replace(' ','')
                # print(core_txt)
                ptn_date=re.compile('\d{2,4}年.*月.*日|')
                ptn_time=re.compile('\d+(?:\:|\：)(?:\d\d|\d)')

                meeting_date=list(filter(None,ptn_date.findall(core_txt)))[0].strip()
                meeting_time=list(filter(None,ptn_time.findall(core_txt)))[0].strip()
                meet_year=re.findall('\d{2,4}年',meeting_date)[0][0:-1]
                this_m=datetime.today().month
                if this_m==1 or this_m==2:            
                    y=datetime.today().year
                    if int(meet_year)!=y:
                        verify_date='yes'
                    else:
                        verify_date='no'

                break
            

        return {'mt_date':meeting_date,'mt_time':meeting_time,'verify_date':verify_date}

    def check_info_dir(self,dirname):
        fns=[]
        for fn in os.listdir(dirname):
            if fn[-4:].lower()=='docx' and '$' not in fn:
                fns.append(os.path.join(dirname,fn))
        for info in fns:
            info_fn=info.split('\\')[-1]
            print('\n正在检查 {} ------>>'.format(info_fn))
            res=self.check_date(info)
            print('结果：')
            print('会议日期：',res['mt_date'])
            print('会议时间：',res['mt_time'])
            if res['verify_date']=='yes':
                print('现在是 {} 年，请检查会议日期。'.format(datetime.today().year))
            print('----------------------------\n')

if __name__=='__main__':
    info=Infos()
    # res=info.check_date('tz01.docx')
    # print(res)
    info.check_info_dir('e:\\py\\lyy\\通知')