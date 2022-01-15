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

    def check_date(self):
        fn=os.path.join(self.dir,'tz01.docx')

        # ptn=re.compile('\d{2,4}年.*月.*日.*\d+(?:\:|\：)\d?.*在.*(?:召开|举行).*会议(?:\.|\,|，|。)')
        for pattern in self.ptns:
        # ?: 是正则括号的非捕获版本。 匹配在括号内的任何正则表达式，但该分组所匹配的子字符串 不能 在执行匹配后被获取或是之后在模式中被引用。
            ptn=re.compile(pattern)
            print(ptn)
            doc=Document(fn)
            txt=''
            for prgh in doc.paragraphs:
                txt=txt+prgh.text.strip()
    
            core_txt=ptn.findall(txt)[0]
            ptn_date=re.compile('\d{2,4}年.*月.*日|')
            ptn_time=re.compile('\d+(?:\:|\：)(?:\d\d|\d)')

            meeting_date=ptn_date.findall(core_txt)[0]
            meeting_time=ptn_time.findall(core_txt)[0]
            
            meet_year=re.findall('\d{2,4}年',meeting_date)[0][0:-1]
            this_m=datetime.today().month
            if this_m==1 or this_m==2:            
                y=datetime.today().year
                if int(meet_year)!=y:
                    verify_date='yes'

            if core_txt:
                break

        return {'mt_date':meeting_date,'mt_time':meeting_time,'verify_date':verify_date}


if __name__=='__main__':
    info=Infos()
    res=info.check_date()
    # print(res)