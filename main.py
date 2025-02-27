import pandas as pd 
import re 
import numpy as np
from itertools import product
from datetime import datetime, date, time, timedelta
from ics import Calendar, Event
from ics.alarm import DisplayAlarm
import streamlit as st


class Item():
    def __init__(self, clsname:str, techer:str, place:str, weekdesc:str, xinqi:int, jiebegin:int, jieend:int):
        '''
        一个Item是表格中一格
        clsname: 课程名
        techer: 教师名
        place: 上课地点
        weekdesc: 周数描述
        xinqi: 星期几   
        jiebegin: 开始节数
        jieend: 结束节数
        '''
        self.name = clsname
        self.techer = techer
        self.place = place
        self.weekdesc = weekdesc
        self.xinqi = xinqi
        self.jiebegin = jiebegin
        self.jieend = jieend
        self.weekls = self.__getWeekls(weekdesc)
       
       
    def run(self):
        ''' 
        返回一个事件列表
        '''  
        res = []
        for ele in self.weekls:
            event = Event()
            event.name = self.name
            event.location = self.place
            event.description = "老师:"+self.techer if self.techer is not None else "无"
            event.begin, event.end = self.__getTime(ele)
            alarm = DisplayAlarm(trigger=timedelta(minutes=-ALARMTIME))
            event.alarms.append(alarm)
            res.append(event)
        return res 
     
    def __getTime(self, week:int):
        td = timedelta(hours=-8)
        h1 = int(HOURTIMEBEGIN[self.jiebegin].split(':')[0])
        m1 = int(HOURTIMEBEGIN[self.jiebegin].split(':')[1])
        h2 = int(HOURTIMEEND[self.jieend].split(':')[0])
        m2 = int(HOURTIMEEND[self.jieend].split(':')[1])
        new_date = BEGINOFTERM + timedelta(days=(week-1)*7 + self.xinqi - 1)
        time_begin = new_date.replace(hour=h1, minute=m1) + td
        time_end = new_date.replace(hour=h2, minute=m2) + td
        # print(time_begin, time_end)
        return time_begin, time_end
        
                
        
    def __getWeekls(self, weekdesc:str):
        '''
        weekdesc: 周数描述
        '''
        ls = weekdesc.strip().split(',')
        res = []
        for ele in ls:
            if ele.endswith('单') or ele.endswith('双'):
                ele = ele[:-1]
                begin, end = map(int, ele.split('-'))
                res.extend(list(range(begin, end+1, 2)))
            elif '-' in ele:
                begin, end = ele.split('-')
                res.extend(list(range(int(begin), int(end)+1)))
            else:
                res.append(int(ele))
        res.sort()
        return res
        
class HITSZICS():
    df:pd.DataFrame = None
    lsevent:list[Event] = []
    calendar = Calendar()
    
    
    def __init__(self, dfname:str):
        '''
        读取表格
        '''
        HITSZICS.df = pd.read_excel(dfname, skiprows=2)
        col_dict = {'Unnamed: 0': '节数', '星期一':1, '星期二':2, '星期三':3, '星期四':4, '星期五':5, '星期六':6, '星期日':7}
        HITSZICS.df.rename(col_dict, inplace=True, axis='columns')
        # display(HITSZICS.df)
        self.run()
            
        
        
    def run(self):
        '''
        填充lsitem
        '''
        idx = 6
        col = 7
        for i in range(idx):
            for j in range(1, 8):
                if pd.isna(HITSZICS.df.iloc[i, j]):
                    continue
                # if str(HITSZICS.df.iloc[i,j]).count('\n') >=4:   ###!!!!这里需要修改
                self.processItem(HITSZICS.df.iloc[i, j], i, j)
        for ele in HITSZICS.lsevent:
            HITSZICS.calendar.events.add(ele)
        with open(f'{FILENAME}.ics', 'w', encoding='utf-8') as f:
            f.writelines(HITSZICS.calendar) 
    

    def processItem(self, item:str, idx, col):
        '''
        从一格中提取消息
        '''            
        s = item.split('##')
        s = [_ for _ in s if len(_.strip()) > 0]
        # print(s)
        for ele in s:
            temp =' '.join(ele.split('\n')[1:])
            print('*'*60)
            # print(ele)
            print('文字描述\n', ele)
            xinqi = col
            print("星期", xinqi)
            weekdesc = HITSZICS.__getWeekDesc(temp)
            print('周数描述', weekdesc)
            techname = HITSZICS.__getTechName(temp)
            print('老师名', techname)
            resJie = HITSZICS.__getJie(temp)
            if resJie is None:
                jiebegin, jieend =  2*idx+1, 2*idx+2 
            else:
                jiebegin, jieend = resJie
            print('节开始', jiebegin)
            print('节结束', jieend)
            place = HITSZICS.__getPlace(temp)
            print('地点', place)
            # clsname = HITSZICS.__getClsName(ele.split('\n'), weekdesc, place, str(jiebegin)+'-'+str(jieend)+'节', techname)
            clsname = ele.split('\n')[0].strip()
            print('课程', clsname)
            item = Item(clsname, techname, place, weekdesc, xinqi, jiebegin, jieend)
            HITSZICS.lsevent.extend(item.run())
        print(len(HITSZICS.lsevent))
        
        
    @classmethod    
    def __getWeekDesc(cls, s:str):
        # pattern = '(((\d)+-(\d)+|(\d)+),)*((\d)+-(\d)+|(\d)+)周'
        pattern = r'(((\d+)-(\d+)|(\d+)),)*((\d+)-(\d+)|(\d+))(单|双)*周'
        result = re.search(pattern, s)
        if result is None:
            return None
        else:
            return s[result.span()[0]:result.span()[1]-1].strip()
    
    @classmethod
    def __getTechName(cls, s:str):
        pattern = '\[\D*\]'
        result = re.findall(pattern, s)
        # print(result)
        if result is None or len(result) == 0:
            return 'default teacher name'
        else:
            return ' '.join(result)
        
    @classmethod
    def __getJie(cls, s:str):
        pattern = '(\d)*-(\d)*节'       
        result = re.search(pattern, s)
        if result is None:
            return None
        else:
            res = s[result.span()[0]:result.span()[1]-1]
            begin, end = map(int, res.split('-'))
            return begin, end
        
    @classmethod
    def __getPlace(cls, s:str):
        pattern = f'\[([THGK]\d+)\]|\[(哈工大\D+)\]|\[[THGK]\d+-[THGK]\d+\]|\[(大学城\D+)\]'
        result = re.search(pattern, s)
        if result is None:
            return None
        else:
            return s[result.span()[0]+1:result.span()[1]-1].strip()
        
    @classmethod
    def __getClsName(cls, ele:list[str],  weekdesc:str, place:str, jie:str, techname:str):
        if techname is None:
            techname = 'default name'
        for s in ele:
            if len(s.strip())==0:
                continue
            if techname not in s and weekdesc not in s and place not in s and jie not in s:
                return s.strip()
            else:
                continue 
        raise ValueError("找不到课程名字")
    
    

HOURTIMEBEGIN = {1:'8:30', 3:'10:30', 5:'14:00', 7:'16:00', 9:'18:45', 11:'20:45',
                 2:'9:25', 4:'11:25', 6:'14:55', 8:'16:55', 10:'19:40', 12:'21:40'}
HOURTIMEEND   = {2:'10:15', 4:'12:15', 6:'15:45', 8:'17:45', 10:'20:30', 12:'22:30',
                 1:'9:20', 3:'11:20', 5:'14:50', 7:'16:50', 9:'19:35', 11:'21:35'}
BEGINOFTERM = datetime(2025, 2, 24)
ALARMTIME = 15
FILENAME = 'DefaultName'


st.title('课表ICS生成🥰🥰')
st.markdown('''
1. 目前仅支持HITSZ本科生课表😭
2. 使用正则表达式对表格内容进行解析，**解析结果可能存在一些问题**。**建议先新建一个日历本再导入下载好的ics文件。如果课表内容有错误，删除该日历本即可🤗**
3. 因为测试用例较少，如果有问题或者有一些改进的意见，欢迎大家与我联系😋
''')
st.subheader('准备工作')
st.markdown('''
1.**从本研平台上下载xlsx课表即可**，如下图
''')
st.image("guide.png", caption="下载课表xlsx")
st.markdown('''
2.**对于一个单元格内多个课程信息的情况，需要人工在课程名前添加**`##`**进行分隔**。如下图。
''')
cols = st.columns(2)
with cols[0]:
    st.image("p1.png", caption="1个单元格内1个课程信息")
    st.markdown('无需理会')
with cols[1]:
    st.image("p2.png", caption="1个单元格内多个课程信息")
    st.markdown('需要修改.xlsx文件，在课程名前添加`##`进行分隔')
st.subheader('开始使用')
FILENAME = st.text_input("1.输入生成的ics文件的名称（无需包含.ics后缀）", 
                         "")
BEGINOFTERM = str(st.date_input("2.输入学期开始日期(即第一周第一天)"))
ls = (BEGINOFTERM.split('-'))
BEGINOFTERM = datetime(int(ls[0]), int(ls[1]), int(ls[2]))
st.write('你选择的是:', BEGINOFTERM.date())
ALARMTIME = st.slider("3.提前提醒时间(min)", 0, 30, 15)
st.write('将在上课前', abs(ALARMTIME), 'min提醒你(弹窗方式)')
st.write('4.上传教务下载的xlsx课表')
uploaded_file = st.file_uploader("上传下载的xlsx文件", type="xlsx")
if uploaded_file is not None:
    ics = HITSZICS(uploaded_file)
    st.dataframe(ics.df, use_container_width=True)
    with open(f'{FILENAME}.ics', 'r', encoding='utf-8') as my_file:
        btn = st.download_button(
            label="下载ics文件",
            data=my_file,
            file_name=f'{FILENAME}.ics',
            mime='text/calendar',
            key='download_calendar_button'
        )