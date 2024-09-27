import streamlit as st 
import pandas as pd
from ics import Calendar, Event
from ics.alarm import DisplayAlarm
from datetime import datetime, timedelta

class Item():
    classname = ''
    teachername = ''
    weeklist = []
    daylist = []
    icstimebegin = []
    icstimeend = []
    beginhourtime = ''
    endhourtime = ''
    xinqi = -1
    location = ''
    def __init__(self, itemstr:str, xinqi:int) -> None:
        self.strlist = itemstr.strip().split('\n')
        self.classname = ''
        self.teachername = ''
        self.weeklist = []
        self.daylist = []
        self.icstimebegin = []
        self.icstimeend = []
        self.beginhourtime = ''
        self.endhourtime = ''
        self.xinqi = -1
        self.location = ''
        if self.strlist.__len__() == 4:
            self.classname = self.strlist[0]
            self.teachername = self.strlist[1][1:-1]
            self.weekdescribe = self.strlist[2][1: self.strlist[2].index('周]')]
            self.location = self.strlist[2][self.strlist[2].index('周]')+3:-1]
            self.hourdescribe = self.strlist[3][1:-1]
            self.xinqi = xinqi
        elif self.strlist.__len__() == 3:
            self.classname = self.strlist[0]
            self.xinqi = xinqi
            self.hourdescribe = self.strlist[1][1:self.strlist[1].index('节]')]
            self.weekdescribe = self.strlist[1][self.strlist[1].index('节]')+3:-2]
            self.location = self.strlist[2][1:-1]
        else:
            raise ValueError()
        self.__getHourtime()
        self.__getWeeklist()
        self.__getdaylist()
        self.__gteICStime()
        self.addtoCal()
        
    def __getHourtime(self):
        hourtime = self.hourdescribe.split('-')
        self.beginhourtime = HOURTIMEBEGIN[int(hourtime[0])]
        self.endhourtime   = HOURTIMEEND[int(hourtime[1])]
    def __getWeeklist(self):
        dotsplit = self.weekdescribe.split(',')
        for el in dotsplit:
            _split = el.split('-')
            self.__extendWeeklist(_split)
        self.weeklist = list(set(self.weeklist))
    def __extendWeeklist(self, l:list[str]):
        if l[-1].endswith('单'):
            l[-1] = l[-1][:-1]
            self.weeklist.extend(list(range(int(l[0]), int(l[-1])+1, 2)))
        elif l[-1].endswith('双'):
            l[-1] = l[-1][:-1]
            self.weeklist.extend(list(range(int(l[0]), int(l[-1])+1, 2)))
        else:
            self.weeklist.extend(list(range(int(l[0]), int(l[-1])+1))) 
    def __getdaylist(self):
        for week in self.weeklist:
            time_delta = (week-1)*7 + (self.xinqi-1)
            time_delta = timedelta(days=time_delta)
            new_data = BEGIN_OF_TERM + time_delta
            self.daylist.append(new_data)
        if self.teachername == '叶允明':
            print(self.daylist.__len__())
            print(self.weeklist.__len__())
            
    def __gteICStime(self):
        td = timedelta(hours=-8)
        h1 = int(self.beginhourtime.split(':')[0])
        m1 = int(self.beginhourtime.split(':')[1])
        h2 = int(self.endhourtime.split(':')[0])
        m2 = int(self.endhourtime.split(':')[1])
        self.icstimebegin = [ele.replace(hour=h1, minute=m1)+td  for ele in self.daylist]
        self.icstimeend = [ele.replace(hour=h2, minute=m2)+td for ele in self.daylist]
     
    def __str__(self) -> str:
        return f'课程名称:{self.classname}\n授课老师:{self.teachername}\n地点:{self.location}\n星期:{self.xinqi}\n周:{self.weeklist}\n日期:{self.daylist}\n上课时间{self.beginhourtime}\n下课时间:{self.endhourtime}'
    
    def addtoCal(self):
        global cal
        for i in range(len(self.icstimebegin)):
            event = Event()
            event.name = self.classname
            event.location = self.location
            event.begin = self.icstimebegin[i]
            event.end   = self.icstimeend[i]
            alarm = DisplayAlarm(trigger=timedelta(minutes=ALARMTIME))  # 提前30分钟提醒
            event.description = f'老师:{self.teachername}'
            event.alarms.append(alarm)
            cal.events.add(event)




HOURTIMEBEGIN = {1:'8:30', 3:'10:30', 5:'14:00', 7:'16:00', 9:'18:45', 11:'20:45',
                 2:'9:25', 4:'11:25', 6:'14:55', 8:'16:55', 10:'19:40', 12:'21:40'}
HOURTIMEEND   = {2:'10:15', 4:'12:15', 6:'15:45', 8:'17:45', 10:'20:30', 12:'22:30',
                 1:'9:20', 3:'11:20', 5:'14:50', 7:'16:50', 9:'19:35', 11:'21:35'}

st.title('课表ICS生成🥰🥰')
st.subheader('使用说明')
st.markdown('''
1. 目前仅支持HITSZ本科生课表😭
2. 对xlsx内容解析为硬解析，即按照定格式进行解析。默认实验课占一个单元格的3行，理论课占一个单元格的4行
3. 由于是硬解析，**不能确保对所有专业的情况都适用**（目前仅用本人的课表进行了测试），请先确认ics文件是否正确再进行导入，避免污染手机原本的日历日程🤓
   **推荐：先新建一个日历本再导入下载好的ics文件。如果课表内容有错误，删除该日历本即可🤗**。
4. **仅需从本研平台上加载xlsx课表即可**，如下
''')
st.image("guide.png", caption="下载课表xlsx")


st.subheader('开始使用')
FILENAME = st.text_input("1.输入生成的ics文件的名称（无需包含.ics后缀）", 
                         "")
BEGIN_OF_TERM = str(st.date_input("2.输入学期开始日期(即第一周第一天)"))
ls = (BEGIN_OF_TERM.split('-'))
BEGIN_OF_TERM = datetime(int(ls[0]), int(ls[1]), int(ls[2]))
st.write('你选择的是:', BEGIN_OF_TERM.date())
ALARMTIME = st.slider("3.提前提醒时间(min)", 0, 30, 15)
ALARMTIME = -ALARMTIME
st.write('将在上课前', abs(ALARMTIME), 'min提醒你(弹窗方式)')
st.write('4.上传教务下载的xlsx课表')
uploaded_file = st.file_uploader("上传下载的xlsx文件", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, skiprows=2)
    st.dataframe(df)
    cnt = 0
    cal = Calendar()
    for x in range(1,8):
        for y in range(0, 6):
            content = df.iloc[y, x]
            if pd.isna(content):
                continue
            if content.split('\n').__len__()==3 or content.split('\n').__len__()==4:
                course = Item(content, x)
            elif '【' in content:
                s = content.split('\n')
                index3 = []
                for i in range(0, len(s)):
                    if "【" in s[i]:
                        index3.append(i)
                index4 = []
                for ele in index3:
                    if ele+3 not in index3 and ele+3<s.__len__():
                        index4.append(ele+3)
                for ele in index3:
                    content_splice = '\n'.join(s[ele:ele+3])
                    course = Item(content_splice, x)
                for ele in index4:
                    content_splice = '\n'.join(s[ele:ele+4])
                    course = Item(content_splice, x)
            elif '【'  not in content:
                s = content.split('\n')
                if s.__len__() % 4 != 0:
                    raise ValueError()
                index4 = list(range(0, s.__len__(), 4))
                for ele in index4:
                    content_splice = '\n'.join(s[ele:ele+4])
                    course = Item(content_splice, x)
    with open(f'{FILENAME}.ics', 'w', encoding='utf-8') as my_file:
        my_file.writelines(cal)
    with open(f'{FILENAME}.ics', 'r', encoding='utf-8') as my_file:
        btn = st.download_button(
            label="下载ics文件",
            data=my_file,
            file_name=f'{FILENAME}.ics',
            mime='text/calendar',
            key='download_calendar_button'
        )

    
