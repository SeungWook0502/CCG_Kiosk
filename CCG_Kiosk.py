from kivy.app import App
from kivy.lang import Builder
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.label import Label
from kivy.uix.dropdown import DropDown
from kivy.uix.textinput import TextInput
import pandas as pd
import openpyxl
from datetime import datetime

Title = [['Monitor','monitor','<MonitorScreen>:'],['Hub','hub','<HubScreen>:'],['Lan Card','lanCard','<LanCardScreen>:'],['Keyborad&Mouse','kNm','<KnMScreen>:'],['Mini PC','miniPC','<MiniPCScreen>:'],['Camera','camera','<CameraScreen>:'],['ETC','etc','<ETCScreen>:']]
excel_path = './list.xlsx'
fontName = './font.ttf'

monitor_excel = pd.read_excel(excel_path, sheet_name = Title[0][1], index_col=None)
hub_excel = pd.read_excel(excel_path, sheet_name = Title[1][1], index_col=None)
lanCard_excel = pd.read_excel(excel_path, sheet_name = Title[2][1], index_col=None)
kNm_excel = pd.read_excel(excel_path, sheet_name = Title[3][1], index_col=None)
miniPC_excel = pd.read_excel(excel_path, sheet_name = Title[4][1], index_col=None)
camera_excel = pd.read_excel(excel_path, sheet_name = Title[5][1], index_col=None)
etc_excel = pd.read_excel(excel_path, sheet_name = Title[6][1], index_col=None)


light_gray = ["2.5, 2.5, 2.5, 255", "192, 192, 192, .9", ".9, .9, .9, 6"]
dark_gray = ["2, 2, 2, 255", "128, 128, 128, .6", "0.7, 0.7, 0.7, 2"]
check_white = "0, 0, 0, 1"

#####################################################################
# main screen

Main = """
<MainScreen>:
    GridLayout:
        cols: 1
        size_hint: (0.8, 1)
        background_color: (255, 255, 255, 1)
        pos_hint: {"center_x": 0.5}
        BoxLayout:
            background_color: (1, 1, 1, 1)
            Image:
                source: './CCG_logo_nonBackground_top.png'
"""

Main_content = ""
for title in Title:
    Main_content = Main_content + """
        BoxLayout:
            padding_x: (20, 20)
            Button:
                text: '""" + title[0] + """'
                on_press: 
                    root.manager.current = '""" + title[1] + """'
                    root.manager.transition.direction = 'left'
"""

Main_bottom = """
            
        BoxLayout:
            Image:
                source: './CCG_logo_nonBackground_bot.png'

"""

#####################################################################
# subScreen

Screen_top = """
    GridLayout:
        col: 1
        rows: """

Screen_back = """
        BoxLayout:
            name: 'title_bl'
            spacing: 9

            Button:
                text: 'Back'
                size_hint: (0.1, 1)
                on_press: 
                    root.manager.current = 'main'
                    root.manager.transition.direction = 'right'

            Label:
                text: '"""

Screen_save = """'
                size_hint: (0.7, 1)
      
            Button:
                text: 'Save'
                size_hint: (0.1, 1)
                on_press:
                    root.save_file()
                    root.manager.current = 'main'
                    root.manager.transition.direction = 'right'

        BoxLayout:
            name: 'subTitle_bl'
"""

Screen_bottom = """
            
        BoxLayout:
            Label:
                text: ' '
                size_hint: (1, 1)
                text_size: root.width/2, None
                halign: 'center'
                walign: 'center'

"""

#####################################################################
# monitor content

monitor_sheet = ""
for i in range(0,len(monitor_excel.index)):
    monitor_sheet = monitor_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]
        if j in [4,7]: 
            monitor_sheet = monitor_sheet+ """
            Spinner:
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            monitor_sheet = monitor_sheet + """
            Button:
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('monitor_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            monitor_sheet = monitor_sheet + """
            CheckBox:
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'monitor_cell""" + str(i) + str(j) + """')
"""
            if not(str(monitor_excel.iloc[i,j]).replace("nan","") in ['False','']):
                monitor_sheet = monitor_sheet + """
                active: 'down'
"""

        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            monitor_sheet = monitor_sheet+ """
            TextInput:
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (""" + w + """, 1)
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                
"""

for i in range(len(monitor_excel.index), 30):
    monitor_sheet = monitor_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]
        if j in [4, 7]:
            monitor_sheet = monitor_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            monitor_sheet = monitor_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('monitor_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            monitor_sheet = monitor_sheet + """
            CheckBox:
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                text: ''
                size_hint: (0.05, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'monitor_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            monitor_sheet = monitor_sheet+ """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                
"""

monitor_subTitle = ""
sub_i = 0
for sub_title in monitor_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    monitor_subTitle = monitor_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_subtitle_""" + sub_title + """
"""


#####################################################################
# hub content

hub_sheet = ""
for i in range(0, len(hub_excel.index)):
    hub_sheet = hub_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            hub_sheet = hub_sheet+ """
            Spinner:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            hub_sheet = hub_sheet + """
            Button:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('hub_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            hub_sheet = hub_sheet + """
            CheckBox:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'hub_cell""" + str(i) + str(j) + """')
"""
            if not(str(hub_excel.iloc[i,j]).replace("nan","") in ['False','']):
                hub_sheet = hub_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            hub_sheet = hub_sheet+ """
            TextInput:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: hub_cell""" + str(i) + str(j) + """
"""
for i in range(len(hub_excel.index), 30):
    hub_sheet = hub_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            hub_sheet = hub_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            hub_sheet = hub_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('hub_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            hub_sheet = hub_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'hub_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            hub_sheet = hub_sheet+ """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: hub_cell""" + str(i) + str(j) + """
"""

hub_subTitle = ""
sub_i = 0
for sub_title in hub_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    hub_subTitle = hub_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: hub_subtitle_""" + sub_title + """
"""

#####################################################################
# lancard content

lanCard_sheet = ""
for i in range(0,len(lanCard_excel.index)):
    lanCard_sheet = lanCard_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            lanCard_sheet = lanCard_sheet+ """
            Spinner:
                text: '""" + str(lanCard_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            lanCard_sheet = lanCard_sheet + """
            Button:
                text: '""" + str(lanCard_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('lanCard_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            lanCard_sheet = lanCard_sheet + """
            CheckBox:
                text: '""" + str(lanCard_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'lanCard_cell""" + str(i) + str(j) + """')
"""
            if not(str(lanCard_excel.iloc[i,j]).replace("nan","") in ['False','']):
                lanCard_sheet = lanCard_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            lanCard_sheet = lanCard_sheet+ """
            TextInput:
                text: '""" + str(lanCard_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: lanCard_cell""" + str(i) + str(j) + """
"""

for i in range(len(lanCard_excel.index), 30):
    lanCard_sheet = lanCard_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            lanCard_sheet = lanCard_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            lanCard_sheet = lanCard_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('lanCard_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            lanCard_sheet = lanCard_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'lanCard_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            lanCard_sheet = lanCard_sheet+ """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: lanCard_cell""" + str(i) + str(j) + """
"""

lanCard_subTitle = ""
sub_i = 0
for sub_title in lanCard_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    lanCard_subTitle = lanCard_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: ("""  + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: lanCard_subtitle_""" + sub_title + """
"""

#####################################################################
# keyboard&mouse content

kNm_sheet = ""
for i in range(0, len(kNm_excel.index)):
    kNm_sheet = kNm_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            kNm_sheet = kNm_sheet+ """
            Spinner:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            kNm_sheet = kNm_sheet + """
            Button:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('kNm_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            kNm_sheet = kNm_sheet + """
            CheckBox:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'kNm_cell""" + str(i) + str(j) + """')
"""
            if not(str(kNm_excel.iloc[i,j]).replace("nan","") in ['False','']):
                kNm_sheet = kNm_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            kNm_sheet = kNm_sheet+ """
            TextInput:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: kNm_cell""" + str(i) + str(j) + """
"""
for i in range(len(kNm_excel.index), 30):
    kNm_sheet = kNm_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            kNm_sheet = kNm_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            kNm_sheet = kNm_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('kNm_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            kNm_sheet = kNm_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'kNm_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            kNm_sheet = kNm_sheet+ """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: kNm_cell""" + str(i) + str(j) + """
"""

kNm_subTitle = ""
sub_i = 0
for sub_title in kNm_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    kNm_subTitle = kNm_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: kNm_subtitle_""" + sub_title + """
"""

#####################################################################
# miniPC content

miniPC_sheet = ""
for i in range(0, len(miniPC_excel.index)):
    miniPC_sheet = miniPC_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            miniPC_sheet = miniPC_sheet+ """
            Spinner:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            miniPC_sheet = miniPC_sheet + """
            Button:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('miniPC_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            miniPC_sheet = miniPC_sheet + """
            CheckBox:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'miniPC_cell""" + str(i) + str(j) + """')
"""
            if not(str(miniPC_excel.iloc[i,j]).replace("nan","") in ['False','']):
                miniPC_sheet = miniPC_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            miniPC_sheet = miniPC_sheet+ """
            TextInput:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: miniPC_cell""" + str(i) + str(j) + """
"""
for i in range(len(miniPC_excel.index), 30):
    miniPC_sheet = miniPC_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            miniPC_sheet = miniPC_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            miniPC_sheet = miniPC_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('miniPC_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            miniPC_sheet = miniPC_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'miniPC_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            miniPC_sheet = miniPC_sheet+ """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: miniPC_cell""" + str(i) + str(j) + """
"""

miniPC_subTitle = ""
sub_i = 0
for sub_title in miniPC_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    miniPC_subTitle = miniPC_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: miniPC_subtitle_""" + sub_title + """
"""

#####################################################################
# camera content

camera_sheet = ""
for i in range(0, len(camera_excel.index)):
    camera_sheet = camera_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            camera_sheet = camera_sheet+ """
            Spinner:
                text: '""" + str(camera_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            camera_sheet = camera_sheet + """
            Button:
                text: '""" + str(camera_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('camera_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            camera_sheet = camera_sheet + """
            CheckBox:
                text: '""" + str(camera_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'camera_cell""" + str(i) + str(j) + """')
"""
            if not(str(camera_excel.iloc[i,j]).replace("nan","") in ['False','']):
                camera_sheet = camera_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            camera_sheet = camera_sheet + """
            TextInput:
                text: '""" + str(camera_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: camera_cell""" + str(i) + str(j) + """
"""
for i in range(len(camera_excel.index), 30):
    camera_sheet = camera_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            camera_sheet = camera_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            camera_sheet = camera_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('camera_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            camera_sheet = camera_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'camera_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 1:
                w = '0.07'
            elif j == 2:
                w = '0.18'
            elif j == 3:
                w = '0.05'
            else:
                w = '0.1'
            camera_sheet = camera_sheet + """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: camera_cell""" + str(i) + str(j) + """
"""

camera_subTitle = ""
sub_i = 0
for sub_title in camera_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    camera_subTitle = camera_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: camera_subtitle_""" + sub_title + """
"""

#####################################################################
# etc content

etc_sheet = ""
for i in range(0, len(etc_excel.index)):
    etc_sheet = etc_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            etc_sheet = etc_sheet+ """
            Spinner:
                text: '""" + str(etc_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            etc_sheet = etc_sheet + """
            Button:
                text: '""" + str(etc_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('etc_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            etc_sheet = etc_sheet + """
            CheckBox:
                text: '""" + str(etc_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'etc_cell""" + str(i) + str(j) + """')
"""
            if not(str(etc_excel.iloc[i,j]).replace("nan","") in ['False','']):
                etc_sheet = etc_sheet + """
                active: 'down'
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            etc_sheet = etc_sheet + """
            TextInput:
                text: '""" + str(etc_excel.iloc[i,j]).replace("nan","") + """'
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: etc_cell""" + str(i) + str(j) + """
"""
for i in range(len(etc_excel.index), 30):
    etc_sheet = etc_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        background_rgba = ""
        if i%2==0:
            button_rgba = light_gray[0]
            Check_rgba = light_gray[1]
            inputText_rgba = light_gray[2]
        else:
            button_rgba = dark_gray[0]
            Check_rgba = dark_gray[1]
            inputText_rgba = dark_gray[2]

        if j in [4,7]: 
            etc_sheet = etc_sheet+ """
            Spinner:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                values: ["이현기","노종구","강복구","최락현","바산트 라구","하호건","이종철","이승준","윤재성","구경모","김진수","박장훈",""]
                
"""
        elif j in [5, 8]:
            etc_sheet = etc_sheet + """
            Button:
                text: ''
                size_hint: (0.1, 1)
                background_color: (""" + button_rgba + """)
                color: 0, 0, 0, 1
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                on_press: root.get_dateNow('etc_cell""" + str(i) + str(j) + """')
"""
        elif j == 3:
            etc_sheet = etc_sheet + """
            CheckBox:
                text: ''
                size_hint: (0.05, 1)
                canvas.before:
                    Color:
                        rgba: """ + Check_rgba + """
                    Rectangle:
                        pos: self.pos
                        size: self.size
                color: """ + check_white + """
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_cell""" + str(i) + str(j) + """
                on_active: root.on_check(self, self.active, 'etc_cell""" + str(i) + str(j) + """')
"""
        else:
            if j == 2:
                w = '0.18'
            elif j == 1:
                w = '0.07'
            else:
                w = '0.1'
            etc_sheet = etc_sheet + """
            TextInput:
                text: ''
                background_color: (""" + inputText_rgba + """)
                color: 0, 0, 0, 1
                size_hint: (""" + w + """, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: etc_cell""" + str(i) + str(j) + """
"""

etc_subTitle = ""
sub_i = 0
for sub_title in etc_excel.columns:
    if sub_i == 1:
        w = '0.07'
    elif sub_i == 2:
        w = '0.18'
    elif sub_i == 3:
        w = '0.05'
    else:
        w = '0.1'
    sub_i = sub_i + 1
    etc_subTitle = etc_subTitle + """
            TextInput:
                text: '""" + sub_title + """'
                size_hint: (""" + w + """, 1)
                disabled_foreground_color: (0, 0, 0, 1)
                background_color: (0, 0, 0, 1)
                foreground_color: (1, 1, 1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: etc_subtitle_""" + sub_title + """
"""


#####################################################################
# moni hub lan km pc camera etc

Monitor = Title[0][2]+Screen_top+str(len(monitor_excel)+    3+(30-len(monitor_excel.index)))+   Screen_back+Title[0][0]+Screen_save
hub     = Title[1][2]+Screen_top+str(len(hub_excel)+        3+(30-len(hub_excel.index)))+       Screen_back+Title[1][0]+Screen_save
lanCard = Title[2][2]+Screen_top+str(len(lanCard_excel)+    3+(30-len(lanCard_excel.index)))+   Screen_back+Title[2][0]+Screen_save
KnM     = Title[3][2]+Screen_top+str(len(kNm_excel)+        3+(30-len(kNm_excel.index)))+       Screen_back+Title[3][0]+Screen_save
miniPC  = Title[4][2]+Screen_top+str(len(miniPC_excel)+     3+(30-len(miniPC_excel.index)))+    Screen_back+Title[4][0]+Screen_save
camera  = Title[5][2]+Screen_top+str(len(camera_excel)+     3+(30-len(camera_excel.index)))+    Screen_back+Title[5][0]+Screen_save
etc     = Title[6][2]+Screen_top+str(len(etc_excel)+        3+(30-len(etc_excel.index)))+       Screen_back+Title[6][0]+Screen_save

Builder.load_string(
                    Main+       Main_content+       Main_bottom+
                    Monitor+    monitor_subTitle+   monitor_sheet+  Screen_bottom+
                    KnM+        kNm_subTitle+       kNm_sheet+      Screen_bottom+
                    hub+        hub_subTitle+       hub_sheet+      Screen_bottom+
                    lanCard+    lanCard_subTitle+   lanCard_sheet+  Screen_bottom+
                    miniPC+     miniPC_subTitle+    miniPC_sheet+   Screen_bottom+
                    camera+     camera_subTitle+    camera_sheet+   Screen_bottom+
                    etc+        etc_subTitle+       etc_sheet+      Screen_bottom
)


class MainScreen(Screen):
    pass

class MonitorScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0


        monitor_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        monitor_df = monitor_df.rename(columns=monitor_df.iloc[0])
        monitor_df = monitor_df.drop(monitor_df.index[0])


        monitor_df.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        monitor_excel = pd.read_excel(excel_path, sheet_name = Title[0][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)

class HubScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        hub_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        hub_df = hub_df.rename(columns=hub_df.iloc[0])
        hub_df = hub_df.drop(hub_df.index[0])

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_df.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        hub_excel = pd.read_excel(excel_path, sheet_name = Title[1][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)

class LanCardScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        lanCard_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        lanCard_df = lanCard_df.rename(columns=lanCard_df.iloc[0])
        lanCard_df = lanCard_df.drop(lanCard_df.index[0])

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_df.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        lanCard_excel = pd.read_excel(excel_path, sheet_name = Title[2][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)

class KnMScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        kNm_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        kNm_df = kNm_df.rename(columns=kNm_df.iloc[0])
        kNm_df = kNm_df.drop(kNm_df.index[0])

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_df.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        kNm_excel = pd.read_excel(excel_path, sheet_name = Title[3][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)

class MiniPCScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        miniPC_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        miniPC_df = miniPC_df.rename(columns=miniPC_df.iloc[0])
        miniPC_df = miniPC_df.drop(miniPC_df.index[0])

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_df.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        miniPC_excel = pd.read_excel(excel_path, sheet_name = Title[4][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)

class CameraScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        camera_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        camera_df = camera_df.rename(columns=camera_df.iloc[0])
        camera_df = camera_df.drop(camera_df.index[0])


        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_df.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()
        
        camera_excel = pd.read_excel(excel_path, sheet_name = Title[5][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)
        
class EtcScreen(Screen):
    def save_file(self):

        sheet_row = []
        sheet = []
        i = 0

        for a in self.ids:
            i=i+1
            if i<9:
                sheet_row.append(self.ids[a].text)
            else:
                sheet_row.append(self.ids[a].text)
                sheet.append(sheet_row)
                sheet_row = []
                i=0

        etc_df = pd.DataFrame(sheet)

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        etc_df = etc_df.rename(columns=etc_df.iloc[0])
        etc_df = etc_df.drop(etc_df.index[0])

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_df.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

        etc_excel = pd.read_excel(excel_path, sheet_name = Title[6][1], index_col=None)

    def get_dateNow(self, btnId):

        if self.ids[btnId].text == '':
            now_date = datetime.now()
            now_year = str(now_date.year%100)
            if now_date.month < 10:
                now_month = '0' + str(now_date.month)
            else:
                now_month = str(now_date.month)
            if now_date.day < 10:
                now_day = '0' + str(now_date.day)
            else:
                now_day = str(now_date.day)
            self.ids[btnId].text = now_year + '.' + now_month + '.' + now_day
        else:
            self.ids[btnId].text = ''

    def on_check(self, instance, value, chId):
        if value:
            self.ids[chId].text = ''
        else:
            self.ids[chId].text = str(value)
        
class CCG(App):
    def build(self):

        self.CCG = ScreenManager()
        self.CCG.add_widget(MainScreen(name='main'))
        self.CCG.add_widget(MonitorScreen(name='monitor'))
        self.CCG.add_widget(KnMScreen(name='kNm'))
        self.CCG.add_widget(MiniPCScreen(name='miniPC'))
        self.CCG.add_widget(HubScreen(name='hub'))
        self.CCG.add_widget(LanCardScreen(name='lanCard'))
        self.CCG.add_widget(CameraScreen(name='camera'))
        self.CCG.add_widget(EtcScreen(name='etc'))

        return self.CCG

if __name__ == '__main__':
    CCG().run()