from kivy.app import App
from kivy.lang import Builder
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.label import Label
from kivy.uix.dropdown import DropDown
import pandas as pd

from kivy.uix.textinput import TextInput
excel_path = 'D:/lab/kiosk/list.xlsx'
monitor_excel = pd.read_excel(excel_path, sheet_name = 'monitor', index_col=None).replace('nan','-')
hub_excel = pd.read_excel(excel_path, sheet_name = 'hub', index_col=None)
lanCard_excel = pd.read_excel(excel_path, sheet_name = 'lanCard', index_col=None)
kNm_excel = pd.read_excel(excel_path, sheet_name = 'kNm', index_col=None)
miniPC_excel = pd.read_excel(excel_path, sheet_name = 'miniPC', index_col=None)
fontName = 'D:/lab/kiosk/font.ttf'

Title = [['Monitor','monitor','<MonitorScreen>:'],['Hub','hub','<HubScreen>:'],['Lan Card','lanCard','<LanCardScreen>:'],['Keyborad&Mouse','kNm','<KnMScreen>:'],['Mini PC','miniPC','<MiniPCScreen>:']]
subTitle = ['Num','Place','Position','Dgist','Owner','Tack out Date','Place','Owner','Collection Date']

#####################################################################
# main screen

Main = """
<MainScreen>:
    GridLayout:
        cols: 1
        size_hint: (0.8, 1)
        pos_hint: {"center_x": 0.5}
        BoxLayout:
            Label:
                text: "CCG"
                color: "#FFFAFF"
"""
#           Image:
#               source: ''
Main_Button = ""
for title in Title:
    Main_Button = Main_Button + """
        BoxLayout:
            padding_x: (20, 20)
            Button:
                text: '""" + title[0] + """'
                on_press: 
                    root.manager.current = '""" + title[1] + """'
                    root.manager.transition.direction = 'left'
"""

#####################################################################
# subScreen

Screen_top = """
    GridLayout:
        col: 1
        rows: """

Screen_middle = """
        BoxLayout:
            spacing: 9

            Button:
                text: 'Back'
                size_hint: (0.1, 1)
                on_press: 
                    root.manager.current = 'main'
                    root.manager.transition.direction = 'right'

            Label:
                text: '"""

Screen_bottom = """'
                size_hint: (0.7, 1)
      
            Button:
                text: 'Save'
                size_hint: (0.1, 1)
        BoxLayout:
"""

Screen_subTitle = ""
for sub_title in subTitle:
    Screen_subTitle = Screen_subTitle + """
            Label:
                text: '""" + sub_title + """'
                size_hint: (0.1, 1)
"""

#####################################################################
# monitor content

monitor_sheet = ""
for i in range(0,len(monitor_excel.index)):
    monitor_sheet = monitor_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        monitor_sheet = monitor_sheet+ """
            TextInput:
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
                
"""

#####################################################################
# hub content

hub_sheet = ""
for i in range(0,len(hub_excel.index)):
    hub_sheet = hub_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        hub_sheet = hub_sheet+ """
            TextInput:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
"""

#####################################################################
# lancard content

lanCard_sheet = ""
for i in range(0,len(hub_excel.index)):
    lanCard_sheet = lanCard_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        lanCard_sheet = lanCard_sheet+ """
            TextInput:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
"""

#####################################################################
# keyboard&mouse content

kNm_sheet = ""
for i in range(0,len(kNm_excel.index)):
    kNm_sheet = kNm_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        kNm_sheet = kNm_sheet+ """
            TextInput:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
"""
#####################################################################
# miniPC content

miniPC_sheet = ""
for i in range(0,len(miniPC_excel.index)):
    miniPC_sheet = miniPC_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        miniPC_sheet = miniPC_sheet+ """
            TextInput:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
"""

#####################################################################
# moni hub lan km pc

Monitor = Title[0][2]+Screen_top+str(len(monitor_excel)+2)+   Screen_middle+Title[0][0]+Screen_bottom
hub     = Title[1][2]+Screen_top+str(len(hub_excel)+2)+       Screen_middle+Title[1][0]+Screen_bottom
lanCard = Title[2][2]+Screen_top+str(len(lanCard_excel)+2)+   Screen_middle+Title[2][0]+Screen_bottom
KnM     = Title[3][2]+Screen_top+str(len(kNm_excel)+2)+       Screen_middle+Title[3][0]+Screen_bottom
miniPC  = Title[4][2]+Screen_top+str(len(miniPC_excel)+2)+    Screen_middle+Title[4][0]+Screen_bottom


Builder.load_string(
                    Main    +Main_Button+
                    Monitor +Screen_subTitle+monitor_sheet+
                    KnM     +Screen_subTitle+kNm_sheet+
                    hub     +Screen_subTitle+hub_sheet+
                    lanCard +Screen_subTitle+lanCard_sheet+
                    miniPC  +Screen_subTitle+miniPC_sheet
)


class MainScreen(Screen):
    pass

class MonitorScreen(Screen):
    pass

class KnMScreen(Screen):
    pass

class MiniPCScreen(Screen):
    pass

class HubScreen(Screen):
    pass

class LanCardScreen(Screen):
    pass

class CCG(App):
    def build(self):

        self.CCG = ScreenManager()
        self.CCG.add_widget(MainScreen(name='main'))
        self.CCG.add_widget(MonitorScreen(name='monitor'))
        self.CCG.add_widget(KnMScreen(name='kNm'))
        self.CCG.add_widget(MiniPCScreen(name='miniPC'))
        self.CCG.add_widget(HubScreen(name='hub'))
        self.CCG.add_widget(LanCardScreen(name='lanCard'))

        return self.CCG

if __name__ == '__main__':
    CCG().run()