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

Title = [['Monitor','monitor','<MonitorScreen>:'],['Hub','hub','<HubScreen>:'],['Lan Card','lanCard','<LanCardScreen>:'],['Keyborad&Mouse','kNm','<KnMScreen>:'],['Mini PC','miniPC','<MiniPCScreen>:']]
subTitle = ['Num','Place','Position','Dgist','Owner','Tack out Date','Place','Owner','Collection Date']
subTitle_kr = ['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜']
excel_path = 'D:/lab/kiosk/CCG_Kiosk/list.xlsx'
fontName = 'D:/lab/kiosk/CCG_Kiosk/font.ttf'

monitor_excel = pd.read_excel(excel_path, sheet_name = Title[0][1], index_col=None)
hub_excel = pd.read_excel(excel_path, sheet_name = Title[1][1], index_col=None)
lanCard_excel = pd.read_excel(excel_path, sheet_name = Title[2][1], index_col=None)
kNm_excel = pd.read_excel(excel_path, sheet_name = Title[3][1], index_col=None)
miniPC_excel = pd.read_excel(excel_path, sheet_name = Title[4][1], index_col=None)


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
                font_size: '40sp'
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

Screen_bottom = """'
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

Screen_subTitle = ""
for sub_title in subTitle_kr:
    Screen_subTitle = Screen_subTitle + """
            Label:
                text: '""" + sub_title + """'
                font_name: '""" + fontName + """'
                size_hint: (0.1, 1)
"""

#####################################################################
# monitor content

monitor_sheet = ""
for i in range(0,len(monitor_excel.index)):
    monitor_sheet = monitor_sheet + """
        BoxLayout:
            name: 'content_bl'
            background_color: (1, 1, 1, 1)
"""
    for j in range(0,9):
        monitor_sheet = monitor_sheet+ """
            TextInput:
                text: '""" + str(monitor_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                
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
                halign: 'center'
                walign: 'center'
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
                halign: 'center'
                walign: 'center'
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
                halign: 'center'
                walign: 'center'
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
                halign: 'center'
                walign: 'center'
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



        monitor_sheet = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_sheet.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        save.save()

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

        hub_sheet = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_sheet.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        save.save()

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

        lanCard_sheet = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_sheet.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        save.save()

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

        kNm_sheet = pd.DataFrame(sheet, columns=['번호','장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_sheet.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        save.save()

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

        miniPC_sheet = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_sheet.to_excel(save, sheet_name = 'miniPC', index = False)
        save.save()


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