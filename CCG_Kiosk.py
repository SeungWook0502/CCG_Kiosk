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

Title = [['Monitor','monitor','<MonitorScreen>:'],['Hub','hub','<HubScreen>:'],['Lan Card','lanCard','<LanCardScreen>:'],['Keyborad&Mouse','kNm','<KnMScreen>:'],['Mini PC','miniPC','<MiniPCScreen>:'],['Camera','camera','<CameraScreen>:'],['ETC','etc','<ETCScreen>:']]
subTitle_kr = ['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜']
excel_path = './list.xlsx'
fontName = './font.ttf'

monitor_excel = pd.read_excel(excel_path, sheet_name = Title[0][1], index_col=None)
hub_excel = pd.read_excel(excel_path, sheet_name = Title[1][1], index_col=None)
lanCard_excel = pd.read_excel(excel_path, sheet_name = Title[2][1], index_col=None)
kNm_excel = pd.read_excel(excel_path, sheet_name = Title[3][1], index_col=None)
miniPC_excel = pd.read_excel(excel_path, sheet_name = Title[4][1], index_col=None)
camera_excel = pd.read_excel(excel_path, sheet_name = Title[5][1], index_col=None)
etc_excel = pd.read_excel(excel_path, sheet_name = Title[6][1], index_col=None)

#####################################################################
# main screen

Main = """
<MainScreen>:
    GridLayout:
        cols: 1
        size_hint: (0.8, 1)
        background_color: (1, 1, 1, 1)
        pos_hint: {"center_x": 0.5}
        BoxLayout:
            background_color: (1, 1, 1, 1)
            Image:
                source: './CCG_logo.png'
"""
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
for i in range(len(monitor_excel.index), 30):
    monitor_sheet = monitor_sheet + """
        BoxLayout:
            name: 'content_bl'
            background_color: (1, 1, 1, 1)
"""
    for j in range(0,9):
        monitor_sheet = monitor_sheet+ """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                font_name: '""" + fontName + """'
                halign: 'center'
                walign: 'center'
                id: monitor_cell""" + str(i) + str(j) + """
                
"""


#####################################################################
# hub content

hub_sheet = ""
for i in range(0, len(hub_excel.index)):
    hub_sheet = hub_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        hub_sheet = hub_sheet+ """
            TextInput:
                text: '""" + str(hub_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        hub_sheet = hub_sheet+ """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: hub_cell""" + str(i) + str(j) + """
"""


#####################################################################
# lancard content

lanCard_sheet = ""
for i in range(0,len(lanCard_excel.index)):
    lanCard_sheet = lanCard_sheet + """
        BoxLayout:
"""
    for j in range(0,9):
        lanCard_sheet = lanCard_sheet+ """
            TextInput:
                text: '""" + str(lanCard_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        lanCard_sheet = lanCard_sheet+ """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: lanCard_cell""" + str(i) + str(j) + """
"""

#####################################################################
# keyboard&mouse content

kNm_sheet = ""
for i in range(0, len(kNm_excel.index)):
    kNm_sheet = kNm_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        kNm_sheet = kNm_sheet+ """
            TextInput:
                text: '""" + str(kNm_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        kNm_sheet = kNm_sheet+ """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: kNm_cell""" + str(i) + str(j) + """
"""

#####################################################################
# miniPC content

miniPC_sheet = ""
for i in range(0, len(miniPC_excel.index)):
    miniPC_sheet = miniPC_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        miniPC_sheet = miniPC_sheet+ """
            TextInput:
                text: '""" + str(miniPC_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        miniPC_sheet = miniPC_sheet+ """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: miniPC_cell""" + str(i) + str(j) + """
"""

#####################################################################
# camera content

camera_sheet = ""
for i in range(0, len(camera_excel.index)):
    camera_sheet = camera_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        camera_sheet = camera_sheet + """
            TextInput:
                text: '""" + str(camera_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        camera_sheet = camera_sheet + """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: camera_cell""" + str(i) + str(j) + """
"""

#####################################################################
# etc content

etc_sheet = ""
for i in range(0, len(etc_excel.index)):
    etc_sheet = etc_sheet + """
        BoxLayout:
"""
    for j in range(0, 9):
        etc_sheet = etc_sheet + """
            TextInput:
                text: '""" + str(etc_excel.iloc[i,j]).replace("nan","") + """'
                size_hint: (0.1, 1)
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
        etc_sheet = etc_sheet + """
            TextInput:
                text: ''
                size_hint: (0.1, 1)
                halign: 'center'
                walign: 'center'
                font_name: '""" + fontName + """'
                id: etc_cell""" + str(i) + str(j) + """
"""



#####################################################################
# moni hub lan km pc camera etc

Monitor = Title[0][2]+Screen_top+str(len(monitor_excel)+3+(30-len(monitor_excel.index)))+   Screen_back+Title[0][0]+Screen_save
hub     = Title[1][2]+Screen_top+str(len(hub_excel)+3+(30-len(hub_excel.index)))+       Screen_back+Title[1][0]+Screen_save
lanCard = Title[2][2]+Screen_top+str(len(lanCard_excel)+3+(30-len(lanCard_excel.index)))+   Screen_back+Title[2][0]+Screen_save
KnM     = Title[3][2]+Screen_top+str(len(kNm_excel)+3+(30-len(kNm_excel.index)))+       Screen_back+Title[3][0]+Screen_save
miniPC  = Title[4][2]+Screen_top+str(len(miniPC_excel)+3+(30-len(miniPC_excel.index)))+    Screen_back+Title[4][0]+Screen_save
camera  = Title[5][2]+Screen_top+str(len(camera_excel)+3+(30-len(camera_excel.index)))+    Screen_back+Title[5][0]+Screen_save
etc     = Title[6][2]+Screen_top+str(len(etc_excel)+3+(30-len(etc_excel.index)))+       Screen_back+Title[6][0]+Screen_save


Builder.load_string(
                    Main+   Main_Button+       Screen_bottom+
                    Monitor+Screen_subTitle+   monitor_sheet+  Screen_bottom+
                    KnM+    Screen_subTitle+   kNm_sheet+      Screen_bottom+
                    hub+    Screen_subTitle+   hub_sheet+      Screen_bottom+
                    lanCard+Screen_subTitle+   lanCard_sheet+  Screen_bottom+
                    miniPC+ Screen_subTitle+   miniPC_sheet+   Screen_bottom+
                    camera+ Screen_subTitle+   camera_sheet+   Screen_bottom+
                    etc+    Screen_subTitle+   etc_sheet+      Screen_bottom
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
        print(len(self.ids)/9)


        monitor_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_df.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

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

        hub_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_df.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

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

        lanCard_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_df.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

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

        kNm_df = pd.DataFrame(sheet, columns=['번호','장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_df.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

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

        miniPC_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_df.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

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

        camera_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_df.to_excel(save, sheet_name = 'camera', index = False)
        etc_excel.to_excel(save, sheet_name = 'etc', index = False)

        save.save()

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

        etc_df = pd.DataFrame(sheet, columns=['번호', '장소', '위치', '자산성 물품', '사용자', '반출날짜', '반출위치', '책임자', '회수날짜'])

        save = pd.ExcelWriter(excel_path, engine='xlsxwriter')

        monitor_excel.to_excel(save, sheet_name = 'monitor', index = False)
        hub_excel.to_excel(save, sheet_name = 'hub', index = False)
        lanCard_excel.to_excel(save, sheet_name = 'lanCard', index = False)
        kNm_excel.to_excel(save, sheet_name = 'kNm', index = False)
        miniPC_excel.to_excel(save, sheet_name = 'miniPC', index = False)
        camera_excel.to_excel(save, sheet_name = 'camera', index = False)
        etc_df.to_excel(save, sheet_name = 'etc', index = False)

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
        self.CCG.add_widget(CameraScreen(name='camera'))
        self.CCG.add_widget(EtcScreen(name='etc'))

        return self.CCG

if __name__ == '__main__':
    CCG().run()