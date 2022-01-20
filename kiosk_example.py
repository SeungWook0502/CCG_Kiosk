from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.image import Image
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.togglebutton import ToggleButton
# from kivy.uix.popup import Popup
# from KivyCalendar import CalenderWidget

class CCG(App):
    def build(self): #run match method
        self.window = GridLayout() #run layout
        self.window.cols = 1 #columns count number
        self.window.size_hint = (0.6, 0.7) #margin
        self.window.pos_hint = {"center_x": 0.5, "center_y":0.5} #position
        #add widgets to window

        #image show widget
        self.window.add_widget(Image(source="SeengWook.jpg")) #add widget

        # #Calendar Popup widget
        # popup = Popup(
        #     title='Insert Old Date',
        #     content=CalendarWidget(),
        #     size_hint=(.9, .5)).open()

        self.check = ToggleButton(text='Check', group='DGIST', state='down')
        self.window.add_widget(self.check)

        #text output widget
        self.greeting = Label( 
                            text="What's your name?",
                            font_size = 18,
                            color = '#00FFCE'
                        ) #Title
        self.window.add_widget(self.greeting) #add widget

        #text typing widget
        self.user = TextInput(
                        multiline=False,
                        padding_y = (20, 20), #pading pixel
                        size_hint = (1, 0.5) #margin 상대적
                    )
        self.window.add_widget(self.user)

        #button
        self.button = Button(
                            text="GREET",
                            size_hint = (1, 0.5),
                            bold = True,
                            background_color = '#00FFCE',
                            background_normal = "",
                            font_size = 30,
                            color = '#F4F4F4'
                            )
        self.button.bind(on_press = self.callback) #extivate button
        self.window.add_widget(self.button)


        return self.window

    def callback(self, instance):
        self.greeting.text = "Hello " + self.user.text + "!"


if __name__ == "__main__":
    CCG().run() #run app