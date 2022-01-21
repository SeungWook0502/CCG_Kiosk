import kivy

kivy.require('2.0.0')

from kivy.app import App

from kivy.uix.vkeyboard import VKeyboard
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout

class Test(App):
	def build(self):

		layout = GridLayout(cols=1)
		keyboard = VKeyboard(on_key_up = self.key_up)
		self.label = Label(text = "Selected key : ", font_size = "50sp")

		layout.add_widget(self.label)
		layout.add_widget(keyboard)

		return layout

	def key_up(self, keyboard, keycode, *args):
		if isinstance(keycode, tuple):
			keycode = keycode[1]

		if keycode == 'backspace':
			self.label.text = self.label.text[0:-1]
		elif keycode == 'spacebar':
			self.label.text = self.label.text + " "
		else:
			self.label.text = self.label.text + str(keycode)

		# print(type(self.label.text), self.label.text[0:-1])


if __name__ == '__main__':
	Test().run()