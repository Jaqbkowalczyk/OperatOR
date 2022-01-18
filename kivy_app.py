from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.widget import Widget
from kivy.uix.button import Button
from kivymd.uix.boxlayout import MDBoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.properties import ObjectProperty
from kivy.lang import Builder
from kivy.core.window import Window
from kivymd.app import MDApp

# designate our kivy design file:
Window.size = (500, 700)
# Builder.load_file('operator.kv')

class MyBoxLayout(MDBoxLayout):

    def checkbox_click(self, instance, value):
        print(value)

    def on_enter(instance, value):
        print('User pressed enter in', instance)
        print(f'Value: {value}')

class Operator(MDApp):
    def build(self):
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "BlueGray"
        MyBoxLayout()
        return Builder.load_file('operator.kv')

