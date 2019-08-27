# from ReportFunctions import *
import kivy

kivy.require('1.10.0')
import random
from kivy.app import App
from kivy.lang import Builder
from kivy.uix.button import Button, Label
from bewbz import *


class MainApp(App):
    def build(self):
        """ Builds the GUI """
        self.title = 'Time Statistics v1.0'
        self.root = Builder.load_file('app2.kv')
        self.homescreen()
        return self.root

    def random_colour(self):
        r = random.uniform(0, 1)
        g = random.uniform(0, 1)
        b = random.uniform(0, 1)
        rgba = r, g, b, 1
        self.root.ids.do_nothing.background_color = rgba

    def homescreen(self):
        """ Makes a list of buttons for the required items + sets the colours of said buttons """
        self.root.ids.items_box.clear_widgets()
        homescreen_information = Label(
            text='Welcome.\n\n\nThis was designed by Charles to calculate the hours different\n'
                 'workgroups work at CIPS and WPS.\n\nTo begin, press the Load Data button on'
                 ' the \nleft hand side of this window.', font_size=20, halign='center')
        self.root.ids.items_box.add_widget(homescreen_information)

    def support_screen(self):
        self.root.ids.items_box.clear_widgets()
        support_information = Label(text='Sean (Charles) Robson\n0467 694 844\nsean.robson@territorygeneration.com.au',
                                    size_hint=(1, 1), halign='center', bold='True', font_size='30dp')
        self.root.ids.items_box.add_widget(support_information)

    def load_data_button(self):
        global data, row, denied
        data, row, denied = load_data()
        return data, row, denied

    def run_button(self):
            # check to see if data has been loaded check_data_loaded()
            process_data(data, row)


MainApp().run()
