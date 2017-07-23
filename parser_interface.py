# -*- coding: utf-8 -*-
from Tkinter import *
import parser_with_temp


class MainFrame(object):  # contains all other components inside
    def __init__(self):
        self.root = Tk()
        self.root.geometry('350x250+250+70')
        self.root.title("Bombayparser")
        self.root.resizable(width=True, height=False)


class InputEntry(object):  # Entry for data inputting

    def __init__(self, frame, font):
        self.value = StringVar()
        self.ent = Entry(
            frame.root, font=font, width=12, bd=3, textvariable=self.value,
            justify='left')
        self.ent.focus_set()


class Scene(object):
    def __init__(self):
        self.frame = MainFrame()  # main window
        self.head = Label(self.frame.root, text="Bombayparser",
                          justify='center', font="Arial 22")  # head label
        self.head.grid(column=0, row=0, columnspan=2)
        self.username_label = Label(self.frame.root, text=u'Enter username',
                                    font='Arial 16')
        self.username_label.grid(column=0, row=1)
        self.username = InputEntry(self.frame,
                                      "Arial 16")
        self.username.ent.grid(column=1, row=1)
        self.password_label = Label(self.frame.root, text=u'Enter password',
                                    font='Arial 16')
        self.password_label.grid(column=0, row=2, padx=10, pady=10)
        self.password = InputEntry(self.frame, "Arial 16")
        self.password.ent.grid(column=1, row=2)
        self.run_button = Button(
            self.frame.root, text=u'Start', font="Arial 18", bd=5)
        self.run_button.bind('<Button-1>', self.start_parse)
        self.run_button.grid(column=0, row=3, columnspan=2)
        self.frame.root.mainloop()


    def start_parse(self,*args):
        parser_with_temp.run_parser(
            self.username.value.get(), self.password.value.get())

if __name__ == '__main__':
    run = Scene()

