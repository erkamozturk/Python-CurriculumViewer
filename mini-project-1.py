#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Sun Apr 28 15:54:00 2019

@author: erkamozturk
"""

__author__ = 'erkamozturk'

from Tkinter import *
import Tkconstants, tkFileDialog
from xlrd import open_workbook, cellname
import anydbm
import pickle
import ttk
import tkMessageBox
import os


class Curriculum(Frame):
    def __init__(self, root):
        Frame.__init__(self, root)
        self.root = root
        self.tools()
        self.planning()
        self.browsed = 0

    def tools(self):
        # frames, working with 3 frame one of green part, one of buttons, one of display excel
        self.Input_UP = Frame(self.root, bg="green", width=300, height=150)  # TOP
        self.Input_MIDDLE = Frame(self.root, bg="white", width=300,height=300)  # MIDDLE
        # self.Input_UP.config(background="green")
        self.Input_BOTTOM = Frame(self.root, bg="white", width=300, height=300)  # BOTTOM
        # in frame_UP, label of Curriculum Viewer v1.0
        self.green_part = Label(self.Input_UP, text="Curriculum Viewer v1.0", bg="green", fg="white", font="Times 30")
        # labels of second part, in frame 2
        self.label1 = Label(self.Input_MIDDLE, bg="white", text="Please select curriculum excel file: ")
        self.label2 = Label(self.Input_MIDDLE, bg="white", text="Please select semester that you want to print: ")
        # button of browse in frame 2
        self.button_browse= Button(self.Input_MIDDLE, bg="white", text="Browse", command=self.askopenfile)
        self.file_opt = options = {}  # settings of browse
        options['defaultextension'] = '.xlsx'
        options['filetypes'] = [('Excel Files', ('.xlsx*','.xls')), ('Pdf Files',  '.pdf*')]
        options['initialdir'] = 'C:\Users\erkamozturk\Desktop\miniproject1 f'
        options['initialfile'] = 'cs.xlsx'
        options['parent'] = self.root
        options['title'] = 'Choose a file'
        # button of display in frame 2
        self.button_display = Button(self.Input_MIDDLE, bg="white", text="Display", command=self.display_excel)
        self.box_value = StringVar()  # settings of box
        self.box = ttk.Combobox(self.Input_MIDDLE, textvariable=self.box_value)
        self.box['values'] = ("Semester 1", "Semester 2", "Semester 3", "Semester 4", "Semester 5", "Semester 6", "Semester 7", "Semester 8")
        self.box.current(0)

    def planning(self):
        self.green_part.grid(columnspan=2, padx=50, pady=5, sticky=W+E+N+S)  # Green part
        self.label1.grid(row=3, column=0, columnspan=2, sticky="e",pady=5)  # Please select curriculum excel file
        self.label2.grid(row=4, column=0, columnspan=2, sticky="e")  # Please select semester that you want to print:
        self.button_browse.grid(row=3, column=2, sticky="w")  # button of browse
        self.button_display.grid(row=5, column=2, sticky="w")  # button of display
        self.box.grid(row=4, column=2, sticky="w")  # combobox
        self.Input_UP.pack(fill=BOTH, expand=True)  # 1st frame
        self.Input_MIDDLE.pack(fill=BOTH, expand=True)  # 2st frame. 3th frame will pach when click display button

    def askopenfile(self):
        selected_file = tkFileDialog.askopenfile(mode='r', **self.file_opt)  # select one
        self.browsed = 1  # if selected self.browsed will be 1. it controls we were in or not
        global filename
        filename = selected_file.name
        print filename # name of selected
        if selected_file.name.endswith((".xlsx", ".xls")):  # this is not obligatory. it checks we are working excels or
            pass
        else:
            tkMessageBox.showerror("Error", "This function just works with .xlsx or xlsf files. Please try again.")
            self.askopenfile()  # run again to select excel files

    def display_excel(self):
        if self.browsed == 0:  # if we were not in browse
            if os.path.exists("curriculum.db"): # if we have currunt directory any curriculum.db
                self.dab_secondshow=anydbm.open('curriculum.db',"c")  # with new variable, read to open db
                self.secondshow()  # it is for settings of frame3
                self.writing_onframe(self.dab_secondshow)  # it writes on frame3

            else:  # give specific error
                tkMessageBox.showerror("Error", "A curriculum file should be selected "
                                           "first by clicking on the Browse button")
        else:  # if we browsed before
            self.all_semesters = anydbm.open("curriculum.db", "c")  # create new db to saves info. save_to_db saves info
            self.save_to_db()  # it is for get info and fill it in db

            self.secondshow()  # it is for settings of frame3. it will display frame3 on self.root
            self.writing_onframe(self.all_semesters)  # it writes on frame3

    def save_to_db(self):
        curriculum = open_workbook(filename)  # go into excel
        sheet = curriculum.sheet_by_index(0)  # sheet 0th

# this part I get all info for all semesters. I worked with lists and container lists. at end I directed it in db before
        semester1 = [] # lists                                                                                 I created
        semester2 = []
        semester3 = []
        semester4 = []
        semester5 = []
        semester6 = []
        semester7 = []
        semester8 = []
        for i in range(6, 15):
            cont_semester1 = []  # containers
            cont_semester2 = []

            if sheet.cell(i, 0).value == "":  # if 1st one empty pass it. it will help us get summer practices
                pass
            else:
                for k in [0,1,5]:  # code, title, credit
                    cont_semester1.append(str(sheet.cell(i, k).value))
                semester1.append(cont_semester1)
            if sheet.cell(i, 8).value == "":
                pass
            else:
                for k in [8,9,13]:  # all process same before
                    cont_semester2.append(str(sheet.cell(i, k).value))
                semester2.append(cont_semester2)
        # this loop for semester3,semester4
        for i in range(18, 26):
            cont_semester3 = []
            cont_semester4 = []
            if sheet.cell(i, 0).value == "":
                pass
            else:
                for k in [0,1,5]:
                    cont_semester3.append(str(sheet.cell(i, k).value))
                semester3.append(cont_semester3)
            if sheet.cell(i, 8).value == "":
                pass
            else:
                for k in [8,9,13]:
                    cont_semester4.append(str(sheet.cell(i, k).value))
                semester4.append(cont_semester4)
        # this loop for semester5,semester6
        for i in range(30, 38):
            cont_semester5 = []
            cont_semester6 = []
            if sheet.cell(i, 0).value == "":
                pass
            else:
                for k in [0,1,5]:
                    cont_semester5.append(str(sheet.cell(i, k).value))
                semester5.append(cont_semester5)
            if sheet.cell(i, 8).value == "":
                pass
            else:
                for k in [8,9,13]:
                    cont_semester6.append(str(sheet.cell(i, k).value))
                semester6.append(cont_semester6)
        # this loop for semester7,semester8
        for i in range(41, 47):
            cont_semester7 = []
            cont_semester8 = []
            if sheet.cell(i, 0).value == "":
                pass
            else:
                for k in [0,1,5]:
                    cont_semester7.append(str(sheet.cell(i, k).value))
                semester7.append(cont_semester7)
            if sheet.cell(i, 8).value == "":
                pass
            else:
                for k in [8,9,13]:
                    cont_semester8.append(str(sheet.cell(i, k).value))
                semester8.append(cont_semester8)
        self.all_semesters["S1"] = pickle.dumps(semester1); self.all_semesters["S2"] = pickle.dumps(semester2)
        self.all_semesters["S3"] = pickle.dumps(semester3); self.all_semesters["S4"] = pickle.dumps(semester4)
        self.all_semesters["S5"] = pickle.dumps(semester5); self.all_semesters["S6"] = pickle.dumps(semester6)
        self.all_semesters["S7"] = pickle.dumps(semester7); self.all_semesters["S8"] = pickle.dumps(semester8)
        # this process for we are working with dbs. dbs just get str so we pickled
        

    def writing_onframe(self,data):
        which_s = self.box.get()  # get which_s and write it on frame3 respectively
        for i in range(1,8):
            r = 1
            if which_s == "Semester 1":
                l = data["S1"]
            elif which_s == "Semester 2":
                l = data["S2"]
            elif which_s == "Semester 3":
                l = data["S3"]
            elif which_s == "Semester 4":
                l = data["S4"]
            elif which_s == "Semester 5":
                l = data["S5"]
            elif which_s == "Semester 6":
                l = data["S6"]
            elif which_s == "Semester 7":
                l = data["S7"]
            elif which_s == "Semester 8":
                l = data["S8"]
            for c in pickle.loads(l):
                Label(self.Input_BOTTOM, text=c[0], bg="white", fg="red").grid(row=r, column=0, sticky="w")
                Label(self.Input_BOTTOM, text=c[1], bg="white", fg="red").grid(row=r, column=1, sticky="w")
                Label(self.Input_BOTTOM, text=c[2], bg="white", fg="red").grid(row=r, column=2)
                r = r+1
                # this process does writing on frame3. how? c[o] code, c[1] title, c[2] credits

    def secondshow(self):
        which_s = self.box.get()
        self.Input_BOTTOM.pack(fill=BOTH, expand=True)  # now we are displaying frame3
        for child in self.Input_BOTTOM.winfo_children():  # every click clean it on frame3
                child.destroy()
        self.button_code=Label(self.Input_BOTTOM, text="Course Code", bg="gray").grid(row=0, column=0, pady=10,
                                                                                          sticky="w")
        self.button_title=Label(self.Input_BOTTOM, text="  Course Title  ", bg="gray").grid(row=0, column=1,
                                                                                            padx=50, pady=10)
        self.button_credit=Label(self.Input_BOTTOM, text="Credit", bg="gray").grid(row=0, column=2,
                                                                                   padx=100, pady=10)


def main():
    root = Tk()
    root.wm_title("Curriculum Viewer v1.0")
    # root.geometry("800x300+250+200")
    app_erkam = Curriculum(root)
    root.mainloop()
if __name__ == "__main__":
    main()
#
