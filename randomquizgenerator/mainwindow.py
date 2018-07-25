from tkinter import *
from tkinter.filedialog import askopenfilename
import tkinter.messagebox as msgbox
from randomquizgenerator.randomquiz import RandomQuiz


class MainWindow(Frame):

    def browse_file_func(self, file_label):

        filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
        self.filename.set(filename)
        file_label["text"] = filename.split('/')[-1]

    def valid_fields(self):

        if self.filename.get() == "":
            self.err_title = "Invalid Input File"
            self.err_msg = "Please select a valid excel sheet."
            return None

        if self.title.get() == "":
            self.err_title = "Invalid Paper Title"
            self.err_msg = "Please enter title to be printed on every question paper."
            return None

        try:
            if self.num_files.get() == 0:
                self.err_title = "Invalid Question Paper Count"
                self.err_msg = "How many question papers needed?"
                return None
        except:
            self.err_title = "Invalid Question Paper Count"
            self.err_msg = "Question Paper count should be a Non-Zero Integer."
            return None

        return True

    def generate_files_func(self):

        if self.valid_fields():
            random_quiz = RandomQuiz(self.title.get(), self.filename.get(), self.num_files.get())
            generated, self.err_title, self.err_msg = random_quiz.generate_random_quiz()
            if not generated:
                msgbox.showerror(self.err_title, self.err_msg)
            else:
                msgbox.showinfo(self.err_title, self.err_msg)
        else:
            msgbox.showwarning(self.err_title, self.err_msg)

    def createWidgets(self):

        # Quit Button
        self.QUIT = Button(self)
        self.QUIT["text"] = "QUIT"
        self.QUIT["fg"] = "red"
        self.QUIT["command"] = self.quit
        self.QUIT.grid(row=10, columnspan=2, sticky=S)

        # File Label
        self.file_label = Label(self)
        self.file_label.grid(row=0, column=1)

        # Browse Button
        self.browse_file = Button(self)
        self.browse_file["text"] = "Browse Question/Answer File"
        self.browse_file["command"] = lambda: self.browse_file_func(self.file_label)
        self.browse_file.grid(row=0, column=0)

        # Title Label
        self.title_label = Label(self)
        self.title_label["text"] = "Paper Title"
        self.title_label.grid(row=1, column=0)

        # Title Entry
        self.title_entry = Entry(self)
        self.title_entry["textvariable"] = self.title
        self.title_entry.grid(row=1, column=1)

        # Number of output files Label
        self.num_files_label = Label(self)
        self.num_files_label["text"] = "Number of Question Papers"
        self.num_files_label.grid(row=2, column=0)

        # Number of output files Entry
        self.num_files_entry = Entry(self)
        self.num_files_entry["textvariable"] = self.num_files
        self.num_files_entry.grid(row=2, column=1)

        # Generate Files Button
        self.generate_files = Button(self)
        self.generate_files["text"] = "Generate Papers"
        self.generate_files["command"] = lambda: self.generate_files_func()
        self.generate_files.grid(row=3, columnspan=2)

    def __init__(self, master=None):

        Frame.__init__(self, master)
        self.pack()

        self.filename = StringVar()
        self.filename.set("")
        self.title = StringVar()
        self.title.set("")
        self.num_files = IntVar()
        self.num_files.set(0)

        self.err_title = ""
        self.err_msg = ""

        self.createWidgets()
