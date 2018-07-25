from randomquizgenerator import root
from randomquizgenerator.mainwindow import MainWindow


if __name__ == "__main__":
    app = MainWindow(master=root)
    app.mainloop()