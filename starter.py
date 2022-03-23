import sys
from PyQt5.QtWidgets import QApplication
from main import Main

class App(QApplication):
    # this is an entry point for the program, LET IT BE
    def __init__(self, sys_argv):
        super(App, self).__init__(sys_argv)
        self.main_controller = Main()

if __name__ == '__main__':
    app = App(sys.argv)
    sys.exit(app.exec_())