import DataClasses
import logging
import json,sys,config

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QPushButton
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

FONT_MID = QFont("Times New Roman", 14)
FONT_BIG = QFont("Times New Roman", 20)
WINDOW_TITLE = "SPC/PSR Score"

class window(QMainWindow):
    def __init__(self,type_):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0
        self.type = type_

        self.initUI()

    def initUI(self):
        with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_without_deduction.json') as json_file:  
            self.Labs = DataClasses.Labs.from_dict(json.load(json_file))
        for group, labs in self.Labs.Groups.items():
            label_group = QLabel(group + ":")
            label_group.setFont(FONT_BIG)
            self.grid.addWidget(label_group,self.row_counter,0,1,1)
            self.row_counter += 1
            for lab in labs.values():
                label_id = QLabel(lab.ID + ":")
                self.grid.addWidget(label_id,self.row_counter,1,1,1)
                if (self.type == "SPC"):
                    lab_score = lab.SPC_score
                else:
                    lab_score = lab.PSR_score
                label_lab_score = QLabel(self.type + ": " + str(lab_score) + "%")
                self.grid.addWidget(label_lab_score,self.row_counter,2,1,1)
                self.row_counter += 1
                for sample in lab.Samples.values():
                    if (self.type == "SPC"):
                        sample_score = sample.SPC_score
                    else:
                        sample_score = sample.PSR_score
                    if (sample_score == -1):
                        continue
                    label_sample = QLabel("Sample " + str(sample.sample_id) + ": " + str(sample_score) + "%")
                    self.grid.addWidget(label_sample,self.row_counter,3,1,1)
                    self.row_counter += 1
                    tests = sample.Tests[self.type]
                    for count,test in enumerate(tests):
                        label_test_name = QLabel("Test " + str(count+1) + ": (Name: " + test.name + " Method: " + test.method + ")")
                        self.grid.addWidget(label_test_name,self.row_counter,4,1,1)
                        self.row_counter += 1
                        label_test_score = QLabel("Test Score: " + str(round(test.average_score,2)) + "/2")
                        self.grid.addWidget(label_test_score,self.row_counter,5,1,1)
                        self.row_counter += 1
            self.row_counter += 1
                    
            
        self.show_window()
        
        

    def show_window(self):
        self.widget.setLayout(self.grid)
        self.setWindowTitle(WINDOW_TITLE)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.widget)
        self.setCentralWidget(self.scroll)
        self.showMaximized()
        self.show()


def main():
    logging.info("Entered view_score.py")
    app = QApplication(sys.argv)
    ex = window()
    #logging.info("Leaveing new_project.py")
    sys.exit(app.exec())
    
if __name__ == '__main__':
    main()
