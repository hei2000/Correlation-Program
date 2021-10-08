import DataClasses
import logging
import json,sys,config

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QPushButton
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

FONT_MID = QFont("Times New Roman", 14)
FONT_BIG = QFont("Times New Roman", 20)
WINDOW_TITLE = "Score with deduction overview"

class window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0

        self.initUI()

    def initUI(self):
        with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:  
            self.Labs = DataClasses.Labs.from_dict(json.load(json_file))

        for group,labs in self.Labs.Groups.items():
            label_group = QLabel(group + ":")
            label_group.setFont(FONT_BIG)
            self.grid.addWidget(label_group,self.row_counter,0,1,1)
            self.row_counter += 1
            for lab in labs.values():
                label_id = QLabel(lab.ID + ":")
                self.grid.addWidget(label_id,self.row_counter,1,1,1)
                label_without_deduction = QLabel("Score without deduction: ")
                self.grid.addWidget(label_without_deduction,self.row_counter,2,1,1)
                lab_score_without_deduction = lab.average_score * 50
                label_lab_score_without_deduction = QLabel(str(round(lab_score_without_deduction,2)) + "%")
                self.grid.addWidget(label_lab_score_without_deduction,self.row_counter,3,1,1)
                label_with_deduction = QLabel("Score with deduction: ")
                self.grid.addWidget(label_with_deduction,self.row_counter,4,1,1)
                #print(type(lab.total_test_error),type(lab.deduction_timeliness),type(lab.deduction_revision))
                lab_total_error = lab.total_test_error + int(lab.deduction_timeliness) + int(lab.deduction_revision)
                lab_score_with_deduction = lab_score_without_deduction - lab_total_error * 0.5
                label_lab_score_with_deduction = QLabel(str(round(lab_score_with_deduction,2)) + "%")
                self.grid.addWidget(label_lab_score_with_deduction,self.row_counter,5,1,1)
                self.row_counter += 1
                label_error = QLabel("Total number of error: " + str(lab_total_error))
                self.grid.addWidget(label_error,self.row_counter,2,1,1)
                self.row_counter += 1
                label_timeliness = QLabel("Timeliness: " + str(lab.deduction_timeliness))
                self.grid.addWidget(label_timeliness,self.row_counter,3,1,1)
                self.row_counter += 1
                label_revision = QLabel("Revision: " + str(lab.deduction_revision))
                self.grid.addWidget(label_revision,self.row_counter,3,1,1)
                self.row_counter += 1
                for sample in lab.Samples.values():
                    label_sample_error = QLabel("Sample " + str(sample.sample_id) + ": " + str(sample.total_error) + " errors")
                    self.grid.addWidget(label_sample_error,self.row_counter,3,1,1)
                    self.row_counter += 1
                    for tests in sample.Tests.values():
                        for count,test in enumerate(tests):
                            label_test_name = QLabel("Test " + str(count+1) + ": (Name: " + test.name + " Method: " + test.method + ")")
                            self.grid.addWidget(label_test_name,self.row_counter,4,1,3)
                            label_test_error = QLabel(str(test.total_error) + " errors")
                            self.grid.addWidget(label_test_error,self.row_counter,7,1,1)
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
