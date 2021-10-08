import DataClasses
import logging
import json,sys,config

import view_score_without_deduction_overview,view_score_with_deduction_overview,view_score_SPC_PSR

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QPushButton
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

FONT_MID = QFont("Times New Roman", 14)
FONT_BIG = QFont("Times New Roman", 20)
WINDOW_TITLE = "View Score"

class window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0

        self.initUI()

    def initUI(self):
        with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:  
            self.Labs = DataClasses.Labs.from_dict(json.load(json_file))
            
        label_score_without_deduction_overview = QLabel("Score_without_deduction_overview")
        self.grid.addWidget(label_score_without_deduction_overview,self.row_counter,0,1,5)
        
        button_score_without_deduction_overview = QPushButton("View")
        button_score_without_deduction_overview.clicked.connect(self.button_score_without_deduction_overview_clicked)
        self.grid.addWidget(button_score_without_deduction_overview,self.row_counter,8,1,2)
        self.row_counter += 1
        
        label_score_with_deduction_overview = QLabel("Score_with_deduction_overview")
        self.grid.addWidget(label_score_with_deduction_overview,self.row_counter,0,1,5)
        
        button_score_with_deduction_overview = QPushButton("View")
        button_score_with_deduction_overview.clicked.connect(self.button_score_with_deduction_overview_clicked)
        self.grid.addWidget(button_score_with_deduction_overview,self.row_counter,8,1,2)
        self.row_counter += 1
        
        label_score_SPC = QLabel("SPC_score")
        self.grid.addWidget(label_score_SPC,self.row_counter,0,1,5)
        
        button_score_SPC = QPushButton("View")
        button_score_SPC.clicked.connect(self.button_score_SPC_clicked)
        self.grid.addWidget(button_score_SPC,self.row_counter,8,1,2)
        self.row_counter += 1
        
        label_score_PSR = QLabel("PSR_score")
        self.grid.addWidget(label_score_PSR,self.row_counter,0,1,5)
        
        button_score_PSR = QPushButton("View")
        button_score_PSR.clicked.connect(self.button_score_PSR_clicked)
        self.grid.addWidget(button_score_PSR,self.row_counter,8,1,2)
        self.row_counter += 1

        self.show_window()
        
    def button_score_without_deduction_overview_clicked(self):
        self.window_view_score_without_deduction_overview = view_score_without_deduction_overview.window()
        
    def button_score_with_deduction_overview_clicked(self):
        self.window_view_score_with_deduction_overview = view_score_with_deduction_overview.window()
        
    def button_score_SPC_clicked(self):
        self.window_view_score_SPC = view_score_SPC_PSR.window("SPC")
        
    def button_score_PSR_clicked(self):
        self.window_view_score_PSR = view_score_SPC_PSR.window("PSR")
        

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
