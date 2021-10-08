import DataClasses
import logging
import json,sys,config

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

FONT_MID = QFont("Times New Roman", 14)
FONT_BIG = QFont("Times New Roman", 20)
WINDOW_TITLE = "View Project Sample"
REQUIREMENT_START = 2
RESULT_START = 2

class window(QMainWindow):
    def __init__(self,sample_id:int,Project):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0
        self.sample_id = sample_id
        self.Project = Project

        self.initUI()

    def initUI(self):

        self.create_sample(self.Project.Samples[str(self.sample_id)])

        self.show_window()
        
    def create_sample(self,sample):
        label_sample = QLabel("<b>Sample " + str(sample.sample_id) + ":</b>")
        label_sample.setFont(FONT_BIG)
        self.grid.addWidget(label_sample,self.row_counter,0,1,10)
        self.row_counter += 1
        for test_type in sample.Tests:
            tests = sample.Tests[test_type]
            self.create_test_type(test_type,tests)
            
    def create_test_type(self,test_type:str,tests):
        label_test_type = QLabel("<b>" + test_type + "</b>")
        label_test_type.setFont(FONT_MID)
        self.grid.addWidget(label_test_type,self.row_counter,0,1,10)
        self.row_counter += 1
        for idx,test in enumerate(tests):
            self.create_test(test,idx)

    def create_test(self,test:DataClasses.Test,idx):
        self.create_test_info(test,idx)
        self.create_test_requirements(test.Requirements)
        self.create_test_results(test.Results)
        
    def create_test_info(self,test:DataClasses.Test,idx):
        label_test = QLabel(str(idx+1) + ": ")
        label_test.setFont(FONT_MID)
        self.grid.addWidget(label_test,self.row_counter,0,1,1)
        label_name = QLabel(test.name)
        label_name.setFont(FONT_MID)
        self.grid.addWidget(label_name,self.row_counter,1,1,2)
        label_method = QLabel(test.method)
        label_method.setFont(FONT_MID)
        self.grid.addWidget(label_method,self.row_counter,3,1,2)
        label_fabric_type = QLabel(test.fabric_type)
        label_fabric_type.setFont(FONT_MID)
        self.grid.addWidget(label_fabric_type,self.row_counter,5,1,1)
        label_test_type = QLabel(test.type)
        label_test_type.setFont(FONT_MID)
        self.grid.addWidget(label_test_type,self.row_counter,6,1,1)
        label_group = QLabel(test.group)
        label_group.setFont(FONT_MID)
        self.grid.addWidget(label_group,self.row_counter,7,1,1)
        self.row_counter += 1
        
    def create_test_requirements(self,requirements:list):
        label_requirements = QLabel("Requirements:")
        self.grid.addWidget(label_requirements,self.row_counter,1,1,1)
        for idx,requirement in enumerate(requirements):
            self.create_test_requirement(requirement,idx)
        self.row_counter += 1
            
    def create_test_requirement(self,requirement,idx):
        label_requirement = QLabel(requirement)
        self.grid.addWidget(label_requirement,self.row_counter,REQUIREMENT_START + idx,1,1)
        
    def create_test_results(self,Results):
        label_results = QLabel("Result:")
        self.grid.addWidget(label_results,self.row_counter,1,1,1)
        for idx,result in enumerate(Results):
            self.create_test_result(result,idx)
        self.row_counter += 2
        
    def create_test_result(self,result:DataClasses.Result,idx):
        label_test_name = QLabel(str(idx+1) + ". " + result.name)
        self.grid.addWidget(label_test_name,self.row_counter,RESULT_START + idx,1,1)
        label_test_type = QLabel("   " + result.type)
        self.grid.addWidget(label_test_type,self.row_counter+1,RESULT_START + idx,1,1)
        
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