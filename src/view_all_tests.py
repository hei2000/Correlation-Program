import DataClasses
import logging
import json,sys,config,xlsxwriter

from PyQt5.QtWidgets import QLabel,QMainWindow,QWidget,QScrollArea,QGridLayout,QApplication,QPushButton
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')


WINDOW_TITLE = "View all tests"
FONT_MID = QFont("Times New Roman", 14)
REQUIREMENT_START = 2
RESULT_START = 2

class window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0

        self.initUI()

    def initUI(self):
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
                self.Tests = DataClasses.Tests.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from /Data/Tests, closing 'View all tests'")
            self.close()
            
        convert2excel = QPushButton("Convert to excel")
        convert2excel.clicked.connect(self.convert2excel_clicked)
        self.grid.addWidget(convert2excel,self.row_counter,0,1,2)
        self.row_counter += 1
        
        self.create_label()
        for idx,test in enumerate(self.Tests.tests):
            self.create_test(test,idx)

        self.show_window()
        
    def convert2excel_clicked(self):
        workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + 'Data/Tests2excel.xlsx')
        normal = workbook.add_format({'text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        worksheet = workbook.add_worksheet()
        worksheet.write(0,0,"Test Name",normal)
        worksheet.write(0,1,"Test Method",normal)
        worksheet.write(0,2,"Fabric Type",normal)
        worksheet.write(0,3,"Type of Test",normal)
        worksheet.write(0,4,"Test Group",normal)
        excel_row_counter = 2
        for test in self.Tests.tests:
            worksheet.write(excel_row_counter,0,test.name,normal)
            worksheet.write(excel_row_counter,1,test.method,normal)
            worksheet.write(excel_row_counter,2,test.fabric_type,normal)
            worksheet.write(excel_row_counter,3,test.type,normal)
            worksheet.write(excel_row_counter,4,test.group,normal)
            excel_row_counter += 1
            worksheet.write(excel_row_counter,1,"Requirements",normal)
            for requirement in test.Requirements:
                worksheet.write(excel_row_counter,2,requirement,normal)
                excel_row_counter += 1
            worksheet.write(excel_row_counter,1,"Results",normal)
            for result in test.Results:
                worksheet.write(excel_row_counter,2,result.name,normal)
                worksheet.write(excel_row_counter,3,result.type,normal)
                excel_row_counter += 1
            self.row_counter += 1
        workbook.close()
        
    def create_test(self,test: DataClasses.Test,test_count):
        self.create_info(test,test_count)
        self.create_requirements(test.Requirements)
        self.create_results(test.Results)
        
    def create_label(self):
        label_name = QLabel("<b>Test Name</b>")
        self.grid.addWidget(label_name,self.row_counter,1,1,2)
        label_method = QLabel("<b>Test Method</b>")
        self.grid.addWidget(label_method,self.row_counter,3,1,2)
        label_fabric_type = QLabel("<b>Fabric Type</b>")
        self.grid.addWidget(label_fabric_type,self.row_counter,5,1,1)
        label_test_type = QLabel("<b>Type of Test</b>")
        self.grid.addWidget(label_test_type,self.row_counter,6,1,1)
        label_group = QLabel("<b>Test Group</b>")
        self.grid.addWidget(label_group,self.row_counter,7,1,1)
        self.row_counter += 1
        
    def create_info(self,test: DataClasses.Test,test_count):
        label_test_number = QLabel("<b>" + str(test_count + 1) + ":</b>")
        label_test_number.setFont(FONT_MID)
        self.grid.addWidget(label_test_number,self.row_counter,0,1,1)
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

    def create_requirements(self,requirements: list):
        label_requirements = QLabel("<b>Requirements:</b>")
        self.grid.addWidget(label_requirements,self.row_counter,1,1,1)
        for count,requirement in enumerate(requirements):
            self.create_requirement(requirement,count)
            self.row_counter += 1
        self.row_counter += 1
            
    def create_requirement(self,requirement,count):
        label_requirement = QLabel(str(count+1) + ". " + requirement)
        self.grid.addWidget(label_requirement,self.row_counter,REQUIREMENT_START,1,1)
        
    def create_results(self,results: DataClasses.Result):
        label_results = QLabel("<b>Results:</b>")
        self.grid.addWidget(label_results,self.row_counter,1,1,1)
        for count,result in enumerate(results):
            self.create_result(result.name,result.type,count)
            self.row_counter += 1
        self.row_counter += 2
        
    def create_result(self,name,type_,count):
        label_name = QLabel(str(count+1) + ". " + name)
        self.grid.addWidget(label_name,self.row_counter,RESULT_START,1,1)
        label_type = QLabel("   " + type_)
        self.grid.addWidget(label_type,self.row_counter,RESULT_START+1,1,1)
        

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
    logging.info("Entered view_all_tests.py")
    app = QApplication(sys.argv)
    ex = window()
    #logging.info("Leaving view_all_tests.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
