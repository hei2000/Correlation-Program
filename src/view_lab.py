import DataClasses
import logging
import helper,json,sys,config

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "View Lab"
FONT_BIG = QFont("Times New Roman", 22)
FONT_MID = QFont("Times New Roman", 16)
TEST_START = 2
REQUIREMENT_START = 4
RESULT_START = 4

class window(QMainWindow):
    def __init__(self,lab_id:str):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()
        
        self.row_counter = 0
        
        self.lab_id = lab_id
        self.initUI()
        
    def initUI(self):
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:  
                self.Labs = DataClasses.Labs.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from Labs,closing 'View Lab'")
            self.close()
            
        lab = self.find_lab(self.lab_id)
        if lab == -1:
            logging.fatal("Cannot find '" + self.lab_id + "' in ./Data/Labs.json,closing view_lab.py")
            self.close()
        self.create_lab(lab)

        self.show_window()
    
    def find_lab(self,lab_id:str):
        for labs in self.Labs.Groups.values():
            for lab in labs.values():
                if lab.ID == lab_id:
                    return lab
        return -1
    
    def create_lab(self,lab: DataClasses.Lab):
        self.create_info()
        self.create_lab_info(lab)
        label_assign = QLabel("<b>Assisnged Correlation Test Samples:</b>")
        label_assign.setFont(FONT_BIG)
        self.grid.addWidget(label_assign,self.row_counter,0,1,10)
        self.row_counter += 1
        for sample in lab.Samples.values():
            self.create_lab_sample(sample)
        
    def create_info(self):
        label_id = QLabel("Lab ID")
        self.grid.addWidget(label_id,self.row_counter,0,1,1)
        label_group = QLabel("Lab Group")
        self.grid.addWidget(label_group,self.row_counter,1,1,1)
        label_location = QLabel("Lab Location")
        self.grid.addWidget(label_location,self.row_counter,2,1,2)
        label_fullname = QLabel("Lab Fullname")
        self.grid.addWidget(label_fullname,self.row_counter,4,1,2)
        label_third_party = QLabel("Third-Party")
        self.grid.addWidget(label_third_party,self.row_counter,6,1,1)
        self.row_counter += 1
        
    def create_lab_info(self,lab: DataClasses.Lab):
        label_id = QLabel(lab.ID)
        self.grid.addWidget(label_id,self.row_counter,0,1,1)
        label_group = QLabel(lab.group)
        self.grid.addWidget(label_group,self.row_counter,1,1,1)
        label_location = QLabel(helper.location_city_combine(lab.location,lab.city))
        self.grid.addWidget(label_location,self.row_counter,2,1,2)
        label_fullname = QLabel(lab.fullname)
        self.grid.addWidget(label_fullname,self.row_counter,4,1,2)
        label_third_party = QLabel("YES" if lab.third_party else "NO")
        self.grid.addWidget(label_third_party,self.row_counter,6,1,1)
        self.row_counter += 1
        
    def create_lab_sample(self,sample: DataClasses.Sample):
        label_sample_id = QLabel("<b>Sample " + str(sample.sample_id) + ":</b>")
        label_sample_id.setFont(FONT_MID)
        self.grid.addWidget(label_sample_id,self.row_counter,0,1,4)
        label_sample_fabric_type = QLabel("Fabric Type: " + sample.fabric_type)
        label_sample_fabric_type.setFont(FONT_MID)
        self.grid.addWidget(label_sample_fabric_type,self.row_counter,4,1,5)
        self.row_counter += 1
        
        for test_type,tests in sample.Tests.items():
            self.create_lab_sample_type(tests,test_type)
            
    def create_lab_sample_type(self,tests,test_type:str):
        label_test_type = QLabel("<b>" + test_type + "</b>")
        self.grid.addWidget(label_test_type,self.row_counter,1,1,1)
        for count,test in enumerate(tests):
            self.create_lab_sample_test(test,count)
        self.row_counter += 1
            
    def create_lab_sample_test(self,test: DataClasses.Test,count:int):
        self.create_lab_sample_test_info(test,count)
        label_requirements = QLabel("Requirements:")
        self.grid.addWidget(label_requirements,self.row_counter,REQUIREMENT_START - 1,1,1)
        for requirement_count,requirement in enumerate(test.Requirements):
            self.create_lab_sample_requirement(requirement,requirement_count)
        self.row_counter += 1
        label_results = QLabel("Results:")
        self.grid.addWidget(label_results,self.row_counter,RESULT_START - 1,1,1)
        for results_count,result in enumerate(test.Results):
            self.create_lab_sample_result(result,results_count)
        self.row_counter += 2
        
    def create_lab_sample_test_info(self,test: DataClasses.Test,count:int):
        label_test_count = QLabel(str(count+1) + ": ")
        self.grid.addWidget(label_test_count,self.row_counter,TEST_START,1,1)
        label_name = QLabel(test.name)
        self.grid.addWidget(label_name,self.row_counter,TEST_START + 1,1,2)
        label_method = QLabel(test.method)
        self.grid.addWidget(label_method,self.row_counter,TEST_START + 3,1,2)
        label_fabric_type = QLabel(test.fabric_type)
        self.grid.addWidget(label_fabric_type,self.row_counter,TEST_START + 5,1,1)
        label_test_type = QLabel(test.type)
        self.grid.addWidget(label_test_type,self.row_counter,TEST_START + 6,1,1)
        label_group = QLabel(test.group)
        self.grid.addWidget(label_group,self.row_counter,TEST_START + 7,1,1)
        self.row_counter += 1
        
    def create_lab_sample_requirement(self,requirement:str,count:int):
        label_requirement = QLabel(str(count+1) + ": " + requirement)
        self.grid.addWidget(label_requirement,self.row_counter,REQUIREMENT_START + count,1,1)
        
    def create_lab_sample_result(self,result: DataClasses.Result,count:int):
        label_test_name = QLabel(str(count+1) + ": " + result.name)
        self.grid.addWidget(label_test_name,self.row_counter,RESULT_START + count,1,1)
        label_test_type = QLabel("   " + result.type)
        self.grid.addWidget(label_test_type,self.row_counter+1,RESULT_START + count,1,1)
        
        
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
    lab_id = sys.argv[1]
    logging.info("Entered view_lab.py '" + lab_id + "'")
    app = QApplication(sys.argv)
    ex = window(lab_id)
    #logging.info("Leaving create_project.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
