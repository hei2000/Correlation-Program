import DataClasses
import logging
import json,sys,config,helper

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

FONT_MID = QFont("Times New Roman", 16)
FONT_BIG = QFont("Times New Roman", 20)
WINDOW_TITLE = "Find Missing Labs"

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

        self.show_info()
        self.show_missing_labs()
        self.show_imported_labs()
        
        self.show_window()
        
    def show_info(self):
        label_id = QLabel("<b>Lab ID</b>")
        self.grid.addWidget(label_id,self.row_counter,2,1,1)
        label_location = QLabel("<b>Lab Location</b>")
        self.grid.addWidget(label_location,self.row_counter,3,1,2)
        label_fullname = QLabel("<b>Lab Name</b>")
        self.grid.addWidget(label_fullname,self.row_counter,5,1,2)
        label_import_datetime = QLabel("<b>Import Datetime</b>")
        self.grid.addWidget(label_import_datetime,self.row_counter,7,1,2)
        self.row_counter += 1
        
    def show_missing_labs(self):
        label_missing = QLabel("<b>Missing Labs:</b>")
        label_missing.setFont(FONT_BIG)
        self.grid.addWidget(label_missing,self.row_counter,0,1,10)
        self.row_counter += 1
        for group,labs in self.Labs.Groups.items():
            self.create_group_missing_labs(group,labs)
        
    def create_group_missing_labs(self,group,labs):
        label_group = QLabel("<b>" + group + ":</b>")
        label_group.setFont(FONT_MID)
        self.grid.addWidget(label_group,self.row_counter,1,1,9)
        #self.row_counter += 1
        not_imported_lab_count = 1;
        for lab in labs.values():
            if not lab.imported:
                self.create_group_missing_lab(lab,not_imported_lab_count)
                not_imported_lab_count += 1
        self.row_counter += 1
    
    def create_group_missing_lab(self,lab,count):
        label_lab_id = QLabel(str(count) + ": " + lab.ID)
        self.grid.addWidget(label_lab_id,self.row_counter,2,1,1)
        label_lab_location = QLabel(helper.location_city_combine(lab.location,lab.city))
        self.grid.addWidget(label_lab_location,self.row_counter,3,1,2)
        label_lab_fullname = QLabel(lab.fullname)
        self.grid.addWidget(label_lab_fullname,self.row_counter,5,1,2)
        self.row_counter += 1
        
    def show_imported_labs(self):
        label_imported = QLabel("<b>Imoprted Labs:</b>")
        label_imported.setFont(FONT_BIG)
        self.grid.addWidget(label_imported,self.row_counter,0,1,10)
        self.row_counter += 1
        for group, labs in self.Labs.Groups.items():
            self.create_group_imported_labs(group,labs)
            
    def create_group_imported_labs(self,group,labs):
        label_group = QLabel("<b>" + group + ":</b>")
        label_group.setFont(FONT_MID)
        self.grid.addWidget(label_group,self.row_counter,1,1,9)
        #self.row_counter += 1
        imported_lab_count = 1
        for lab in labs.values():
            if lab.imported:
                self.create_group_imported_lab(lab,imported_lab_count)
                imported_lab_count += 1
        self.row_counter += 1
                
    def create_group_imported_lab(self,lab,count):
        label_lab_id = QLabel(str(count) + ": " + lab.ID)
        self.grid.addWidget(label_lab_id,self.row_counter,2,1,1)
        label_lab_location = QLabel(helper.location_city_combine(lab.location,lab.city))
        self.grid.addWidget(label_lab_location,self.row_counter,3,1,2)
        label_lab_fullname = QLabel(lab.fullname)
        self.grid.addWidget(label_lab_fullname,self.row_counter,5,1,2)
        label_lab_import_datetime = QLabel(lab.import_date)
        self.grid.addWidget(label_lab_import_datetime,self.row_counter,7,1,2)
        self.row_counter += 1
        
        
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
    logging.info("Entered missing_lab.py")
    app = QApplication(sys.argv)
    ex = window()
    #logging.info("Leaveing new_project.py")
    sys.exit(app.exec())
    
if __name__ == '__main__':
    main()
