import DataClasses
import logging
import helper,json,sys,config
import xlsxwriter

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QComboBox,QCheckBox,QPushButton,QLineEdit,QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "Add New Lab"
FONT_MID = QFont("Times New Roman", 16)
COLUMN_WIDTH_UNIT = 12


class window(QMainWindow):
    def __init__(self,stage:str,Lab="",sample_list = list()):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()
        
        self.row_counter = 0
        
        self.sample_list = sample_list
        self.stage = stage
        self.Lab = Lab
        
        self.initUI()
        
    def initUI(self):
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:  
                self.Project = DataClasses.Project.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from /Data/Project.json,closing 'Add New Lab'")
            self.close()
            
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
                self.Tests = DataClasses.Tests.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from /Data/Tests.json,closing 'Add New Lab'")
            self.close()
            
        if self.stage == "Info":
            self.create_info_page()
        else:
            self.create_capabitility_selection_page()

        self.show_window()
        
    def create_info_page(self):
        self.samples = list()
        self.create_info()
        self.create_assigned_samples()
        self.info_page_next()
        
    def create_info(self):
        label_info = QLabel("<b>Basic information:<b>")
        label_info.setFont(FONT_MID)
        self.grid.addWidget(label_info,self.row_counter,0,1,10)
        self.row_counter += 1
        
        label_id = QLabel("Lab ID: (must be unique)")
        self.grid.addWidget(label_id,self.row_counter,0,1,3)
        self.line_id = QLineEdit()
        self.grid.addWidget(self.line_id,self.row_counter,3,1,7)
        self.row_counter += 1
        
        label_location = QLabel("Lab Location:")
        self.grid.addWidget(label_location,self.row_counter,0,1,3)
        self.combo_location = QComboBox()
        self.combo_location.addItems(config.COUNTRY_LIST)
        self.grid.addWidget(self.combo_location,self.row_counter,3,1,7)
        self.row_counter += 1
        
        label_city = QLabel("City: ")
        self.grid.addWidget(label_city,self.row_counter,0,1,3)
        self.line_city = QLineEdit()
        self.grid.addWidget(self.line_city,self.row_counter,3,1,7)
        self.row_counter += 1
        
        label_fullname = QLabel("Lab Full Name:")
        self.grid.addWidget(label_fullname,self.row_counter,0,1,3)
        self.line_fullname = QLineEdit()
        self.grid.addWidget(self.line_fullname,self.row_counter,3,1,7)
        self.row_counter += 1
        
        label_group = QLabel("Lab Group:")
        self.grid.addWidget(label_group,self.row_counter,0,1,3)
        self.combo_group = QComboBox()
        self.combo_group.addItems(config.LAB_GROUPS)
        #self.combo_group.currentIndexChanged.connect(self.combo_group_changed)
        self.grid.addWidget(self.combo_group,self.row_counter,3,1,7)
        self.row_counter += 1
        self.group_inhouse_row_count = self.row_counter
        self.row_counter += 1
        
        label_third_party = QLabel("Third_party:")
        self.grid.addWidget(label_third_party,self.row_counter,0,1,3)
        self.combo_third_party = QComboBox()
        self.combo_third_party.addItems(["NO","YES"])
        self.combo_third_party.currentIndexChanged.connect(self.combo_third_party_changed)
        self.grid.addWidget(self.combo_third_party,self.row_counter,3,1,7)
        self.row_counter += 1
        
    #def combo_group_changed(self):
    #    if (self.combo_group.currentText() == "In-House"):
    #        self.line_group = QLineEdit()
    #        self.grid.addWidget(self.line_group,self.group_inhouse_row_count,3,1,7)
        
    def combo_third_party_changed(self):
        for checkbox in self.checkbox_samples:
            bool_yes = (self.combo_third_party.currentText() == "YES")
            checkbox.setChecked(bool_yes)
                
    def create_assigned_samples(self):
        label_assigned_samples = QLabel("<b>Assigned Correlation test samples:</b>")
        label_assigned_samples.setFont(FONT_MID)
        self.grid.addWidget(label_assigned_samples,self.row_counter,0,1,10)
        self.row_counter += 1
        
        number_sample = self.Project.number_sample
        self.checkbox_samples = list()
        for sample_number in range(1,number_sample+1):
            self.create_assigned_sample(sample_number)
            
    def create_assigned_sample(self,sample_number):
        checkbox_sample = QCheckBox("Sample " + str(sample_number))
        checkbox_sample.stateChanged.connect(lambda: self.checkbox_sample_clicked(checkbox_sample))
        self.grid.addWidget(checkbox_sample,self.row_counter,0,1,10)
        self.checkbox_samples.append(checkbox_sample)
        self.row_counter += 1
    
    def checkbox_sample_clicked(self,box):
        sample_number = int(box.text()[7:])
        if box.isChecked():
            self.samples.append(sample_number)
        else:
            self.samples.remove(sample_number)
        #print(self.samples)
        
    def info_page_next(self):
        button_info_next = QPushButton("Next")
        button_info_next.clicked.connect(self.button_info_next_clicked)
        self.grid.addWidget(button_info_next,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_info_next_clicked(self):
        id_ = self.line_id.text()
        group = self.combo_group.currentText()
        if (group == "In-House"):
            group = self.line_fullname.text()
        location = self.combo_location.currentText()
        city = self.line_city.text()
        fullname = self.line_fullname.text()
        third_party = (self.combo_third_party.currentText() == "YES")
        
        Samples = dict()
        self.samples.sort()
        for sample_number in self.samples:
            sample = self.Project.Samples[str(sample_number)]
            if not third_party:
                sample.Tests = {"PSR":list(),"SPC":list(),"Performance":list()}
            Samples[str(sample_number)] = sample
        
        if not self.info_data_validation(id_,location,fullname,Samples):
            return
            
        self.Lab = DataClasses.Lab(ID=id_,group=group,location=location,city=city,fullname=fullname,third_party=third_party,Samples=Samples)
        self.close()
        self.__init__("0",self.Lab,self.samples)
        
    def info_data_validation(self,lab_id,location,fullname,Samples):
        validate = True
        if not self.lab_id_validation(lab_id):
            validate = False
        if not self.filled_validation(lab_id,"Lab ID"):
            validate = False
        if not self.filled_validation(location,"Location"):
            validate = False
        if not self.filled_validation(fullname,"Full name"):
            validate = False
        if not self.sample_validation(Samples):
            validate = False
        return validate
    
    def lab_id_validation(self,lab_id):
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:
                Labs = DataClasses.Labs.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from /Data/Labs.json,closing 'Add New Lab'")
            self.close()
        lab_id_list = list()
        for labs_dict in Labs.Groups.values():
            for exist_lab_id in labs_dict:
                lab_id_list.append(exist_lab_id)
        if lab_id in lab_id_list:
            message = QMessageBox()
            title = "Lab ID Duplication"
            message.setWindowTitle(title)
            text = "Cannot add existing lab.\nLab with identical Lab ID existed already!"
            message.setText(text)
            x = message.exec_()
            return False
        else:
            return True
        
    def filled_validation(self,value,text):
        if len(value) == 0:
            message = QMessageBox()
            title = "Missing " + text
            message.setWindowTitle(title)
            text = text + " field cannot be empty! Please fill in " + text + "!"
            message.setText(text)
            x = message.exec_()
            return False
        else:
            return True
        
    def sample_validation(self,Samples):
        if len(Samples) == 0:
            message = QMessageBox()
            title = "Missing Sample"
            message.setWindowTitle(title)
            text = "None of the samples is selected! You hvae to select AT LEAST ONE sample per lab"
            message.setText(text)
            x = message.exec_()
            return False
        else:
            return True
        
    def get_sample_id(self):
        idx = int(self.stage)
        sample_id = self.sample_list[idx]
        return sample_id
        #sample = self.Lab.Samples[str(sample_id)]
        #return sample.sample_id
        
    def create_capabitility_selection_page(self):
        sample_number = self.get_sample_id()
        label_selection = QLabel("<b>Sample " + str(sample_number) + " Capability Selection:</b>")
        label_selection.setFont(FONT_MID)
        self.grid.addWidget(label_selection,self.row_counter,0,1,10)
        self.row_counter += 1
        
        for test_type in self.Project.Samples[str(sample_number)].Tests:
            self.create_test_type(test_type,sample_number)
            
        if (int(self.stage) + 1) != len(self.Lab.Samples):
            selection_button = QPushButton("Next")
            selection_button.clicked.connect(self.selection_next_clicked)
        else:
            selection_button = QPushButton("Apply")
            selection_button.clicked.connect(self.selection_apply_clicked)
            
        self.grid.addWidget(selection_button,self.row_counter,8,1,2)
        self.row_counter += 1
            
    def create_test_type(self,test_type,sample_number):
        label_test_type = QLabel("<b>" + test_type + ":</b>")
        self.grid.addWidget(label_test_type,self.row_counter,0,1,10)
        self.row_counter += 1
        for count,test in enumerate(self.Project.Samples[str(sample_number)].Tests[test_type]):
            self.create_checkbox_test(test,count)
        self.row_counter += 1
            
    def create_checkbox_test(self,test,count):
        name = test.name
        method = test.method
        name_method = name + "{" + method + "}"
        checkbox_selection = QCheckBox(name_method)
        
        tests_sample = list()
        for tests_type in self.Lab.Samples[str(self.get_sample_id())].Tests.values():
            tests_sample += tests_type
        if (test in tests_sample):
            checkbox_selection.setChecked(True)
        checkbox_selection.stateChanged.connect(lambda: self.checkbox_selection_clicked(checkbox_selection))
        self.grid.addWidget(checkbox_selection,self.row_counter,0 + (count%2) * 5,1,5)
        if (count % 2):
            self.row_counter += 1
            
    def checkbox_selection_clicked(self,box):
        text = box.text()
        name,method = text.split("{")
        method = method[:-1]
        test = helper.find_test(self.Tests,name,method)
        if box.isChecked():
            self.Lab.Samples[str(self.get_sample_id())].Tests[test.type].append(test)
        else:
            self.Lab.Samples[str(self.get_sample_id())].Tests[test.type].remove(test)
            
    def selection_next_clicked(self):
        if not self.test_selection_validation():
            return
        self.close()
        self.__init__(str(int(self.stage) + 1),self.Lab,self.sample_list)
        
    def selection_apply_clicked(self):
        if not self.test_selection_validation():
            return
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:
                Labs = DataClasses.Labs.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from /Data/Labs.json,closing 'Add New Lab'")
            self.close()
        group = self.Lab.group
        if (group not in config.LAB_GROUPS):
            group = "In-House"
        if (group not in Labs.Groups):
            Labs.Groups[group] = dict()
        Labs.Groups[group][self.Lab.ID] = self.Lab
        Labs_json = Labs.to_dict()
        with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json','w',newline='') as fp:
            json.dump(Labs_json,fp)
            
        self.create_input_form()
        self.close()
        
    def test_selection_validation(self):
        sample_id = self.get_sample_id()
        sample = self.Lab.Samples[str(self.get_sample_id())]
        test_list = list()
        for test_list_type in sample.Tests.values():
            for test in test_list_type:
                test_list.append(test)
        if (len(test_list) == 0):
            message = QMessageBox()
            title = "Missing Test"
            message.setWindowTitle(title)
            text = "None of the tests is selected! You have to select AT LEAST ONE test per sample"
            message.setText(text)
            x = message.exec_()
            return False
        else:
            return True
        
        
    def create_input_form(self):
        self.workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + 'Data/Input Form/' + self.Lab.ID + ".xlsx")
        self.worksheet = self.workbook.add_worksheet()
        self.excel_row_counter = 0
        self.input_format = self.workbook.add_format({'locked': False,'bg_color':'gray','num_format': '@','text_wrap': True})
        self.number_input_format = self.workbook.add_format({'num_format': '###,##0.0','locked': False,'bg_color':'silver','text_wrap': True})
        self.center = self.workbook.add_format({'align': 'center','text_wrap': True})
        self.text_wrap = self.workbook.add_format({'text_wrap': True})
        self.worksheet.protect()
        self.create_form_info()
        for sample in self.Lab.Samples.values():
            for tests in sample.Tests.values():
                for test in tests:
                    self.create_form_test(test,sample.sample_id)
        self.workbook.close()
        
    def create_form_info(self):
        self.worksheet.write(self.excel_row_counter, 0,"Lab ID")
        self.worksheet.write(self.excel_row_counter, 1,"Lab Group")
        self.worksheet.set_column(0, 1, COLUMN_WIDTH_UNIT * 2.5)
        self.worksheet.write(self.excel_row_counter, 2,"Lab Location")
        self.worksheet.write(self.excel_row_counter, 3,"Lab Name")
        self.worksheet.write(self.excel_row_counter, 4,"Third_party")
        self.excel_row_counter += 1
        self.worksheet.write(self.excel_row_counter, 0,self.Lab.ID, self.text_wrap)
        self.worksheet.write(self.excel_row_counter, 1,self.Lab.group, self.text_wrap)
        self.worksheet.write(self.excel_row_counter, 2,helper.location_city_combine(self.Lab.location,self.Lab.city), self.text_wrap)
        self.worksheet.write(self.excel_row_counter, 3,self.Lab.fullname, self.text_wrap)
        self.worksheet.write(self.excel_row_counter, 4,"YES" if self.Lab.third_party else "NO")
        self.excel_row_counter += 2
        
    def create_form_test(self,test: DataClasses.Test,sample_id:int):
        self.worksheet.write(self.excel_row_counter, 0,"Test Name")
        self.worksheet.write(self.excel_row_counter, 1,"Test Method")
        self.worksheet.write(self.excel_row_counter, 2,"Sample")
        self.worksheet.set_column(0, 1, COLUMN_WIDTH_UNIT * 2)
        self.worksheet.set_column(2,2,COLUMN_WIDTH_UNIT)
        self.worksheet.write(self.excel_row_counter + 1, 0,test.name, self.text_wrap)
        self.worksheet.write(self.excel_row_counter + 1, 1,test.method, self.text_wrap)
        self.worksheet.write(self.excel_row_counter + 1, 2,str(sample_id), self.text_wrap)
        column_counter = 3
        if len(test.Requirements) > 1:
            self.worksheet.merge_range(self.excel_row_counter,column_counter,self.excel_row_counter,column_counter + len(test.Requirements) - 1,"Requirements",self.center)
        elif len(test.Requirements) == 1:
            self.worksheet.write(self.excel_row_counter,column_counter,"Requirement")
        for requirement in test.Requirements:
            self.worksheet.write(self.excel_row_counter + 1,column_counter,requirement,self.text_wrap)
            self.worksheet.write_string(self.excel_row_counter + 2,column_counter,"",self.input_format)
            self.worksheet.set_column(column_counter, column_counter, COLUMN_WIDTH_UNIT * 3)
            column_counter += 1
        column_counter += 1
        if len(test.Results) > 1:
            self.worksheet.merge_range(self.excel_row_counter,column_counter,self.excel_row_counter,column_counter + len(test.Results) - 1,"Results",self.center)
        elif len(test.Results) == 1:
            self.worksheet.write(self.excel_row_counter,column_counter,"Result")
        #self.worksheet.set_column(column_counter, column_counter + len(test.Results) - 1, COLUMN_WIDTH_UNIT * 1.2)
        for result in test.Results:
            self.worksheet.write(self.excel_row_counter + 1,column_counter,result.name,self.text_wrap)
            if (result.type != "Observation"):
                self.worksheet.write_string(self.excel_row_counter + 2,column_counter,"",self.number_input_format)
                self.worksheet.data_validation(self.excel_row_counter + 2,column_counter,self.excel_row_counter + 2,column_counter, {'validate': 'decimal','criteria': 'less than','value': 99999999})
            else:
                self.worksheet.write_string(self.excel_row_counter + 2,column_counter,"",self.input_format)
            self.worksheet.set_column(column_counter, column_counter, COLUMN_WIDTH_UNIT * 3)
            column_counter += 1
            
        self.worksheet.write(self.excel_row_counter + 1,column_counter,"Result Rating")
        self.worksheet.set_column(column_counter, column_counter, COLUMN_WIDTH_UNIT * 3)
        self.worksheet.write(self.excel_row_counter + 2,column_counter,"",self.input_format)
        self.worksheet.data_validation(self.excel_row_counter + 2,column_counter,self.excel_row_counter + 2,column_counter, {'validate': 'list', 'source': ['Pass', 'Fail', 'Data']})
            
        self.excel_row_counter += 4

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
    logging.info("Entered add_new_lab.py")
    app = QApplication(sys.argv)
    ex = window(stage = "Info")
    #logging.info("Leaving create_project.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
