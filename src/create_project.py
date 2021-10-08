import DataClasses
import logging
import helper,config,json,sys

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QComboBox,QCheckBox,QCompleter,QPushButton,QLineEdit
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIntValidator

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "Create new Project"

class window(QMainWindow):
    def __init__(self,stage:str,Project = ""):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()
        
        self.row_counter = 0
        
        self.stage = str(stage)
        self.Project = Project
        self.initUI()
        
    def initUI(self):
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
                self.Tests = DataClasses.Tests.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from Tests,closing 'Create Project'")
            self.close()
        if self.stage == "sample_number":
            self.create_sample_number()
        elif int(self.stage) <= self.Project.number_sample:
            self.create_sample()
        self.show_window()
        
    def create_sample_number(self):
        label_number = QLabel("Please input the number of sample: ")
        self.grid.addWidget(label_number,self.row_counter,0,1,2)

        self.line_number = QLineEdit()
        self.line_number.setValidator(QIntValidator(1,30))
        self.grid.addWidget(self.line_number,self.row_counter,2,1,2)
        self.row_counter += 1

        button_next_number = QPushButton("Next")
        button_next_number.clicked.connect(self.button_next_number_clicked)
        self.grid.addWidget(button_next_number,self.row_counter,4,1,2)

    def button_next_number_clicked(self):
        sample_number = int(self.line_number.text())
        samples = dict()
        for sample_id in range(1,sample_number+1):
            sample = DataClasses.Sample(fabric_type = "Knit",sample_id=sample_id,Tests = {"PSR":list(),"SPC":list(),"Performance":list()})
            samples[str(sample_id)] = sample
        #print(len(samples))
        Project = DataClasses.Project(number_sample=sample_number,Samples=samples)
        self.close()
        self.__init__('1',Project)

    def create_sample(self):
        label_sample_text = "<b>Sample " + str(self.stage) + ":</b>"
        label_sample = QLabel(label_sample_text)
        self.grid.addWidget(label_sample,self.row_counter,0,1,10)
        self.row_counter += 1
        
        label_type = QLabel("Fabric type:")
        self.grid.addWidget(label_type,self.row_counter,0,1,2)

        self.combo_type = QComboBox()
        self.combo_type.addItems(config.FABRIC_TYPE)
        self.combo_type.setCurrentText(self.Project.Samples[str(self.stage)].fabric_type)
        self.combo_type.currentIndexChanged.connect(self.combo_type_changed)
        self.grid.addWidget(self.combo_type,self.row_counter,2,1,2)
        self.row_counter += 1
        
        self.create_search()
        
        label_selected = QLabel("<b>Selected:</b>")
        self.grid.addWidget(label_selected,self.row_counter,0,1,5)

        label_not_selected = QLabel("<b>Not Selected:</b>")
        self.grid.addWidget(label_not_selected,self.row_counter,5,1,5)
        self.row_counter += 1
        
        self.create_selected_checkboxs_tests()
        self.create_non_selected_checkboxs_tests()

        if str(self.stage) != str(self.Project.number_sample):
            button_next = QPushButton("Next")
            button_next.clicked.connect(self.button_next_clicked)
            self.grid.addWidget(button_next,1000,8,1,2)
        else:
            button_apply = QPushButton("Apply")
            button_apply.clicked.connect(self.button_apply_clicked)
            self.grid.addWidget(button_apply,1000,8,1,2)
            
    def combo_type_changed(self):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        self.close()
        self.__init__(self.stage,self.Project)
            
    def create_search(self):
        uni_name_method_name = set()
        current_type = self.combo_type.currentText()
        for test in self.Tests.tests:
            tests_type = test.fabric_type.split(",")
            if (test.fabric_type == "All") or (current_type in tests_type):
                name = test.name
                method = test.method
                name_method = name + "{" + method + "}"
                method_name = method + "{" + name + "}"
                uni_name_method_name.add(name_method)
                uni_name_method_name.add(method_name)
        
        self.search = QLineEdit()
        completer_name_method_name = QCompleter(list(uni_name_method_name))
        completer_name_method_name.setCaseSensitivity(Qt.CaseInsensitive)
        self.search.setCompleter(completer_name_method_name)
        self.grid.addWidget(self.search,self.row_counter,0,1,5)
        button_search_add = QPushButton("Add")
        button_search_add.clicked.connect(lambda: self.button_search_clicked(True))
        self.grid.addWidget(button_search_add,self.row_counter,6,1,2)
        button_search_remove = QPushButton("Remove")
        button_search_remove.clicked.connect(lambda: self.button_search_clicked(False))
        self.grid.addWidget(button_search_remove,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_search_clicked(self,add:bool):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        search_text = self.search.text()
        name,method = helper.enforce_name_method(self.Tests,search_text)
        if name == -1:
            return
        test = helper.find_test(self.Tests,name,method)
        tests = list()
        for type_test in self.Project.Samples[self.stage].Tests.values():
            tests += type_test
        if add:
            if test not in tests:
                self.Project.Samples[self.stage].Tests[test.type].append(test)
        else:
            if test in tests:
                self.Project.Samples[self.stage].Tests[test.type].remove(test)
        self.close()
        self.__init__(self.stage,self.Project)
        
    def create_selected_checkboxs_tests(self):
        selected_row_counter = self.row_counter
        for test_type in self.Project.Samples[self.stage].Tests:
            test_type_tests = self.Project.Samples[self.stage].Tests[test_type]
            self.create_label_test_type(test_type,selected_row_counter)
            selected_row_counter += 1
            for test in test_type_tests:
                self.create_selected_checkbox_test(test,selected_row_counter)
                selected_row_counter += 1
                
    def create_label_test_type(self,test_type,row_counter):
        label_test_type = QLabel("<b>" + test_type + ":</b>")
        self.grid.addWidget(label_test_type,row_counter,0,1,5)

    def create_selected_checkbox_test(self,test,counter:int):
        name = test.name
        method = test.method
        name_method = name + "{" + method + "}"
        checkbox = QCheckBox(name_method)
        checkbox.setChecked(True)
        checkbox.stateChanged.connect(lambda: self.selected_checkbox_test_clicked(checkbox))
        self.grid.addWidget(checkbox,counter,0,1,5)

    def selected_checkbox_test_clicked(self,checkbox):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        name,method = helper.name_method_split(checkbox.text())
        test = helper.find_test(self.Tests,name,method)
        self.Project.Samples[self.stage].Tests[test.type].remove(test)
        self.close()
        self.__init__(self.stage,self.Project)
        
    def create_non_selected_checkboxs_tests(self): 
        non_selected_row_counter = self.row_counter
        current_type = self.combo_type.currentText()
        tests = list()
        for type_test in self.Project.Samples[self.stage].Tests.values():
            tests += type_test        
        for test in self.Tests.tests:
            if test in tests:
                continue
            if (test.fabric_type != "All") and (current_type not in test.fabric_type.split(",")):
                continue
            self.create_non_selected_checkbox_test(test,non_selected_row_counter)
            non_selected_row_counter += 1
            
    def create_non_selected_checkbox_test(self,test,counter:int):
        name = test.name
        method = test.method
        name_method = name + "{" + method + "}"
        checkbox = QCheckBox(name_method)
        checkbox.stateChanged.connect(lambda: self.non_selected_checkbox_test_clicked(checkbox))
        self.grid.addWidget(checkbox,counter,5,1,5)
        
    def non_selected_checkbox_test_clicked(self,checkbox):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        name,method = helper.name_method_split(checkbox.text())
        test = helper.find_test(self.Tests,name,method)
        self.Project.Samples[self.stage].Tests[test.type].append(test)
        self.close()
        self.__init__(self.stage,self.Project)

    def button_next_clicked(self):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        self.close()
        self.__init__(str(int(self.stage)+1),self.Project)

    def button_apply_clicked(self):
        type_ = self.combo_type.currentText()
        self.Project.Samples[self.stage].fabric_type = type_
        project_json = self.Project.to_dict()
        with open(config.BASE_DIRECTORY_PATH + 'Data/Project_tem.json','w',newline='') as fp:
            json.dump(project_json,fp)
        self.close()

        
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
    logging.info("Entered create_project.py")
    app = QApplication(sys.argv)
    ex = window("sample_number")
    #logging.info("Leaving create_project.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
