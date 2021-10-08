import DataClasses
import logging
import helper,json,sys,config
import collections

import view_project_sample

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QComboBox,QCheckBox,QCompleter,QPushButton,QLineEdit,QInputDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "Edit Porject Scope"
FONT_MID = QFont("Times New Roman", 16)


class window(QMainWindow):
    def __init__(self,Project,bool_add = False,sample_id = 0):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()
        
        self.sample_id = sample_id
        
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
                self.Tests = DataClasses.Tests.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot load data from Tests,closing 'Create Project Sample'")
            self.close()
        
        if Project == "read":
            try:
                with open(config.BASE_DIRECTORY_PATH + 'Data/Project_tem.json') as json_file:  
                    self.Project = DataClasses.Project.from_dict(json.load(json_file))
            except:
                logging.fatal("Cannot load data from /Data/Project_tem.json,closing 'Add New Lab'")
                self.close()
        else:
            self.Project = Project
        self.bool_add = bool_add

        self.row_counter = 0
        self.initUI()
        
    def initUI(self):
        if self.bool_add:
            self.create_add_sample(self.sample_id)
        else:
            self.create_edit_samples()
            self.create_menu_add_sample()
            self.create_save()
        
        self.show_window()
        
    #copied from create_project.py START
    def create_add_sample(self,sample_id):
        self.sample_id = sample_id
        label_sample_text = "<b>Sample " + str(self.sample_id) + ":</b>"
        label_sample = QLabel(label_sample_text)
        self.grid.addWidget(label_sample,self.row_counter,0,1,10)
        self.row_counter += 1
        
        label_type = QLabel("Fabric type:")
        self.grid.addWidget(label_type,self.row_counter,0,1,2)

        self.combo_type = QComboBox()
        self.combo_type.addItems(config.FABRIC_TYPE)
        self.combo_type.setCurrentText(self.Project.Samples[str(self.sample_id)].fabric_type)
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

        button_finish_text = "Update"
        button_finish = QPushButton(button_finish_text)
        button_finish.clicked.connect(self.button_finish_clicked)
        self.grid.addWidget(button_finish,100000,8,1,2)
        self.row_counter += 1
        
    def combo_type_changed(self):
        type_ = self.combo_type.currentText()
        self.Project.Samples[str(self.sample_id)].fabric_type = type_
        self.close()
        self.__init__(self.Project,True,self.sample_id)
            
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
        self.Project.Samples[str(self.sample_id)].fabric_type = type_
        search_text = self.search.text()
        name,method = helper.enforce_name_method(self.Tests,search_text)
        if name == -1:
            return
        test = helper.find_test(self.Tests,name,method)
        tests = list()
        for type_test in self.Project.Samples[str(self.sample_id)].Tests.values():
            tests += type_test
        if add:
            if test not in tests:
                self.Project.Samples[str(self.sample_id)].Tests[test.type].append(test)
        else:
            if test in tests:
                self.Project.Samples[str(self.sample_id)].Tests[test.type].remove(test)
        self.close()
        self.__init__(self.Project,True,self.sample_id)
        
    def create_selected_checkboxs_tests(self):
        selected_row_counter = self.row_counter
        for test_type in self.Project.Samples[str(self.sample_id)].Tests:
            test_type_tests = self.Project.Samples[str(self.sample_id)].Tests[test_type]
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
        self.Project.Samples[str(self.sample_id)].fabric_type = type_
        name,method = helper.name_method_split(checkbox.text())
        test = helper.find_test(self.Tests,name,method)
        self.Project.Samples[str(self.sample_id)].Tests[test.type].remove(test)
        self.close()
        self.__init__(self.Project,True,self.sample_id)
        
    def create_non_selected_checkboxs_tests(self):
        non_selected_row_counter = self.row_counter
        current_type = self.combo_type.currentText()
        tests = list()
        for type_test in self.Project.Samples[str(self.sample_id)].Tests.values():
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
        self.Project.Samples[str(self.sample_id)].fabric_type = type_
        name,method = helper.name_method_split(checkbox.text())
        test = helper.find_test(self.Tests,name,method)
        self.Project.Samples[str(self.sample_id)].Tests[test.type].append(test)
        self.close()
        self.__init__(self.Project,True,self.sample_id)
        
    #copied from create_project.py END
    
    def button_finish_clicked(self):
        type_ = self.combo_type.currentText()
        self.Project.Samples[str(self.sample_id)].fabric_type = type_
        self.close()
        self.__init__(self.Project)
        
    def create_edit_samples(self):
        for sample in self.Project.Samples.values():
            self.edit_sample(sample)
            
    def edit_sample(self,sample):
        label_sample_text = "Sample " + str(sample.sample_id) + "(" + sample.fabric_type + ")"
        label_sample = QLabel(label_sample_text)
        self.grid.addWidget(label_sample,self.row_counter,0,1,4)
        
        button_drop = QPushButton("Drop")
        button_drop.clicked.connect(lambda: self.button_drop_clicked(sample.sample_id))
        self.grid.addWidget(button_drop,self.row_counter,4,1,2)
        
        button_edit = QPushButton("Edit")
        button_edit.clicked.connect(lambda: self.button_edit_clicked(sample.sample_id))
        self.grid.addWidget(button_edit,self.row_counter,6,1,2)
        
        button_view = QPushButton("View")
        button_view.clicked.connect(lambda: self.button_view_clicked(sample.sample_id))
        self.grid.addWidget(button_view,self.row_counter,8,1,2)
        
        self.row_counter += 1

    def button_edit_clicked(self,sample_id):
        self.close()
        self.__init__(self.Project,True,sample_id)
        
    def button_drop_clicked(self,sample_id: int):
        text,ok = QInputDialog.getText(self,"Drop Sample " + str(sample_id),"Type 'DROP' to confirm!")
        if ok and (text == "DROP"):
            self.Project.Samples.pop(str(sample_id))
            self.close()
            self.__init__(self.Project)
            
    def button_view_clicked(self,sample_id:int):
        self.view_sample = view_project_sample.window(sample_id,self.Project)
        
    def create_menu_add_sample(self):
        new_sample_id = len(self.Project.Samples) + 1
        label_sample_text = "<b>Add Sample: </b>"
        label_sample = QLabel(label_sample_text)
        self.grid.addWidget(label_sample,self.row_counter,0,1,4)
        
        add_sample_list = [str(x) for x in range(1,31)]
        for sample in self.Project.Samples.values():
            add_sample_list.remove(str(sample.sample_id))
        self.combo_add_sample_number = QComboBox()
        self.combo_add_sample_number.addItems(add_sample_list)
        self.grid.addWidget(self.combo_add_sample_number,self.row_counter,6,1,2)
        
        button_add = QPushButton("Add")
        button_add.clicked.connect(self.button_add_clicked)
        self.grid.addWidget(button_add,self.row_counter,8,1,2)
        #button_back = QPushButton("Back")
        #button_back.clicked.connect(self.button_back_clicked)
        #self.grid.addWidget(button_back,self.row_counter,6,1,2)
        
        self.row_counter += 1
        
    def button_add_clicked(self):
        new_sample_id = str(self.combo_add_sample_number.currentText())
        sample = DataClasses.Sample(sample_id = new_sample_id,fabric_type = "Knit",Tests = {"PSR":list(),"SPC":list(),"Performance":list()})
        self.Project.Samples[str(new_sample_id)] = sample
        self.Project.Samples = collections.OrderedDict(sorted(self.Project.Samples.items()))
        self.close()
        self.__init__(self.Project,True,new_sample_id)
        
    def create_save(self):
        button_save = QPushButton("Save and quit")
        button_save.clicked.connect(self.button_save_clicked)
        self.grid.addWidget(button_save,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_save_clicked(self):
        project_json = self.Project.to_dict()
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Project_tem.json','w',newline='') as fp:
                json.dump(project_json,fp)
            self.close()
        except:
            logging.fatal("Cannot save the updated Project Scope")
            

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
    logging.info("Entered edit_project.py")
    app = QApplication(sys.argv)
    ex = window("read")
    #logging.info("Leaving create_project.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()