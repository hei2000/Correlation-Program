import DataClasses
import logging
import helper,config,sys,json

from itertools import chain,combinations

from PyQt5.QtWidgets import QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QLabel,QCompleter,QLineEdit,QPushButton,QComboBox,QMessageBox
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "Add new test"
LABEL_WIDTH = 3
LINEEDIT_WIDTH = 5

class window(QMainWindow):
    def __init__(self,name_method = ""):
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()

        self.row_counter = 0
        self.name_method = name_method
        
        self.edit = False
        if name_method != "":
            self.edit = True

        self.initUI()

    def initUI(self):
        #Get setup data
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
                self.Tests = DataClasses.Tests.from_dict(json.load(json_file))
        except:
            logging.fatal("Cannot read Data/Tests.json file, closiong 'Add New Test'")
            self.close()
            #self.Tests = DataClasses.Tests()

        uni_method,uni_name,uni_requirement,uni_result,uni_group = helper.search_completer(self.Tests)
        
        if self.edit:
            self.current_test_name,self.current_test_method = helper.name_method_split(self.name_method)
            self.current_test = helper.find_test(self.Tests,self.current_test_name,self.current_test_method)
  
        self.create_name(uni_name)
        self.create_method(uni_method)
        self.create_fabric_type()
        self.create_test_type()
        self.create_group(uni_group)
        self.create_requirements(uni_requirement)
        self.create_results(uni_result)
        if self.edit:
            self.create_edit_button()
        else:
            self.create_add_button()

        self.show_window()
        
    def create_name(self,uni_name):
        label_name = QLabel("<b>Test name:</b> ")
        self.grid.addWidget(label_name,self.row_counter,0,1,LABEL_WIDTH)
        completer_name = QCompleter(uni_name)
        self.line_name = QLineEdit()
        self.line_name.setCompleter(completer_name)
        if self.edit:
            self.line_name.setText(self.current_test_name)
        self.grid.addWidget(self.line_name,self.row_counter,LABEL_WIDTH,1,LINEEDIT_WIDTH)
        self.row_counter += 1
        
    def create_method(self,uni_method):
        label_method = QLabel("<b>Test method:</b>")
        self.grid.addWidget(label_method,self.row_counter,0,1,LABEL_WIDTH)
        completer_method = QCompleter(uni_method)
        self.line_method = QLineEdit()
        self.line_method.setCompleter(completer_method)
        if self.edit:
            self.line_method.setText(self.current_test_method)
        self.grid.addWidget(self.line_method,self.row_counter,LABEL_WIDTH,1,LINEEDIT_WIDTH)
        self.row_counter += 1
        
    def create_fabric_type(self):
        label_fabric_type = QLabel("<b>Fabric type:</b>")
        self.grid.addWidget(label_fabric_type,self.row_counter,0,1,LABEL_WIDTH)
        self.combo_fabric_type = QComboBox()
        fabric_type_combination = self.get_fabric_type_combination(config.FABRIC_TYPE)
        fabric_type_combination.remove(','.join(config.FABRIC_TYPE))
        fabric_type_combination.remove('')
        self.combo_fabric_type.addItem("All")
        self.combo_fabric_type.addItems(fabric_type_combination)
        if self.edit:
            self.combo_fabric_type.setCurrentText(self.current_test.fabric_type)
        self.grid.addWidget(self.combo_fabric_type,self.row_counter,LABEL_WIDTH,1,LINEEDIT_WIDTH)
        self.row_counter += 1
        
    def get_fabric_type_combination(self,fabric_type):
        list_of_tuple = chain.from_iterable(combinations(fabric_type, r) for r in range(len(fabric_type)+1))
        return_list = list()
        for tuple_combination in list_of_tuple:
            return_list.append(','.join(tuple_combination))
        return return_list
        
    def create_test_type(self):
        label_test_type = QLabel("<b>Type of Test:</b>")
        self.grid.addWidget(label_test_type,self.row_counter,0,1,LABEL_WIDTH)
        self.combo_test_type = QComboBox()
        self.combo_test_type.addItems(config.TYPE_OF_TEST)
        if self.edit:
            self.combo_test_type.setCurrentText(self.current_test.type)
        self.grid.addWidget(self.combo_test_type,self.row_counter,LABEL_WIDTH,1,LINEEDIT_WIDTH)
        self.row_counter += 1
        
    def create_group(self,uni_group):
        label_group = QLabel("<b>Test Group:</b>")
        self.grid.addWidget(label_group,self.row_counter,0,1,LABEL_WIDTH)
        self.line_group = QLineEdit()
        completer_group = QCompleter(uni_group)
        self.line_group.setCompleter(completer_group)
        if self.edit:
            self.line_group.setText(self.current_test.group)
        self.grid.addWidget(self.line_group,self.row_counter,LABEL_WIDTH,1,LINEEDIT_WIDTH)
        self.row_counter += 1
        
    def create_requirements(self,uni_requirement):
        label_requirement = QLabel("<b>Requirement:</b>")
        self.grid.addWidget(label_requirement,self.row_counter,0,1,10)
        self.row_counter += 1
        number_requirement = config.MAX_REQUIREMENT
        self.line_requirements = list()
        for requirement_count in range(1,number_requirement+1):
            self.create_requirement(requirement_count,uni_requirement)
            
    def create_requirement(self,requirement_count:int,uni_requirement):
        label_requirement = QLabel("Requirement " + str(requirement_count) + ":")
        self.grid.addWidget(label_requirement,self.row_counter,0,1,LABEL_WIDTH)
        line_requirement = QLineEdit()
        completer_requirement = QCompleter(uni_requirement)
        line_requirement.setCompleter(completer_requirement)
        if self.edit:
            if requirement_count <= len(self.current_test.Requirements):
                line_requirement.setText(self.current_test.Requirements[requirement_count-1])
        self.line_requirements.append(line_requirement)
        self.grid.addWidget(line_requirement,self.row_counter,5,1,5)
        self.row_counter += 1
        
    def create_results(self,uni_result):
        label_result = QLabel("<b>Result:</b>")
        self.grid.addWidget(label_result,self.row_counter,0,1,10)
        self.row_counter += 1
        number_result = config.MAX_RESULT
        self.line_results = list()
        self.combo_results = list()
        for result_count in range(1,number_result+1):
            self.create_result(result_count,uni_result)
            
    def create_result(self,result_count:int,uni_result):
        label_result = QLabel("Result "+str(result_count) + ":")
        self.grid.addWidget(label_result,self.row_counter,0,1,LABEL_WIDTH)
        
        combo_result = QComboBox()
        combo_result.addItems(config.RESULT_TYPE)
        if self.edit:
            if result_count <= len(self.current_test.Results):
                combo_result.setCurrentText(self.current_test.Results[result_count-1].type)
        self.combo_results.append(combo_result)
        self.grid.addWidget(combo_result,self.row_counter,3,1,2)
        
        line_result = QLineEdit()
        completer_result = QCompleter(uni_result)
        line_result.setCompleter(completer_result)
        if self.edit:
            if result_count <= len(self.current_test.Results):
                line_result.setText(self.current_test.Results[result_count-1].name)
        self.line_results.append(line_result)
        self.grid.addWidget(line_result,self.row_counter,5,1,5)
        self.row_counter += 1
        
    def create_add_button(self):
        button_add = QPushButton("Add")
        button_add.clicked.connect(self.button_add_clicked)
        self.grid.addWidget(button_add,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_add_clicked(self):
        name = self.line_name.text()
        method = self.line_method.text()
        fabric_type = self.combo_fabric_type.currentText()
        test_type = self.combo_test_type.currentText()
        group = self.line_group.text()
        requirements = list()
        for line_requirement in self.line_requirements:
            if line_requirement.text() != "":
                requirements.append(line_requirement.text())
        results_type = list()
        results_name = list()
        for idx,line_result in enumerate(self.line_results):
            if line_result.text() != "":
                results_name.append(line_result.text())
                results_type.append(self.combo_results[idx].currentText())
        
        if not self.data_validation(name,method,requirements,results_name):
            return
        
        test = DataClasses.Test(fabric_type=fabric_type,group=group,type=test_type,name=name,method=method,Requirements=requirements)
        for idx in range(len(results_type)):
            result_name = results_name[idx]
            result_type = results_type[idx]
            result = DataClasses.Result(name=result_name,type=result_type)
            test.Results.append(result)
        #print(test)
        self.Tests.tests.append(test)
        tests_json = self.Tests.to_dict()
        with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json','w',newline="") as fp:
            json.dump(tests_json,fp)
            
        self.close()
        self.__init__()
        
    def create_edit_button(self):
        button_edit = QPushButton("Edit")
        button_edit.clicked.connect(self.button_edit_clicked)
        self.grid.addWidget(button_edit,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_edit_clicked(self):
        name = self.line_name.text()
        method = self.line_method.text()
        fabric_type = self.combo_fabric_type.currentText()
        test_type = self.combo_test_type.currentText()
        group = self.line_group.text()
        requirements = list()
        for line_requirement in self.line_requirements:
            if line_requirement.text() != "":
                requirements.append(line_requirement.text())
        results_type = list()
        results_name = list()
        for idx,line_result in enumerate(self.line_results):
            if line_result.text() != "":
                results_name.append(line_result.text())
                results_type.append(self.combo_results[idx].currentText())
                
        if not self.data_validation(name,method,requirements,results_name):
            return
        
        self.current_test.name = name
        self.current_test.method = method
        self.current_test.fabric_type = fabric_type
        self.current_test.test_type = test_type
        self.current_test.group = group
        self.current_test.Requirements = requirements
        self.current_test.Results = list()
        for idx in range(len(results_type)):
            result_name = results_name[idx]
            result_type = results_type[idx]
            result = DataClasses.Result(name=result_name,type=result_type)
            self.current_test.Results.append(result)
        
        tests_json = self.Tests.to_dict()
        with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json','w',newline="") as fp:
            json.dump(tests_json,fp)
            
        self.close()
        #self.__init__()
        
    def data_validation(self,name,method,requirements,results_name):
        if (len(name) == 0):
            self.pop_message_missing("name")
            return False
        if (len(method) == 0):
            self.pop_message_missing("method")
            return False
        if not self.validate(name):
            self.pop_message_invalid("name")
            return False
        if not self.validate(method):
            self.pop_message_invalid("method")
            return False
        if name[-1] == " ":
            self.pop_message_last_space("name")
            return False
        if method[-1] == " ":
            self.pop_message_last_space("name")
            return False
        if self.edit:
            test_found = helper.find_test(self.Tests,name,method)
            if test_found != -1 and test_found != self.current_test:
                self.pop_message_duplicate()
                return False
        else:
            if helper.find_test(self.Tests,name,method) != -1:
                self.pop_message_duplicate()
                return False
        if (len(requirements) == 0):
            self.pop_message_missing("requirements")
            return False
        if (len(results_name) == 0):
            self.pop_message_missing("results")
            return False
        return True
        
    def pop_message_last_space(self,text):
        message = QMessageBox()
        title = "Last digit space"
        message.setWindowTitle(title)
        message.setText("Last digit of " + text + " cannot be a space!")
        x = message.exec_()
        
    def validate(self,text):
        for char in text:
            if char in config.RESTRICTED_CHAR:
                return False
        return True
        
    def pop_message_invalid(self,type_):
        message = QMessageBox()
        title = "Invalid " + type_ +" input!"
        message.setWindowTitle(title)
        message.setText(' '.join(config.RESTRICTED_CHAR) + " should not to be used in test " + type_)
        x = message.exec_()
        
    def pop_message_missing(self,type_):
        message = QMessageBox()
        title = "Missing " + type_ + "!"
        message.setWindowTitle(title)
        message.setText("You must fill in at least one " + type_ + "!")
        x = message.exec_()
        
    def pop_message_duplicate(self):
        message = QMessageBox()
        title = "Duplication"
        message.setWindowTitle(title)
        text = "Cannot add existing test.\nTest with identical name and method existed already!"
        message.setText(text)
        x = message.exec_()
        
        
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
    logging.info("Entered new_test.py")
    app = QApplication(sys.argv)
    ex = window()
    #logging.info("Leaving new_test.py")
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
