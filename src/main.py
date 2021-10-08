import DataClasses
import logging
import config,json,xlsxwriter
import pandas as pd

import new_test,view_all_tests,create_project,view_project,add_new_lab,view_lab,missing_lab,edit_project,calculate_without_deduction,view_score,import_data,generate_error_input_form,calculate_with_deduction
import appendix_10,appendix_1,appendix_3,appendix_4,appendix_6_7,appendix_8_9,appendix_2

import sys,os,os.path,shutil,time
from PyQt5.QtWidgets import QDialog,QLabel,QPushButton,QFileDialog,QComboBox,QApplication,QMainWindow,QWidget,QScrollArea,QGridLayout,QInputDialog,QMessageBox
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

WINDOW_TITLE = "Correlation Program"
FONT_BIG = QFont("Times New Roman", 24)
FONT_MID = QFont("Times New Roman", 14)
FONT_SMALL = QFont("Times New Roman", 12)

class window(QMainWindow):
    def __init__(self):    
        super().__init__()
        self.widget = QWidget()
        self.scroll = QScrollArea()
        self.grid = QGridLayout()
        self.row_counter = 0
        
        try:
            with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:  
                self.Labs = DataClasses.Labs.from_dict(json.load(json_file))
                self.success_read_labs = True
        except:
            logging.fatal("Cannot load data from Labs,diabled 'View Lab function'")
            self.success_read_labs = False

        self.initUI()

    def initUI(self):
        button_refresh = QPushButton("REFRESH")
        button_refresh.clicked.connect(self.button_refresh_clicked)
        self.grid.addWidget(button_refresh,0,8,1,2)
        self.create_setup()
        self.create_dataCollection()
        self.create_scoreCalculation()
        self.create_appendixGeneration()
        self.create_helper()
   
        self.show_window()
        
    def button_refresh_clicked(self):
        self.close()
        self.__init__()

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

    def create_setup(self):
        label_setup = QLabel("<b>Setup</b>")
        label_setup.setFont(FONT_BIG)
        self.grid.addWidget(label_setup,self.row_counter,0,1,8)
        self.row_counter += 1

        self.create_add_new_test()
        self.create_project()
        self.row_counter += 1

    def create_add_new_test(self):
        label_add_new_test = QLabel("<b>Add new test:</b>")
        label_add_new_test.setFont(FONT_MID)
        self.grid.addWidget(label_add_new_test,self.row_counter,0,1,10)
        self.row_counter += 1

        label_add_new_test_input = QLabel("Input:\t\t1). Test code, test name,testing results and testing requirement")
        label_add_new_test_input.setFont(FONT_SMALL)
        self.grid.addWidget(label_add_new_test_input,self.row_counter,0,1,10)
        self.row_counter += 1

        label_add_new_test_function = QLabel("Functionality:\t1). Add new test to the database for test selection.")
        label_add_new_test_function.setFont(FONT_SMALL)
        self.grid.addWidget(label_add_new_test_function,self.row_counter,0,1,10)

        button_add_new_test = QPushButton("Add")
        button_add_new_test.clicked.connect(self.add_new_test_clicked)
        self.grid.addWidget(button_add_new_test,self.row_counter,6,1,2)
        
        with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:
            Tests = DataClasses.Tests.from_dict(json.load(json_file))
        tests_list = list()
        for test in Tests.tests:
            name_method = test.name + "{" + test.method + "}"
            tests_list.append(name_method)
        self.combo_edit_test = QComboBox()
        tests_list.sort()
        self.combo_edit_test.addItems(tests_list)
        self.grid.addWidget(self.combo_edit_test,self.row_counter-1,4,1,5)
        
        button_edit_test = QPushButton("Edit")
        button_edit_test.clicked.connect(self.edit_test_clicked)
        self.grid.addWidget(button_edit_test,self.row_counter-1,9,1,1)

        button_view_all_test = QPushButton("View")
        button_view_all_test.clicked.connect(self.view_all_test_clicked)
        self.grid.addWidget(button_view_all_test,self.row_counter,8,1,2)
        self.row_counter += 1

    def add_new_test_clicked(self):
        logging.info("Entering new_test.py")
        self.window_new_test = new_test.window()
        
    def edit_test_clicked(self):
        logging.info("Entering new_test.py")
        name_method = self.combo_edit_test.currentText()
        self.window_edit_test = new_test.window(name_method)

    def view_all_test_clicked(self):
        logging.info("Entering view_all_tests.py")
        self.window_view_all_tests = view_all_tests.window()

    def create_project(self):
        new_project_text = "<b>Create New Project: "
        bool_created_project = os.path.exists(config.BASE_DIRECTORY_PATH + "Data/Project_tem.json")
        self.bool_confirmed_project = os.path.exists(config.BASE_DIRECTORY_PATH + "Data/Project.json")
        if bool_created_project:
            new_project_text += "(Created)"
        else:
            new_project_text += "(Not created)"
        new_project_text += "</b>"
        label_create_new_project = QLabel(new_project_text)
        label_create_new_project.setFont(FONT_MID)
        self.grid.addWidget(label_create_new_project,self.row_counter,0,1,5)
        self.row_counter += 1

        label_create_new_project_input = QLabel("Input:\t\t1). Number of sample and the corresponding fabric type\n\t\t2). Select tests for each sample from the corresponding package (Performance protocol+PSR+SPC)")
        label_create_new_project_input.setFont(FONT_SMALL)
        self.grid.addWidget(label_create_new_project_input,self.row_counter,0,1,10)
        self.row_counter += 1

        label_create_new_project_function = QLabel("Functionality:\t1). Automatically generate the data templates for storing the test result")
        label_create_new_project_function.setFont(FONT_SMALL)
        self.grid.addWidget(label_create_new_project_function,self.row_counter,0,1,7)

        if not bool_created_project:
            button_create_project = QPushButton("Create")
            button_create_project.clicked.connect(self.create_project_clicked)
            self.grid.addWidget(button_create_project,self.row_counter,4,1,2)
            
        elif not self.bool_confirmed_project:
            button_edit_project = QPushButton("Edit")
            button_edit_project.clicked.connect(self.edit_project_clicked)
            self.grid.addWidget(button_edit_project,self.row_counter,4,1,2)
            
            button_confirm_project = QPushButton("Confirm")
            button_confirm_project.clicked.connect(self.confirm_project_clicked)
            self.grid.addWidget(button_confirm_project,self.row_counter,6,1,2)

        if bool_created_project:
            button_view_project = QPushButton("View")
            button_view_project.clicked.connect(self.view_project_clicked)
            self.grid.addWidget(button_view_project,self.row_counter,8,1,2)

        self.row_counter += 1

    def create_project_clicked(self):
        logging.info("Entering create_project.py")
        self.window_create_project = create_project.window("sample_number")
        
    def edit_project_clicked(self):
        logging.info("Editing project")
        self.window_edit_project = edit_project.window("read")
        
    def confirm_project_clicked(self):
        logging.info("Confirming project, note that the project scope can no longer be changed again in the future")
        text,ok = QInputDialog.getText(self,"Confirm Project Scope","Type 'CONFIRM' to confirm the project scope.\nNote that the project scope can no longer be changed after confirmation!")
        if ok and (text == "CONFIRM"):
            try:
                with open(config.BASE_DIRECTORY_PATH + 'Data/Project_tem.json') as json_file:  
                    Project = DataClasses.Project.from_dict(json.load(json_file))
            except:
                logging.fatal("Cannot load data from Project,failed to create result template")
            self.create_samples_result_template(Project)
            shutil.copyfile(config.BASE_DIRECTORY_PATH + "Data/Project_tem.json", config.BASE_DIRECTORY_PATH + "Data/Project.json")
            time.sleep(0.5)
            self.close()
            self.__init__()
            
    #slow version but doesnt matter if the size is small O(n^2) , could be reduced to O(nlogn)
    def create_samples_result_template(self,Project):
        for sample in Project.Samples.values():
            self.create_sample_result_template(sample)
                
    def create_sample_result_template(self,sample: DataClasses.Sample):
        os.mkdir(config.BASE_DIRECTORY_PATH + "Data/Result/Sample " + str(sample.sample_id))
        tests = list()
        for type_tests in sample.Tests.values():
            tests += type_tests
        for test in tests:
            self.create_csv_template(test,sample.sample_id)

    def create_csv_template(self,test:DataClasses.Test,sample_id: int):
        method = test.method
        name = test.name
        method_name = method + "_" + name
        filepath = config.BASE_DIRECTORY_PATH + 'Data/Result/Sample ' + str(sample_id) + "/" + method_name + ".xlsx"
        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet()
        template_format = workbook.add_format({'num_format': '@'})
        column = 0
        for template in config.RESULT_TEMPLATE:
            worksheet.write(0,column,template,template_format)
            column += 1
        for requirement in test.Requirements:
            worksheet.write(0,column,"requirement_" + requirement,template_format)
            column += 1
        for result in test.Results:
            worksheet.write(0,column,"result_" + result.name,template_format)
            column += 1
        worksheet.write(0,column,"Result Rating",template_format)
        workbook.close()

    def view_project_clicked(self):
        logging.info("Entering view_project.py")
        self.window_view_project = view_project.window()

    def create_dataCollection(self):
        label_text = "<b>Data Collection"
        if not self.bool_confirmed_project:
            label_text += " (RESTRICTED)"
        label_text += "</b>"
        label_data_collection = QLabel(label_text)
        
        label_data_collection.setFont(FONT_BIG)
        self.grid.addWidget(label_data_collection,self.row_counter,0,1,10)
        self.row_counter += 1

        self.add_new_lab()
        self.import_lab()
        #self.create_view_data()

    def add_new_lab(self):
        label_new_lab = QLabel("<b>Add new lab:</b> (and generate lab input excel)")
        label_new_lab.setFont(FONT_MID)
        self.grid.addWidget(label_new_lab,self.row_counter,0,1,10)
        self.row_counter += 1
        label_new_lab_input = QLabel("Input:\t\t1). Lab info (Lab ID,Group,etc.)\n\t\t2). Assigned test samples\n\t\t3). Capability selection for each assigned sample.")
        label_new_lab_input.setFont(FONT_SMALL)
        self.grid.addWidget(label_new_lab_input,self.row_counter,0,1,10)
        self.row_counter += 1
        label_new_lab_function = QLabel("Functionality:\t1). Store lab info and data for appendix generation\n\t\t2). Create lab template for storing test result.")
        label_new_lab_function.setFont(FONT_SMALL)
        self.grid.addWidget(label_new_lab_function,self.row_counter,0,1,8)
        button_add_new_lab_add = QPushButton("Add")
        button_add_new_lab_add.clicked.connect(self.button_add_new_lab_add_clicked)
        self.grid.addWidget(button_add_new_lab_add,self.row_counter,5,1,2)
        if self.success_read_labs:
            self.combo_add_new_lab_view_labID = QComboBox()
            labsID = self.get_labsID()
            self.combo_add_new_lab_view_labID.addItems(labsID)
            self.grid.addWidget(self.combo_add_new_lab_view_labID,self.row_counter,7,1,1)
            button_add_new_lab_view = QPushButton("View")
            button_add_new_lab_view.clicked.connect(self.button_add_new_lab_view_clicked)
            self.grid.addWidget(button_add_new_lab_view,self.row_counter,8,1,2)
        self.row_counter += 1
        
        label_edit_lab = QLabel("<b>Edit Existing Lab: </b>(and generate updated lab input excel)")
        label_edit_lab.setFont(FONT_MID)
        self.grid.addWidget(label_edit_lab,self.row_counter,0,1,5)
        if self.success_read_labs:
            self.combo_add_new_lab_edit_labID = QComboBox()
            labsID = self.get_labsID()
            self.combo_add_new_lab_edit_labID.addItems(labsID)
            self.grid.addWidget(self.combo_add_new_lab_edit_labID,self.row_counter,7,1,1)
            button_add_new_lab_edit = QPushButton("Edit")
            button_add_new_lab_edit.clicked.connect(self.button_add_new_lab_edit_clicked)
            self.grid.addWidget(button_add_new_lab_edit,self.row_counter,8,1,2)
            
        self.row_counter += 1
        
    def button_add_new_lab_add_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering add_new_lab.py")
        self.window_add_new_lab = add_new_lab.window("Info")
        
    def get_labsID(self):
        labs_id = list()
        for labs in self.Labs.Groups.values():          
            for lab in labs.values():
                labs_id.append(lab.ID)
        return sorted(labs_id)
    
    def button_add_new_lab_view_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering view_lab.py")
        labID = self.combo_add_new_lab_view_labID.currentText()
        self.window_view_lab = view_lab.window(labID)
        
    def button_add_new_lab_edit_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering edit_lab.py")
        labID = self.combo_add_new_lab_edit_labID.currentText()
        with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:
            Labs = DataClasses.Labs.from_dict(json.load(json_file))
        for labs_group in Labs.Groups.values():
            for lab in labs_group.values():
                if lab.ID == labID:
                    lab_edit = lab
                    sample_list = list()
                    for sample in lab.Samples.values():
                        sample_list.append(sample.sample_id)
                    self.window_edit_lab = add_new_lab.window(stage="0",Lab = lab_edit,sample_list = sample_list)
                    return

    def import_lab(self):
        label_import = QLabel("<b>Import test result: </b>")
        label_import.setFont(FONT_MID)
        self.grid.addWidget(label_import,self.row_counter,0,1,10)
        self.row_counter += 1

        label_import_input = QLabel("Input:\t\t1). Select the excel file returned from a lab")
        label_import_input.setFont(FONT_SMALL)
        self.grid.addWidget(label_import_input,self.row_counter,0,1,10)
        self.row_counter += 1

        label_import_function = QLabel("Functionality:\t1). Import the returned test result to corresponding data file\n\t\t2). Data validation and backup")
        label_import_function.setFont(FONT_SMALL)
        self.grid.addWidget(label_import_function,self.row_counter,0,1,10)

        button_import = QPushButton("Import")
        button_import.clicked.connect(self.import_clicked)
        self.grid.addWidget(button_import,self.row_counter,8,1,2)
        self.row_counter += 1

    def import_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"Select file", "c:\\",'XLSX(*.xlsx)', options=options)
        if not fileName:
            return
        #print(fileName)
        self.window_import_data = import_data.window(fileName)

    def create_appendixGeneration(self):
        label_data_analysis = QLabel("<b>Appendix Generation</b>")
        label_data_analysis.setFont(FONT_BIG)
        self.grid.addWidget(label_data_analysis,self.row_counter,0,1,10)
        self.row_counter += 1

        self.create_generate_report()

    def create_view_data(self):
        label_view = QLabel("<b>View test result:</b>")
        label_view.setFont(FONT_MID)
        self.grid.addWidget(label_view,self.row_counter,0,1,10)
        self.row_counter += 1

        label_view_des = QLabel("Description:\t  It provides a flexible selection. You could glance the statistics summary of whatever combination you like, as long as they all have the same test code")
        label_view_des.setFont(FONT_SMALL)
        self.grid.addWidget(label_view_des,self.row_counter,0,1,10)
        self.row_counter += 1

        label_view_code = QLabel("Test code")
        label_view_code.setFont(FONT_SMALL)
        self.grid.addWidget(label_view_code,self.row_counter,1,1,4)
        self.row_counter += 1

        self.select_view_code = QComboBox()
        tests_str = self.get_tests_str()
        self.select_view_code.addItems(tests_str)
        self.grid.addWidget(self.select_view_code,self.row_counter,1,1,4)

        button_view_test_result = QPushButton("View")
        #button_view_test_result.clicked.connect(self.button_view_test_result_clicked)
        self.grid.addWidget(button_view_test_result,self.row_counter,8,1,2)
        self.row_counter += 1

    def get_tests_str(self):
        filelist = []
        for root, dirs, files in os.walk(config.BASE_DIRECTORY_PATH + 'Data/Result'):
            for file in files:
                filelist.append(file[:-4])
        return filelist

    def button_view_test_result_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        code_name = self.select_view_code.currentText()
        filename = code_name + ".csv"
        run_script = 'python view_test_result.py "' + filename + '"'
        os.system(run_script)

    def create_generate_report(self):

        label_report_note_text = "Note: the corresponding appendix will be generated in the Data/Report folder.\n     Please click UPDATE everytime after you changed the result/error data"
        label_report_note = QLabel(label_report_note_text)
        label_report_note.setFont(FONT_SMALL)
        self.grid.addWidget(label_report_note,self.row_counter,0,1,6)
        button_check_missing = QPushButton("Check missing")
        button_check_missing.clicked.connect(self.button_check_missing_clicked)
        self.grid.addWidget(button_check_missing,self.row_counter,7,1,1)
        button_update_data_report = QPushButton("UPDATE")
        button_update_data_report.clicked.connect(self.button_update_data_report_clicked)
        self.grid.addWidget(button_update_data_report,self.row_counter,8,1,2)
        self.row_counter += 1
        
        for appendix_info in config.APPENDIX:
            self.create_appendix_row(appendix_info)
            
    def button_update_data_report_clicked(self):
        self.window_appendix_calculate_without_deduction = calculate_without_deduction.main()
        self.window_appendix_generate_error_input_form = generate_error_input_form.main()
        self.window_appendix_calculate_with_deduction = calculate_with_deduction.main()
            
    def button_check_missing_clicked(self):
        with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
            Project = DataClasses.Project.from_dict(json.load(json_file))
        miss = False
        for sample in Project.Samples.values():
            sample_id = sample.sample_id
            for test_type,tests in sample.Tests.items():
                for test in tests:
                    filepath = config.BASE_DIRECTORY_PATH + "Data/Result/Sample " + str(sample_id) + "/" + test.method + "_" + test.name + ".xlsx"
                    df = pd.read_excel(filepath,index_col = False,converters = {'Lab ID':str})
                    count = df.isnull().values.sum()
                    if (count):
                        miss = True
                        text = str(count) + " value(s) is missing in Sample " + str(sample_id) + " " + test.method + "_" + test.name
                        dlg = QDialog()
                        dlg.setWindowTitle(text)
                        dlg.setFixedWidth(600)
                        dlg.setFixedHeight(100)
                        dlg.exec()
        if not miss:
            dlg = QDialog()
            dlg.setWindowTitle("No data is missing")
            dlg.setFixedWidth(600)
            dlg.setFixedHeight(100)
            dlg.exec()
        
    def create_appendix_row(self,appendix_info):
        appendix_number = int(appendix_info[0])
        appendix_name = appendix_info[1]

        label_appendix_number_text = "Appendix " + str(appendix_number) + ": "
        label_appendix_number = QLabel(label_appendix_number_text)
        self.grid.addWidget(label_appendix_number,self.row_counter,0,1,1)

        label_appendix_name = QLabel(appendix_name)
        self.grid.addWidget(label_appendix_name,self.row_counter,1,1,2)

        button_generate_appendix = QPushButton("Generate Appendix " + str(appendix_number))
        button_generate_appendix.clicked.connect(lambda: self.button_generate_appendix_clicked(appendix_number))
        self.grid.addWidget(button_generate_appendix,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_generate_appendix_clicked(self,number:int):
        try:
            logging.info("Entering 'appendix_" + str(number) + ".py'")
            if (number == 1):
                appendix_1.main()
            elif (number == 2):
                appendix_2.main()
            elif (number == 3):
                appendix_3.main()
            elif (number == 4):
                appendix_4.main()
            elif (number == 6):
                appendix_6_7.main(True)
            elif (number == 7):
                appendix_6_7.main(False)
            elif (number == 8):
                appendix_8_9.main(True)
            elif (number == 9):
                appendix_8_9.main(False)
            elif (number == 10):
                appendix_10.main()
            dlg = QDialog(self)
            dlg.setWindowTitle("GENERATED APPENDIX " + str(number) + "!")
            dlg.setFixedWidth(300)
            dlg.setFixedHeight(50)
            dlg.exec()
        except:
            dlg = QDialog(self)
            dlg.setWindowTitle("FAILED!")
            dlg.setFixedWidth(300)
            dlg.setFixedHeight(50)
            dlg.exec()
        
    def create_scoreCalculation(self):
        label_calculate_score = QLabel("<b>Score Calculation</b>")
        label_calculate_score.setFont(FONT_BIG)
        self.grid.addWidget(label_calculate_score,self.row_counter,0,1,10)
        self.row_counter += 1
        
        self.create_calculate_without_deduction()
        self.create_generate_error_form()
        self.create_calculate_with_deduction()
        self.create_view_score()
        
    def create_calculate_without_deduction(self):
        label_calculate_without_deduction = QLabel("<b>Calculate score without deduction:</b>")
        label_calculate_without_deduction.setFont(FONT_MID)
        self.grid.addWidget(label_calculate_without_deduction,self.row_counter,0,1,10)
        
        button_calculate_without_deduction = QPushButton("Calculate")
        button_calculate_without_deduction.clicked.connect(self.button_calculate_without_deduction_clicked)
        self.grid.addWidget(button_calculate_without_deduction,self.row_counter,8,1,2)
        
        self.row_counter += 1
        
    def button_calculate_without_deduction_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering calculate_without_deduction.py")
        self.window_calculate_without_deduction = calculate_without_deduction.main()
        
    def create_generate_error_form(self):
        label_generate_error_form = QLabel("<b>Generate error form:</b> (By lab)")
        label_generate_error_form.setFont(FONT_MID)
        self.grid.addWidget(label_generate_error_form,self.row_counter,0,1,10)
        self.row_counter += 1
        label_generate_error_form_function = QLabel("Note:\t\t1). Generate error input form (xlsx)\n\t\t2). Just save in the same directory")
        label_generate_error_form_function.setFont(FONT_SMALL)
        self.grid.addWidget(label_generate_error_form_function,self.row_counter,0,1,10)
        button_generate_error_form = QPushButton("Generate")
        button_generate_error_form.clicked.connect(self.button_generate_error_form_clicked)
        self.grid.addWidget(button_generate_error_form,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_generate_error_form_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering generate_error_input_form.py")
        self.window_generate_error_input_form_calculate_without_deduction = calculate_without_deduction.main()
        self.window_generate_error_input_form = generate_error_input_form.main()
        
    def create_calculate_with_deduction(self):
        label_calculate_with_deduction = QLabel("<b>Calculate score with deduction:</b>")
        label_calculate_with_deduction.setFont(FONT_MID)
        self.grid.addWidget(label_calculate_with_deduction,self.row_counter,0,1,10)
        
        button_calculate_with_deduction = QPushButton("Calculate")
        button_calculate_with_deduction.clicked.connect(self.button_calculate_with_deduction_clicked)
        self.grid.addWidget(button_calculate_with_deduction,self.row_counter,8,1,2)
        
        self.row_counter += 1
        
    def button_calculate_with_deduction_clicked(self):
        logging.info("Entering calculate_with_deduction.py")
        self.window_calculate_with_deduction_calculate_without_deduction = calculate_without_deduction.main()
        self.window_calculate_with_deduction_generate_error_input_form = generate_error_input_form.main()
        self.window_calculate_with_deduction = calculate_with_deduction.main()
        
    def create_view_score(self):
        label_view_score = QLabel("<b>View Score:</b>")
        label_view_score.setFont(FONT_MID)
        self.grid.addWidget(label_view_score,self.row_counter,0,1,2)
        
        button_view_score = QPushButton("View")
        button_view_score.clicked.connect(self.button_view_score_clicked)
        self.grid.addWidget(button_view_score,self.row_counter,8,1,2)

        self.row_counter += 1
        
    def button_view_score_clicked(self):
        if not self.bool_confirmed_project:
            message = QMessageBox()
            message.setText("This function is disabled before confirming the project scope!")
            self.message_lab = message.exec_()
            return
        logging.info("Entering view_score.py")
        self.window_view_project = view_score.window()
        
    def create_helper(self):
        label_create_helper = QLabel("<b>Helper Functions</b>")
        label_create_helper.setFont(FONT_BIG)
        self.grid.addWidget(label_create_helper,self.row_counter,0,1,10)
        self.row_counter += 1
        
        self.create_find_missing_data()
        
    def create_find_missing_data(self):
        label_find_missing_data = QLabel("<b>Check Missing Data:</b>")
        label_find_missing_data.setFont(FONT_MID)
        self.grid.addWidget(label_find_missing_data,self.row_counter,0,1,10)
        self.row_counter += 1
        
        label_find_missing_data_description = QLabel("Description:\tFind out the labs without imported data")
        label_find_missing_data_description.setFont(FONT_SMALL)
        self.grid.addWidget(label_find_missing_data_description,self.row_counter,0,1,8)
        
        button_find_missing_data = QPushButton("Search")
        button_find_missing_data.clicked.connect(self.button_find_missing_data_clicked)
        self.grid.addWidget(button_find_missing_data,self.row_counter,8,1,2)
        self.row_counter += 1
        
    def button_find_missing_data_clicked(self):
        self.window_missing_lab = missing_lab.window()
        
def main():
        logging.info("Opened Correlation Program")
        app = QApplication(sys.argv)
        ex = window()
        #logging.info("Leaving Correlation Program\n")
        sys.exit(app.exec())

if __name__ == '__main__':
        main()
