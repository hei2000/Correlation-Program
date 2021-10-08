import xlsxwriter,openpyxl
import DataClasses
import logging
import config,json,os,helper

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
    
ROW_COUNTER = 0
    
def create_info(lab,worksheet,formats):
    global ROW_COUNTER
    worksheet.write(ROW_COUNTER,0,"Lab ID")
    worksheet.write(ROW_COUNTER,1,"Lab Group")
    worksheet.write(ROW_COUNTER,2,"Lab Location")
    worksheet.write(ROW_COUNTER,3,"Lab Name")
    worksheet.write(ROW_COUNTER,4,"Third_party")
    worksheet.write(ROW_COUNTER,5,"Import datetime")
    worksheet.write(ROW_COUNTER,6,"Timeliness")
    worksheet.write(ROW_COUNTER,7,"Revision")
    ROW_COUNTER += 1
    worksheet.write(ROW_COUNTER,0,lab.ID)
    worksheet.write(ROW_COUNTER,1,lab.group)
    worksheet.write(ROW_COUNTER,2,helper.location_city_combine(lab.location,lab.city))
    worksheet.write(ROW_COUNTER,3,lab.fullname)
    worksheet.write(ROW_COUNTER,4,"YES" if lab.third_party else "NO")
    worksheet.write(ROW_COUNTER,5,lab.import_date)
    worksheet.write(ROW_COUNTER,6,str(lab.deduction_timeliness),formats["integer"])
    worksheet.data_validation("G2", {'validate': 'integer','criteria': 'between','minimum': 1,'maximum': 200})
    worksheet.write(ROW_COUNTER,7,str(lab.deduction_revision),formats["integer"])
    worksheet.data_validation("H2", {'validate': 'integer','criteria': 'between','minimum': 1,'maximum': 200})
    ROW_COUNTER += 2
    
def create_test(test,sample_id,worksheet,formats):
    global ROW_COUNTER
    worksheet.write(ROW_COUNTER,0,"Test Name")
    worksheet.write(ROW_COUNTER,1,"Test Method")
    worksheet.write(ROW_COUNTER,2,"Sample")
    worksheet.write(ROW_COUNTER,3,"Type of test")
    worksheet.write(ROW_COUNTER + 1,0,test.name)
    worksheet.write(ROW_COUNTER + 1,1,test.method)
    worksheet.write(ROW_COUNTER + 1,2,str(sample_id))
    worksheet.write(ROW_COUNTER + 1,3,test.type)
    
    column_counter = 4
    if len(test.Requirements) > 1:
        worksheet.merge_range(ROW_COUNTER,column_counter,ROW_COUNTER,column_counter + len(test.Requirements) - 1,"Requirements",formats["center"])
    elif len(test.Requirements) == 1:
        worksheet.write(ROW_COUNTER,column_counter,"Requirement")
        
    for count,requirement in enumerate(test.Requirements):
        worksheet.write(ROW_COUNTER + 1,column_counter+count,requirement)
        try:
            worksheet.write(ROW_COUNTER + 2,column_counter+count,test.Requirements_data[count],formats["non_numeric_data"])
        except:
            tem = 1
    
    column_counter += len(test.Requirements) + 1
    
    if len(test.Results) > 1:
        worksheet.merge_range(ROW_COUNTER,column_counter,ROW_COUNTER,column_counter + len(test.Results) - 1, "Results",formats["center"])
    elif len(test.Results) == 1:
        worksheet.write(ROW_COUNTER,column_counter,"Result")
    for count,result in enumerate(test.Results):
        worksheet.write(ROW_COUNTER + 1,column_counter + count,result.name)
        if result.type == "Observation":
            worksheet.write(ROW_COUNTER + 2,column_counter + count,str(result.value))
        elif result.type == "Quantitative":
            worksheet.write(ROW_COUNTER + 2,column_counter + count,str(result.value),formats["quan_numeric_data"])
        else:
            worksheet.write(ROW_COUNTER + 2,column_counter + count,str(result.value),formats["semi_numeric_data"])
        if result.type == "Quantitative":
            if result.Z_score >= 2:
                worksheet.write(ROW_COUNTER + 3,column_counter + count,str(result.Z_score),formats['red'])
            else:
                worksheet.write(ROW_COUNTER + 3,column_counter + count,str(result.Z_score))
        elif result.type == "Semi-Quantitative":
            if result.away_from_mode >= 0.5:
                worksheet.write(ROW_COUNTER + 3,column_counter + count,str(result.away_from_mode) + " Unit",formats['red'])
            else:
                worksheet.write(ROW_COUNTER + 3,column_counter + count,str(result.away_from_mode) + " Unit")
        else:
            worksheet.write(ROW_COUNTER + 3,column_counter + count,"Observation")
    column_counter += len(test.Results)
    worksheet.write(ROW_COUNTER + 1,column_counter,"Result Rating")
    worksheet.write(ROW_COUNTER + 2,column_counter,test.result_rating)
    column_counter += 2
    worksheet.merge_range(ROW_COUNTER,column_counter,ROW_COUNTER,column_counter + len(config.ERRORS),"Errors",formats["center"])
    for count,error in enumerate(config.ERRORS):
        worksheet.write(ROW_COUNTER + 1,column_counter + count,error)
    if len(test.Errors) > 0:
        for count,error in enumerate(test.Errors.values()):
            worksheet.write(ROW_COUNTER + 2,column_counter + count,str(error.number),formats["integer"])
            worksheet.data_validation(ROW_COUNTER + 2,column_counter + count,ROW_COUNTER + 2,column_counter + count, {'validate': 'integer','criteria': 'between','minimum': 1,'maximum': 200})
            worksheet.write(ROW_COUNTER + 3,column_counter + count ,error.description,formats["description"])
    else:
        for count in range(6):
            worksheet.write(ROW_COUNTER + 2,column_counter + count,"0",formats["integer"])
            worksheet.data_validation(ROW_COUNTER + 2,column_counter + count,ROW_COUNTER + 2,column_counter + count, {'validate': 'integer','criteria': 'between','minimum': 1,'maximum': 200})
            worksheet.write(ROW_COUNTER + 3,column_counter + count , "",formats["description"])
    
def create_lab_error_input_form(lab):
    workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Error/" + lab.ID + ".xlsx")
    worksheet = workbook.add_worksheet()
    global ROW_COUNTER
    ROW_COUNTER = 0
    formats = dict()
    formats["non_numeric_data"] = workbook.add_format({'bg_color':'gray','num_format': '@'})
    formats["semi_numeric_data"] = workbook.add_format({'num_format': '@','bg_color':'silver'})
    formats["quan_numeric_data"] = workbook.add_format({'num_format': '@','bg_color':'blue'})
    formats["red"] = workbook.add_format({'num_format': '@','bg_color':'red'})
    formats["outlier"] = workbook.add_format({"bg_color":'red','num_format':'@'})
    formats["integer"] = workbook.add_format({"bg_color":'yellow','num_format':'###','locked':False})
    formats["description"] = workbook.add_format({'bg_color':'green','num_format':"@",'locked':False})
    formats["center"] = workbook.add_format({"align": "center"})
    worksheet.protect()
    create_info(lab,worksheet,formats)
    for sample in lab.Samples.values():
        for tests in sample.Tests.values():
            for test in tests:
                create_test(test,sample.sample_id,worksheet,formats)
                ROW_COUNTER += 4
    workbook.close()
    
def write_to_lab(lab,path):
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active
    timeliness = int(sheet.cell(2,7).value)
    revision = int(sheet.cell(2,8).value)
    lab.deduction_timeliness = int(timeliness)
    lab.deduction_revision = int(revision)
    current_row = 4
    max_row = sheet.max_row
    while (current_row < max_row):
        sample = sheet.cell(current_row + 1,3).value
        type_ = sheet.cell(current_row + 1,4).value
        test_name = sheet.cell(current_row + 1,1).value
        test_method = sheet.cell(current_row + 1,2).value
        #print(sample,type_,test_name,test_method)
        tests = lab.Samples[sample].Tests[type_]
        for test in tests:
            if (test.name == test_name and test.method == test_method):
                column_error = len(test.Requirements) + len(test.Results) + 8
                for count in range(6):
                    error_name = sheet.cell(current_row + 1,column_error).value
                    error_type_number = count+1
                    error_number = int(sheet.cell(current_row + 2,column_error).value)
                    error_description = sheet.cell(current_row + 3,column_error).value
                    error = DataClasses.Error(name = error_name,type_number = error_type_number,number = error_number,description = error_description)
                    test.Errors[str(count+1)] = error
                    column_error += 1
                continue
        current_row += 4
    
def write_data_to_labs(Labs):
    for labs in Labs.Groups.values():
        for lab in labs.values():
            path = config.BASE_DIRECTORY_PATH + "Data/Error/" + lab.ID + ".xlsx"
            if os.path.exists(path):
                print("Writing " + lab.ID + " data to Labs")
                write_to_lab(lab,path)
    
def save_lab(Labs):
    Labs_json = Labs.to_dict()
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_errors.json','w',newline='') as fp:
        json.dump(Labs_json,fp)

def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_without_deduction.json') as json_file:
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    write_data_to_labs(Labs)
    save_lab(Labs)
    for labs in Labs.Groups.values():
        for lab in labs.values():
            if lab.imported:
                create_lab_error_input_form(lab)
    
if __name__ == '__main__':
    logging.info("Entered 'generate_error_input_form.py'")
    main()  #skip read old data first