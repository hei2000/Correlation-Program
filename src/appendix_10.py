import xlsxwriter,os,config,helper,logging,json,DataClasses
import pandas as pd

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
FOLDER_PATH = config.BASE_DIRECTORY_PATH + "Data/Result/"
START_ROW = 4

def create_header(test,formats,worksheet,sample_fabric_type,sample_id):
    worksheet.write(0,0,"GAP Global Correlation Study Program (2021)",formats['bold_underline'])
    worksheet.write(1,0,"Program ID: ",formats['bold'])
    worksheet.write(2,0,"Sample ID: ",formats['bold'])
    worksheet.write(1,2,"GAP Global PT_2021",formats['underline'])
    worksheet.write(2,2,sample_fabric_type + " (Sample " + sample_id + ")" ,formats['underline'])
    
    worksheet.merge_range(START_ROW,0,START_ROW + 2,0,"Lab ID",formats["header"])
    worksheet.merge_range(START_ROW,1,START_ROW + 2,1,"Lab Group",formats["header"])
    worksheet.merge_range(START_ROW,2,START_ROW + 2,2,"Lab Location",formats["header"])
    worksheet.merge_range(START_ROW,3,START_ROW + 2,3,"Lab Name",formats["header"])
    requirement_length = len(test.Requirements)
    result_length = 0
    for result in test.Results:
        if result.type == "Quantitative":
            result_length += 2
        else:
            result_length += 1
    total_length = result_length + requirement_length + 2
    worksheet.merge_range(START_ROW,4,START_ROW,4 + total_length - 1,test.name,formats["header"])
    worksheet.merge_range(START_ROW + 1,4,START_ROW + 2,4,"Test Method",formats["header"])
    if (requirement_length > 1):
        worksheet.merge_range(START_ROW + 1,5,START_ROW + 1,5 + requirement_length - 1,"Requirement Specified by Lab",formats["header"])
    else:
        worksheet.write(START_ROW + 1,5,"Requirement Specified by Lab",formats["header"])
    column_counter = 5
    for requirement in test.Requirements:
        worksheet.write(START_ROW + 2,column_counter,requirement,formats["header"])
        column_counter += 1
    for result in test.Results:
        if result.type == "Quantitative":
            worksheet.merge_range(START_ROW + 1,column_counter,START_ROW + 1,column_counter + 1,result.name,formats["header"])
            worksheet.write(START_ROW + 2,column_counter,"Result",formats["header"])
            worksheet.write(START_ROW + 2,column_counter + 1,"Z-Score",formats["header"])
            column_counter += 2
        else:
            worksheet.write(START_ROW + 1,column_counter,result.name,formats["header"])
            worksheet.write(START_ROW + 2,column_counter,"Result",formats["header"])
            column_counter += 1
    worksheet.merge_range(START_ROW + 1,column_counter,START_ROW + 2,column_counter,"Result Rating",formats["header"])
    if test.Results[0].type == "Quantitative":
        worksheet.merge_range(2,column_counter - 3,2,column_counter - 2,"Questionable",formats['questionable'])
    else:
        worksheet.merge_range(2,column_counter - 3,2,column_counter - 2,"0.5 Grade from Mode",formats['questionable'])
    worksheet.merge_range(2,column_counter -1,2,column_counter,"Outlier",formats['outlier'])
    
def write_row(row_counter,test,lab,worksheet,formats):
    worksheet.write(row_counter,0,lab.ID,formats["normal"])
    worksheet.write(row_counter,1,lab.group,formats["normal"])
    worksheet.write(row_counter,2,helper.location_city_combine(lab.location,lab.city),formats["normal"])
    worksheet.write(row_counter,3,lab.group + '-' + helper.location_city_combine(lab.location,lab.city),formats["normal"])
    worksheet.write(row_counter,4,test.method,formats["normal"])
    column_counter = 5
    for requirement in test.Requirements_data:
        worksheet.write(row_counter,column_counter,requirement,formats["requirement"])
        column_counter += 1
    for result in test.Results:
        #print(test)
        if result.type == "Semi-Quantitative":
            if abs(result.away_from_mode) > 0.5:
                worksheet.write(row_counter,column_counter,result.value,formats["outlier"])
            elif abs(result.away_from_mode) == 0.5:
                worksheet.write(row_counter,column_counter,result.value,formats["questionable"])
            else:
                worksheet.write(row_counter,column_counter,result.value,formats["normal"])
        else:
            #print(result)
            worksheet.write(row_counter,column_counter,result.value,formats["normal"])
        column_counter += 1
        if result.type == "Quantitative":
            if abs(result.Z_score) >= 3:
                worksheet.write(row_counter,column_counter,result.Z_score,formats["outlier"])
            elif abs(result.Z_score) > 2:
                worksheet.write(row_counter,column_counter,result.Z_score,formats["questionable"])
            else:
                worksheet.write(row_counter,column_counter,result.Z_score,formats["normal"])
            column_counter += 1
    worksheet.write(row_counter,column_counter,test.result_rating,formats["normal"])

def worksheet_print_setting(worksheet):
    worksheet.set_landscape()
    worksheet.set_paper(9)
    #fits all columns in one page
    #workbook: print all worksheets
    
def write_xlsx(df,test,workbook,count,formats,sample_id,Labs,sample_fabric_type,Statistics,third_party):
    method_name = test.method + "_" + test.name
    worksheet_name = ""
    if count < 10:
        worksheet_name += "0"
        worksheet_name += str(count) + "_" + method_name[:28]
    else:
        worksheet_name += str(count) + "_" + method_name[:28]
    worksheet = workbook.add_worksheet(worksheet_name)
    create_header(test,formats,worksheet,sample_fabric_type,sample_id)
    row_counter = START_ROW + 3
    for labs in Labs.Groups.values():
        for lab in labs.values():
            if (third_party):
                if lab.group == "In-House":
                    continue
            else:
                if lab.group != "In-House":
                    continue
            if sample_id in lab.Samples:
                sample = lab.Samples[sample_id]
                if test.type in sample.Tests:
                    tests = sample.Tests[test.type]
                    for cur_test in tests:
                        if cur_test.name == test.name and cur_test.method == test.method:
                            write_row(row_counter,cur_test,lab,worksheet,formats)
                            row_counter += 1
                            continue
    row_counter += 1
    worksheet.write(row_counter,3,"Data from the Correlation Pool",formats['bold'])
    row_counter += 1
    if test.Results[0].type == "Quantitative":
        worksheet.write(row_counter,3,"Median",formats["bold_center"])
        worksheet.write(row_counter + 1,3,"Q1",formats["bold_center"])
        worksheet.write(row_counter + 2,3,"Q3",formats["bold_center"])
        worksheet.write(row_counter + 3,3,"NIQR",formats["bold_center"])
        worksheet.write(row_counter + 4,3,"Maximum",formats["bold_center"])
        worksheet.write(row_counter + 5,3,"Minimum",formats["bold_center"])
        worksheet.write(row_counter + 6,3,"Range",formats["bold_center"])
    else:
        worksheet.write(row_counter,3,"Mode",formats["bold_center"])
        worksheet.write(row_counter + 1,3,"Maximum",formats["bold_center"])
        worksheet.write(row_counter + 2,3,"Minimum",formats["bold_center"])
        worksheet.write(row_counter + 3,3,"Range",formats["bold_center"])
    worksheet.write(row_counter,3,"")
    unique_test = DataClasses.Unique_Test(name=test.name,method=test.method,sample = int(sample_id))
    Statistic = Statistics.Statistics[unique_test.__hash__()]
    column_counter = 5 + len(test.Requirements)
    for count,result in enumerate(test.Results):
        if result.type == "Quantitative":
            Median = Statistic[count].median
            Q1 = Statistic[count].Q1
            Q3 = Statistic[count].Q3
            NIQR = round(Statistic[count].NIQR,3)
            Maximum = Statistic[count].maximum
            Minimum = Statistic[count].minimum
            Range = Statistic[count].range
            worksheet.write(row_counter,column_counter,Median,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 1,column_counter,Q1,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 2,column_counter,Q3,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 3,column_counter,NIQR,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 4,column_counter,Maximum,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 5,column_counter,Minimum,formats["bold_center_stat_quan"])
            worksheet.write(row_counter + 6,column_counter,Range,formats["bold_center_stat_quan"])
            
            column_counter += 1
        elif result.type == "Semi-Quantitative":
            Mode = Statistic[count].mode
            Maximum = Statistic[count].maximum
            Minimum = Statistic[count].minimum
            Range = Statistic[count].range
            worksheet.write(row_counter,column_counter,Mode,formats["bold_center_stat_semi"])
            worksheet.write(row_counter + 1,column_counter,Maximum,formats["bold_center_stat_semi"])
            worksheet.write(row_counter + 2,column_counter,Minimum,formats["bold_center_stat_semi"])
            worksheet.write(row_counter + 3,column_counter,Range,formats["bold_center_stat_semi"])
        column_counter += 1
    worksheet_print_setting(worksheet)
    
    

def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_without_deduction.json') as json_file:  
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Tests.json') as json_file:  
        Tests = DataClasses.Tests.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:  
        Project = DataClasses.Project.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Statistics.json') as json_file:  
        Statistics = DataClasses.Statistics.from_dict(json.load(json_file))
    for folder in os.listdir(FOLDER_PATH):  #assume all folders
        sample_id = folder[7:]
        sample_fabric_type = Project.Samples[sample_id].fabric_type
        workbook  = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + 'Data/Appendix/Appendix_10(3rd)/Sample ' + sample_id + ".xlsx")
        formats = dict()
        formats["bold_underline"] = workbook.add_format({'bold': True,'underline': True})
        formats["bold"] = workbook.add_format({'bold': True})
        formats["bold_center"] = workbook.add_format({'bold': True,'align': 'center'})
        formats["bold_center_stat_quan"] = workbook.add_format({'bold': True,'align': 'center','num_format': '###0.00'})
        formats["bold_center_stat_semi"] = workbook.add_format({'bold': True,'align': 'center','num_format': '###0.0'})
        formats["underline"] = workbook.add_format({'underline': True})
        formats["header"] = workbook.add_format({'bg_color': '#808080','border': 1,'text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["normal"] = workbook.add_format({'border': 1,'text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["requirement"] = workbook.add_format({'border': 1,'bg_color': '#808080','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["outlier"] = workbook.add_format({'border': 1,'bg_color': '#595959','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["questionable"] = workbook.add_format({'border': 1,'bg_color': '#B8B8B8','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        
        for count,file in enumerate(os.listdir(FOLDER_PATH + folder)):
            file_path = FOLDER_PATH + folder + "/" + file
            df = pd.read_excel(file_path)
            test_method,test_name = file.split("_")
            test_name = test_name[:-5]
            test = helper.find_test(Tests,test_name,test_method)
            write_xlsx(df,test,workbook,count,formats,sample_id,Labs,sample_fabric_type,Statistics,third_party = True)
        workbook.close()
        
    for folder in os.listdir(FOLDER_PATH):  #assume all folders
        sample_id = folder[7:]
        sample_fabric_type = Project.Samples[sample_id].fabric_type
        workbook  = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + 'Data/Appendix/Appendix_10(in-house)/Sample ' + sample_id + ".xlsx")
        formats = dict()
        formats["bold_underline"] = workbook.add_format({'bold': True,'underline': True})
        formats["bold"] = workbook.add_format({'bold': True})
        formats["bold_center"] = workbook.add_format({'bold': True,'align': 'center'})
        formats["bold_center_stat_quan"] = workbook.add_format({'bold': True,'align': 'center','num_format': '###0.00'})
        formats["bold_center_stat_semi"] = workbook.add_format({'bold': True,'align': 'center','num_format': '###0.0'})
        formats["underline"] = workbook.add_format({'underline': True})
        formats["header"] = workbook.add_format({'bg_color': '#808080','border': 1,'text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["normal"] = workbook.add_format({'border': 1,'text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["requirement"] = workbook.add_format({'border': 1,'bg_color': '#808080','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["outlier"] = workbook.add_format({'border': 1,'bg_color': '#595959','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        formats["questionable"] = workbook.add_format({'border': 1,'bg_color': '#B8B8B8','text_wrap': True,'align':"center",'valign':"vcenter",'num_format': '@'})
        
        for count,file in enumerate(os.listdir(FOLDER_PATH + folder)):
            file_path = FOLDER_PATH + folder + "/" + file
            df = pd.read_excel(file_path)
            test_method,test_name = file.split("_")
            test_name = test_name[:-5]
            test = helper.find_test(Tests,test_name,test_method)
            write_xlsx(df,test,workbook,count,formats,sample_id,Labs,sample_fabric_type,Statistics,third_party = False)
        workbook.close()
            
        
    
if __name__ == '__main__':
    main()