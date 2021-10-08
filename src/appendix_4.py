import xlsxwriter,os,config,helper,logging,json,DataClasses

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

def create_overview_table(workbook,formats,Labs,third_party):
    if (third_party):
        worksheet = workbook.add_worksheet("Overview_table_third_party")
    else:
        worksheet = workbook.add_worksheet("Overview_table_in_house")
    helper.worksheet_print_setting(worksheet)
    table_title = "Overview of Laboratory Performance Analysis (Laboratory Errors) - By Country "
    if (third_party):
        table_title += "(Third-Party)"
    else:
        table_title += "(In-House)"
    worksheet.write(0,0,table_title,formats["bold"])
    
    worksheet.write(1,0,"Country",formats["bold_border_gray_vcenter"])
    worksheet.write(1,1,"No. of Errors\n(Least errors being the top)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,2,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter"])
    worksheet.set_column(0,2,25)
    
    worksheet.write(1,4,"Country",formats["bold_border_gray_vcenter"])
    worksheet.write(1,5,"No. of Errors\n(Least errors being the top)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,6,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter"])
    worksheet.set_column(4,6,25)
    
    lab_by_location = dict()
    lab_count = 0
    for group,labs in Labs.Groups.items():
        if (third_party):
            if (group == "In-House"):
                continue
        else:
            if (group != "In-House"):
                continue
        for lab in labs.values():
            lab_count += 1
            location = lab.location
            if location not in lab_by_location:
                lab_by_location[location] = list()
            lab_by_location[location].append(lab)
    lab_by_location_list = sorted(list(lab_by_location))
    lab_count += len(lab_by_location_list)
    left_row_counter = 2
    right_row_counter = 2
    for count,country in enumerate(lab_by_location_list):
        left = ((left_row_counter-2)*2 < lab_count)
        if (left):
            column = 0
        else:
            column = 4
        labs = lab_by_location[country]
        labs_error = sorted(labs,key = lambda lab:(lab.deduction_timeliness+lab.deduction_revision+lab.total_test_error))
        #for lab in labs_error:
            #print(lab.deduction_timeliness+lab.deduction_revision+lab.total_test_error)
        if (left):
            row_counter = left_row_counter
        else:
            row_counter = right_row_counter
        if (len(labs) == 1):
            worksheet.write(row_counter,column,country,formats['border'])
        else:
            worksheet.merge_range(row_counter,column,row_counter + len(labs) - 1,column,country,formats['border_vcenter'])
        for lab in labs_error:
            if (left):
                row_counter = left_row_counter
            else:
                row_counter = right_row_counter
            worksheet.write(row_counter,column+1,str(lab.deduction_timeliness+lab.deduction_revision+lab.total_test_error),formats['border'])
            worksheet.write(row_counter,column+2,lab.group + "-" + helper.location_city_combine(lab.location,lab.city) + " (" + lab.ID + ")",formats['border'])
            if (left):
                left_row_counter += 1
            else:
                right_row_counter += 1
        if (left):
            worksheet.write(left_row_counter,column,"",formats['border'])
            worksheet.write(left_row_counter,column+1,"",formats['border'])
            worksheet.write(left_row_counter,column+2,"",formats['border'])
            left_row_counter += 1
        else:
            worksheet.write(right_row_counter,column,"",formats['border'])
            worksheet.write(right_row_counter,column+1,"",formats['border'])
            worksheet.write(right_row_counter,column+2,"",formats['border'])
            right_row_counter += 1
    while (right_row_counter < left_row_counter):
        worksheet.write(right_row_counter,4,"",formats['border'])
        worksheet.write(right_row_counter,5,"",formats['border'])
        worksheet.write(right_row_counter,6,"",formats['border'])
        right_row_counter += 1
            
        
def create_overview_graph(group,labs,workbook,formats,workbook_name,group_count):
    worksheet_name = "Overview_graph_" + group
    worksheet = workbook.add_worksheet(worksheet_name)
    helper.worksheet_print_setting(worksheet)
    chart = workbook.add_chart({'type': 'column','subtype':'stacked'})
    chart.set_title({'name':'GAP Correlation Study (2021) - Laboratory Errors in Reporting (All Tests)'})
    chart.set_table()
    chart.set_legend({'position': 'bottom'})
    chart.set_x_axis({'num_font':  {'rotation': -45,'italic': True,'bold': True}})
    chart.set_y_axis({'name': 'No. of Reporting Errors Observed'})
    cat = list()
    for i in range(1,7):
        cat.append("Error Type #" + str(i))
    lab_list = list()
    datum = list()
    for lab in labs.values():
        lab_list.append(lab.fullname)
        data = [0 for i in range(6)]
        for sample in lab.Samples.values():
            for tests in sample.Tests.values():
                for test in tests:
                    for cnt,error in enumerate(test.Errors.values()):
                        data[cnt] += error.number
        datum.append(data)
    
    row_counter = 0
    worksheet.write_row(row_counter,0,lab_list)
    row_counter += 1
    for data in datum:
        worksheet.write_row(row_counter,0,data)
        row_counter += 1
    color_list = ['red','orange','yellow','green','blue','purple']
    for cnt,name in enumerate(cat):
        chart.add_series({'line':{'color': color_list[cnt]},'fill':{'color': color_list[cnt]},'name':name,'values':[worksheet_name,1,cnt,len(datum),cnt],'categories': [worksheet_name,0,0,0,len(datum[0])]})
    worksheet.write(row_counter+1,0,"Fig. A4.1" + chr(97 + group_count) + " - Laboratory Errors observed in Reporting (" + group + ")",formats['bold'])
    worksheet.insert_chart(row_counter+2,0,chart)
    
def create_header(worksheet,formats):
    worksheet.write(1,0,"Lab ID",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,1,"Lab Name",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,2,"Source of Error",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,3,"Test Item",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,4,"Sample ID",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,5,"Error Type",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,6,"Details/Description of Error",formats["bold_border_gray_vcenter_center"])
        
def create_analysis(group,labs,group_count,workbook,formats):
    worksheet = workbook.add_worksheet("Analysis -" + group)
    helper.worksheet_print_setting(worksheet)
    worksheet.write(0,0,"Table A4." + str(group_count) + " - Laboratory Performance Analysis (Laboratory Errors) - By Lab (" + group + ")",formats['bold'])
    create_header(worksheet,formats)
    row_counter = 2
    for lab in labs.values():
        error_count = 0
        error_group = dict()
        for sample in lab.Samples.values():
            for tests in sample.Tests.values():
                for test in tests:
                    for error in test.Errors.values():
                        if (error.number > 0):
                            error_count += 1
                            error_source = test.group + " (" + test.type + ")"
                            if error_source not in error_group:
                                error_group[error_source] = list()
                            error_data = [test.name,str(sample.sample_id),str(error.type_number),error.description]
                            error_group[error_source].append(error_data)
        if error_count == 0:
            continue
        if error_count == 1:
            worksheet.write(row_counter,0,lab.ID,formats["border_center_vcenter"])
            worksheet.write(row_counter,1,helper.location_city_combine(lab.location,lab.city),formats["border_center_vcenter"])
        else:
            worksheet.merge_range(row_counter,0,row_counter+error_count-1,0,lab.ID,formats["border_center_vcenter"])
            worksheet.merge_range(row_counter,1,row_counter+error_count-1,1,lab.fullname,formats["border_center_vcenter"])
        for error_source,error_datum in error_group.items():
            error_count_source = len(error_datum)
            if error_count_source == 1:
                worksheet.write(row_counter,2,error_source,formats["center_border"])
            else:
                worksheet.merge_range(row_counter,2,row_counter + error_count_source-1,2,error_source,formats["center_border"])
            for error_data in error_datum:
                for cnt in range(4):
                    worksheet.write(row_counter,3+cnt,error_data[cnt],formats['border'])
                row_counter += 1
    
            
def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:      #change to Labs_with_deduction.json
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    #with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        #Project = DataClasses.Project.from_dict(json.load(json_file))
    workbook_name = config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_4.xlsx"
    workbook  = xlsxwriter.Workbook(workbook_name)
    formats = dict()
    formats["bold"] = workbook.add_format({'bold': True})
    formats["bold_border_gray_vcenter"] = workbook.add_format({'valign':'vcenter','bold': True,'border': 1,'bg_color': 'gray','text_wrap': True})
    formats["bold_border_gray_vcenter_center"] = workbook.add_format({'align':'center','valign':'vcenter','bold': True,'border': 1,'bg_color': 'gray','text_wrap': True})
    formats["border_vcenter"] = workbook.add_format({'valign':'vcenter','border': 1,'text_wrap': True})
    formats["bold_underline"] = workbook.add_format({'bold': True,'underline' : True})
    formats["bold_center_border_gray"] = workbook.add_format({'bg_color': 'gray','bold': True,'align': 'center','border' : 1,'text_wrap': True})
    formats["bold_center_border"] = workbook.add_format({'bold': True,'align': 'center','border' : 1,'text_wrap': True})
    formats["border"] = workbook.add_format({'border': 1,'text_wrap': True})
    formats["border_vcenter"] = workbook.add_format({'border': 1,'text_wrap': True,'valign' : 'vcenter'})
    formats["border_center_vcenter"] = workbook.add_format({'border': 1,'text_wrap': True,'align': 'center','valign' : 'vcenter'})
    formats["center_border"] = workbook.add_format({'align': 'center','border': 1,'text_wrap': True})
    formats["bold_center_vertical_border_gray"] = workbook.add_format({'bg_color': 'gray','bold': True,'align': 'center','border': 1,'text_wrap': True})
    formats["bold_center_vertical_border_gray"].set_rotation(90)
    
    create_overview_table(workbook,formats,Labs,third_party = True)
    create_overview_table(workbook,formats,Labs,third_party = False)
    group_count = 0
    for group,labs in Labs.Groups.items():
        create_overview_graph(group,labs,workbook,formats,workbook_name,group_count)
        group_count += 1
    group_count = 2
    for group,labs in Labs.Groups.items():
        create_analysis(group,labs,group_count,workbook,formats)
        group_count += 1

    workbook.close()

if __name__ == '__main__':
    main()