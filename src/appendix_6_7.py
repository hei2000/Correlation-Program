import xlsxwriter,os,config,helper,logging,json,DataClasses

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
    
def create_test_str(name,method,id_):
    return name + " (" + method + ") (Sample " + id_ + ")"

def create_overview_header(worksheet,formats):
    worksheet.merge_range(1,0,2,0,"Country",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,1,1,2,"Test Performance (Top to Bottom)\n(WITHOUT Score Deduction)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,1,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,2,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,3,1,4,"Test Performance (Top to Bottom)\n(WITH Score Deduction)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,3,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,4,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter_center"])
    
    worksheet.merge_range(1,6,2,6,"Country",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,7,1,8,"Test Performance (Top to Bottom)\n(WITHOUT Score Deduction)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,7,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,8,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,9,1,10,"Test Performance (Top to Bottom)\n(WITH Score Deduction)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,9,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,10,"Lab Name (Lab ID)",formats["bold_border_gray_vcenter_center"])

def create_overview(workbook,Labs,PSR,formats,third_party):
    worksheet_name = "Overview ("
    if (PSR):
        worksheet_name += "PSR)"
    else:
        worksheet_name += "SPC)"
    if (third_party):
        worksheet_name += " (Third-Party)"
    else:
        worksheet_name += " (In-House)"
    worksheet = workbook.add_worksheet(worksheet_name)
    helper.worksheet_print_setting(worksheet)
    table_title = "Tabel A6.1 - Overview of Laboratory Performance Analysis on "
    if (PSR):
        table_title += "PSR"
    else:
        table_title += "SPC"
    table_title += " Test Items (Score) - By Country"
    if (third_party):
        table_title += " (Third-Party)"
    else:
        table_title += " (In-House)"
    worksheet.write(0,0,table_title,formats['bold'])
    create_overview_header(worksheet,formats)
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
    left_row_counter = 3
    right_row_counter = 3
    for count,country in enumerate(lab_by_location_list):
        left = ((left_row_counter-3) * 2<lab_count)
        if (left):
            column = 0
        else:
            column = 6
        labs = lab_by_location[country]
        if (PSR):
            labs_without_deduction = sorted(labs,key = lambda lab:lab.PSR_score,reverse = True)
            labs_with_deduction = sorted(labs,key = lambda lab:(lab.PSR_score - lab.PSR_error * 0.5),reverse = True)
        else:
            labs_without_deduction = sorted(labs,key = lambda lab:lab.SPC_score,reverse = True)
            labs_with_deduction = sorted(labs,key = lambda lab:(lab.SPC_score - lab.SPC_error * 0.5),reverse = True)
        if (left):
            row_counter = left_row_counter
        else:
            row_counter = right_row_counter
        if (len(labs) == 1):
            worksheet.write(row_counter,column,country,formats['border'])
        else:
            worksheet.merge_range(row_counter,column,row_counter + len(labs) - 1,column,country,formats['border_vcenter'])
        for cnt in range(len(labs)):
            if (left):
                row_counter = left_row_counter
            else:
                row_counter = right_row_counter
            lab_without_deduction = labs_without_deduction[cnt]
            lab_with_deduction = labs_with_deduction[cnt]
            if (PSR):
                worksheet.write(row_counter,column+1,str(lab_without_deduction.PSR_score) + "%",formats['border'])
                worksheet.write(row_counter,column+3,str(lab_without_deduction.PSR_score - lab_without_deduction.PSR_error * 0.5) + "%",formats['border'])
            else:
                worksheet.write(row_counter,column+1,str(lab_without_deduction.SPC_score) + "%",formats['border'])
                worksheet.write(row_counter,column+3,str(lab_without_deduction.SPC_score - lab_without_deduction.SPC_error * 0.5) + "%",formats['border'])
            worksheet.write(row_counter,column+2,lab_without_deduction.group+"-"+helper.location_city_combine(lab_without_deduction.location,lab_without_deduction.city)+" ("+lab_without_deduction.ID+")",formats['border'])
            worksheet.write(row_counter,column+4,lab_with_deduction.group+"-"+helper.location_city_combine(lab_with_deduction.location,lab_with_deduction.city)+" ("+lab_with_deduction.ID+")",formats['border'])
            if (left):
                left_row_counter += 1
            else:
                right_row_counter += 1
        for i in range(5):
            if (left):
                worksheet.write(left_row_counter,column+i,"",formats['border'])
            else:
                worksheet.write(right_row_counter,column+i,"",formats['border'])
        if (left):
            left_row_counter += 1
        else:
            right_row_counter += 1
    while (right_row_counter < left_row_counter):
        for i in range(5):
            worksheet.write(right_row_counter,6+i,"",formats['border'])
        right_row_counter += 1
    
def create_analysis_header(worksheet,formats,group):
    worksheet.merge_range(1,0,2,0,"Lab ID",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,1,2,1,"Lab",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,2,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,3,"Test Performance",formats["bold_border_gray_vcenter_center"])
    if (group == "In-House"):
        worksheet.merge_range(1,2,1,3,"Test Performance",formats["bold_border_gray_vcenter_center"])
        column = 4
    else:
        worksheet.merge_range(1,2,1,4,"Test Performance",formats["bold_border_gray_vcenter_center"])
        worksheet.write(2,4,"No Capability Test Item(s)",formats["bold_border_gray_vcenter_center"])
        column = 5
    worksheet.merge_range(1,column,1,column+3,"Lab Errors / Observations",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,column,"Test",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,column+1,"Sample ID",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,column+2,"Error type",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,column+3,"Details / Description of Error",formats["bold_border_gray_vcenter_center"])
    
def create_analysis(workbook,group,labs,PSR,formats,project_tests):
    worksheet_name = "Analysis "
    if (PSR):
        worksheet_name += "PSR"
    else:
        worksheet_name += "SPC"
    worksheet_name += " " + group
    worksheet = workbook.add_worksheet(worksheet_name)
    helper.worksheet_print_setting(worksheet)
    create_analysis_header(worksheet,formats,group)
    
    row_counter = 3
    chart_A_data = [[],[],[]]
    for lab in labs.values():
        chart_A_data[2].append(lab.fullname)
        if (PSR):
            chart_A_data[0].append(lab.PSR_score)
            chart_A_data[1].append(lab.PSR_score - lab.PSR_error * 0.5)
        else:
            chart_A_data[0].append(lab.SPC_score)
            chart_A_data[1].append(lab.SPC_score - lab.SPC_error * 0.5)
        if (group != "In-House"):
            no_cap_list = project_tests.copy()
            for sample in lab.Samples.values():
                if (PSR):
                    tests = sample.Tests["PSR"]
                else:
                    tests = sample.Tests["SPC"]
                for test in tests:
                    no_cap_list.remove(create_test_str(test.name,test.method,str(sample.sample_id)))
            no_cap_text = ""
            if (len(no_cap_list) == 0):
                no_cap_text = "---"
            for no_cap in no_cap_list:
                no_cap_text += "- " + no_cap + "\n"
                
        stragglers = list()
        outliers = list()
        errors = list()
        for sample in lab.Samples.values():
            if (PSR):
                tests = sample.Tests["PSR"]
            else:
                tests = sample.Tests["SPC"]
            for test in tests:
                if (test.outlier):
                    outliers.append(create_test_str(test.name,test.method,str(sample.sample_id)))
                elif (test.straggler):
                    stragglers.append(create_test_str(test.name,test.method,str(sample.sample_id)))
                    
                for error in test.Errors.values():
                    if (error.number > 0):
                        error_data = [test.name,str(sample.sample_id),str(error.type_number),error.description]
                        errors.append(error_data)
        details_text = "Stragglers\n"
        if (len(stragglers) == 0):
            details_text += "- Nil\n"
        for text in stragglers:
            details_text += "- " + text + "\n"
        details_text += "\n"
        details_text += "Outliers\n"
        if (len(outliers) == 0):
            details_text += "- Nil\n"
        for text in outliers:
            details_text += "- " + text + "\n"
            
        score_text = "Overall\n"
        if (PSR):
            score_text += str(lab.PSR_score)
        else:
            score_text += str(lab.SPC_score)
        score_text += "%\n\nIndividual\n"
        for sample in lab.Samples.values():
            if (PSR):
                if (sample.PSR_score != -1):
                    score_text += "Sample " + str(sample.sample_id) + ": "
                    score_text += str(sample.PSR_score) + "%\n"
            else:
                if (sample.SPC_score != -1):
                    score_text += "Sample " + str(sample.sample_id) + ": "
                    score_text += str(sample.SPC_score) + "%\n"
                    
        lab_row_number = len(errors)
        if (lab_row_number <= 1):
            worksheet.write(row_counter,0,lab.ID,formats["border"])
            worksheet.write(row_counter,1,helper.location_city_combine(lab.location,lab.city),formats["border"])
            worksheet.write(row_counter,2,score_text,formats["border"])
            worksheet.write(row_counter,3,details_text,formats["border"])
            column = 4
            if (group != "In-House"):
                worksheet.write(row_counter,column,no_cap_text,formats["border"])
                column += 1
            if (lab_row_number == 0):
                errors.append(["---" for i in range(4)])
            for i in range(4):
                worksheet.write(row_counter,column + i,errors[0][i],formats["border"])
            row_counter += 1
        else:
            worksheet.merge_range(row_counter,0,row_counter + lab_row_number - 1,0,lab.ID,formats["border"])
            worksheet.merge_range(row_counter,1,row_counter + lab_row_number - 1,1,lab.fullname,formats["border"])
            worksheet.merge_range(row_counter,2,row_counter + lab_row_number - 1,2,score_text,formats["border"])
            worksheet.merge_range(row_counter,3,row_counter + lab_row_number - 1,3,details_text,formats["border"])
            column = 4
            if (group != "In-House"):
                worksheet.merge_range(row_counter,column,row_counter + lab_row_number - 1,column,no_cap_text,formats["border"])
                column += 1
            for error in errors:
                for i in range(4):
                    worksheet.write(row_counter,column + i,error[i],formats["border"])
                row_counter += 1
                
    chart_A = workbook.add_chart({'type': 'column'})
    chart_A.set_legend({'position': 'bottom'})
    worksheet.write_row(row_counter,0,chart_A_data[0])
    worksheet.write_row(row_counter + 1,0,chart_A_data[1])
    worksheet.write_row(row_counter + 2,0,chart_A_data[2])
    if (PSR):
        tem_text = "PSR"
    else:
        tem_text = "SPC"
    chart_A.add_series({'name':tem_text + " (Overall Performance WITHOUT Score Deduction)",'values': [worksheet_name,row_counter,0,row_counter,len(chart_A_data[0]) - 1],'data_labels': {'value': True, 'num_format': '#.0\%'}})
    chart_A.add_series({'name':tem_text + " (Overall Performance WITH Score Deduction)",'values': [worksheet_name,row_counter + 1,0,row_counter + 1,len(chart_A_data[0]) - 1],'data_labels': {'value': True, 'num_format': '#.0\%'},'categories': [worksheet_name,row_counter+2,0,row_counter+2,len(chart_A_data[0]) - 1]})
    chart_A.set_title({'name': "GAP Correlation Study (2021): Test Performance on " + tem_text + " Test Items"})
    worksheet.insert_chart(row_counter + 4,0,chart_A)
    
    row_counter += 20
    start_row = row_counter
    
    chart_B = workbook.add_chart({'type': 'column','subtype':'stacked'})
    chart_B.set_title({'name':'GAP Correlation Study (2021) - Laboratory Errors in Reporting (SPC Test Items)'})
    chart_B.set_table()
    chart_B.set_legend({'position': 'bottom'})
    chart_B.set_x_axis({'num_font':  {'rotation': -45,'italic': True,'bold': True}})
    chart_B.set_y_axis({'name': 'No. of Reporting Errors Observed'})
    cat = list()
    for i in range(1,7):
        cat.append("Error Type #" + str(i))
    lab_list = list()
    datum = list()
    for lab in labs.values():
        lab_list.append(lab.fullname)
        data = [0 for i in range(6)]
        for sample in lab.Samples.values():
            if (PSR):
                tests = sample.Tests["PSR"]
            else:
                tests = sample.Tests["SPC"]
            for test in tests:
                for cnt,error in enumerate(test.Errors.values()):
                    data[cnt] += error.number
        datum.append(data)
    worksheet.write_row(row_counter,0,lab_list)
    row_counter += 1
    for data in datum:
        worksheet.write_row(row_counter,0,data)
        row_counter += 1
    color_list = ['red','orange','yellow','green','blue','purple']
    for cnt,name in enumerate(cat):
        chart_B.add_series({'line':{'color': color_list[cnt]},'fill':{'color': color_list[cnt]},'name':name,'values':[worksheet_name,start_row + 1,cnt,start_row + len(datum),cnt],'categories': [worksheet_name,start_row,0,start_row,len(datum[0])]})
    worksheet.insert_chart(row_counter + 4,0,chart_B)
            
def main(PSR = False):
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:      #change to Labs_with_deduction.json
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        Project = DataClasses.Project.from_dict(json.load(json_file))
    project_tests = list()
    for sample in Project.Samples.values():
        if (PSR):
            tests = sample.Tests["PSR"]
        else:
            tests = sample.Tests["SPC"]
        for test in tests:
            test_str = create_test_str(test.name,test.method,str(sample.sample_id))
            project_tests.append(test_str)
    if (PSR):
        workbook_name = config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_6.xlsx"
    else:
        workbook_name = config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_7.xlsx"
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
    
    create_overview(workbook,Labs,PSR,formats,third_party = True)
    if (not PSR):
        create_overview(workbook,Labs,PSR,formats,third_party = False)
    for group,labs in Labs.Groups.items():
        if (PSR and group == "In-House"):
            continue
        create_analysis(workbook,group,labs,PSR,formats,project_tests)


    workbook.close()

if __name__ == '__main__':
    main()