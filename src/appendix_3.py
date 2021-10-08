import xlsxwriter,os,config,helper,logging,json,DataClasses

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

def create_test_str(name,method,id_):
    return name + " (" + method + ") (Sample " + id_ + ")"

def create_overview(workbook,Labs,formats,third_party):
    if (third_party):
        worksheet = workbook.add_worksheet("Overview_third_party")
    else:
        worksheet = workbook.add_worksheet("Overview_in_house")
    helper.worksheet_print_setting(worksheet)
    table_title = "Table A3.1 - Overview of Laboratory Performance Analysis (Score) - By Country"
    if (third_party):
        table_title += " (Third-Party)"
    else:
        table_title += " (In-House)"
    worksheet.write(0,0,table_title,formats["bold"])
    worksheet.merge_range(1,0,2,0,"Country",formats["bold_center_border_gray"])
    worksheet.merge_range(1,1,1,2,"Overall Performance (Top to Bottom) (Without Score Deduction)",formats["bold_center_border_gray"])
    worksheet.merge_range(1,3,1,4,"Overall Performance (Top to Bottom) (With Score Deduction)",formats["bold_center_border_gray"])
    worksheet.write(2,1,"Score",formats["bold_center_border_gray"])
    worksheet.write(2,2,"Lab Name (Lab ID)",formats["bold_center_border_gray"])
    worksheet.write(2,3,"Score",formats["bold_center_border_gray"])
    worksheet.write(2,4,"Lab Name (Lab ID)",formats["bold_center_border_gray"])
    
    lab_by_location = dict()
    for group,labs in Labs.Groups.items():
        if (third_party):
            if (group == "In-House"):
                continue
        else:
            if (group != "In-House"):
                continue
        for lab in labs.values():
            location = lab.location
            if location in lab_by_location:
                lab_by_location[location].append(lab)
            else:
                lab_by_location[location] = [lab]
    lab_by_location_list = sorted(list(lab_by_location))
    row_counter = 3
    for country in lab_by_location_list:
        labs = lab_by_location[country]
        labs_without_deduction = sorted(labs,key = lambda lab:lab.score_without_deduction,reverse = True)
        labs_with_deduction = sorted(labs,key = lambda lab:lab.score_with_deduction,reverse = True)
        if (len(labs) == 1):
            worksheet.write(row_counter,0,country,formats['border'])
        else:
            worksheet.merge_range(row_counter,0,row_counter + len(labs) - 1,0,country,formats["border_vcenter"])
        for cnt in range(len(labs)):
            worksheet.write(row_counter,1,str(labs_without_deduction[cnt].score_without_deduction)+"%",formats['border'])
            worksheet.write(row_counter,2,labs_without_deduction[cnt].group + "-" + helper.location_city_combine(country,labs_without_deduction[cnt].city) + " (" + labs_without_deduction[cnt].ID + ")",formats['border'])
            worksheet.write(row_counter,3,str(labs_with_deduction[cnt].score_with_deduction) + "%",formats['border'])
            worksheet.write(row_counter,4,labs_with_deduction[cnt].group + "-" + helper.location_city_combine(country,labs_with_deduction[cnt].city) + " (" + labs_with_deduction[cnt].ID + ")",formats['border'])
            row_counter += 1
        worksheet.write(row_counter,0,"",formats['border'])
        worksheet.write(row_counter,1,"",formats['border'])
        worksheet.write(row_counter,2,"",formats['border'])
        worksheet.write(row_counter,3,"",formats['border'])
        worksheet.write(row_counter,4,"",formats['border'])
        row_counter += 1
       
    worksheet.set_column(0,0,20)
    worksheet.set_column(1,1,10)
    worksheet.set_column(2,2,30)
    worksheet.set_column(3,3,10)
    worksheet.set_column(4,4,30)
    #print_setting_overview(worksheet)
    
def create_by_group(workbook,group,labs,formats,project_tests):
    worksheet = workbook.add_worksheet(group)
    helper.worksheet_print_setting(worksheet)
    worksheet.write(0,0,"Table A3.2 - Laboratory Performance Analysis (Score) - By Lab (" + group + ") (Cont'd)",formats["bold"])
    worksheet.merge_range(1,0,2,0,"Lab ID",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,1,2,1,"Lab Name",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,2,2,2,"Overall Performance",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,3,1,4,"Test Performance (A)",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,3,"Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(2,4,"Details",formats["bold_border_gray_vcenter_center"])
    worksheet.merge_range(1,5,2,5,"Score Deduction (B)",formats["bold_border_gray_vcenter_center"])
    if (group != "In-House"):
        worksheet.merge_range(1,6,2,6,"No Capability Test Item(s)",formats["bold_border_gray_vcenter_center"])
    row_counter = 3
    chart_data = [[],[],[]]
    for lab in labs.values():
        chart_data[0].append(lab.score_without_deduction)
        chart_data[1].append(lab.score_with_deduction)
        chart_data[2].append(lab.fullname)
        worksheet.write(row_counter,0,lab.ID,formats['border_vcenter'])
        worksheet.write(row_counter,1,group + "-" + helper.location_city_combine(lab.location,lab.city),formats['border_vcenter'])
        performance_text = "Lab Score\n(A)\n" + str(lab.score_without_deduction) + "%\n\nLab Score\n(A-B)\n" + str(lab.score_with_deduction) + "%"
        worksheet.write(row_counter,2,performance_text,formats['border_center_vcenter'])
        score_text = "Overall:\n" + str(lab.score_without_deduction) + "%\n\nIndividual:\n"
        for sample in lab.Samples.values():
            score_text += "Sample " + str(sample.sample_id) + ": " + str(sample.score) + "%\n"
        worksheet.write(row_counter,3,score_text,formats['border_vcenter'])
        
        stragglers = list()
        outliers = list()
        for sample in lab.Samples.values():
            for tests in sample.Tests.values():
                for test in tests:
                    if (test.outlier):
                        outliers.append(create_test_str(test.name,test.method,str(sample.sample_id)))
                    elif (test.straggler):
                        stragglers.append(create_test_str(test.name,test.method,str(sample.sample_id)))
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
        worksheet.write(row_counter,4,details_text,formats['border'])
                    
        deduction_text = "Total No. of Lab Errors:\n" + str(lab.total_test_error) + "\n\nLabte Report / Result\n"
        deduction_text += str(lab.deduction_timeliness) + "\n\nReport / Result Revision\n" + str(lab.deduction_revision)
        worksheet.write(row_counter,5,deduction_text,formats['border'])
        
        if (group != "In-House"):
            no_cap_list = project_tests.copy()
            for sample in lab.Samples.values():
                for tests in sample.Tests.values():
                    for test in tests:
                        no_cap_list.remove(create_test_str(test.name,test.method,str(sample.sample_id)))
            no_cap_text = ""
            if (len(no_cap_list) == 0):
                no_cap_text = "---"
            for no_cap in no_cap_list:
                no_cap_text += "- " + no_cap + "\n"
            worksheet.write(row_counter,6,no_cap_text,formats['border'])
        row_counter += 1
        
        
    chart = workbook.add_chart({'type': 'column'})
    chart.set_legend({'position': 'bottom'})
    chart.set_x_axis({'num_font':  {'rotation': -45,'italic': True,'bold': True}})
    worksheet.write_row(row_counter,0,chart_data[0])
    worksheet.write_row(row_counter + 1,0,chart_data[1])
    worksheet.write_row(row_counter + 2,0,chart_data[2])
    chart.add_series({'name':"Lab Performance (without Score Deduction)",'values': [group,row_counter,0,row_counter,len(chart_data[0]) - 1],'data_labels': {'value': True, 'num_format': '#.0\%'}})
    chart.add_series({'name':"Lab Performance (with Score Deduction)",'values': [group,row_counter + 1,0,row_counter + 1,len(chart_data[0]) - 1],'data_labels': {'value': True, 'num_format': '#.0\%'},'categories': [group,row_counter+2,0,row_counter+2,len(chart_data[0]) - 1]})
    chart.set_title({'name': "GAP Correlation Study (2021) - Laboratory Performance (Lab Score Without Score Deduction vs. Lab Score With Score Deduction)"})
    #worksheet.set_row(row_counter, None, None, {'hidden': True})
    #worksheet.set_row(row_counter + 1, None, None, {'hidden': True})
    worksheet.insert_chart(row_counter + 4,0,chart)
    
    worksheet.set_column(0,0,6)
    worksheet.set_column(1,1,10)
    worksheet.set_column(2,2,13)
    worksheet.set_column(3,3,20)
    worksheet.set_column(4,4,30)
    worksheet.set_column(5,5,25)
    if (group != "In-House"):
        worksheet.set_column(6,6,90)
            
def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:      #change to Labs_with_deduction.json
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        Project = DataClasses.Project.from_dict(json.load(json_file))
    project_tests = list()
    for sample in Project.Samples.values():
        for tests in sample.Tests.values():
            for test in tests:
                test_str = create_test_str(test.name,test.method,str(sample.sample_id))
                project_tests.append(test_str)
    workbook  = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_3.xlsx")
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
    create_overview(workbook,Labs,formats,third_party = True)
    create_overview(workbook,Labs,formats,third_party = False)
    for group,labs in Labs.Groups.items():
        create_by_group(workbook,group,labs,formats,project_tests)
    workbook.close()

if __name__ == '__main__':
    main()