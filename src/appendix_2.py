import xlsxwriter,os,config,helper,logging,json,DataClasses

def convert(name,method):
    return name + " (" + method + ")"

def create_type(workbook,Labs,Project,formats,fabric_type,samples_id,table_count):
    row_counter = 0
    worksheet = workbook.add_worksheet(fabric_type)
    helper.worksheet_print_setting(worksheet)
    worksheet.set_column(0,0,50)
    worksheet.set_row(row_counter + 2,90)
    table_title = "Table A2." + str(table_count) + " - Test Items selected by Laboratories on Samples "
    for count,sample in enumerate(samples_id):
        if (count < len(samples_id) - 1 and count != 0):
            table_title += ", "
        elif (count == len(samples_id) -1):
            table_title += " & "
        table_title += str(sample)
    worksheet.write(row_counter,0,table_title,formats['bold'])
    worksheet.merge_range(row_counter + 1,0,row_counter + 2,0,"Test Item",formats["bold_border_gray_vcenter"])
    column_counter = 1
    for sample in samples_id:
        worksheet.merge_range(row_counter + 1,column_counter,row_counter + 1,column_counter + 1,"Sample " + str(sample),formats["bold_border_gray_vcenter_center"])
        worksheet.write(row_counter + 2,column_counter,"Tested",formats["bold_center_vcenter_vertical_border_gray"])
        worksheet.write(row_counter + 2,column_counter + 1,"No Capability",formats["bold_center_vcenter_vertical_border_gray"])
        column_counter += 2
    row_counter += 3
    
    sample2test = dict()
    for sample_id in samples_id:
        test2number = dict()
        project_sample = Project.Samples[str(sample_id)]
        for tests in project_sample.Tests.values():
            for test in tests:
                test2number[convert(test.name,test.method)] = [0,0]
        part_lab_count = 0
        for labs in Labs.Groups.values():
            for lab in labs.values():
                if str(sample_id) in lab.Samples:
                    part_lab_count += 1
                    sample = lab.Samples[str(sample_id)]
                    for tests in sample.Tests.values():
                        for test in tests:
                            test2number[convert(test.name,test.method)][0] += 1
        for test in test2number:
            test2number[test][1] = part_lab_count - test2number[test][0]
        sample2test[sample_id] = test2number
    #print(sample2test)
    all_test = set()
    for sample,tests in sample2test.items():
        for test in tests:
            all_test.add(test)
    #print(all_test)
    for test in all_test:
        worksheet.write(row_counter,0,test,formats["border_center_vcenter"])
        column_counter = 1
        for sample,tests in sample2test.items():
            if test in tests:
                worksheet.write(row_counter,column_counter,tests[test][0],formats["border_center_vcenter"])
                worksheet.write(row_counter,column_counter + 1,tests[test][1],formats["border_center_vcenter"])
            else:
                worksheet.write(row_counter,column_counter,"",formats["border_center_vcenter"])
                worksheet.write(row_counter,column_counter + 1,"",formats["border_center_vcenter"])
            column_counter += 2
        row_counter += 1
    return row_counter + 3
            
        
def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        Project = DataClasses.Project.from_dict(json.load(json_file))
    workbook  = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_2.xlsx")
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
    formats["bold_center_vcenter_vertical_border_gray"] = workbook.add_format({'bg_color': 'gray','bold': True,'align': 'center','valign': 'vcenter','border': 1,'text_wrap': True})
    formats["bold_center_vcenter_vertical_border_gray"].set_rotation(90)
    
    type2sample = dict()
    for sample in Project.Samples.values():
        fabric_type = sample.fabric_type
        if fabric_type not in type2sample:
            type2sample[fabric_type] = list()
        type2sample[fabric_type].append(sample.sample_id)
    table_count = 1
    for fabric_type,samples_id in type2sample.items():
        create_type(workbook,Labs,Project,formats,fabric_type,samples_id,table_count)
        table_count += 1
    workbook.close()

if __name__ == '__main__':
    main()