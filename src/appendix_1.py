import xlsxwriter,os,config,helper,logging,json,DataClasses

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

def worksheet_print_setting(worksheet,page_break_list):
    worksheet.set_landscape()
    worksheet.set_paper(9)
    logo_path = config.BASE_DIRECTORY_PATH + "Data/Appendix/UL-Logo.jpg"
    #workshet.repeat_rows(first_row,last_row)
    worksheet.set_header("&L&G&RGAP - Global Correlation Program (2021)",{'image_left':logo_path})
    worksheet.fit_to_pages(1, 0)
    worksheet.set_h_pagebreaks(page_break_list)
    
def create_group_lab(count,worksheet,formats,row_counter,group):
    table_title = "Table A1." + str(count) + ": " + group + " - Participating Laboratories by Lab ID"
    worksheet.write(row_counter,0,table_title,formats["bold"])
    row_counter += 1
    
def create_table_header(Samples_id,worksheet,formats,row_counter):
    worksheet.merge_range(row_counter,0,row_counter + 1,0,"Lab ID",formats['bold_border_gray_vcenter'])
    worksheet.merge_range(row_counter,1,row_counter + 1,1,"Lab Location",formats['bold_border_gray_vcenter'])
    worksheet.merge_range(row_counter,2,row_counter + 1,2,"Participating Laboratory",formats['bold_border_gray_vcenter'])
    worksheet.merge_range(row_counter,3,row_counter + 1,3,"Participating Laboratory (Full Name)",formats['bold_border_gray_vcenter'])
    sample_number = len(Samples_id)
    worksheet.merge_range(row_counter,4,row_counter,4 + sample_number - 1,"Assigned Correlation Test Samples",formats['bold_center_border_gray'])
    for count,sample_id in enumerate(Samples_id):
        worksheet.write(row_counter + 1,4 + count,"Sample " + str(sample_id),formats['bold_center_vertical_border_gray'])
    
    

def create_table(group,labs,Samples_id,worksheet,formats,row_counter):
    create_table_header(Samples_id,worksheet,formats,row_counter)
    row_counter += 2
    for lab in labs.values():
        worksheet.write(row_counter,0,lab.ID,formats['border'])
        worksheet.write(row_counter,1,helper.location_city_combine(lab.location,lab.city),formats['border'])
        worksheet.write(row_counter,2,group + "-" + helper.location_city_combine(lab.location,lab.city),formats['border'])
        worksheet.write(row_counter,3,lab.fullname,formats['border'])
        for count,sample_id in enumerate(Samples_id):
            worksheet.write(row_counter,4 + count,"",formats['center_border'])
        for sample in lab.Samples.values():
            worksheet.write(row_counter,3 + sample.sample_id,"X",formats['center_border'])
        row_counter += 1
    return row_counter

def create_by_country_5(workbook,formats,Labs):
    worksheet = workbook.add_worksheet("By Country")
    worksheet.set_landscape()
    worksheet.set_paper(9)
    logo_path = config.BASE_DIRECTORY_PATH + "Data/Appendix/UL-Logo.jpg"
    worksheet.set_header("&L&G&RGAP - Global Correlation Program (2021)",{'image_left':logo_path})
    worksheet.fit_to_pages(1, 0)
    worksheet.write(0,0,"Table A1.5: Participating Laboratories by Countries",formats["bold"])
    worksheet.write(1,0,"Country",formats["bold_border_gray_vcenter"])
    worksheet.write(1,2,"Country",formats["bold_border_gray_vcenter"])
    worksheet.write(1,4,"Country",formats["bold_border_gray_vcenter"])
    worksheet.write(1,1,"No. of Labs",formats["bold_border_gray_vcenter"])
    worksheet.write(1,3,"No. of Labs",formats["bold_border_gray_vcenter"])
    worksheet.write(1,5,"No. of Labs",formats["bold_border_gray_vcenter"])
    
    location_count = dict()
    for labs in Labs.Groups.values():
        for lab in labs.values():
            location = lab.location
            if location in location_count:
                location_count[location] += 1
            else:
                location_count[location] = 1
                
    row_counter = 2
    column = 0
    location_count_list = sorted(list(location_count))
    for location in location_count_list:
        worksheet.write(row_counter,column,location,formats['border'])
        worksheet.write(row_counter,column + 1,location_count[location],formats['border'])
        row_counter += (column == 4)
        column = (column + 2) % 6
    while (column != 6):
        worksheet.write(row_counter,column,"",formats["border"])
        worksheet.write(row_counter,column + 1,"",formats["border"])
        column += 2

def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:  
        Project = DataClasses.Project.from_dict(json.load(json_file))
    Samples_id = Project.Samples.keys()
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:  
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    workbook  = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_1.xlsx")
    worksheet = workbook.add_worksheet("By Group")
    formats = dict()
    formats["bold"] = workbook.add_format({'bold': True})
    formats["bold_border_gray_vcenter"] = workbook.add_format({'valign':'vcenter','bold': True,'border': 1,'bg_color': 'gray','text_wrap': True})
    formats["bold_underline"] = workbook.add_format({'bold': True,'underline' : True})
    formats["bold_center_border_gray"] = workbook.add_format({'bg_color': 'gray','bold': True,'align': 'center','border' : 1,'text_wrap': True})
    formats["bold_center_border"] = workbook.add_format({'bold': True,'align': 'center','border' : 1,'text_wrap': True})
    formats["border"] = workbook.add_format({'border': 1,'text_wrap': True})
    formats["center_border"] = workbook.add_format({'align': 'center','border': 1,'text_wrap': True})
    formats["bold_center_vertical_border_gray"] = workbook.add_format({'bg_color': 'gray','bold': True,'align': 'center','border': 1,'text_wrap': True})
    formats["bold_center_vertical_border_gray"].set_rotation(90)
    worksheet.write(0,0,"Appendix 1 - List of Participating Laboratories",formats["bold_underline"])
    count = 1
    row_counter = 2
    page_break_list = list()
    for group,labs in Labs.Groups.items():
        create_group_lab(count,worksheet,formats,row_counter,group)
        row_counter += 1
        row_counter = create_table(group,labs,Samples_id,worksheet,formats,row_counter)
        page_break_list.append(row_counter)
        row_counter += 3
        count += 1
    create_by_country_5(workbook,formats,Labs)
        
    worksheet.set_row(1, 25)
    worksheet.set_column(0,0,7)
    worksheet.set_column(1,1,20)
    worksheet.set_column(2,2,25)
    worksheet.set_column(3,3,40)
    worksheet_print_setting(worksheet,page_break_list)
    workbook.close()

if __name__ == '__main__':
    main()