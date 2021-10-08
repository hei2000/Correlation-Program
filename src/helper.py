import DataClasses
import config

def search_completer(Tests: DataClasses.Tests):
    uni_method = set()
    uni_name = set()
    uni_requirement = set()
    uni_result = set()
    uni_group = set()
    for test in Tests.tests:
        uni_method.add(test.method)
        uni_name.add(test.name)
        for requirement in test.Requirements:
            uni_requirement.add(requirement)
        for result in test.Results:
            uni_result.add(result.name)
        uni_group.add(test.group)

    return list(uni_method),list(uni_name),list(uni_requirement),list(uni_result),list(uni_group)

def find_test(Tests: DataClasses.Tests,test_name,test_method):
    for test in Tests.tests:
        if (test.method == test_method and test.name == test_name):
            return test
    return -1

def name_method_split(name_method):
    try:
        name,method = name_method.split("{")
        method = method[:-1]
        return name,method
    except:
        return -1,-1

def enforce_name_method(Tests:DataClasses.Tests,name_method_name):
    name,method = name_method_split(name_method_name)
    if name == -1:
        return -1,-1
    uni_method,uni_name,uni_requirement,uni_result,uni_group = search_completer(Tests)
    if not (name in uni_name and method in uni_method):
        name,method = method,name
    if (name in uni_name and method in uni_method):
        return name,method
    return -1,-1
    
def find_type_index(type_):
    return config.FABRIC_TYPE.index(type_)

def worksheet_print_setting(worksheet):
    worksheet.set_landscape()
    worksheet.set_paper(9)
    logo_path = config.BASE_DIRECTORY_PATH + "Data/Appendix/UL-Logo.jpg"
    worksheet.set_header("&L&G&RGAP - Global Correlation Program (2021)",{'image_left':logo_path})
    worksheet.fit_to_pages(1, 0)
    #workshet.repeat_rows(first_row,last_row)
    
def location_city_combine(location,city):
    if (city == ""):
        return location
    return location + " (" + city + ")"

#def location_city_combine(lab):
#    group = lab.group
#    if (group == "In-House"):
#        return lab.fullname
#    city = lab.city
#    location = lab.location
#    if (city == ""):
#        return lab.group + "-" + location
#    return lab.group + "-" + location + " (" + city + ")"