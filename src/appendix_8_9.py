import matplotlib.pyplot as plt
import os,config,json,DataClasses,xlsxwriter,math,scipy,helper
import numpy as np
import scipy.stats as si

def find_outlier(arr):
    q1 = np.quantile(arr,0.25)
    q3 = np.quantile(arr,0.75)
    iqr = q3-q1
    median = round(np.median(arr),1)
    #upper = q3 + (1.5 * iqr)
    lower = q1 - (1.5 * iqr)
    answer = list()
    for num in arr:
        if (num < lower):
            answer.append(num)
    return answer,median

def create_sample(Labs,sample_id,save_directory):
    datum = dict()
    for group,labs in Labs.Groups.items():
        datum[group] = list()
        for lab in labs.values():
            if str(sample_id) in lab.Samples:
                sample = lab.Samples[str(sample_id)]
                datum[group].append(sample.score)
            
    all_lab = list()
    for data in datum.values():
        all_lab += data
    bp_data = list()
    bp_data.append(np.array(all_lab))
    for data in datum.values():
        bp_data.append(np.array(data))
    xlabels = ['All Labs'] + list(datum)
    
    outliers = list()
    median_list = list()
    out,median = find_outlier(all_lab)
    median_list.append(median)
    outliers.append(out)
    for data in datum.values():
        if (len(data) == 0):
            continue
        (out,median) = find_outlier(data)
        outliers.append(sorted(out))
        median_list.append(median)
        
    outlier_labs = list()
    outlier_all_lab = list()
    for labs in Labs.Groups.values():
        tem = list()
        for lab in labs.values():
            if str(sample_id) in lab.Samples:
                tem.append(lab)
        outlier_labs.append(tem)
        outlier_all_lab += tem
    outlier_labs = [outlier_all_lab] + outlier_labs
    #print(outlier_labs)
    fig = plt.figure()
    ax = fig.add_subplot(111)
    bp = ax.boxplot(bp_data)
    ax.set_xticklabels(xlabels)
    ax.set_title("GAP Global Correlation (2021) - Lab Performance on Testing Performance (Sample " + str(sample_id) + ")")
    ax.title.set_size(9)
    ax.set_xlabel('Lab Group')
    ax.set_ylabel("Sample " + str(sample_id) + " (Lab Score)")
    
    ret_outlier_labs = list()
    for count in range(len(outliers)):
        ret_outlier_labs.append([])
        outlier_lab = outlier_labs[count]
        outliers_list = outliers[count]
        outlier_lab.sort(key = lambda lab:lab.Samples[str(sample_id)].score)
        if (count == 0):
            min_score = outlier_lab[0].Samples[str(sample_id)].score
        for i in range(len(outliers_list)):
            lab = outlier_lab[i]
            plt.text(count + 1,outliers_list[i],"   " + helper.location_city_combine(lab.location,lab.city),fontsize = 7)
            ret_outlier_labs[count].append([helper.location_city_combine(lab.location,lab.city),outlier_lab[i].Samples[str(sample_id)].score])
    for cnt,median in enumerate(median_list):
        plt.text(cnt + 1,median + 0.3,str(median),fontsize = 6)
        
    plt.plot([0.5,len(datum) + 1.5],[median_list[0],median_list[0]],linestyle = '--',color = "0.4")
    plt.text(len(datum) + 1.6,median_list[0],"Median: " + str(median_list[0]),fontsize = 6)
    
    ax.set_ylim([min_score-10,100])
    plt.savefig(save_directory + "appendix_8_sample" + str(sample_id) + ".png")
    return ret_outlier_labs
    #plt.show()
    
def create_overall(Labs,save_directory,app_9 = False):
    datum_without = dict()
    datum_with = dict()
    for group,labs in Labs.Groups.items():
        if (app_9):
            if (group == "In-House"):
                continue
        datum_without[group] = list()
        datum_with[group] = list()
        for lab in labs.values():
            datum_without[group].append(lab.score_without_deduction)
            datum_with[group].append(lab.score_with_deduction)
    bp_datum_without = []
    bp_datum_with = []
    all_without = list()
    all_with = list()
    for data in datum_without.values():
        bp_datum_without.append(np.array(data))
        all_without += data
    for data in datum_with.values():
        bp_datum_with.append(np.array(data))
        all_with += data
    bp_datum_without = [np.array(all_without)] + bp_datum_without
    bp_datum_with = [np.array(all_with)] + bp_datum_with
    xlabels = ["All Labs"] + list(Labs.Groups)
    if (app_9):
        if "In-House" in xlabels:
            xlabels.remove("In-House")
    outliers_without = list()
    median_list_without = list()
    out,median = find_outlier(all_without)
    outliers_without.append(out)
    median_list_without.append(median)
    
    outliers_with = list()
    median_list_with = list()
    out,median = find_outlier(all_with)
    outliers_with.append(out)
    median_list_with.append(median)
    for data in datum_without.values():
        if (len(data) == 0):
            continue
        out,median = find_outlier(data)
        outliers_without.append(out)
        median_list_without.append(median)
    for data in datum_with.values():
        if (len(data) == 0):
            continue
        out,median = find_outlier(data)
        outliers_with.append(out)
        median_list_with.append(median)
        
    outlier_labs = [[]]
    for group,labs in Labs.Groups.items():
        if (app_9):
            if (group == "In-House"):
                continue
        outlier_labs.append(labs.values())
        outlier_labs[0] += labs.values()
    
    fig = plt.figure()
    ax = fig.add_subplot(111)
    bp_without = ax.boxplot(bp_datum_without,positions = [1.1 + 2 * i for i in range(len(bp_datum_without))])
    bp_with = ax.boxplot(bp_datum_with,positions = [1.9 + 2 * i for i in range(len(bp_datum_with))])
    ax.set_xticklabels(["Score 1","Score 2"] * len(bp_datum_without))
    plt.tick_params(axis='x', which='major', labelsize=7)
    ax.set_title("GAP Global Correlation (2021) - Overall Lab Performance (WITH and WITHOUT Score Deduction)")
    ax.title.set_size(9)
    #ax.set_xlabel('Lab Group')
    ax.set_ylabel("Lab Score")
    outlier_labs_without = list()
    outlier_labs_with = list()
    for count in range(len(outliers_without)):
        outlier_labs_without.append([])
        outlier_lab = outlier_labs[count]
        outliers_list = outliers_without[count]
        tem_outlier_lab = sorted(outlier_lab,key = lambda lab:lab.score_without_deduction)
        if (count == 0):
            min_score = tem_outlier_lab[0].score_without_deduction
        for i in range(len(outliers_list)):
            lab = tem_outlier_lab[i]
            plt.text(1.1 + 2 * count,outliers_list[i],"   " + helper.location_city_combine(lab.location,lab.city),fontsize = 7)
            outlier_labs_without[count].append([helper.location_city_combine(lab.location,lab.city),tem_outlier_lab[i].score_without_deduction])
            
    for count in range(len(outliers_with)):
        outlier_labs_with.append([])
        outlier_lab = outlier_labs[count]
        outliers_list = outliers_with[count]
        tem_outlier_lab = sorted(outlier_lab,key = lambda lab:lab.score_with_deduction)
        if (count == 0):
            min_score = min(min_score,tem_outlier_lab[0].score_with_deduction)
        for i in range(len(outliers_list)):
            lab = tem_outlier_lab[i]
            plt.text(1.9 + 2 * count,outliers_list[i],"   " + helper.location_city_combine(lab.location,lab.city),fontsize = 7)
            outlier_labs_with[count].append([helper.location_city_combine(lab.location,lab.city),tem_outlier_lab[i].score_with_deduction])
            
    for cnt,median in enumerate(median_list_without):
        plt.text(1.1 + 2 * cnt,median + 0.3,str(median),fontsize = 6)
    for cnt,median in enumerate(median_list_with):
        plt.text(1.9 + 2 * cnt,median + 0.3,str(median),fontsize = 6)
        
    plt.plot([0,len(datum_without) * 2 + 2.7],[median_list_without[0],median_list_without[0]],linestyle = '--',color = "0.6")
    plt.text(len(datum_without) * 2 + 2.5,median_list_without[0] + 0.5,"Median(S1):" + str(median_list_without[0]),fontsize = 5)
    
    plt.plot([0,len(datum_with) * 2 + 2.7],[median_list_with[0],median_list_with[0]],linestyle = '--',color = "0.6")
    plt.text(len(datum_with) * 2 + 2.5,median_list_with[0] - 0.5,"Median(S2):" + str(median_list_with[0]),fontsize = 5)
    
    for cnt,xlabel in enumerate(xlabels):
        plt.text(1.5 + 2 * cnt,min_score-17,xlabel,fontsize = 9,ha='center')
    plt.text(-0.3,min_score-17,"Lab Group",fontsize=8,ha='center')
    plt.text(len(xlabels) * 2 + 0.6,min_score + 2,"Score 1 (S1): \nLab Score WITHOUT\nScore Deduction",fontsize = 4)
    plt.text(len(xlabels) * 2 + 0.6,min_score - 5,"Score 2 (S2): \nLab Score WITH\nScore Deduction",fontsize = 4)
    ax.set_xlim([0,len(xlabels) * 2 + 0.5])
    
    ax.set_ylim([min_score-10,100])
    if (app_9):
        plt.savefig(save_directory + "appendix_9_overall.png")
    else:
        plt.savefig(save_directory + "appendix_8_overall.png")
        
    return outlier_labs_without,outlier_labs_with

def create_header_8(worksheet,formats,text,groups,START_ROW):
    worksheet.merge_range(START_ROW,0,START_ROW + 1,0,"Lab Performance Score",formats["bold_border_gray_vcenter_center"])
    worksheet.write(START_ROW,1,"Outlier(s) within the Pool",formats["bold_border_gray_vcenter_center"])
    combined_text = "Outlier(s) within Lab Group for " + text
    if (len(groups) == 1):
        worksheet.write(START_ROW,2,combined_text,formats["bold_border_gray_vcenter_center"])
    else:
        worksheet.merge_range(START_ROW,2,START_ROW,2 + len(groups) - 1,combined_text,formats["bold_border_gray_vcenter_center"])
    worksheet.write(START_ROW + 1,1,"All Labs",formats["bold_border_gray_vcenter_center"])
    for cnt,group in enumerate(groups):
        worksheet.write(START_ROW + 1,2 + cnt,group,formats["bold_border_gray_vcenter_center"])

def create_appendix_8_overview(workbook,formats,outlier_overall_without,outlier_overall_with,groups):
    worksheet = workbook.add_worksheet("Overview")
    helper.worksheet_print_setting(worksheet)
    worksheet.write(0,0,"Fig. A8.1A - Overall Lab Performance (Lab Score WITH and WITHOUT Score Deduction) (by Lab Group)",formats['bold'])
    worksheet.insert_image(1,0,config.BASE_DIRECTORY_PATH + "Data/Appendix/boxplot/appendix_8_overall.png")
    START_ROW = 30
    create_header_8(worksheet,formats,"Overall Lab Performance",groups,START_ROW)
    worksheet.write(START_ROW + 2,0,"Score 1 (WITHOUT Score Deduction)",formats["bold_center_border"])
    for cnt,outlier in enumerate(outlier_overall_without):
        if (len(outlier) == 0):
            worksheet.write(START_ROW + 2,cnt + 1,"---",formats["center_border"])
            continue
        text = ""
        for out in outlier:
            text += out[0] + "(" + str(out[1]) + "%)\n"
        worksheet.write(START_ROW + 2,cnt + 1,text,formats["border"])
    worksheet.write(START_ROW + 3,0,"Score 2 (WITH Score Deduction)",formats["bold_center_border"])
    for cnt,outlier in enumerate(outlier_overall_with):
        if (len(outlier) == 0):
            worksheet.write(START_ROW + 3,cnt + 1,"---",formats["center_border"])
            continue
        text = ""
        for out in outlier:
            text += out[0] + "(" + str(out[1]) + "%)\n"
        worksheet.write(START_ROW + 3,cnt + 1,text,formats["border"])
    
def create_appendix_8_sample(workbook,formats,outlier_sample,sample_id,groups):
    worksheet = workbook.add_worksheet("Sample " + str(sample_id))
    helper.worksheet_print_setting(worksheet)
    worksheet.write(0,0,"Fig. A8.2." + str(sample_id) + " - Lab Performance (by Lab Group) - Sample " + str(sample_id),formats['bold'])
    worksheet.insert_image(1,0,config.BASE_DIRECTORY_PATH + "Data/Appendix/boxplot/appendix_8_sample" + str(sample_id) + ".png")
    START_ROW = 30
    create_header_8(worksheet,formats,"Sample " + str(sample_id),groups,START_ROW)
    worksheet.write(START_ROW + 2,0,"Test Performance",formats["bold_center_border"])
    for cnt,outlier in enumerate(outlier_sample):
        if (len(outlier) == 0):
            worksheet.write(START_ROW + 2,cnt + 1,'---',formats["center_border"])
            continue
        text = ""
        for out in outlier:
            text += out[0] + "(" + str(out[1]) + "%)\n"
        worksheet.write(START_ROW + 2,cnt + 1,text,formats['border'])
    if (len(outlier_sample) < len(groups) + 1):
        worksheet.write(START_ROW + 2,len(groups) + 1,"",formats['border'])

def create_appendix_8(outlier_overall_without,outlier_overall_with,outlier_samples,sample_number,groups):
    workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_8.xlsx")
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
    
    create_appendix_8_overview(workbook,formats,outlier_overall_without,outlier_overall_with,groups)
    for sample_id in range(1,sample_number + 1):
        create_appendix_8_sample(workbook,formats,outlier_samples[sample_id-1],sample_id,groups)
    
    workbook.close()
    
def grade_str(lab,score):
    return helper.location_city_combine(lab.location,lab.city) + "(" + lab.ID + ") (" + str(score) +"%)\n"
    

def get_graph_stat(data):
    answer = dict()
    length = len(data)
    AD, crit, sig = si.anderson(data, dist='norm')
    print("Values = ", data)
    print("Significance Levels:", sig)
    print("Critical Values:", crit)
    print("\nA^2 = ", AD)
    AD_adjusted = AD*(1 + (.75/length) + 2.25/(length**2))
    print("Adjusted A^2 = ", AD_adjusted)
    if AD_adjusted >= .6:
        p = math.exp(1.2937 - 5.709*AD_adjusted - .0186*(AD_adjusted**2))
    elif AD_adjusted >=.34:
        p = math.exp(.9177 - 4.279*AD_adjusted - 1.38*(AD_adjusted**2))
    elif AD_adjusted >.2:
        p = 1 - math.exp(-8.318 + 42.796*AD_adjusted - 59.938*(AD_adjusted**2))
    else:
        p = 1 - math.exp(-13.436 + 101.14*AD_adjusted - 223.73*(AD_adjusted**2))
    print("p = ", p)
    mean = np.mean(data)
    print("mean = ",mean)
    stdv = np.std(data)
    print("Stdv = ",stdv)
    var = np.var(data)
    print("Var = ",var)
    skew = scipy.stats.skew(data)
    print("Skewness = ",skew)
    kurt = scipy.stats.kurtosis(data)
    print("Kurtosis = ",kurt)
    print("N = ",length)
    min_ = min(data)
    print("Minimum = ",min_)
    q1 = np.quantile(data,0.25)
    print("1st Quartile = ",q1)
    median = np.median(data)
    print("Median = ",median)
    q3 = np.quantile(data,0.75)
    print('3rd Quartile = ',q3)
    max_ = max(data)
    print('Maximum = ',max_)
    
    CI_mean = list(si.t.interval(0.95,len(data)-1,loc=mean,scale = si.sem(data)))
    print('95% CI of mean: ',CI_mean)
    CI_median = list(si.t.interval(0.95,len(data)-1,loc=median,scale = si.sem(data)))
    print('95% CI of median: ',CI_median)
    CI_stdv = list(si.t.interval(0.95,len(data)-1,loc=stdv,scale = si.sem(data)))
    print('95% CI of StDev: ',CI_stdv)
    
    
    answer['A-Squared'] = AD
    answer['P-Value'] = p
    answer['Mean'] = mean
    answer['StDev'] = stdv
    answer['Variance'] = var
    answer['Skewness'] = skew
    answer['Kurtosis'] = kurt
    answer['N'] = length
    answer['Minimum'] = min_
    answer['Q1'] = q1
    answer['median'] = median
    answer['Q3'] = q3
    answer['Maximum'] = max_
    answer['CI_mean'] = CI_mean
    answer['CI_median'] = CI_median
    answer['CI_stdv'] = CI_stdv
    
    return answer

def app_9_write_stat(worksheet,row_counter,stat_col,graph_stat_without):
    worksheet.write(row_counter+3,stat_col,"Anderson-Darling Normalitry Test")
    worksheet.write(row_counter+4,stat_col+1,"A-Squared")
    try:
        worksheet.write(row_counter+4,stat_col+2,graph_stat_without["A-Squared"])
    except:
        worksheet.write(row_counter+4,stat_col+2,-1)
    worksheet.write(row_counter+5,stat_col+1,"P-Value")
    try:
        worksheet.write(row_counter+5,stat_col+2,graph_stat_without["P-Value"])
    except:
        worksheet.write(row_counter+5,stat_col+2,-1)
    worksheet.write(row_counter+6,stat_col+1,"Mean")
    worksheet.write(row_counter+6,stat_col+2,graph_stat_without["Mean"])
    worksheet.write(row_counter+7,stat_col+1,"StDev")
    worksheet.write(row_counter+7,stat_col+2,graph_stat_without["StDev"])
    worksheet.write(row_counter+8,stat_col+1,"Variance")
    worksheet.write(row_counter+8,stat_col+2,graph_stat_without["Variance"])
    worksheet.write(row_counter+9,stat_col+1,"Skewness")
    worksheet.write(row_counter+9,stat_col+2,graph_stat_without["Skewness"])
    worksheet.write(row_counter+10,stat_col+1,"Kurtosis")
    worksheet.write(row_counter+10,stat_col+2,graph_stat_without["Kurtosis"])
    worksheet.write(row_counter+11,stat_col+1,"N")
    worksheet.write(row_counter+11,stat_col+2,graph_stat_without["N"])
    worksheet.write(row_counter+12,stat_col+1,"Minimum")
    worksheet.write(row_counter+12,stat_col+2,graph_stat_without["Minimum"])
    worksheet.write(row_counter+13,stat_col+1,"1st Quartile")
    worksheet.write(row_counter+13,stat_col+2,graph_stat_without["Q1"])
    worksheet.write(row_counter+14,stat_col+1,"Median")
    worksheet.write(row_counter+14,stat_col+2,graph_stat_without["median"])
    worksheet.write(row_counter+15,stat_col+1,"3rd Quartile")
    worksheet.write(row_counter+15,stat_col+2,graph_stat_without["Q3"])
    worksheet.write(row_counter+16,stat_col+1,"Maximum")
    worksheet.write(row_counter+17,stat_col+2,graph_stat_without["Maximum"])
    worksheet.write(row_counter + 18,stat_col,"95% Confidence Interval for Mean")
    try:
        worksheet.write(row_counter + 19,stat_col+1,graph_stat_without["CI_mean"][0])
        worksheet.write(row_counter + 19,stat_col+2,graph_stat_without["CI_mean"][1])
    except:
        worksheet.write(row_counter + 19,stat_col+1,-1)
        worksheet.write(row_counter + 19,stat_col+2,-1)
    worksheet.write(row_counter + 20,stat_col,"95% Confidence Interval for Median")
    try:
        worksheet.write(row_counter + 21,stat_col+1,graph_stat_without["CI_median"][0])
        worksheet.write(row_counter + 21,stat_col+2,graph_stat_without["CI_median"][1])
    except:
        worksheet.write(row_counter + 21,stat_col+1,-1)
        worksheet.write(row_counter + 21,stat_col+2,-1)
    worksheet.write(row_counter + 22,stat_col,"95% Confidence Interval for StDev")
    try:
        worksheet.write(row_counter + 23,stat_col+1,graph_stat_without["CI_stdv"][0])
        worksheet.write(row_counter + 23,stat_col+2,graph_stat_without["CI_stdv"][1])
    except:
        worksheet.write(row_counter + 23,stat_col+1,-1)
        worksheet.write(row_counter + 23,stat_col+2,-1)
        
    
def create_appendix_9(Labs,save_directory,third_party):
    score_without = list()
    score_with = list()
    for group,labs in Labs.Groups.items():
        if (third_party):
            if (group == "In-House"):
                continue
        else:
            if (group != "In-House"):
                continue
        for lab in labs.values():
            score_without.append(lab.score_without_deduction)
            score_with.append(lab.score_with_deduction)
    q1_without = np.quantile(score_without,0.25)
    q1_with = np.quantile(score_with,0.25)
    q3_without = np.quantile(score_without,0.75)
    q3_with = np.quantile(score_with,0.75)
    iqr_without = q3_without - q1_with
    iqr_with = q3_with - q1_with
    median_without = round(np.median(score_without),1)
    median_with = round(np.median(score_with),1)
    upper_whisker_without = q3_without + (1.5 * iqr_without)
    upper_whisker_with = q3_with + (1.5 * iqr_with)
    lower_whisker_without = q1_without - (1.5 * iqr_without)
    lower_whisker_with = q1_with - (1.5 * iqr_with)
    outliers_without = list()
    outliers_with = list()
    
    min_score_without = 100
    min_score_with = 100
    for num in score_without:
        min_score_without = min(num,min_score_without)
        if (num < lower_whisker_without):
            outliers_without.append(num)
    for num in score_with:
        min_score_with = min(num,min_score_with)
        if (num < lower_whisker_with):
            outliers_with.append(num)
            
            
    
    outliers_label_without = list()
    outliers_label_with = list()
    for outlier in outliers_without:
        for group, labs in Labs.Groups.items():
            if (third_party):
                if (group == "In-House"):
                    continue
            else:
                if (group != "In-House"):
                    continue
            for lab in labs.values():
                if (lab.score_without_deduction == outlier):
                    outliers_label_without.append(str(lab.score_without_deduction) + "% (" + helper.location_city_combine(lab.location,lab.city) + ")")
                    continue
    for outlier in outliers_with:
        for group, labs in Labs.Groups.items():
            if (third_party):
                if (group == "In-House"):
                    continue
            else:
                if (group != "In-House"):
                    continue
            for lab in labs.values():
                if (lab.score_with_deduction == outlier):
                    outliers_label_with.append(str(lab.score_with_deduction) + "% (" + helper.location_city_combine(lab.location,lab.city) + ")")
                    continue
                
                
    fig2 = plt.figure()
    ax2 = fig2.add_subplot(111)
    bp2 = ax2.boxplot([score_without,score_with])
    ax2.set_xticklabels(["Lab Score WITHOUT Score Deduction","Lab Score WITH Score Deduction"])
    plt.tick_params(axis='x', which='major', labelsize=7)
    ax2.set_title("GAP Global Correlation (2021) - Overall Lab Performance (WITH and WITHOUT Score Deduction)")
    ax2.title.set_size(9)
    ax2.set_ylabel("Lab Score")                

    for count in range(len(outliers_without)):
        plt.text(1,outliers_without[count],"   " + outliers_label_without[count])
    for count in range(len(outliers_with)):
        plt.text(2,outliers_with[count],"   " + outliers_label_with[count])
        
    plt.text(1.1,median_without,str(median_without) + "%")
    plt.text(2.1,median_with,str(median_with) + "%")
    
    ax2.set_ylim([min(min_score_without,min_score_with)-10,100])
                    
    
    fig2_filename = save_directory + "appendix_9_fig2"
    if (third_party):
        fig2_filename += "_Third-party" + ".png"
    else:
        fig2_filename += "_In-house" + ".png"
    plt.savefig(fig2_filename)
    
    
    #####fig3
    fig3 = plt.figure()
    ax3 = fig3.add_subplot(111)
    bp3 = ax3.boxplot([score_without])
    ax3.set_xticklabels(["All Labs"])
    ax3.set_xlabel('Lab Group')
    plt.tick_params(axis='x', which='major', labelsize=7)
    ax3.set_title("GAP Global Correlation (2021) - Lab Performance (WITHOUT Score Deduction)")
    ax3.title.set_size(9)
    ax3.set_ylabel("Lab Score")
    
    plt.text(1.1,median_without,"Median: " + str(median_without) + "%")
    Q3_text = "Q3 - 75% of the\ndata values are less\nthan or equal to this\n value (Q3 = " + str(q3_without) + "%)"
    Q1_text = "Q1 - 25% of the\ndata values are less\nthan or equal to this\n value (Q1 = " + str(q1_without) + "%)"
    plt.text(0.7,q3_without + 3,Q3_text,fontsize=7)
    plt.text(0.7,q1_without - 4,Q1_text,fontsize=7)
    
    fig3_filename = save_directory + "appendix_9_fig3"
    if (third_party):
        fig3_filename += "_Third-party" + ".png"
    else:
        fig3_filename += "_In-house" + ".png"
    plt.savefig(fig3_filename)
    
    #####fig4
    fig4 = plt.figure()
    ax4 = fig4.add_subplot(111)
    bp4 = ax4.boxplot([score_with])
    ax4.set_xticklabels(["All Labs"])
    ax4.set_xlabel('Lab Group')
    plt.tick_params(axis='x', which='major', labelsize=7)
    ax4.set_title("GAP Global Correlation (2021) - Lab Performance (WITH Score Deduction)")
    ax4.title.set_size(9)
    ax4.set_ylabel("Lab Score")
    
    plt.text(1.1,median_without,"Median: " + str(median_without) + "%")
    Q3_text = "Q3 - 75% of the\ndata values are less\nthan or equal to this\n value (Q3 = " + str(q3_with) + "%)"
    Q1_text = "Q1 - 25% of the\ndata values are less\nthan or equal to this\n value (Q1 = " + str(q1_with) + "%)"
    plt.text(0.7,q3_with + 1.3,Q3_text,fontsize=7)
    plt.text(0.7,q1_with - 4,Q1_text,fontsize=7)
    
    fig4_filename = save_directory + "appendix_9_fig4"
    if (third_party):
        fig4_filename += "_Third-party" + ".png"
    else:
        fig4_filename += "_In-house" + ".png"
    plt.savefig(fig4_filename)
    
    #####fig5
    fig5 = plt.figure()
    #fig5.suptitle("Summary Report for Overall Lab Performance (WITHOUT Score Deduction)")
    ax5 = fig5.add_subplot(211)
    bin_heights, bin_borders, _ = ax5.hist(score_without)
    #def gaussian(x, mean, amplitude, standard_deviation):
    #    return amplitude * np.exp( - (x - mean)**2 / (2*standard_deviation ** 2))
    #bin_centers = bin_borders[:-1] + np.diff(bin_borders) / 2
    #print(bin_heights,bin_borders)
    #popt, _ = curve_fit(gaussian, bin_centers, bin_heights)
    #x_interval_for_fit = np.linspace(bin_borders[0], bin_borders[-1], 10000)
    #plt.plot(x_interval_for_fit, gaussian(x_interval_for_fit, *popt))
    #bp5_fit = ax5.
    
    graph_stat_without = get_graph_stat(score_without)
    
    ax5.get_yaxis().set_visible(False)
    
    ax5_2 = fig5.add_subplot(212)
    ax5_2.get_yaxis().set_visible(False)
    ax5_2.get_xaxis().set_visible(False)
    bp5_2 = ax5_2.boxplot(score_without,vert = False)
    
    fig5_filename = save_directory + "appendix_9_fig5_1"
    if (third_party):
        fig5_filename += "_Third-party" + ".png"
    else:
        fig5_filename += "_In-house" + ".png"
    plt.savefig(fig5_filename)
    
    fig5_3 = plt.figure()
    ax5_3 = fig5_3.add_subplot(111)
    ax5_3.set_title("95% Confidence Interval",fontsize = 14)
    ax5_3.errorbar(x=graph_stat_without['Mean'],y=2,xerr=[[graph_stat_without['CI_mean'][0]],[graph_stat_without['CI_mean'][1]]],fmt='o')
    ax5_3.errorbar(x=graph_stat_without['median'],y=1,xerr=[[graph_stat_without['CI_median'][0]],[graph_stat_without['CI_median'][1]]],fmt='o')
    #ax5_3.set_yticks(('Median','Mean'))
    ax5_3.get_yaxis().set_visible(False)
    ax5_3.text(graph_stat_without['CI_mean'][0],2.1,'Mean',fontsize = 7)
    ax5_3.text(graph_stat_without['CI_mean'][0],1.1,'Median',fontsize=7)
    ax5_3.set_ylim(0.25,2.25)
    
    fig5_3_filename = save_directory + "appendix_9_fig5_3"
    if (third_party):
        fig5_3_filename += "_Third-party" + ".png"
    else:
        fig5_3_filename += "_In-house" + ".png"
    plt.savefig(fig5_3_filename)
    
    ##fig 6
    fig6 = plt.figure()
    ax6 = fig6.add_subplot(211)
    bin_heights, bin_borders, _ = ax6.hist(score_without)

    graph_stat_with = get_graph_stat(score_with)
    
    ax6.get_yaxis().set_visible(False)
    
    ax6_2 = fig6.add_subplot(212)
    ax6_2.get_yaxis().set_visible(False)
    ax6_2.get_xaxis().set_visible(False)
    bp6_2 = ax6_2.boxplot(score_without,vert = False)
    
    fig6_filename = save_directory + "appendix_9_fig6_1"
    if (third_party):
        fig6_filename += "_Third-party" + ".png"
    else:
        fig6_filename += "_In-house" + ".png"
    plt.savefig(fig6_filename)
    
    fig6_3 = plt.figure()
    ax6_3 = fig6_3.add_subplot(111)
    ax6_3.set_title("95% Confidence Interval",fontsize = 14)
    ax6_3.errorbar(x=graph_stat_with['Mean'],y=2,xerr=[[graph_stat_with['CI_mean'][0]],[graph_stat_with['CI_mean'][1]]],fmt='o')
    ax6_3.errorbar(x=graph_stat_with['median'],y=1,xerr=[[graph_stat_with['CI_median'][0]],[graph_stat_with['CI_median'][1]]],fmt='o')
    ax6_3.get_yaxis().set_visible(False)
    ax6_3.text(graph_stat_with['CI_mean'][0],2.1,'Mean',fontsize = 7)
    ax6_3.text(graph_stat_with['CI_mean'][0],1.1,'Median',fontsize=7)
    ax6_3.set_ylim(0.25,2.25)
    
    fig6_3_filename = save_directory + "appendix_9_fig6_3"
    if (third_party):
        fig6_3_filename += "_Third-party" + ".png"
    else:
        fig6_3_filename += "_In-house" + ".png"
    plt.savefig(fig6_3_filename)
    
    if (third_party):
        workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_9_third_party.xlsx")
    else:
        workbook = xlsxwriter.Workbook(config.BASE_DIRECTORY_PATH + "Data/Appendix/Appendix_9_in_house.xlsx")
    
    formats = dict()
    formats["bold"] = workbook.add_format({'bold': True})
    formats["bold_underline"] = workbook.add_format({'bold': True,'underline': True})
    formats["bold_center"] = workbook.add_format({'bold': True,'align': 'center'})
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
    formats["border_center_vcenter_bold_A"] = workbook.add_format({'bg_color': 'FFD970','bold': True,'align': 'center','valign':'vcenter','border':1})
    formats["border_center_vcenter_bold_B"] = workbook.add_format({'bg_color': 'BFBFBF','bold': True,'align': 'center','valign':'vcenter','border':1})
    formats["border_center_vcenter_bold_C"] = workbook.add_format({'bg_color': '833C0C','bold': True,'align': 'center','valign':'vcenter','border':1})
    formats["border_center_vcenter_bold_D"] = workbook.add_format({'bg_color': '375623','bold': True,'align': 'center','valign':'vcenter','border':1})
    
    worksheet = workbook.add_worksheet("first")
    helper.worksheet_print_setting(worksheet)
    worksheet.write(0,0,"Appendix 9 - Laboratory Performance (Laboratory Grading)",formats['bold_underline'])
    worksheet.write(2,0,'Laboratories are categorized based on the "Overall Lab Score" according to the following Grading System.')
    worksheet.merge_range(3,0,3,3,"Table 9.1 - Laboratory Grading System",formats['bold_center'])
    worksheet.write(4,0,"Grade",formats["bold_center_border_gray"])
    worksheet.write(4,1,"Performance Level",formats["bold_center_border_gray"])
    worksheet.write(4,2,"Overall Lab Score (X)",formats["bold_center_border_gray"])
    worksheet.write(4,3,"Remark",formats["bold_center_border_gray"])
    worksheet.write(5,0,"A",formats["border_center_vcenter_bold_A"])
    worksheet.write(5,1,"Excellent",formats['center_border'])
    worksheet.write(5,2,"X > Q3",formats['center_border'])
    worksheet.write(6,0,"B",formats["border_center_vcenter_bold_B"])
    worksheet.write(6,1,"Good",formats['center_border'])
    worksheet.write(6,2,"Median < X <= Q3",formats['center_border'])
    worksheet.write(7,0,"C",formats["border_center_vcenter_bold_C"])
    worksheet.write(7,1,"Acceptable",formats['center_border'])
    worksheet.write(7,2,"Q1 < X <= Median",formats['center_border'])
    worksheet.write(8,0,"D",formats["border_center_vcenter_bold_D"])
    worksheet.write(8,1,"Need Improvement",formats['center_border'])
    worksheet.write(8,2,"X <= Q1",formats['center_border'])
    worksheet.merge_range(5,3,8,3,"Q1 = 1st Quartile\nQ3 = 3rd Quartile",formats["border_center_vcenter"])
    
    row_counter = 10
    if (third_party):
        worksheet.write(row_counter,0,"Fig. A9.1 - Overall Lab Performance (Lab Score WITH and WITHOUT Score Deduction) (by Lab Group)",formats['bold'])
        worksheet.insert_image(row_counter + 1,0,config.BASE_DIRECTORY_PATH + "Data/Appendix/boxplot/appendix_9_overall.png")
        row_counter += 30
    worksheet.write(row_counter,0,"Fig. A9.2 - Overall Lab Performance (Lab Score WITH and WITHOUT Score Deduction) (All Labs)",formats['bold'])
    worksheet.insert_image(row_counter + 1,0,fig2_filename)
    
    
    #second page
    worksheet = workbook.add_worksheet("Second")
    helper.worksheet_print_setting(worksheet)
    row_counter = 0
    worksheet.write(row_counter,0,"Fig. A9.3 - Overall Lab Performance (Lab Score WITHOUT Score Deduction) (All Labs)",formats['bold'])
    worksheet.insert_image(row_counter + 1,0,fig3_filename)
    row_counter += 30
    worksheet.write(row_counter,0,"Fig. A9.4 - Overall Lab Performance (Lab Score WITH Score Deduction ) (All Labs)",formats['bold'])
    worksheet.insert_image(row_counter + 1,0,fig4_filename)
    
    ###thid page
    worksheet = workbook.add_worksheet("Third")
    helper.worksheet_print_setting(worksheet)
    row_counter = 0
    worksheet.write(row_counter,0,"Fig. A9.5 - Graphical Summary of Overall Performance (Lab Score WITHOUT Score Deduction) (All Labs)",formats['bold'])
    stat_col = 14
    app_9_write_stat(worksheet,row_counter,stat_col,graph_stat_without)
    worksheet.insert_image(row_counter + 1,0,fig5_filename)
    row_counter += 30
    worksheet.insert_image(row_counter + 1,0,fig5_3_filename)
    row_counter += 30
    worksheet.write(row_counter,0,"Fig. A9.6 - Graphical Summary of Overall Lab Performance (Lab Score WITH Score Deduction) (All Labs)",formats['bold'])
    app_9_write_stat(worksheet,row_counter,stat_col,graph_stat_with)
    worksheet.insert_image(row_counter + 1,0,fig6_filename)
    row_counter += 30
    worksheet.insert_image(row_counter + 1,0,fig6_3_filename)
    row_counter += 30
    
    
    #grading
        #without score deduction
    worksheet = workbook.add_worksheet("Grading")
    helper.worksheet_print_setting(worksheet)
    worksheet.merge_range(0,0,0,4,"Table 9.2 - Grading of Participating Laboratories (Lab Performance WITHOUT Score Deduction)",formats['bold_center'])
    worksheet.write(1,0,"Lab Group",formats["bold_border_gray_vcenter_center"])
    worksheet.write(1,1,"Grade A \n(Excellent)",formats["border_center_vcenter_bold_A"])
    worksheet.write(1,2,"Grade B \n(Good)",formats["border_center_vcenter_bold_B"])
    worksheet.write(1,3,"Grade C \n(Acceptable)",formats["border_center_vcenter_bold_C"])
    worksheet.write(1,4,"Grade D \n(Need Improvement)",formats["border_center_vcenter_bold_D"])
    
    worksheet.set_column(0,0,10)
    worksheet.set_column(1,4,40)
    row_counter = 2
    A_count = 0
    B_count = 0
    C_count = 0
    D_count = 0
    for group,labs in Labs.Groups.items():
        if (third_party):
            if (group == "In-House"):
                continue
        else:
            if (group != "In-House"):
                continue
        worksheet.write(row_counter,0,"BV",formats["border_center_vcenter"])
        A_text = ""
        B_text = ""
        C_text = ""
        D_text = ""
        for lab in labs.values():
            score = lab.score_without_deduction
            if score >= q3_without:
                A_count += 1
                A_text += grade_str(lab,score)
            elif score >= median_without:
                B_count += 1
                B_text += grade_str(lab,score)
            elif score >= q1_without:
                C_count += 1
                C_text += grade_str(lab,score)
            else:
                D_count += 1
                D_text += grade_str(lab,score)
        worksheet.write(row_counter,1,A_text,formats['border'])
        worksheet.write(row_counter,2,B_text,formats['border'])
        worksheet.write(row_counter,3,C_text,formats['border'])
        worksheet.write(row_counter,4,D_text,formats['border'])
            
        row_counter += 1
    worksheet.write(row_counter,0,'Total',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,1,str(A_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,2,str(B_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,3,str(C_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,4,str(D_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    row_counter += 3
    
        #with score deduciton
    worksheet.merge_range(row_counter,0,row_counter,4,"Table 9.3 - Grading of Participating Laboratories (Lab Performance WITH Score Deduction)",formats['bold_center'])
    worksheet.write(row_counter + 1,0,"Lab Group",formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter + 1,1,"Grade A \n(Excellent)",formats["border_center_vcenter_bold_A"])
    worksheet.write(row_counter + 1,2,"Grade B \n(Good)",formats["border_center_vcenter_bold_B"])
    worksheet.write(row_counter + 1,3,"Grade C \n(Acceptable)",formats["border_center_vcenter_bold_C"])
    worksheet.write(row_counter + 1,4,"Grade D \n(Need Improvement)",formats["border_center_vcenter_bold_D"])
    row_counter += 2
    A_count = 0
    B_count = 0
    C_count = 0
    D_count = 0
    for group,labs in Labs.Groups.items():
        if (third_party):
            if (group == "In-House"):
                continue
        else:
            if (group != "In-House"):
                continue
        worksheet.write(row_counter,0,"BV",formats["border_center_vcenter"])
        A_text = ""
        B_text = ""
        C_text = ""
        D_text = ""
        for lab in labs.values():
            score = lab.score_with_deduction
            if score >= q3_with:
                A_count += 1
                A_text += grade_str(lab,score)
            elif score >= median_with:
                B_count += 1
                B_text += grade_str(lab,score)
            elif score >= q1_with:
                C_count += 1
                C_text += grade_str(lab,score)
            else:
                D_count += 1
                D_text += grade_str(lab,score)
        worksheet.write(row_counter,1,A_text,formats['border'])
        worksheet.write(row_counter,2,B_text,formats['border'])
        worksheet.write(row_counter,3,C_text,formats['border'])
        worksheet.write(row_counter,4,D_text,formats['border'])
            
        row_counter += 1
    worksheet.write(row_counter,0,'Total',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,1,str(A_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,2,str(B_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,3,str(C_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    worksheet.write(row_counter,4,str(D_count) + ' Labs',formats["bold_border_gray_vcenter_center"])
    
    workbook.close()
    
def main(app_8 = False):
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json') as json_file:
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        Project = DataClasses.Project.from_dict(json.load(json_file))
        
    save_directory = config.BASE_DIRECTORY_PATH + "Data/Appendix/boxplot/"
    outlier_overall_without,outlier_overall_with = create_overall(Labs,save_directory)
        
    outlier_samples = list()
    for sample_id in range(1,Project.number_sample+1):
        outlier_samples.append(create_sample(Labs,sample_id,save_directory))
        
    create_overall(Labs,save_directory,app_9 = True)
    
    if (app_8):
        create_appendix_8(outlier_overall_without,outlier_overall_with,outlier_samples,Project.number_sample,groups = list(Labs.Groups))
    else:
        create_appendix_9(Labs,save_directory,third_party = True)
        create_appendix_9(Labs,save_directory,third_party = False)
    
    
if __name__ == '__main__':
    main()