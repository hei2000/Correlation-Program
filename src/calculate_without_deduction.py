import DataClasses
import logging
import config,json
import pandas as pd
import numpy as np

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
    
Statistics = DataClasses.Statistics()
    
def calculate_test_statistics(df,test,sample_id):
    df = df[df["Lab Group"] != "In-House"]
    unique_test = DataClasses.Unique_Test(name=test.name,method=test.method,sample=sample_id)
    #print(type(unique_test.__hash__()) == type(1))
    #print(isinstance(str(hash(unique_test)), int))
    #print(type(str(hash(unique_test))))
    Statistics.Statistics[unique_test.__hash__()] = list()
    for result in test.Results:
        statistic_result = DataClasses.Statistic_Result(name = result.name,type = result.type)
        if result.type == "Quantitative":
            median = df["result_" + result.name].median()
            X = list(df.dropna()["result_" + result.name])
            if (len(X) == 1):
                statistic_result.Q3 = round(X[0],2)
                statistic_result.Q1 = round(X[0],2)
                statistic_result.median = round(X[0],2)
                statistic_result.NIQR = round(0.01,2)
                Statistics.Statistics[unique_test.__hash__()].append(statistic_result)
                continue
            X.sort()
            Q1 = np.median(X[:len(X)//2])
            Q3 = np.median(X[len(X)//2 + (len(X)%2):])
            #Q3 = float(df["result_" + result.name].quantile(.75,interpolation='midpoint'))
            #Q1 = float(df["result_" + result.name].quantile(.25,interpolation='midpoint'))
            IQR = Q3 - Q1
            #print(result.name," Q3:" + str(df["result_" + result.name].quantile(.25,interpolation='midpoint')))
            #print(result.name," Q1:" + str(df["result_" + result.name].quantile(.75,interpolation='midpoint')))
            #print(result.name," IQR:" + str(IQR))
            NIQR = IQR * 0.7413
            statistic_result.Q3 = round(Q3,2)
            statistic_result.Q1 = round(Q1,2)
            statistic_result.median = round(median,2)
            statistic_result.NIQR = round(NIQR,2)
        elif result.type == "Semi-Quantitative":
            #print(df[result.name].mode())
            try:
                X = list(df.dropna()["result_" + result.name])
                vals,counts = np.unique(X, return_counts=True)
                index = np.argmax(counts)
                mode = vals[index]
                #mode = df["result_" + result.name].mode().iat[0]
                statistic_result.mode = round(float(mode),1)
            except:
                print("No data")
        if result.type != "Observation":
            maximum = float(df["result_" + result.name].max())
            minimum = float(df["result_" + result.name].min())
            range_ = maximum - minimum
            statistic_result.maximum = maximum
            statistic_result.minimum = minimum
            statistic_result.range = range_
        #print(statistic_result)
        Statistics.Statistics[unique_test.__hash__()].append(statistic_result)
        
def get_semi_score(value,unique_test,count):
    statistic = Statistics.Statistics[unique_test.__hash__()][count]
    mode = statistic.mode
    away_from_mode = value - mode
    if (abs(away_from_mode) == 0):
        return away_from_mode,2
    elif (abs(away_from_mode) <= 0.5):
        return away_from_mode,1.5
    else:
        return away_from_mode,0

def calculate_z_score(Xi,X,NIQR):
    #print(Xi,X,NIQR)
    diff = (Xi-X)
    if NIQR == 0:
        return 100 if diff > 0 else 0
    return diff/NIQR


def get_quan_score(value,unique_test,count):
    statistic = Statistics.Statistics[unique_test.__hash__()][count]
    median = statistic.median
    NIQR = statistic.NIQR
    z_score = round(calculate_z_score(value,median,NIQR),2)
    abs_score = abs(z_score)
    if (abs_score <= 2):
        score = 2
    elif (abs_score < 3):
        score = 1
    else:
        score = 0
    return z_score,score
        
def update_lab(df,test,sample_id,Labs):
    for idx,lab_test_data in df.iterrows():
        test_sample = lab_test_data["Sample"]
        group = lab_test_data["Lab Group"]
        id_ = lab_test_data["Lab ID"]
        update_sample = Labs.Groups[group][id_].Samples[str(test_sample)]
        for test_type,tests in update_sample.Tests.items():
            for test in tests:
                test_name = lab_test_data["Test Name"]
                test_method = lab_test_data["Test Method"]
                #print(test_name,test_method)
                #print(test.name,test.method)
                if (test.name == test_name) and (test.method == test_method):
                    test.result_rating = lab_test_data["Result Rating"]
                    #print(test)
                    for requirement in test.Requirements:
                        test.Requirements_data.append(lab_test_data["requirement_" + requirement])
                    for count,result in enumerate(test.Results):
                        result.value = lab_test_data["result_" + result.name]
                        unique_test = DataClasses.Unique_Test(name=test_name,method=test_method,sample=test_sample)
                        if result.type == "Quantitative":
                            z_score,score = get_quan_score(result.value,unique_test,count)
                            result.Z_score = z_score
                            result.score = score
                            if (score == 0):
                                test.outlier = True
                            elif (score == 1):
                                test.straggler = True
                        elif result.type == "Semi-Quantitative":
                            away_from_mode, score = get_semi_score(result.value,unique_test,count)
                            result.score = score
                            result.away_from_mode = round(float(away_from_mode),1)
                            if (score == 0):
                                test.outlier = True
                            elif (score == 1.5):
                                test.straggler = True
                    continue
        
    
def calculate_test(test:DataClasses.Test,sample_id,Labs):
    filepath = config.BASE_DIRECTORY_PATH + "Data/Result/Sample " + str(sample_id) + "/" + test.method + "_" + test.name + ".xlsx"
    df = pd.read_excel(filepath,index_col = False,converters = {'Lab ID':str})
    calculate_test_statistics(df,test,sample_id)
    update_lab(df,test,sample_id,Labs)    #value,z_score,score, result rating
    
def save_statistics():
    Statistics_json = Statistics.to_dict()
    with open(config.BASE_DIRECTORY_PATH + 'Data/Statistics.json','w',newline='') as fp:
        json.dump(Statistics_json,fp)

def calculate_average_score_test(test):
    accumulate_score = 0
    test_count = 0
    for result in test.Results:
        if result.type != "Observation":
            test_count += 1
            accumulate_score += result.score
    #assume at least one test
    average_score = accumulate_score/test_count
    #print(average_score)
    test.average_score = average_score
    return

def calculate_score_sample(sample):
    accumulate_score = 0
    test_count = 0
    PSR_accumulate_score = 0
    PSR_test_count = 0
    SPC_accumulate_score = 0
    SPC_test_count = 0
    for type_,tests in sample.Tests.items():
        for test in tests:
            accumulate_score += test.average_score
            test_count += 1
            if (type_ == "PSR"):
                PSR_accumulate_score += test.average_score
                PSR_test_count += 1
            if (type_ == "SPC"):
                SPC_accumulate_score += test.average_score
                SPC_test_count += 1
    score = round(accumulate_score / test_count * 50,2)
    sample.score = score
    if (PSR_test_count != 0):
        PSR_score = round(PSR_accumulate_score / PSR_test_count * 50,2)
        sample.PSR_score = PSR_score
    if (SPC_test_count != 0):
        SPC_score = round(SPC_accumulate_score / SPC_test_count * 50,2)
        sample.SPC_score = SPC_score
    return

def calculate_average_score_lab(lab):
    accumulate_score = 0
    test_count = 0
    PSR_accumulate_score = 0
    PSR_test_count = 0
    SPC_accumulate_score = 0
    SPC_test_count = 0
    for sample in lab.Samples.values():
        for type_,tests in sample.Tests.items():
            for test in tests:
                accumulate_score += test.average_score
                #print(lab.ID,": " , accumulate_score)
                test_count += 1
                if (type_ == "PSR"):
                    PSR_accumulate_score += test.average_score
                    PSR_test_count += 1
                if (type_ == "SPC"):
                    SPC_accumulate_score += test.average_score
                    SPC_test_count += 1
    average_score = accumulate_score/test_count
    #print(average_score)
    lab.average_score = average_score
    lab.score_without_deduction = round(average_score * 50 , 1)
    if (PSR_test_count != 0):
        PSR_score = round(PSR_accumulate_score / PSR_test_count * 50,1)
        lab.PSR_score = PSR_score
    if (SPC_test_count != 0):
        SPC_score = round(SPC_accumulate_score / SPC_test_count * 50,1)
        lab.SPC_score = SPC_score
    return
    
def save_Labs(Labs):
    Labs_json = Labs.to_dict()
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_without_deduction.json','w',newline='') as fp:
        json.dump(Labs_json,fp)

def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Project.json') as json_file:
        Project = DataClasses.Project.from_dict(json.load(json_file))
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs.json') as json_file:  
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    for sample in Project.Samples.values():
        print("Processing sample " + str(sample.sample_id))
        for test_type,tests in sample.Tests.items():
            print("\tProcessing " + test_type + " tests")
            for test in tests:
                print("\t\tProcessing test " + test.method + "_" + test.name)
                calculate_test(test,sample.sample_id,Labs)
        print("Finished processing sample " + str(sample.sample_id))
        print()
        
    for labs in Labs.Groups.values():
        for lab in labs.values():
            for sample in lab.Samples.values():
                for tests in sample.Tests.values():
                    for test in tests:
                        calculate_average_score_test(test)
                calculate_score_sample(sample)
            calculate_average_score_lab(lab)
    
    save_statistics()   #cannot debug unhashable type 'dict'
    save_Labs(Labs)
    #print(Labs)
    #print(Statistics)

if __name__ == '__main__':
    logging.info("Entered 'calculate_without_deduction.py'")
    main()
