import DataClasses
import logging
import config,json

logging.basicConfig(filename="logging.log",level = logging.DEBUG,format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

def update_lab(lab):
    lab_error_count = 0
    PSR_error_count = 0
    SPC_error_count = 0
    for sample in lab.Samples.values():
        sample_error_count = 0
        for tests in sample.Tests.values():
            for test in tests:
                test_error_count = 0
                for error in test.Errors.values():
                    test_error_count += error.number
                    if (test.type == "SPC"):
                        SPC_error_count += error.number
                    if (test.type == "PSR"):
                        PSR_error_count += error.number
                test.total_error = test_error_count
                sample_error_count += test.total_error
        sample.total_error = sample_error_count
        lab_error_count += sample.total_error
    lab.total_test_error = lab_error_count
    lab.score_with_deduction = lab.score_without_deduction - 0.5 * (lab.total_test_error + lab.deduction_timeliness + lab.deduction_revision)
    lab.SPC_error = SPC_error_count
    lab.PSR_error = PSR_error_count
    
def save_labs(Labs):
    Labs_json = Labs.to_dict()
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_deduction.json','w',newline='') as fp:
        json.dump(Labs_json,fp)
    
def main():
    with open(config.BASE_DIRECTORY_PATH + 'Data/Labs_with_errors.json') as json_file:  
        Labs = DataClasses.Labs.from_dict(json.load(json_file))
    for labs in Labs.Groups.values():
        for lab in labs.values():
            update_lab(lab)
    save_labs(Labs)
    
if __name__ == '__main__':
    logging.info("Entered 'calculate_without_deduction.py'")
    main()