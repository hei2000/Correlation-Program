MAX_REQUIREMENT = 15
MAX_RESULT = 15
FABRIC_TYPE = ['Denim','Sweater','Woven','Knit']
TYPE_OF_TEST = ["PSR","SPC","Performance"]
RESULT_TYPE = ["Quantitative","Semi-Quantitative","Observation"]
RESTRICTED_CHAR = ['\\','/',':','*',"?","<",">","|","_"]
LAB_GROUPS = ['ITS','BV','ABC','In-House']

#Report Generation
APPENDIX = [(1,'List of Participating Laboratories'),
            (2,'Test Selection by Participating Laboratories'),
            (3,'Laboratory Performance (Overall)'),
            (4,'Laboratory Performance (Laboratory Errors)'),
            #(5,'Laboratory Performance (Observations)'),
            (6,'Laboratory Performance (PSR Test Items)'),
            (7,'Laboratory Performance (SPC Test Items)'),
            (8,'Laboratory Performance (Boxplot Analysis)'),
            (9,'Laboratory Performance (Laboratory Grading)'),
            (10,'Laboratory Test Data')
            ]

ERRORS = ["Type 1 :Missing / incorrect individual test rating","Type 2 :Missing / incorrect test requirement","Type 3 :Fail to correctly perform test, report test info., etc. in accordance with GAP protocol or test standard","Type 4: Extra test results","Type 5: Other test related errors (Wrong Test Unit Applied, Wrong Calculation","Type 6 : Other non-test related errors (Typo Mistake)"]

#Result template
RESULT_TEMPLATE = ['Lab ID','Lab Group','Lab Location','Lab Name','Third-party','Test Name','Test Method','Sample']

#Absolute Paths
#BASE_DIRECTORY_PATH =  "C:/Users/matthew.chan/Desktop/python3.8.10 version/"
#BASE_DIRECTORY_PATH = "U:/Correlation Program/"
BASE_DIRECTORY_PATH = "./"

PROGRAM_ID = "GAP GLOBAL PT_2021"
YEAR = "2021"