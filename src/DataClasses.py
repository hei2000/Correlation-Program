#import dataclasses
from dataclasses import dataclass,field
from dataclasses_json import dataclass_json
from typing import List,Dict
import json

@dataclass_json
@dataclass
class Error:
    name: str
    type_number: int
    number: int = 0
    description: str = ""
    
@dataclass_json
@dataclass
class Result:
    name: str
    type: str
    value: str = ""
    Z_score: float = -100 #-100 if not applicable
    score: float = -1
    away_from_mode: float = -100

@dataclass_json
@dataclass
class Test:
    fabric_type: str
    group: str
    type: str
    name: str
    method: str
    Requirements: List[str] = field(default_factory = list)
    Requirements_data: List[str] = field(default_factory = list)
    Results: List[Result] = field(default_factory = list)
    Errors: Dict[str,Error] = field(default_factory = dict)
    average_score: float = -1
    result_rating: str = ""
    total_error: int = -1
    outlier: bool = False
    straggler: bool = False
    
@dataclass_json
@dataclass
class Tests:
    tests: List[Test] = field(default_factory = list)
    
@dataclass_json
@dataclass
class Sample:
    fabric_type: str
    sample_id: int
    Tests: Dict[str,List[Test]] = field(default_factory = dict) #TYPE_OF_TEST = ["PSR","SPC","Performance"]
    total_error:int = -1
    score: float = -1
    PSR_score: float = -1
    SPC_score: float = -1
    
@dataclass_json
@dataclass
class Project:
    number_sample: int
    Samples: Dict[str,Sample] = field(default_factory = dict)
    name: str = ""
    
@dataclass_json
@dataclass
class Lab:
    ID: str
    group: str
    location: str
    fullname: str
    third_party: bool
    city: str = ""
    Samples: Dict[str,Sample] = field(default_factory = dict)
    deduction_timeliness: int=0 #days
    deduction_revision: int=0 #times
    imported: bool = False
    import_date: str = ""
    average_score: float = -1
    score_without_deduction: float = -1
    score_with_deduction: float = -1
    PSR_score: float = -1
    SPC_score: float = -1
    total_test_error: int = -1
    PSR_error: int = -1
    SPC_error: int = -1
    
@dataclass_json
@dataclass(eq = True,frozen=True)
#@dataclass(unsafe_hash = True)
class Unique_Test:
    name: str
    method: str
    sample: int
    def __hash__(self):
        return self.method + "_" + self.name + "_" + str(self.sample)
    
@dataclass_json
@dataclass
class Statistic_Result:
    name: str
    type: str
    median: float = -100
    NIQR: float = -100
    mode: float = -100
    maximum: float = -100
    minimum: float = -100
    range: float = -100
    Q1: float = -100
    Q3: float = -100
    
@dataclass_json
@dataclass
class Statistics:
    Statistics: Dict[str,List[Statistic_Result]] = field(default_factory = dict)
    
@dataclass_json
@dataclass                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
class Labs:
    Groups: Dict[str,Dict[str,Lab]] = field(default_factory = dict)
    
def test():
    Test_lab = Lab(ID = "L601",group="ITS",location = "Hong Kong",fullname="HONG KONG-ITS",third_party=True)
    Test_labs = Labs()
    if ("ITS" not in Test_labs.Groups):
        Test_labs.Groups["ITS"] = list()
    Test_labs.Groups["ITS"].append(Test_lab)
    #print(Test_labs.Groups)
    Test_labs_json = Test_labs.to_json()
    #print(Test_labs_json)
    with open("./Test_Data/Labs.json", "w") as outfile:
        outfile.write(Test_labs_json)
        
    with open('./Test_Data/Labs.json') as json_file:
        Test_labs_from_json_file = Labs.from_dict(json.load(json_file))
    print(Test_labs_from_json_file)
    
#test()

def test2():
    Test_Statistics = Statistics()
    unique_test = Unique_Test(name="1",method="2",sample=3)
    #statistic_result = Statistic_Result(name = "results name",type = "type")
    #print(hash(Test_Statistics))
    Test_Statistics.Statistics[unique_test] = list()
    print(Test_Statistics)
    Test_Statistics_json = Test_Statistics.to_json()

#test2()