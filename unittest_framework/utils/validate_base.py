import unittest
import json, time, logging, re, inspect
from datetime import datetime
from utils.common_utils import set_logger
from utils.EmailReportCreation import HtmlReport
import pandas as pd


class MyTestBase(unittest.TestCase):
    @classmethod
    def setUpClass(cls, modulename):
        with open("config/test_config.json", "r") as f:
            cls.test_config = json.load(f)
        module_name = "First Unit test"
        cls.overall_status = dict()
        cls.overall_status["StartTime"] = datetime.now()
        log_file = cls.test_config["log_file"].replace(".log",
                                                       f'_{cls.test_config["app_name"]}_{module_name}_{cls.overall_status["StartTime"].strftime("%Y_%m_%d_%H_%M_%S")}.log')
        set_logger(cls.test_config["log_config"], log_file_name=log_file)
        cls.logger = logging.getLogger(__name__)
        cls.logger.info("starting testsuite")
        cls.status_data = list()
        cls.overall_status["APPName"] = cls.test_config["app_name"]
        cls.overall_status["ModuleName"] = module_name
        cls.overall_status["env_config"] = cls.test_config["env_config"]
        cls.overall_status["report_template"] = cls.test_config["report_template"]

    @classmethod
    def get_exec_methods(cls):
        return [attr for attr in dir(cls) if re.match("test_(\d)+", attr) is not None]

    @classmethod
    def tearDownClass(cls):
        cls.logger.info("executed all testcases")
        cls.overall_status["EndTime"] = datetime.now()
        cls.overall_status["PassedPercent"] = (len(list(
            filter(lambda x: x["Status"] == "PASS", cls.status_data))) / len(cls.status_data)) * 100
        cls.overall_status["FailedPercent"] = (len(list(
            filter(lambda x: x["Status"] == "FAIL", cls.status_data))) / len(cls.status_data)) * 100
        cls.overall_status["ErrorPercent"] = (len(list(
            filter(lambda x: x["Status"] == "ERROR", cls.status_data))) / len(cls.status_data)) * 100
        status_df = pd.DataFrame(cls.status_data)
        HtmlReport(overallData=cls.overall_status, header=status_df.columns, tableData=status_df.values.tolist(),
                   filename=cls.test_config["report_file"].replace(".html",
                                                                   f'_{cls.overall_status["APPName"]}_{cls.overall_status["ModuleName"]}_{cls.overall_status["StartTime"].strftime("%Y_%m_%d_%H_%M_%S")}.html'))
        logging.shutdown()

    def setUp(self):
        self.tc_start_time = time.time()
        self.current_tc_desc = self._testMethodName.replace("_", " ").capitalize()
        self.priority = "High"
        self.logger.info("\n" + "*" * 10 + f" Starting Testcase:{self._testMethodName} " + "*" * 10 + "\n")

    def tearDown(self):
        self.tc_end_time = time.time()
        self.logger.info("\n" + "*" * 10 + f" Ending Testcase:{self._testMethodName} " + "*" * 10 + "\n")
        if hasattr(self, '_outcome'):  # Python 3.4+
            result = self.defaultTestResult()  # This two methods has no side effects
            self._feedErrorsToResult(result, self._outcome.errors)
        else:  # Python 3.0-3.3 and 2.7
            result = getattr(self, '_outcomeForDoCleanups', self._resultForDoCleanups)
        error = self.list2reason(result.errors)
        failure = self.list2reason(result.failures)
        ok = not error and not failure
        if not ok:
            typ, text = ('ERROR', error) if error else ('FAIL', failure)
            msg = [x for x in text.split('\n')[1:] if not x.startswith(' ')][0]
            if typ == "ERROR":
                start_error_msg = f"testcase stopped unexpectedly: \n{typ}: {result.errors[0][0]}"
                start_error_msg = "\n" + "=" * len(start_error_msg) + "\n" + start_error_msg + "\n" + "-" * len(
                    start_error_msg) + "\n"
                self.logger.critical(msg=start_error_msg + "\n".join(result.errors[0][1:]))
            elif typ == "FAIL":
                start_error_msg = f"Validation failed: \n{typ}: {result.failures[0][0]}"
                start_error_msg = "\n" + "=" * len(start_error_msg) + "\n" + start_error_msg + "\n" + "-" * len(
                    start_error_msg) + "\n"
                self.logger.error(msg=start_error_msg + "\n".join(result.failures[0][1:]))
        else:
            msg = ""
            typ = "PASS"
        self.status_data.append(
            {"Testcase ID": self._testMethodName, "Description": self.current_tc_desc,
             "Priority": self.priority.capitalize(), "Status": typ,
             "Error/Failure message(If Any)": msg,
             "Execution Time": f"{self.tc_end_time - self.tc_start_time} seconds"})

    def list2reason(self, exc_list):
        if exc_list and exc_list[-1][0] is self:
            return exc_list[-1][1]


def testSuite(MyTestClass, sysargs):
    args_dd = {}
    for args in range(0, len(sysargs) - 1, 2):
        args_dd[sysargs[args]] = sysargs[args + 1].split(",")
    frame = inspect.stack()[1]
    module = inspect.getmodule(frame[0])
    filename = module.__file__
    tc_tag = {}
    suite = unittest.TestSuite()
    with open(filename) as f:
        tag_tc_map = re.finditer(r"\#\@tag: (.*\w)+|def (test.*)\(", f.read())
        current_tag = []
        current_tc = ""
        for tag_or_tc in tag_tc_map:
            if tag_or_tc.group(1) is None:
                current_tc = tag_or_tc.group(2)
                tc_tag[current_tc] = current_tag
                current_tag = []
            else:
                current_tag.extend(tag_or_tc.group(1).split(" "))
    all_tests = MyTestClass.get_exec_methods()
    to_be_executed = []
    not_to_be_executed = []
    if len(args_dd) > 0:
        for tests in all_tests:
            if any(True for arg in args_dd.get("-i", [-1]) if arg in tc_tag.get(tests, [])):
                to_be_executed.append(tests)
            if any(True for arg in args_dd.get("-e", [-1]) if arg in tc_tag.get(tests, [])):
                not_to_be_executed.append(tests)
            if re.match(args_dd.get("-t", ["-1"])[0], tests) is not None:
                to_be_executed.append(tests)
    else:
        to_be_executed = all_tests
    for tests in to_be_executed:
        if tests not in not_to_be_executed:
            pass
            suite.addTest(MyTestClass(tests))
    return suite
