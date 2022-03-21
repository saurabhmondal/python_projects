import json
import logging
import time
import unittest
from datetime import datetime
from utils.common_utils import set_logger
from utils.EmailReportCreation import HtmlReport
import pandas as pd


class MyTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
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
            {"Testcase ID": self._testMethodName, "Description": self.current_tc_desc, "Status": typ,
             "Error/Failure message(If Any)": msg,
             "Execution Time": f"{self.tc_end_time - self.tc_start_time} seconds"})

    def list2reason(self, exc_list):
        if exc_list and exc_list[-1][0] is self:
            return exc_list[-1][1]

    def test_something(self):
        self.current_tc_desc = "Positive Testcase should pass"
        self.assertEqual(True, False)  # add assertion here


if __name__ == '__main__':
    unittest.main()
