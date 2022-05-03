import unittest
from utils.validate_base import MyTestBase, testSuite
import sys

class MyTest(MyTestBase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass("TestModule")

    def test_01(self):
        self.current_tc_desc = "Positive Testcase should pass"
        self.assertEqual(True, True)  # add assertion here
    #@tag: edge lower
    #@tag: high
    def test_02(self):
        # self.current_tc_desc = "Negative Testcase should pass"
        self.assertEqual(True, False)  # add assertion here

    #@tag: lower upper
    #@tag:
    def test_03(self):
        self.priority = "low"
        self.current_tc_desc = "Error Testcase should pass"
        dict["12"] = 0
        self.assertEqual(True, False)  # add assertion here


if __name__ == '__main__':
    runner = unittest.TextTestRunner()
    runner.run(testSuite(MyTest,sys.argv[1:]))
