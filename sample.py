from util.ParseExcel import ParseExcel
from config.VarConfig import *
from action.PageAction import *
import traceback
from util.log import Logger
import logging
import pytest
import time

log = Logger(__name__, CmdLevel=logging.INFO, FileLevel=logging.INFO)
p = ParseExcel()
sheetName = p.wb.sheetnames  # 获取到excel的所有sheet名称


def code_generation():
    try:
        testCasePassNum = 0

        requiredCase = 0
        isExecuteColumnValues = p.getColumnValue(sheetName[0], testCase_testIsExecute)
        print(isExecuteColumnValues)
        print(len(isExecuteColumnValues))
        # 项目根目录
        projectPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        # 测试用例目录
        Time = datetime.now()
        currentTime = Time.strftime('%H_%M_%S')
        # file_name = time.strftime('')
        global filePath
        filePath = projectPath + r'\testcase\test_dynamic_method_%s.py' % currentTime
        # filePath = 'test_dynamic_method'
        with open(filePath, 'a+', encoding='utf-8') as file:
            # 引入关键字和allure
            file.writelines('from action.PageAction import *' + '\n')
            file.writelines('import allure' + '\n' * 2)

        for index, value in enumerate(isExecuteColumnValues):
            print(index, value)
            # 获取对应的步骤sheet名称
            stepSheetName = p.getCellOfValue(sheetName[0], index + 2, testCase_testStepName)
            # print(stepSheetName)
            if value is not None and value.strip().lower() == 'y':
                testCaseDic = {}
                testCase = []
                testCaseKeyWords = []
                keyWords = []
                requiredCase += 1
                testStepPassNum = 0

                # 添加待执行的测试用例
                testCase.append(stepSheetName)
                print('开始执行测试用例"{}"'.format(stepSheetName))
                log.logger.info('开始执行测试用例"{}"'.format(stepSheetName))
                # 如果用例被标记为执行y，切换到对应的sheet页
                # 获取对应的sheet表中的总步骤数，关键字，定位方式，定位表达式，操作值
                # 步骤总数
                values = p.getColumnValue(stepSheetName, testStep_testNum)  # 第一列数据
                stepNum = len(values)
                print(stepNum)
                for step in range(2, stepNum + 2):
                    rawValue = p.getRowValue(stepSheetName, step)
                    # 执行步骤名称
                    stepName = rawValue[testStep_testStepDescribe - 2]
                    # 关键字
                    keyWord = rawValue[testStep_keyWord - 2]
                    # 定位方式
                    by = rawValue[testStep_elementBy - 2]
                    # 定位表达式
                    locator = rawValue[testStep_elementLocator - 2]
                    # 操作值
                    operateValue = rawValue[testStep_operateValue - 2]

                    if keyWord and by and locator and operateValue:
                        func = keyWord + '(' + '"' + by + '"' + ',' + '"' + locator + '"' + ',' + '"' + operateValue + '"' + ')'
                    elif keyWord and by and locator and operateValue is None:
                        func = keyWord + '(' + '"' + by + '"' + ',' + '"' + locator + '"' + ')'

                    elif keyWord and operateValue and type(operateValue) == str and by is None and locator is None:
                        func = keyWord + '(' + '"' + operateValue + '"' + ')'

                    elif keyWord and operateValue and type(operateValue) == int and by is None and locator is None:
                        func = keyWord + '(' + str(operateValue) + ')'

                    else:
                        func = keyWord + '(' + ')'

                    keyWords.append(func)

                # 添加组装好的关键字
                testCaseKeyWords.append(keyWords)
                testcase_name = p.getCellOfValue(sheetName[0], index + 2, testCase_testCaseName)
                # testcase_des = p.getCellOfValue(sheetName[0], index + 2, testCase_testCaseDes)
                with open(filePath, 'a+', encoding='utf-8') as file:
                    # 写入测试用例名称
                    file.write("@allure.title('%s')" % testcase_name + '\n')
                    # 写入测试用例描述
                    # file.write("@allure.description('%s')" % testcase_des + '\n')

                for k, v in zip(testCase, testCaseKeyWords):
                    with open(filePath, 'a+', encoding='utf-8') as file:
                        # 写入方法名
                        file.write('def ' + k + '():' + '\n')

                    # 将测试方法名与关键字序列一一对应
                    testCaseDic[k] = v
                    print(testCaseDic)

                    # 遍历关键字列表组装成动态测试方法
                    for i in v:
                        with open(filePath, 'a+', encoding='utf-8') as file:
                            # 写入关键字名
                            file.write('\t' + i + '\n')

    except Exception as e:
        print(traceback.format_exc(e))
        log.logger.info(traceback.format_exc(e))
    return filePath


if __name__ == '__main__':
    # global filePath
    file = code_generation()
    if os.path.isfile(file):
        pytest.main(["-v", "-s", "--alluredir=report/allure_results", filePath])
        # 删除动态生成的测试文件
    os.remove(file)