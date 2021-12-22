#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import logging
import os.path
import xlwt, xlrd

from xmind2case.utils import get_xmind_testcase_list

"""
Convert XMind fie to Htp testcase xlsx file 
"""


def xmind_to_htp_xlsx_file(xmind_file):
    """Convert XMind fie to Htp testcase xlsx file """
    xmind_file = get_xmind_testcase_list(xmind_file)
    logging.info('Start converting XMind file(%s) to zentao file...', xmind_file)
    testcases = get_xmind_testcase_list(xmind_file)

    fileheader = ['用例编号', '用例树目录', '用例名称', '摘要', '前置条件', '测试步骤', '预期结果', '用例等级', '自动化覆盖', '状态', '用例类型']
    htp_testcase_rows = [fileheader]
    for testcase in testcases:
        row = gen_a_testcase_row(testcase)
        htp_testcase_rows.append(row)

    htp_file = xmind_file[:-6] + '.xlsx'
    print(fileheader)
    # workbook = xlwt.Workbook(encoding='utf-8')
    # mysheet = mySheet = workbook.add_sheet(htp_file)
    # for row, scores in enumerate(fileheader):
    return htp_file

def gen_a_testcase_row(testcase_dict):
    case_number = ''
    case_tree = get_case_module(testcase_dict['suite'])
    case_title = testcase_dict['name']
    case_summary = ''
    case_precontion = testcase_dict['preconditions']
    case_step, case_expected_result = gen_case_step_and_expected_result(testcase_dict['setps'])
    case_priority = ''
    case_apply_phase = ''
    case_state = ''
    case_type = ''
    row = [case_number, case_tree, case_title, case_summary, case_precontion, case_step, case_expected_result, case_priority, case_apply_phase, case_state, case_type]

def get_case_module(module_name):
    if module_name:
        module_name = module_name.replace('（', '(')
        module_name = module_name.replace('）', ')')
    else:
        module_name = '/'
    return module_name


def gen_case_step_and_expected_result(steps):
    case_step = ''
    case_expected_result = ''
    for step_dict in steps:
        case_step += step_dict['actions'].replace('\n', '').strip() + '\n'
        case_expected_result = step_dict['expectedresults'].replace('\n', '').strip() + '\n'
        # if step_dict.get('expectedresults', '') else ''

    return case_step, case_expected_result
