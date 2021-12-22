#!/usr/bin/env python
# _*_ coding:utf-8 _*_

import os
import json
import logging
import xmindparser
from generalparser import xmind_to_testsuites

xmindparser.config = {
    'showTopicId': False,  # 是否展示主题ID
    'hideEmptyValue': False  # 是否隐藏空值
}


def get_absolute_path(path):
    """
        Return the absolute path of a file

        If path contains a start point (eg Unix '/') then use the specified start point
        instead of the current working directory. The starting point of the file path is
        allowed to begin with a tilde "~", which will be replaced with the user's home directory.
    """
    fp, fn = os.path.split(path)
    if not fp:
        fp = os.getcwd()
    fp = os.path.abspath(os.path.expanduser(fp))
    return os.path.join(fp, fn)


def get_xmind_testsuites(xmind_file):
    """Load the XMind file and parse to `xmind2testcase.metadata.TestSuite` list"""
    xmind_file = get_absolute_path(xmind_file)
    xmind_content_dict = xmindparser.xmind_to_dict(xmind_file)
    logging.debug("loading XMind file(%s) dict data: %s", xmind_file, xmind_content_dict)

    if xmind_content_dict:
        testsuites = xmind_to_testsuites(xmind_content_dict)
        return testsuites
    else:
        logging.error('Invalid XMind file(%s): it is empty!', xmind_file)
        return []


def get_xmind_testcase_list(xmind_file):
    """Load the XMind file and get all testcase in it

    :param xmind_file: the target XMind file
    :return: a list of testcase data
    """
    xmind_file = get_absolute_path(xmind_file)
    logging.info('Start converting XMind file(%s) to testcases dict data...', xmind_file)
    testsuites = get_xmind_testsuites(xmind_file)

    testcases = []

    for testsuite in testsuites:
        product = testsuite.name
        for suite in testsuite.sub_suites:
            for case in suite.testcase_list:
                case_data = case.to_dict()
                case_data['product'] = product
                case_data['suite'] = suite.name
                testcases.append(case_data)

    logging.info('Convert XMind file(%s) to testcases dict data successfully!', xmind_file)
    return testcases
