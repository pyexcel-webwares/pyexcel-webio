# -*- coding: utf-8 -*-
import os
import sys
from common import TestExtendedInput
from nose.tools import eq_


PY2 = sys.version_info[0] == 2


def test_issue_4():
    myinput = TestExtendedInput()
    fixture = os.path.join("tests", "fixtures", "issue4-broken.csv")

    expected = [
        [u'Last Name', u'First Name', u'Company', u'Email', u'Job Title'],
        [u'Test', u'Th\xefs', u'Cool Co',
         u'test.this@example.com', u'Founder']
    ]
    if PY2:
        with open(fixture, "rb") as f:
            array = myinput.get_array(field_name=('csv', f),
                                      encoding='latin1')
            eq_(array, expected)
    else:
        with open(fixture, "r", encoding='latin1') as f:
            array = myinput.get_array(field_name=('csv', f))
            eq_(array, expected)


def test_issue_4_passing_delimiter():
    myinput = TestExtendedInput()
    fixture = os.path.join("tests", "fixtures", "issue4-test.csv")

    expected = [[1, 2, 3], [4, 5, 6]]
    with open(fixture, "rb") as f:
        array = myinput.get_array(field_name=('csv', f),
                                  delimiter='\t')
        eq_(array, expected)
