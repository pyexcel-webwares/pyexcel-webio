# -*- coding: utf-8 -*-
import os
from common import TestExtendedInput
from nose.tools import eq_


def test_issue_4():
    myinput = TestExtendedInput()
    fixture = os.path.join("tests", "fixtures", "issue4-broken.csv")
    with open(fixture, "r") as f:
        array = myinput.get_array(field_name=('csv', f), encoding='latin1')
        expected = [
            [u'Last Name', u'First Name', u'Company', u'Email', u'Job Title'],
            [u'Test', u'Th\xefs', u'Cool Co',
             u'test.this@example.com', u'Founder']
        ]
        eq_(array, expected)
