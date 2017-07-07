import os
import sys
from unittest import TestCase
import pyexcel as pe
import pyexcel_webio as webio
from common import TestInput, TestExtendedInput
from db import Session, Base, Signature, Signature2, engine
from nose.tools import raises, eq_
if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict

FILE_NAME = "response_test"
OUTPUT = "%s.xls" % FILE_NAME


def dumpy_response(content, content_type=None, status=200, file_name=None):
    """A dummy response"""
    with open(file_name, 'wb') as f:
        f.write(content)


webio.ExcelResponse = dumpy_response


class TestExceptions:
    @raises(NotImplementedError)
    def test_load_single_sheet(self):
        testinput = webio.ExcelInput()
        testinput.get_sheet(file_name="test")  # booom

    @raises(NotImplementedError)
    def test_load_book(self):
        testinput = webio.ExcelInput()
        testinput.get_book(file_name="test")  # booom

    @raises(NotImplementedError)
    def test_excel_input_get_file_tuple(self):
        testinput = webio.ExcelInputInMultiDict()
        testinput.get_file_tuple(field_name="test")  # booom

    @raises(pe.exceptions.UnknownParameters)
    def test_get_sheet(self):
        myinput = TestInput()
        myinput.get_sheet(unrelated="foo bar")

    @raises(pe.exceptions.UnknownParameters)
    def test_get_array(self):
        myinput = TestInput()
        myinput.get_array(unrelated="foo bar")

    @raises(pe.exceptions.UnknownParameters)
    def test_get_dict(self):
        myinput = TestInput()
        myinput.get_dict(unrelated="foo bar")

    @raises(pe.exceptions.UnknownParameters)
    def test_get_records(self):
        myinput = TestInput()
        myinput.get_records(unrelated="foo bar")

    @raises(pe.exceptions.UnknownParameters)
    def test_get_book(self):
        myinput = TestInput()
        myinput.get_book(unrelated="foo bar")

    @raises(pe.exceptions.UnknownParameters)
    def test_get_book_dict(self):
        myinput = TestInput()
        myinput.get_book_dict(unrelated="foo bar")

    def test_dummy_function(self):
        result = webio.dummy_func(None, None)
        assert result is None


# excel inputs
class TestExcelInput:
    def setUp(self):
        self.data = [
            ["X", "Y", "Z"],
            [1, 2, 3],
            [4, 5, 6]
        ]
        sheet = pe.Sheet(self.data)
        self.testfile = "testfile.xls"
        sheet.save_as(self.testfile)

    def test_get_sheet(self):
        myinput = TestInput()
        sheet = myinput.get_sheet(file_name=self.testfile)
        assert sheet.to_array() == self.data

    def test_get_array(self):
        myinput = TestInput()
        array = myinput.get_array(file_name=self.testfile)
        assert array == self.data

    def test_iget_array(self):
        myinput = TestInput()
        array = myinput.iget_array(file_name=self.testfile)
        assert list(array) == self.data
        myinput.free_resources()

    def test_get_dict(self):
        myinput = TestInput()
        result = myinput.get_dict(file_name=self.testfile)
        assert result == {
            "X": [1, 4],
            "Y": [2, 5],
            "Z": [3, 6]
        }

    def test_get_records(self):
        myinput = TestInput()
        result = myinput.get_records(file_name=self.testfile)
        assert result == [
            {"X": 1, "Y": 2, "Z": 3},
            {"X": 4, "Y": 5, "Z": 6}
        ]

    def test_iget_records(self):
        myinput = TestInput()
        result = myinput.iget_records(file_name=self.testfile)
        assert list(result) == [
            {"X": 1, "Y": 2, "Z": 3},
            {"X": 4, "Y": 5, "Z": 6}
        ]
        myinput.free_resources()

    def test_save_to_database(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        self.session = Session()
        myinput = TestInput()
        myinput.save_to_database(file_name=self.testfile,
                                 session=self.session,
                                 table=Signature)
        array = pe.get_array(session=self.session, table=Signature)
        assert array == self.data
        self.session.close()

    def tearDown(self):
        os.unlink(self.testfile)


class TestExcelInputInMultiDict:
    def setUp(self):
        self.data = [
            ["X", "Y", "Z"],
            [1, 2, 3],
            [4, 5, 6]
        ]
        sheet = pe.Sheet(self.data)
        self.testfile = "testfile.xls"
        sheet.save_as(self.testfile)

    def tearDown(self):
        os.unlink(self.testfile)

    def test_get_sheet(self):
        myinput = TestExtendedInput()
        f = open(self.testfile, 'rb')
        sheet = myinput.get_sheet(field_name=('xls', f))
        assert sheet.to_array() == self.data
        f.close()

    def test_a_consumed_file_handle(self):
        myinput = TestExtendedInput()
        f = open(self.testfile, 'rb')
        f.read()
        sheet = myinput.get_sheet(field_name=('xls', f))
        assert sheet.to_array() == self.data
        f.close()

    @raises(IOError)
    def test_get_sheet_in_exception(self):
        myinput = TestExtendedInput()
        empty_file = 'empty_file'
        with open(empty_file, 'w') as f:
            f.write('')
        f = open(empty_file, 'rb')
        myinput.get_sheet(field_name=('xls', f))
        os.unlink(empty_file)

    @raises(Exception)
    def test_wrong_file_tuple_returned(self):
        myinput = TestExtendedInput()
        myinput.get_sheet(field_name=('xls', None))


class TestExcelInputOnBook(TestCase):
    def setUp(self):
        self.data = [['X', 'Y', 'Z'], [1, 2, 3], [4, 5, 6]]
        self.data1 = [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]]
        mydict = OrderedDict()
        mydict.update({'signature': self.data})
        mydict.update({'signature2': self.data1})
        book = pe.Book(mydict)
        self.testfile = "testfile.xls"
        book.save_as(self.testfile)

    def test_get_book(self):
        myinput = TestInput()
        result = myinput.get_book(file_name=self.testfile)
        assert result["signature"].to_array() == self.data
        assert result["signature2"].to_array() == self.data1

    def test_get_book_dict(self):
        myinput = TestInput()
        result = myinput.get_book_dict(file_name=self.testfile)
        assert result["signature"] == self.data
        assert result["signature2"] == self.data1

    def test_save_to_database(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        self.session = Session()
        myinput = TestInput()
        myinput.save_book_to_database(file_name=self.testfile,
                                      session=self.session,
                                      tables=[Signature, Signature2])
        array = pe.get_array(session=self.session, table=Signature)
        self.assertEqual(array, self.data)
        array = pe.get_array(session=self.session, table=Signature2)
        assert array == self.data1
        self.session.close()

    def tearDown(self):
        os.unlink(self.testfile)


class TestExcelInput2OnBook:
    def setUp(self):
        self.data = [['X', 'Y', 'Z'], [1, 2, 3], [4, 5, 6]]
        self.data1 = [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]]
        mydict = OrderedDict()
        mydict.update({'sheet1': self.data})
        mydict.update({'sheet2': self.data1})
        book = pe.Book(mydict)
        self.testfile = "testfile.xls"
        book.save_as(self.testfile)

    def test_get_book(self):
        myinput = TestExtendedInput()
        f = open(self.testfile, 'rb')
        result = myinput.get_book(field_name=('xls', f))
        assert result["sheet1"].to_array() == self.data
        assert result["sheet2"].to_array() == self.data1
        f.close()


# responses
class TestResponse:
    def setUp(self):
        self.data = [
            ["X", "Y", "Z"],
            [1, 2, 3],
            [4, 5, 6]
        ]
        self.test_sheet_name = 'custom sheet name'

    def test_make_response_from_sheet(self):
        sheet = pe.Sheet(self.data)
        webio.make_response(sheet, "xls", file_name=FILE_NAME,
                            sheet_name=self.test_sheet_name)
        self.verify()

    def test_make_response_from_array(self):
        webio.make_response_from_array(
            self.data, "xls",
            file_name=FILE_NAME, sheet_name=self.test_sheet_name)
        self.verify()

    def test_make_response_from_records(self):
        records = [
            {"X": 1, "Y": 2, "Z": 3},
            {"X": 4, "Y": 5, "Z": 6}
        ]
        webio.make_response_from_records(
            records, "xls",
            file_name=FILE_NAME, sheet_name=self.test_sheet_name)
        self.verify()

    def test_make_response_from_dict(self):
        adict = {
            "X": [1, 4],
            "Y": [2, 5],
            "Z": [3, 6]
        }
        webio.make_response_from_dict(
            adict, "xls", file_name=FILE_NAME, sheet_name=self.test_sheet_name)
        self.verify()

    def test_make_response_from_table(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        row1 = Signature(X=1, Y=2, Z=3)
        row2 = Signature(X=4, Y=5, Z=6)
        session = Session()
        session.add(row1)
        session.add(row2)
        session.commit()
        webio.make_response_from_a_table(
            session, Signature, "xls",
            file_name=FILE_NAME, sheet_name=self.test_sheet_name)
        self.verify()
        session.close()

    def test_make_response_from_query_sets(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        row1 = Signature(X=1, Y=2, Z=3)
        row2 = Signature(X=4, Y=5, Z=6)
        session = Session()
        session.add(row1)
        session.add(row2)
        session.commit()
        query_sets = session.query(Signature).filter_by(X=1).all()
        column_names = ["X", "Y", "Z"]
        webio.make_response_from_query_sets(
            query_sets, column_names, "xls",
            file_name=FILE_NAME, sheet_name=self.test_sheet_name)
        sheet2 = pe.get_sheet(file_name=OUTPUT)
        assert sheet2.to_array() == [
            ["X", "Y", "Z"],
            [1, 2, 3]
        ]
        session.close()

    def verify(self):
        sheet2 = pe.get_sheet(file_name=OUTPUT)
        assert sheet2.to_array() == self.data
        eq_(sheet2.name, self.test_sheet_name)

    def tearDown(self):
        if os.path.exists(OUTPUT):
            os.unlink(OUTPUT)


class TestBookResponse:
    def setUp(self):
        self.content = OrderedDict()
        self.content.update({
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]]})
        self.content.update({
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]]})
        self.content.update({
            "Sheet3": [[u'X', u'Y', u'Z'], [1, 4, 7], [2, 5, 8], [3, 6, 9]]})

    def test_make_response_from_book(self):
        book = pe.get_book(bookdict=self.content)
        webio.make_response(book, "xls", file_name=FILE_NAME)
        self.verify()

    def test_make_response_from_book_dict(self):
        webio.make_response_from_book_dict(self.content, "xls",
                                           file_name=FILE_NAME)
        self.verify()

    def verify(self):
        book = pe.get_book(file_name=OUTPUT)
        assert book.to_dict() == self.content

    def tearDown(self):
        if os.path.exists(OUTPUT):
            os.unlink(OUTPUT)


class TestBookResponseFromDataBase:
    def test_make_response_from_tables(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        row1 = Signature(X=1, Y=2, Z=3)
        row2 = Signature(X=4, Y=5, Z=6)
        row3 = Signature2(A=1, B=2, C=3)
        row4 = Signature2(A=4, B=5, C=6)
        session = Session()
        session.add(row1)
        session.add(row2)
        session.add(row3)
        session.add(row4)
        session.commit()
        webio.make_response_from_tables(
            session, [Signature, Signature2],
            "xls", file_name=FILE_NAME)
        book = pe.get_book(file_name=OUTPUT)
        expected = OrderedDict()
        expected.update({
            'signature': [['X', 'Y', 'Z'], [1, 2, 3], [4, 5, 6]]})
        expected.update({
            'signature2': [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]]})
        assert book.to_dict() == expected
