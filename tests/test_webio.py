import os
import pyexcel as pe
from pyexcel.ext import webio
from pyexcel.ext import xls
from db import Session, Base, Signature, Signature2, engine
import sys
if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict
from nose.tools import raises
    

OUTPUT = "response_test.xls"
    

class TestInput(webio.ExcelInput):
    """This is sample implementation that read excel source from file"""
    def load_single_sheet(self, filename=None, sheet_name=None, **keywords):
        """Load a single sheet"""
        return pe.get_sheet(file_name=filename, **keywords)

    def load_book(self, filename=None, **keywords):
        """Load a book"""
        return pe.get_book(file_name=filename, **keywords)


def dumpy_response(content, content_type=None, status=200):
    """A dummy response"""
    f = open(OUTPUT, 'wb')
    f.write(content)
    f.close()


webio.ExcelResponse = dumpy_response


class TestExceptions:
    @raises(NotImplementedError)
    def test_load_single_sheet(self):
        testinput = webio.ExcelInput()
        testinput.get_sheet(filename="test") # booom

    @raises(NotImplementedError)
    def test_load_book(self):
        testinput = webio.ExcelInput()
        testinput.get_book(filename="test") # booom

    def test_get_sheet(self):
        myinput = TestInput()
        sheet = myinput.get_sheet(unrelated="foo bar")
        assert sheet == None

    def test_get_array(self):
        myinput = TestInput()
        array = myinput.get_array(unrelated="foo bar")
        assert array == None

    def test_get_dict(self):
        myinput = TestInput()
        result = myinput.get_dict(unrelated="foo bar")
        assert result == None

    def test_get_records(self):
        myinput = TestInput()
        result = myinput.get_records(unrelated="foo bar")
        assert result == None

    def test_get_book(self):
        myinput = TestInput()
        result = myinput.get_book(unrelated="foo bar")
        assert result == None

    def test_get_book_dict(self):
        myinput = TestInput()
        result = myinput.get_book_dict(unrelated="foo bar")
        assert result == None

    def test_dummy_function(self):
        result = webio.dummy_func(None, None)
        assert result == None


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
        sheet = myinput.get_sheet(filename=self.testfile)
        assert sheet.to_array() == self.data

    def test_get_array(self):
        myinput = TestInput()
        array = myinput.get_array(filename=self.testfile)
        assert array == self.data

    def test_get_dict(self):
        myinput = TestInput()
        result = myinput.get_dict(filename=self.testfile)
        assert result == {
            "X": [1, 4],
            "Y": [2, 5],
            "Z": [3, 6]
        }

    def test_get_records(self):
        myinput = TestInput()
        result = myinput.get_records(filename=self.testfile)
        assert result == [
            {"X": 1, "Y": 2, "Z": 3},
            {"X": 4, "Y": 5, "Z": 6}
        ]

    def test_save_to_database(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        self.session = Session()
        myinput = TestInput()
        myinput.save_to_database(filename=self.testfile, session=self.session, table=Signature,)
        array = pe.get_array(session=self.session, table=Signature)
        assert array == self.data
        self.session.close()

    def tearDown(self):
        os.unlink(self.testfile)


class TestExcelInputOnBook:
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
        myinput = TestInput()
        result = myinput.get_book(filename=self.testfile)
        assert result["sheet1"].to_array() == self.data
        assert result["sheet2"].to_array() == self.data1

    def test_get_book_dict(self):
        myinput = TestInput()
        result = myinput.get_book_dict(filename=self.testfile)
        assert result["sheet1"] == self.data
        assert result["sheet2"] == self.data1

    def test_save_to_database(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        self.session = Session()
        myinput = TestInput()
        myinput.save_book_to_database(filename=self.testfile, session=self.session, tables=[Signature, Signature2])
        array = pe.get_array(session=self.session, table=Signature)
        assert array == self.data
        array = pe.get_array(session=self.session, table=Signature2)
        assert array == self.data1
        self.session.close()

    def tearDown(self):
        os.unlink(self.testfile)

## responses

class TestResponse:
    def setUp(self):
        self.data = [
            ["X", "Y", "Z"],
            [1, 2, 3],
            [4, 5, 6]
        ]

    def test_make_response_from_sheet(self):
        sheet = pe.Sheet(self.data)
        webio.make_response(sheet, "xls")
        self.verify()

    def test_make_response_from_array(self):
        webio.make_response_from_array(self.data, "xls")
        self.verify()

    def test_make_response_from_records(self):
        records = [
            {"X": 1, "Y": 2, "Z": 3},
            {"X": 4, "Y": 5, "Z": 6}
        ]
        webio.make_response_from_records(records, "xls")
        self.verify()

    def test_make_response_from_dict(self):
        adict = {
            "X": [1, 4],
            "Y": [2, 5],
            "Z": [3, 6]
        }
        webio.make_response_from_dict(adict, "xls")
        self.verify()

    def test_make_response_from_table(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        row1 = Signature(X=1,Y=2, Z=3)
        row2 = Signature(X=4, Y=5, Z=6)
        session = Session()
        session.add(row1)
        session.add(row2)
        session.commit()
        webio.make_response_from_a_table(session, Signature, "xls")
        self.verify()
        session.close()

    def verify(self):
        sheet2 = pe.load(OUTPUT)
        assert sheet2.to_array() == self.data

    def tearDown(self):
        if os.path.exists(OUTPUT):
            os.unlink(OUTPUT)


class TestBookResponse:
    def setUp(self):
        self.content = OrderedDict()
        self.content.update({"Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]]})
        self.content.update({"Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]]})
        self.content.update({"Sheet3": [[u'X', u'Y', u'Z'], [1, 4, 7], [2, 5, 8], [3, 6, 9]]})

    def test_make_response_from_book(self):
        book = pe.get_book(bookdict=self.content)
        webio.make_response(book, "xls")
        self.verify()

    def test_make_response_from_book_dict(self):
        webio.make_response_from_book_dict(self.content, "xls")
        self.verify()

    def verify(self):
        book = pe.load_book(OUTPUT)
        assert book.to_dict() == self.content

    def tearDown(self):
        if os.path.exists(OUTPUT):
            os.unlink(OUTPUT)


class TestBookResponseFromDataBase:
    def test_make_response_from_tables(self):
        Base.metadata.drop_all(engine)
        Base.metadata.create_all(engine)
        row1 = Signature(X=1,Y=2, Z=3)
        row2 = Signature(X=4, Y=5, Z=6)
        row3 = Signature2(A=1, B=2, C=3)
        row4 = Signature2(A=4, B=5, C=6)
        session =Session()
        session.add(row1)
        session.add(row2)
        session.add(row3)
        session.add(row4)
        session.commit()
        webio.make_response_from_tables(session, [Signature, Signature2], "xls")
        book = pe.load_book(OUTPUT)
        expected = OrderedDict()
        expected.update({'signature': [['X', 'Y', 'Z'], [1, 2, 3], [4, 5, 6]]})
        expected.update({'signature2': [['A', 'B', 'C'], [1, 2, 3], [4, 5, 6]]})
        assert book.to_dict() == expected
