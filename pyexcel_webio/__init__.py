import pyexcel as pe
import sys
PY2 = sys.version_info[0] == 2
if PY2:
    from StringIO import StringIO as BytesIO
else:
    from io import BytesIO

    
FILE_TYPE_MIME_TABLE = {
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
    "csvz": "application/zip",
    "tsvz": "application/zip",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xlsm": "application/vnd.ms-excel.sheet.macroenabled.12"
}


class ExcelInput(object):
    """A generic interface for an excel file to be converted

    The source could be from anywhere, memory or file system
    """
    def load_single_sheet(self, sheet_name=None, **keywords):
        """Abstract method
        
        :param form_field_name: the file field name in the html form for file upload
        :param sheet_name: For an excel book, there could be multiple sheets. If it is left
                         unspecified, the sheet at index 0 is loaded. For 'csv', 'tsv' file,
                         *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A sheet object
        """
        raise NotImplementedError("Please implement this function")

    def load_book(self, **keywords):
        """Abstract method
        
        :param form_field_name: the file field name in the html form for file upload
        :param keywords: additional key words
        :returns: A instance of :class:`Book`
        """
        raise NotImplementedError("Please implement this function")
        
    def get_sheet(self, sheet_name=None, **keywords):
        """
        Get a :class:`Sheet` instance from the file
        
        :param form_field_name: the file field name in the html form for file upload
        :param sheet_name: For an excel book, there could be multiple sheets. If it is left
                         unspecified, the sheet at index 0 is loaded. For 'csv', 'tsv' file,
                         *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A sheet object
        """
        return self.load_single_sheet(sheet_name=sheet_name, **keywords)
        
    def get_array(self, sheet_name=None, **keywords):
        """
        Get a list of lists from the file
        
        :param form_field_name: the file field name in the html form for file upload
        :param sheet_name: For an excel book, there could be multiple sheets. If it is left
                         unspecified, the sheet at index 0 is loaded. For 'csv', 'tsv' file,
                         *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A list of lists
        """
        sheet = self.get_sheet(sheet_name=sheet_name, **keywords)
        if sheet:
            return sheet.to_array()
        else:
            return None

    def get_dict(self, sheet_name=None, name_columns_by_row=0, **keywords):
        """Get a dictionary from the file
        
        :param form_field_name: the file field name in the html form for file upload
        :param sheet_name: For an excel book, there could be multiple sheets. If it is left
                         unspecified, the sheet at index 0 is loaded. For 'csv', 'tsv' file,
                         *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A dictionary
        """
        sheet = self.load_single_sheet(sheet_name=sheet_name, name_columns_by_row=name_columns_by_row, **keywords)
        if sheet:
            return sheet.to_dict()
        else:
            return None

    def get_records(self, sheet_name=None, name_columns_by_row=0, **keywords):
        """Get a list of records from the file
  
        :param form_field_name: the file field name in the html form for file upload
        :param sheet_name: For an excel book, there could be multiple sheets. If it is left
                         unspecified, the sheet at index 0 is loaded. For 'csv', 'tsv' file,
                         *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A list of records
        """
        sheet = self.load_single_sheet(sheet_name=sheet_name, name_columns_by_row=name_columns_by_row, **keywords)
        if sheet:
            return sheet.to_records()
        else:
            return None

    def save_to_database(self, session=None, table=None, sheet_name=None, name_columns_by_row=0, **keywords):
        sheet = self.load_single_sheet(sheet_name=sheet_name, name_columns_by_row=name_columns_by_row, **keywords)
        if sheet:
            sheet.save_to_database(session, table)

    def get_book(self, **keywords):
        """Get a instance of :class:`Book` from the file

        :param form_field_name: the file field name in the html form for file upload
        :param keywords: additional key words
        :returns: A instance of :class:`Book`
        """
        return self.load_book(**keywords)

    def get_book_dict(self, **keywords):
        """Get a dictionary of two dimensional array from the file

        :param form_field_name: the file field name in the html form for file upload
        :param keywords: additional key words
        :returns: A dictionary of two dimensional arrays
        """
        book = self.load_book(**keywords)
        if book:
            return book.to_dict()
        else:
            return None

    def save_book_to_database(self, session=None, tables=None, **keywords):
        book = self.load_book(**keywords)
        if book:
            book.save_to_database(session, tables)


def dummy_func(content, content_type=None, status=200):
    return None


ExcelResponse = dummy_func


def make_response(pyexcel_instance, file_type, status=200, **keywords):
    """Make a http response from a pyexcel instance of :class:`~pyexcel.Sheet` or :class:`~pyexcel.Book`

    :param pyexcel_instance: pyexcel.Sheet or pyexcel.Book
    :param file_type: one of the following strings:
                      
                      * 'csv'
                      * 'tsv'
                      * 'csvz'
                      * 'tsvz'
                      * 'xls'
                      * 'xlsx'
                      * 'xlsm'
                      * 'ods'
                        
    :param status: unless a different status is to be returned.
    """
    io = pe._get_io(file_type)
    pyexcel_instance.save_to_memory(file_type, io, **keywords)
    io.seek(0)
    return ExcelResponse(io.read(), content_type=FILE_TYPE_MIME_TABLE[file_type], status=status)


def make_response_from_array(array, file_type, status=200, **keywords):
    return make_response(pe.Sheet(array), file_type, status, **keywords)

    
def make_response_from_dict(adict, file_type, status=200, **keywords):
    return make_response(pe.load_from_dict(adict), file_type, status, **keywords)


def make_response_from_records(records, file_type, status=200, **keywords):
    return make_response(pe.load_from_records(records), file_type, status, **keywords)


def make_response_from_book_dict(adict, file_type, status=200, **keywords):
    return make_response(pe.Book(adict), file_type, status, **keywords)


def make_response_from_a_table(session, table, file_type, status=200, **keywords):
    sheet = pe.get_sheet(session=session, table=table, **keywords)
    return make_response(sheet, file_type, status, **keywords)


def make_response_from_tables(session, tables, file_type, status=200, **keywords):
    book = pe.get_book(session=session, tables=tables, **keywords)
    return make_response(book, file_type, status, **keywords)