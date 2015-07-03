"""
    pyexcel.ext.webio
    ~~~~~~~~~~~~~~~~~~~

    A generic request and response interface for pyexcel web extensions

    :copyright: (c) 2015 by Onni Software Ltd.
    :license: New BSD License
"""
import pyexcel as pe


FILE_TYPE_MIME_TABLE = {
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
    "csvz": "application/zip",
    "tsvz": "application/zip",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "xls": "application/vnd.ms-excel",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xlsm": "application/vnd.ms-excel.sheet.macroenabled.12",
    "json": "application/json",
    "plain": "text/plain",
    "simple": "text/plain",
    "grid": "text/plain",
    "pipe": "text/plain",
    "orgtbl": "text/plain",
    "rst": "text/plain",
    "mediawiki": "text/plain",
    "latex": "application/x-latex",
    "latex_booktabs": "application/x-latex"
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

    def save_to_database(
            self,
            session=None, table=None, initializer=None, mapdict=None,
            auto_commit=True,
            sheet_name=None, name_columns_by_row=0, name_rows_by_column=-1,
            field_name=None, **keywords):
        sheet = self.load_single_sheet(
            field_name=field_name,
            sheet_name=sheet_name,
            name_columns_by_row=name_columns_by_row,
            name_rows_by_column=name_rows_by_column,
            **keywords)
        if sheet:
            sheet.save_to_database(session,
                                   table,
                                   initializer=initializer,
                                   mapdict=mapdict,
                                   auto_commit=auto_commit)

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

    def save_book_to_database(self,
                              session=None, tables=None,
                              initializers=None, mapdicts=None, auto_commit=True,
                              **keywords):
        book = self.load_book(**keywords)
        if book:
            book.save_to_database(session,
                                  tables,
                                  initializers=initializers,
                                  mapdicts=mapdicts,
                                  auto_commit=auto_commit)



class ExcelInputInMultiDict(ExcelInput):
    """ A generic interface for an upload excel file appearing in a dictionary
    """
    def get_file_tuple(self, field_name):
        raise NotImplementedError("Please implement this function")

    def load_single_sheet(self, field_name=None, sheet_name=None, **keywords):
        file_type, file_handle = self.get_file_tuple(field_name)
        if file_type is not None and file_handle is not None:
            return pe.get_sheet(file_type=file_type,
                                file_content=file_handle.read(),
                                sheet_name=sheet_name,
                                **keywords)
        else:
            return None

    def load_book(self, field_name=None, **keywords):
        file_type, file_handle = self.get_file_tuple(field_name)
        if file_type is not None and file_handle is not None:        
            return pe.get_book(file_type=file_type,
                               file_content=file_handle.read(),
                               **keywords)
        else:
            return None


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
    io = pe.get_io(file_type)
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


def make_response_from_query_sets(query_sets, column_names, file_type, status=200, **keywords):
    sheet = pe.get_sheet(query_sets=query_sets, column_names=column_names)
    return make_response(sheet, file_type, status, **keywords)


def make_response_from_a_table(session, table, file_type, status=200, **keywords):
    sheet = pe.get_sheet(session=session, table=table, **keywords)
    return make_response(sheet, file_type, status, **keywords)


def make_response_from_tables(session, tables, file_type, status=200, **keywords):
    book = pe.get_book(session=session, tables=tables, **keywords)
    return make_response(book, file_type, status, **keywords)
