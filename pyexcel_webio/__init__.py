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
    """A generic interface for an excel file input

    The source could be from anywhere, memory or file system
    """
    def get_params(self, sheet_name=None, **keywords):
        """Abstract method

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv', 'tsv'
                           file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A sheet object
        """
        raise NotImplementedError("Please implement this function")

    def get_sheet(self, **keywords):
        """
        Get a :class:`Sheet` instance from the file

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv',
                           'tsv' file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A sheet object
        """
        params = self.get_params(**keywords)
        return pe.get_sheet(**params)

    def get_array(self, **keywords):
        """
        Get a list of lists from the file

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv',
                           'tsv' file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A list of lists
        """
        params = self.get_params(**keywords)
        return pe.get_array(**params)

    def get_dict(self, **keywords):
        """Get a dictionary from the file

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv',
                           'tsv' file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A dictionary
        """
        params = self.get_params(**keywords)
        if 'name_columns_by_row' not in params:
            params['name_columns_by_row'] = 0
        return pe.get_dict(**params)

    def get_records(self, **keywords):
        """Get a list of records from the file

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv',
                           'tsv' file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A list of records
        """
        params = self.get_params(**keywords)
        if 'name_columns_by_row' not in params:
            params['name_columns_by_row'] = 0
        return pe.get_records(**params)

    def save_to_database(self, session=None, table=None,
                         initializer=None, mapdict=None,
                         auto_commit=True,
                         **keywords):
        """
        Save data from a sheet to database

        :param session: a SQLAlchemy session
        :param table: a database table
        :param initializer: a custom table initialization function if you have one
        :param mapdict: the explicit table column names if your excel data do not
                        have the exact column names
        :param keywords: additional keywords to :meth:`pyexcel.Sheet.save_to_database`
        """
        params = self.get_params(**keywords)
        if 'name_columns_by_row' not in params:
            params['name_columns_by_row'] = 0
        if 'name_rows_by_column' not in params:
            params['name_rows_by_column'] = -1
        params['dest_session'] = session
        params['dest_table'] = table
        params['dest_initializer'] = initializer
        params['dest_mapdict'] = mapdict
        params['dest_auto_commit'] = auto_commit
        pe.save_as(**params)

    def get_book(self, **keywords):
        """Get a instance of :class:`Book` from the file

        :param keywords: additional key words
        :returns: A instance of :class:`Book`
        """
        params = self.get_params(**keywords)
        return pe.get_book(**params)

    def get_book_dict(self, **keywords):
        """Get a dictionary of two dimensional array from the file

        :param keywords: additional key words
        :returns: A dictionary of two dimensional arrays
        """
        params = self.get_params(**keywords)
        return pe.get_book_dict(**params)

    def save_book_to_database(self, session=None, tables=None,
                              initializers=None, mapdicts=None,
                              auto_commit=True, **keywords):
        """
        Save a book into database

        :param session: a SQLAlchemy session
        :param tables: a list of database tables
        :param initializers: a list of model initialization functions.
        :param mapdicts: a list of explicit table column names
                         if your excel data sheets do not have the exact column names
        :param keywords: additional keywords to :meth:`pyexcel.Book.save_to_database`

        """
        params = self.get_params(**keywords)
        params['dest_session'] = session
        params['dest_tables'] = tables
        params['dest_initializers'] = initializers
        params['dest_mapdicts'] = mapdicts
        params['dest_auto_commit'] = auto_commit
        pe.save_book_as(**params)


class ExcelInputInMultiDict(ExcelInput):
    """ A generic interface for an upload excel file appearing in a dictionary
    """
    def get_file_tuple(self, field_name):
        """
        Abstract method to get the file tuple

        It is expected to return file type and a file handle to the
        uploaded file
        """
        raise NotImplementedError("Please implement this function")

    def get_params(self, field_name=None, **keywords):
        """
        Load the single sheet from named form field
        """
        file_type, file_handle = self.get_file_tuple(field_name)
        if file_type is not None and file_handle is not None:
            keywords = {
                'file_type': file_type,
                'file_content': file_handle.read()
            }
            return keywords
        else:
            raise Exception("Invalid parameters")


def dummy_func(content, content_type=None, status=200, file_name=None):
    return None


ExcelResponse = dummy_func


def _make_response(io_stream, file_type,
                   status=200, file_name=None):
    io_stream.seek(0)
    if file_name:
        if not file_name.endswith(file_type):
            file_name = "%s.%s" % (file_name, file_type)
    return ExcelResponse(io_stream.read(),
                         content_type=FILE_TYPE_MIME_TABLE[file_type],
                         status=status, file_name=file_name)



def make_response(pyexcel_instance, file_type,
                  status=200, file_name=None, **keywords):
    """
    Make a http response from a pyexcel instance of
    :class:`~pyexcel.Sheet` or :class:`~pyexcel.Book`

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
    :returns: http response
    """
    file_stream = pyexcel_instance.save_to_memory(file_type, None, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_array(array, file_type,
                             status=200, file_name=None, **keywords):
    """
    Make a http response from an array

    :param array: a list of lists
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: http response
    """
    file_stream = pe.save_as(array=array, dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_dict(adict, file_type,
                            status=200, file_name=None, **keywords):
    """
    Make a http response from a dictionary of lists

    :param dict: a dictinary of lists
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: http response
    """
    file_stream = pe.save_as(adict=adict, dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_records(records, file_type,
                               status=200, file_name=None, **keywords):
    """
    Make a http response from a list of dictionaries

    :param records: a list of dictionaries
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: http response
    """
    file_stream = pe.save_as(records=records, dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_book_dict(adict,
                                 file_type, status=200, file_name=None,
                                 **keywords):
    """
    Make a http response from a dictionary of two dimensional
    arrays

    :param book_dict: a dictionary of two dimensional arrays
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: http response
    """
    file_stream = pe.save_book_as(bookdict=adict, dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_query_sets(query_sets, column_names,
                                  file_type, status=200, file_name=None,
                                  **keywords):
    """
    Make a http response from a dictionary of two dimensional
    arrays

    :param query_sets: a query set
    :param column_names: a nominated column names. It could not be N
                         one, otherwise no data is returned.
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: a http response
    """
    file_stream = pe.save_as(query_sets=query_sets, column_names=column_names,
                             dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_a_table(session, table,
                               file_type, status=200, file_name=None,
                               **keywords):
    """
    Make a http response from sqlalchmey table

    :param session: SQLAlchemy session
    :param table: a SQLAlchemy table
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: a http response
    """
    file_stream = pe.save_as(session=session, table=table,
                             dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)


def make_response_from_tables(session, tables,
                              file_type, status=200, file_name=None,
                              **keywords):
    """
    Make a http response from sqlalchmy tables

    :param session: SQLAlchemy session
    :param tables: SQLAlchemy tables
    :param file_type: same as :meth:`~pyexcel_webio.make_response`
    :param status: same as :meth:`~pyexcel_webio.make_response`
    :returns: a http response
    """
    file_stream = pe.save_book_as(session=session, tables=tables,
                                  dest_file_type=file_type, **keywords)
    return _make_response(file_stream, file_type, status, file_name)
