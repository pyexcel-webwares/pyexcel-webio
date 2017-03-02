import pyexcel_webio as webio


class TestInput(webio.ExcelInput):
    """This is sample implementation that read excel source from file"""
    def get_params(self, **keywords):
        """Load a single sheet"""
        return keywords


class TestExtendedInput(webio.ExcelInputInMultiDict):
    """This is sample implementation that read excel source from file"""
    def get_file_tuple(self, field_name):
        return field_name
