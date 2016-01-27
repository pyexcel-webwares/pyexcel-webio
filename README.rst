==============
pyexcel-webio
==============

.. image:: https://api.travis-ci.org/pyexcel/pyexcel-webio.png
    :target: http://travis-ci.org/pyexcel/pyexcel-webio

.. image:: https://codecov.io/github/pyexcel/pyexcel-webio/coverage.png
    :target: https://codecov.io/github/pyexcel/pyexcel-webio


**pyexcel-webio** is a tiny interface library to unify the web extensions that uses `pyexcel <https://github.com/pyexcel/pyexcel>`__ . You may use it to write a web extension for your faviourite Python web framework.


Installation
============

You can install it via pip::

    $ pip install pyexcel-webio


or clone it and install it::

    $ git clone http://github.com/pyexcel/pyexcel-webio.git
    $ cd pyexcel-webio
    $ python setup.py install


Known extensions
=======================

============== ============================
framework      plugin/middleware/extension
============== ============================
Flask          `Flask-Excel`_
Django         `django-excel`_
Pyramid        `pyramid-excel`_
============== ============================

.. _Flask-Excel: https://github.com/pyexcel/Flask-Excel
.. _django-excel: https://github.com/pyexcel/django-excel
.. _pyramid-excel: https://github.com/pyexcel/pyramid-excel

Usage
=========

This small section outlines the steps to adapt **pyexcel-webio** for your favourite web framework. For illustration purpose, I took **Flask** micro-framework as an example.

1. Inherit **ExcelInput** class and implement **load_single_sheet** and **load_book** methods depending on the parameters you will have. For example, **Flask.Request** put the incoming file in **Flask.Request.files** and the key is the field name in the html form::

    from flask import Flask, Request
    import pyexcel as pe
    from pyexcel.ext import webio

    class ExcelRequest(webio.ExcelInput, Request):
        def _get_file_tuple(self, field_name):
            filehandle = self.files[field_name]
            filename = filehandle.filename
            extension = filename.split(".")[1]
            return extension, filehandle
    
        def load_single_sheet(self, field_name=None, sheet_name=None,
                              **keywords):
            file_type, file_handle = self._get_file_tuple(field_name)
            return pe.get_sheet(file_type=file_type,
                                content=file_handle.read(),
                                sheet_name=sheet_name,
                                **keywords)
    
        def load_book(self, field_name=None, **keywords):
            file_type, file_handle = self._get_file_tuple(field_name)
            return pe.get_book(file_type=file_type,
                               content=file_handle.read(),
                               **keywords)

2. Plugin in a response method that has the following signature::

       def your_func(content, content_type=None, status=200):
           ....

   or a response class has the same signature::

       class YourClass:
           def __init__(self, content, content_type=None, status=200):
           ....

   For example, with **Flask**, it is just a few lines::

       from flask import Response


       webio.ExcelResponse = Response


3. Then make the proxy for **make_response** series by simply copying the following lines to your extension::

    from pyexcel.ext.webio import (
        make_response,
        make_response_from_array,
        make_response_from_dict,
        make_response_from_records,
        make_response_from_book_dict
    )

License
==========

New BSD License

Dependencies
============

* pyexcel >= 0.1.7
