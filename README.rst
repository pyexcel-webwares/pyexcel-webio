================================================================================
pyexcel-webio - Let you focus on data, instead of file formats
================================================================================

.. image:: https://raw.githubusercontent.com/pyexcel/pyexcel.github.io/master/images/patreon.png
   :target: https://www.patreon.com/pyexcel

.. image:: https://api.travis-ci.org/pyexcel/pyexcel-webio.svg?branch=master
   :target: http://travis-ci.org/pyexcel/pyexcel-webio

.. image:: https://codecov.io/github/pyexcel/pyexcel-webio/coverage.png
   :target: https://codecov.io/github/pyexcel/pyexcel-webio

.. image:: https://img.shields.io/gitter/room/gitterHQ/gitter.svg
   :target: https://gitter.im/pyexcel/Lobby


Known constraints
==================

Fonts, colors and charts are not supported.

**pyexcel-webio** is a tiny interface library to unify the web extensions that uses `pyexcel <https://github.com/pyexcel/pyexcel>`__ . You may use it to write a web extension for your faviourite Python web framework.



Installation
================================================================================
You can install it via pip:

.. code-block:: bash

    $ pip install pyexcel-webio


or clone it and install it:

.. code-block:: bash

    $ git clone https://github.com/pyexcel/pyexcel-webio.git
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

Support the project
================================================================================

If your company has embedded pyexcel and its components into a revenue generating
product, please `support me on patreon <https://www.patreon.com/bePatron?u=5537627>`_ to
maintain the project and develop it further.

If you are an individual, you are welcome to support me too on patreon and for however long
you feel like to. As a patreon, you will receive
`early access to pyexcel related contents <https://www.patreon.com/pyexcel/posts>`_.

With your financial support, I will be able to invest
a little bit more time in coding, documentation and writing interesting posts.


Development guide
================================================================================

Development steps for code changes

#. git clone https://github.com/pyexcel/pyexcel-webio.git
#. cd pyexcel-webio

Upgrade your setup tools and pip. They are needed for development and testing only:

#. pip install --upgrade setuptools pip

Then install relevant development requirements:

#. pip install -r rnd_requirements.txt # if such a file exists
#. pip install -r requirements.txt
#. pip install -r tests/requirements.txt

Once you have finished your changes, please provide test case(s), relevant documentation
and update CHANGELOG.rst.

.. note::

    As to rnd_requirements.txt, usually, it is created when a dependent
	library is not released. Once the dependecy is installed
	(will be released), the future
	version of the dependency in the requirements.txt will be valid.


How to test your contribution
------------------------------

Although `nose` and `doctest` are both used in code testing, it is adviable that unit tests are put in tests. `doctest` is incorporated only to make sure the code examples in documentation remain valid across different development releases.

On Linux/Unix systems, please launch your tests like this::

    $ make

On Windows systems, please issue this command::

    > test.bat

How to update test environment and update documentation
---------------------------------------------------------

Additional steps are required:

#. pip install moban
#. git clone https://github.com/pyexcel/pyexcel-commons.git commons
#. make your changes in `.moban.d` directory, then issue command `moban`

What is pyexcel-commons
---------------------------------

Many information that are shared across pyexcel projects, such as: this developer guide, license info, etc. are stored in `pyexcel-commons` project.

What is .moban.d
---------------------------------

`.moban.d` stores the specific meta data for the library.

Acceptance criteria
-------------------

#. Has Test cases written
#. Has all code lines tested
#. Passes all Travis CI builds
#. Has fair amount of documentation if your change is complex
#. Agree on NEW BSD License for your contribution



License
================================================================================

New BSD License
