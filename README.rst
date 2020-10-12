================================================================================
pyexcel-webio - Let you focus on data, instead of file formats
================================================================================

.. image:: https://raw.githubusercontent.com/pyexcel/pyexcel.github.io/master/images/patreon.png
   :target: https://www.patreon.com/chfw

.. image:: https://cdn.rawgit.com/sindresorhus/awesome/d7305f38d29fed78fa85652e3a63e154dd8e8829/media/badge.svg
   :target: https://awesome-python.com/#specific-formats-processing

.. image:: https://travis-ci.org/pyexcel-webwares/pyexcel-webio.svg?branch=master
   :target: http://travis-ci.org/pyexcel-webwares/pyexcel-webio

.. image:: https://codecov.io/gh/pyexcel-webwares/pyexcel-webio/branch/master/graph/badge.svg
   :target: https://codecov.io/gh/pyexcel-webwares/pyexcel-webio

.. image:: https://badge.fury.io/py/pyexcel-webio.svg
   :target: https://pypi.org/project/pyexcel-webio


.. image:: https://pepy.tech/badge/pyexcel-webio/month
   :target: https://pepy.tech/project/pyexcel-webio/month


.. image:: https://img.shields.io/gitter/room/gitterHQ/gitter.svg
   :target: https://gitter.im/pyexcel/Lobby

.. image:: https://img.shields.io/static/v1?label=continuous%20templating&message=%E6%A8%A1%E7%89%88%E6%9B%B4%E6%96%B0&color=blue&style=flat-square
    :target: https://moban.readthedocs.io/en/latest/#at-scale-continous-templating-for-open-source-projects

.. image:: https://img.shields.io/static/v1?label=coding%20style&message=black&color=black&style=flat-square
    :target: https://github.com/psf/black

Support the project
================================================================================

If your company has embedded pyexcel and its components into a revenue generating
product, please support me on github, `patreon <https://www.patreon.com/bePatron?u=5537627>`_
or `bounty source <https://salt.bountysource.com/teams/chfw-pyexcel>`_ to maintain
the project and develop it further.

If you are an individual, you are welcome to support me too and for however long
you feel like. As my backer, you will receive
`early access to pyexcel related contents <https://www.patreon.com/pyexcel/posts>`_.

And your issues will get prioritized if you would like to become my patreon as `pyexcel pro user`.

With your financial support, I will be able to invest
a little bit more time in coding, documentation and writing interesting posts.


Known constraints
==================

Fonts, colors and charts are not supported.

Introduction
================================================================================
**pyexcel-webio** is a tiny interface library to unify the web extensions that uses `pyexcel <https://github.com/pyexcel/pyexcel>`__ . You may use it to write a web extension for your faviourite Python web framework.



Installation
================================================================================

You can install pyexcel-webio via pip:

.. code-block:: bash

    $ pip install pyexcel-webio


or clone it and install it:

.. code-block:: bash

    $ git clone https://github.com/pyexcel-webwares/pyexcel-webio.git
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


Before you commit
------------------------------

Please run::

    $ make format

so as to beautify your code otherwise travis-ci may fail your unit test.




License
================================================================================

New BSD License
