.. pyexcel-webio documentation master file, created by
   sphinx-quickstart on Mon Aug 24 22:32:39 2015.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

.. include:: ../../README.rst

API Reference
=================

.. automodule:: pyexcel_webio

Excel file upload
----------------------

Here are the api for processing excel file upload

.. autoclass:: pyexcel_webio.ExcelInput
   :members:

.. autoclass:: pyexcel_webio.ExcelInputInMultiDict
   :members:

Excel file download
------------------------

Here are the api for converted different data structure into a excel file
download.

.. autofunction:: pyexcel_webio.make_response

.. autofunction:: pyexcel_webio.make_response_from_array
				  
.. autofunction:: pyexcel_webio.make_response_from_dict
				  
.. autofunction:: pyexcel_webio.make_response_from_records
				  
.. autofunction:: pyexcel_webio.make_response_from_book_dict
				  
.. autofunction:: pyexcel_webio.make_response_from_query_sets
				  
.. autofunction:: pyexcel_webio.make_response_from_a_table
				  
.. autofunction:: pyexcel_webio.make_response_from_tables

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

