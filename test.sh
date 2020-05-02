#/bin/bash
pip freeze
nosetests --with-coverage --cover-package pyexcel_webio --cover-package tests tests --with-doctest --doctest-extension=.rst README.rst docs/source pyexcel_webio
