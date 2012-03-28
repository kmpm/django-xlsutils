try:
    from setuptools import setup, find_packages
except ImportError:
    from distutils.core import setup, find_packages

import os
from xlsutils import VERSION


f = open(os.path.join(os.path.dirname(__file__), 'README'))
readme = f.read()
f.close()

setup(name = "django-xlsutils",
      author = "Peter Magnusson",
      url = "http://github.com/kmpm/django-xlsutils",
      version = ".".join(map(str, VERSION)),
      description='django-xlsutils is a resuable Django application ment to help out when dealing with excel files',
      long_description=readme,
      packages = find_packages(),
      package_data = {
            'xlsutils': [
                  'templates/xlsutils/*',
                  'static/*'
      },
      install_requires = [
            'xlrd',
            'xlwt'
      ]
      zip_safe = True,
)