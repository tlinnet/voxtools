#!/usr/bin/env python

################################################################################
# Copyright (C) Voxmeter A/S - All Rights Reserved
#
# Voxmeter A/S
# Borgergade 6, 4.
# 1300 Copenhagen K
# Denmark
#
# Written by Troels Schwarz-Linnet <tsl@voxmeter.dk>, 2018
# 
# Unauthorized copying of this file, via any medium is strictly prohibited.
#
# Any use of this code is strictly unauthorized without the written consent
# by Voxmeter A/S. This code is proprietary of Voxmeter A/S.
# 
################################################################################

from distutils.core import setup
from codecs import open
from os import path
from voxtools import __version__

# Determine install position
package_name = 'voxtools'
import sys
from site import USER_SITE, USER_BASE
rela_path = USER_SITE.split(USER_BASE)[-1][1:]
install_to = path.join(rela_path, package_name)
print(rela_path)
print(install_to)

# get long description from README
here = path.abspath(path.dirname(__file__))
with open(path.join(here, 'README.rst'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name=package_name,
    version=__version__,
    description='A module for working with voxtools.',
    long_description=long_description,
    url='https://github.com/tlinnet/voxtools',
    download_url='https://github.com/tlinnet/voxtools/archive/%s.tar.gz'%__version__,
    author='Troels Schwarz-Linnet',
    author_email='tlinnet@gmail.com',
    license='Apache Software License',
    classifiers=[
        'Intended Audience :: Developers',
        'License :: OSI Approved :: Apache Software License',
        'Programming Language :: Python :: 3.6',
        'Topic :: Scientific/Engineering',
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: POSIX :: Linux'],
    requires=['numpy', 'openpyxl', 'PyQt5'],
    packages=['voxtools'],
    data_files=[(install_to, ['README.rst', 'LICENSE'])],
)