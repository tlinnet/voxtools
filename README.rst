========
voxtools 
========

How to run the GUI

.. code-block:: bash

    python -m voxtools.gui.excel


Important links
---------------

* Source code: https://github.com/tlinnet/voxtools
* PyPI package: https://pypi.python.org/pypi/voxtools


How to install?
---------------
If using conda, these steps will install an environment

.. code-block:: bash

   # With pip from https://pypi.python.org/pypi/voxtools
   pip install voxtools

Developer install for local conda environment:

.. code-block:: bash

   # Create environment
   conda env create -f environment.yml
   
   # Activate environment
   conda env list
   source activate voxtools

Or manual install in root environment:

.. code-block:: bash

   # Manually install package
   python setup.py install --force
   
   #  Manually uninstall
   python setup.py install --record files.txt
   PACK=`dirname $(head -n 1 files.txt)`
   rm -rf $PACK
   #cat files.txt | xargs rm -rf

Developer
---------

Set PYTHONPATH for local development
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

    # In bash, go to developer directory.
    export PYTHONPATH=$PWD

    # Try python from another directory
    cd $HOME
    python -c "import sys;print(sys.path)"
    python -c "import voxtools;print(voxtools.__file__)"

    # Try with module, and run the GUI
    # It needs to have: if __name__ == '__main__':
    python -m voxtools.gui.excel

| For conda environments
| https://conda.io/docs/user-guide/tasks/manage-environments.html#saving-environment-variables

.. code-block:: bash

    # In bash, go to developer directory.
    # Make new directories
    mkdir -p $HOME/anaconda/envs/voxtools/etc/conda/activate.d
    mkdir -p $HOME/anaconda/envs/voxtools/etc/conda/deactivate.d

    # For activate
    echo '#!/bin/bash' > $HOME/anaconda/envs/voxtools/etc/conda/activate.d/env_vars.sh
    echo "export PYTHONPATH='$PWD'" >> $HOME/anaconda/envs/voxtools/etc/conda/activate.d/env_vars.sh
    cat $HOME/anaconda/envs/voxtools/etc/conda/activate.d/env_vars.sh

    # For deactivate
    echo '#!/bin/bash' > $HOME/anaconda/envs/voxtools/etc/conda/deactivate.d/env_vars.sh
    echo "unset PYTHONPATH" >> $HOME/anaconda/envs/voxtools/etc/conda/deactivate.d/env_vars.sh
    cat $HOME/anaconda/envs/voxtools/etc/conda/deactivate.d/env_vars.sh

    # Then start an new terminal, and test
    source activate voxtools
    python -c "import sys;print(sys.path)"
    python -m voxtools.gui.excel


Run test_suite
^^^^^^^^^^^^^^

Run single tests

.. code-block:: bash

    # Get options
    python -m voxtools.test_suite.excel_test -h

    # Run 1 file with test
    python -m voxtools.test_suite.excel_test -b
    python -m voxtools.test_suite.excel_test -b -v

    # Run 1 class from 1 file
    python -m voxtools.test_suite.excel_test Test_excel -b -v
    # Run 1 test, from 1 class, from 1 file
    python -m voxtools.test_suite.excel_test Test_excel.test_copy_excel -b -v

    # Another example
    python -m voxtools.test_suite.wb04_test Test_wb04.test_make_uniq_key

    # With textblob, Run 1 file with test
    python -m voxtools.test_suite.textblob_test
    python -m voxtools.test_suite.kodning01_test

    # With textblob, Run 1 test, from 1 class, from 1 file
    python -m voxtools.test_suite.textblob_test Test_excel.test_copy_excel

Run all tests

.. code-block:: bash

    # From developer directory
    python -m unittest discover voxtools.test_suite -p "*_test.py" -b -v

Update and upload new version
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

* Guide for upload: http://peterdowns.com/posts/first-time-with-pypi.html
* Updated info: https://packaging.python.org/guides/migrating-to-pypi-org/#uploading
* PyPI test account: http://testpypi.python.org/pypi?%3Aaction=register_form 
* PyPI Live account: http://pypi.python.org/pypi?%3Aaction=register_form

.. code-block:: bash

   # Modify version in: voxtools/__init__.py
   
   # Create tag
   VERS=`python -c "from voxtools import __version__; print(__version__)"`
   # Adds a tag so that we can put this on PyPI
   git tag $VERS -m ""
   git push --tags origin master
   
   # Upload your package to PyPI Test
   python setup.py sdist upload -r pypitest
   open https://testpypi.python.org/pypi/voxtools
   
   # Upload to PyPI Live
   # Once you've successfully uploaded to PyPI Test, perform the same steps but point to the live PyPI server instead.
   python setup.py sdist upload -r pypi
   open https://pypi.python.org/pypi/voxtools