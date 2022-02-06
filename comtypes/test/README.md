Running tests
-------------
From the projects root directory, run:

    python -m unittest discover -s ./comtypes/test -t comtypes\test

Or, from PROJECT_ROOT/comtypes/test:

    python -m unittest discover

TODO
----

- [ ] Look at every skipped test and see if it can be fixed and made runnable as a regular
  unit test.
- [ ] Remove the custom test runner stuff. See `comtypes/test/__init__.py`
  and `. /settup.py` for details.
