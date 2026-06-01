import os

import pytest


# generate_excel reads ``template.xlsx`` via a relative path, so the tests must
# run with the working directory set to this ``functions`` directory regardless
# of where pytest was invoked from.
@pytest.fixture(autouse=True)
def _chdir_to_functions(monkeypatch):
    monkeypatch.chdir(os.path.dirname(__file__))
