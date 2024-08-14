"""common fixtures for the dtactions project
"""
import pytest
from natlinkcore import natlinkstatus

@pytest.fixture(scope="module")
def nlstatus():
    status = natlinkstatus.NatlinkStatus()
    return status
