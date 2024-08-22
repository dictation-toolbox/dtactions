"""common fixtures for the dtactions project
"""
import pytest
import natlink
from natlinkcore import natlinkstatus

@pytest.fixture(scope="module")
def nl_stat():
    status = natlinkstatus.NatlinkStatus()
    return status

@pytest.fixture(scope="module")
def nat_conn():
    yield natlink.natConnect()
    natlink.natDisconnect()
