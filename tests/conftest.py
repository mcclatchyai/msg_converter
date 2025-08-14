import pytest
from pathlib import Path

def pytest_addoption(parser):
    parser.addoption(
        "--msg-file",
        action="store",
        default="tests/sample/sample_message1.msg",
        help="Path to the .msg file to test"
    )

@pytest.fixture
def msg_path(request):
    return Path(request.config.getoption("--msg-file"))