# tests/test_converter.py
from pathlib import Path
from converters.msg_to_eml import convert_to_eml
from converters.msg_to_pdf import convert_to_pdf
from converters.msg_to_mbox import convert_to_mbox

def test_msg_to_eml(msg_path):
    #sample_msg = Path("tests/sample/sample_message1.msg")
    output_dir = Path("tests/output")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "test_output.eml"
    convert_to_eml(msg_path, output_path)
    assert output_path.exists()

def test_msg_to_pdf(msg_path):
    #sample_msg = Path("tests/sample/sample_message1.msg")
    output_dir = Path("tests/output")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "test_output.pdf"
    convert_to_pdf(msg_path, output_path)
    assert output_path.exists()

def test_msg_to_mbox(msg_path):
    #sample_msg = Path("tests/sample/sample_message1.msg")
    output_dir = Path("tests/output")
    output_dir.mkdir(parents=True, exist_ok=True)
    eml_path = output_dir / "test_output.eml"
    convert_to_eml(msg_path, eml_path)

    mbox_path = output_dir / "test_output.mbox"
    convert_to_mbox([eml_path], mbox_path)
    assert mbox_path.exists()