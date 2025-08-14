import argparse
import os
import sys
import subprocess
from pathlib import Path
from tqdm import tqdm
from converters.msg_to_eml import convert_to_eml
from converters.msg_to_pdf import convert_to_pdf
from converters.msg_to_mbox import convert_to_mbox

def process_folder(input_folder, output_folder, output_format):
    os.makedirs(output_folder, exist_ok=True)
    msg_files = list(Path(input_folder).rglob("*.msg"))

    if output_format == 'mbox':
        eml_temp_dir = Path(output_folder) / "_temp_eml"
        eml_temp_dir.mkdir(exist_ok=True)
        eml_paths = []

        for msg_file in msg_files:
            eml_path = eml_temp_dir / (msg_file.stem + ".eml")
            convert_to_eml(msg_file, eml_path)
            eml_paths.append(eml_path)

        mbox_path = Path(output_folder) / "output.mbox"
        convert_to_mbox(eml_paths, mbox_path)

    else:
        for msg_file in tqdm(msg_files, desc="Processing MSG files"):
            output_file = Path(output_folder) / (msg_file.stem + f".{output_format}")
            if output_format == 'pdf':
                try:
                    subprocess.run(
                        [sys.executable, "-m", "converters.msg_to_pdf", str(msg_file), str(output_file)],
                        check=True
                    )
                except subprocess.CalledProcessError:
                    print(f"Failed to convert: {msg_file}")
            elif output_format == 'eml':
                convert_to_eml(msg_file, output_file)


def main():
    parser = argparse.ArgumentParser(description="Convert MSG files to PDF, EML, or MBOX")
    parser.add_argument('--input', required=True, help='Input folder containing .msg files')
    parser.add_argument('--output', required=True, help='Output folder for converted files')
    parser.add_argument('--format', required=True, choices=['pdf', 'eml', 'mbox'], help='Output format')
    args = parser.parse_args()

    process_folder(args.input, args.output, args.format)


if __name__ == "__main__":
    main()