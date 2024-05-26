"""
doc2docx.py

Author: hanrai <hanrai@gmail.com>

This module provides functionality to convert .doc files to .docx files using LibreOffice.

The module defines three functions:

- get_doc_files: This function takes a path to a directory and returns a list of .doc files in that directory.
- convert_to_docx: This function takes a path to a LibreOffice executable, a .doc file, an output directory, and a flag indicating whether to overwrite existing files. It converts the .doc file to .docx format using LibreOffice and saves the .docx file in the output directory.
- main: This function parses command line arguments and uses a ThreadPoolExecutor to convert multiple .doc files to .docx format in parallel.

This module can be run as a script from the command line. The command line interface requires three arguments: the source directory containing .doc files, the output directory where .docx files will be saved, and the path to the LibreOffice or soffice executable. It also supports optional arguments for increasing output verbosity, overwriting existing .docx files, and specifying a log file.

Example usage:

    python doc2docx.py -s <source_directory> -o <output_directory> -l <libreoffice_path> [-v] [-w]

License: MIT

MIT License

Copyright (c) 2024 hanrai

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

import argparse
import glob
import logging
import os
import subprocess
from concurrent.futures import ThreadPoolExecutor


def get_doc_files(file_scraper_path):
    """
    Get a list of .doc files in the specified directory.

    Args:
        file_scraper_path (str): The path to the directory to search for .doc files.

    Returns:
        list: A list of file paths to .doc files found in the directory.
    """
    return glob.glob(os.path.join(file_scraper_path, "*.doc"))


def convert_to_docx(libreoffice_path, doc_file, output_path, overwrite):
    """
    Converts a document file to DOCX format using LibreOffice.

    Args:
        libreoffice_path (str): The path to the LibreOffice executable.
        doc_file (str): The path to the input document file.
        output_path (str): The path to the output directory.
        overwrite (bool): Flag indicating whether to overwrite the output file if it already exists.

    Returns:
        None
    """
    output_file = os.path.join(
        output_path, os.path.splitext(os.path.basename(doc_file))[0] + ".docx"
    )
    if os.path.exists(output_file):
        if overwrite:
            logging.info("File %s already exists. Overwriting.", output_file)
        else:
            logging.info("File %s already exists. Skipping conversion.", output_file)
            return
    process = subprocess.Popen(
        [
            libreoffice_path,
            "--headless",
            "--convert-to",
            "docx",
            doc_file,
            "--outdir",
            output_path,
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    stdout, stderr = process.communicate()
    if stdout:
        logging.info("Output: %s", stdout.decode("utf-8"))
    if stderr:
        logging.error("Error: %s", stderr.decode("utf-8"))
        if stderr.strip():
            logging.error("Failed to convert %s to %s.", doc_file, output_file)
            return
    logging.info("Converted: %s to %s", doc_file, output_file)


def main():
    """
    Convert .doc files to .docx files.

    This function takes command line arguments to specify the source directory
    containing .doc files, the output directory where .docx files will be saved,
    and the path to the libreoffice or soffice executable. It also supports
    optional arguments for increasing output verbosity, overwriting existing
    .docx files, and specifying a log file.

    Usage:
        python doc2docx.py -s <source_directory> -o <output_directory> -l <libreoffice_path> [-v] [-w]

    Arguments:
        -s, --source: Required. Path to the directory containing .doc files.
        -o, --output: Required. Path to the directory where .docx files will be saved.
        -l, --libreoffice: Required. Path to the libreoffice or soffice executable.

    Optional Arguments:
        -v, --verbose: Increase output verbosity.
        -w, --overwrite: Overwrite existing .docx files.

    """
    parser = argparse.ArgumentParser(description="Convert .doc to .docx")
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="increase output verbosity"
    )
    parser.add_argument(
        "-l",
        "--libreoffice",
        required=True,
        help="path to the libreoffice or soffice executable",
    )
    parser.add_argument(
        "-s",
        "--source",
        required=True,
        help="path to the directory containing .doc files",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="path to the directory where .docx files will be saved",
    )
    parser.add_argument(
        "-w",
        "--overwrite",
        action="store_true",
        help="overwrite existing .docx files",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler("doc2docx.log"), logging.StreamHandler()],
    )

    executor = ThreadPoolExecutor()
    try:
        libreoffice_path = args.libreoffice
        doc_files = get_doc_files(args.source)
        if not doc_files:
            logging.error("No .doc files found in the source directory.")
            return
        for doc_file in doc_files:
            executor.submit(
                convert_to_docx,
                libreoffice_path,
                doc_file,
                args.output,
                args.overwrite,
            )
    except KeyboardInterrupt:
        logging.info("Interrupted by user. Exiting.")
        executor.shutdown(wait=False)
    except Exception as e:
        logging.error(str(e))
    else:
        executor.shutdown(wait=True)


if __name__ == "__main__":
    main()
