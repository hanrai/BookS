"""
A script to convert .docx files to .md (Markdown) files using Pandoc.

This script recursively searches for .docx files in a specified source directory and converts them to .md format using Pandoc.
The converted .md files are saved in a specified output directory.

The script supports the following command-line arguments:

    -s, --source: The path to the directory containing .docx files. This argument is required.
    -o, --output: The path to the directory where .md files will be saved. This argument is required.
    -p, --pandoc: The path to the Pandoc executable. By default, the script attempts to locate Pandoc using the system's PATH.
    -v, --verbose: If specified, the script prints additional information during the conversion process.
    -w, --overwrite: If specified, the script overwrites existing .md files in the output directory.

The script uses Python's built-in logging module to log information, warnings, and errors. If the --verbose argument is specified,
the script logs additional debug information. The log messages are written to a file named 'docx2md.log' and also printed to the console.

This script requires Python 3.6 or later, and Pandoc.

Author: hanrai (hanrai@gmail.com)

License:
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
import shutil
import subprocess


def get_docx_files(source_directory):
    """
    Retrieves a list of .docx files from the specified source directory and its subdirectories.

    Args:
        source_directory (str): The path to the source directory.

    Returns:
        list: A list of .docx file paths.

    """
    return glob.glob(os.path.join(source_directory, "**", "*.docx"), recursive=True)


def convert_to_md(pandoc_path, docx_file, output_directory, overwrite, verbose):
    """
    Converts a DOCX file to Markdown format using Pandoc.

    Args:
        pandoc_path (str): The path to the Pandoc executable.
        docx_file (str): The path to the input DOCX file.
        output_directory (str): The directory where the output Markdown file will be saved.
        overwrite (bool): If True, overwrite the output file if it already exists. If False, skip conversion.
        verbose (bool): If True, print additional information during the conversion process.

    Returns:
        None

    Raises:
        Exception: If an error occurs during the conversion process.

    """
    output_file = os.path.join(
        output_directory, os.path.splitext(os.path.basename(docx_file))[0] + ".md"
    )
    if not overwrite and os.path.exists(output_file):
        logging.info("Skipping %s because %s already exists", docx_file, output_file)
        return
    logging.debug("Converting %s to %s", docx_file, output_file)
    try:
        process = subprocess.Popen(
            [pandoc_path, "-s", docx_file, "-t", "markdown", "-o", output_file],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        stdout, stderr = process.communicate()
    except Exception as e:
        logging.error("Failed to convert %s to %s: %s", docx_file, output_file, e)
        return

    if process.returncode != 0:
        logging.error(
            "Failed to convert %s to %s: %s",
            docx_file,
            output_file,
            stderr.decode("utf-8"),
        )
        return

    if verbose:
        logging.debug("Pandoc output for %s: %s", docx_file, stdout.decode("utf-8"))

    logging.info("Finished converting %s to %s", docx_file, output_file)


def main():
    """
    Convert .docx files to .md files.

    This function takes command-line arguments to specify the source directory
    containing .docx files, the output directory where .md files will be saved,
    the path to the pandoc executable, and other optional flags.

    Args:
        -s, --source (str): Path to the directory containing .docx files.
        -o, --output (str): Path to the directory where .md files will be saved.
        -p, --pandoc (str, optional): Path to the pandoc executable. Defaults to the
            value returned by `shutil.which("pandoc")`.
        -v, --verbose (bool, optional): Increase output verbosity. Defaults to False.
        -w, --overwrite (bool, optional): Overwrite existing .md files. Defaults to False.

    Returns:
        None
    """
    parser = argparse.ArgumentParser(description="Convert .docx files to .md files")
    parser.add_argument(
        "-s",
        "--source",
        required=True,
        help="Path to the directory containing .docx files.",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Path to the directory where .md files will be saved.",
    )
    parser.add_argument(
        "-p",
        "--pandoc",
        default=shutil.which("pandoc"),
        help="Path to the pandoc executable.",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true", help="Increase output verbosity."
    )
    parser.add_argument(
        "-w", "--overwrite", action="store_true", help="Overwrite existing .md files."
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler("docx2md.log"), logging.StreamHandler()],
    )

    if not args.pandoc:
        logging.error("Pandoc executable not found.")
        return

    for docx_file in get_docx_files(args.source):
        convert_to_md(args.pandoc, docx_file, args.output, args.overwrite, args.verbose)


if __name__ == "__main__":
    main()
