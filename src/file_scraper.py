"""
file_scraper.py

This module provides functionality for scraping files from a given URL. It includes functions to validate the URL, file extensions, and folder path, and to download files from the URL. The files to be downloaded can be filtered based on their extensions.

Functions:
- valid_url(string): Validates if a given string is a valid URL.
- valid_extension(string): Checks if a given string is a valid file extension.
- valid_folder(string, create_folder): Checks if a given folder path is valid.
- download_files(url, extensions, project_folder, create_folder, verbose, overwrite, timeout): Downloads all files with the specified extensions from the given URL and saves them to the project folder.
- download_file(url, filename, verbose, overwrite, timeout): Downloads a file from the given URL and saves it to the specified filename.

This script uses the requests library to send HTTP requests and the BeautifulSoup library to parse the HTML response. It also uses the argparse library to handle command-line arguments.

Example:
    $ python file_scraper.py --url http://example.com --extensions .txt .pdf --verbose

This will download all .txt and .pdf files from http://example.com and print verbose output.

Author: hanrai (hanrai@gmail.com)
Date: 20240526

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
import json
import os
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


def valid_url(string):
    """
    Validates if a given string is a valid URL.

    Args:
        string (str): The string to be validated.

    Raises:
        argparse.ArgumentTypeError: If the string is not a valid URL.

    Returns:
        str: The validated URL string.
    """
    if not string.startswith("http://") and not string.startswith("https://"):
        raise argparse.ArgumentTypeError(
            f"Invalid URL '{string}'. URL should start with 'http://' or 'https://'."
        )
    return string


def valid_extension(string):
    """
    Check if the given string is a valid file extension.

    Args:
        string (str): The string to be checked.

    Returns:
        str: The valid file extension.

    Raises:
        argparse.ArgumentTypeError: If the string is not a valid file extension.

    """
    if not string.startswith("."):
        raise argparse.ArgumentTypeError(
            f"Invalid extension '{string}'. Extension should start with a dot."
        )
    return string


def valid_folder(string, create_folder):
    """
    Check if the given folder path is valid.

    Args:
        string (str): The folder path to be checked.
        create_folder (bool): Flag indicating whether to create the folder if it doesn't exist.

    Returns:
        str: The validated folder path.

    Raises:
        argparse.ArgumentTypeError: If the folder path is invalid and create_folder is False.
    """
    if not os.path.isdir(string):
        if create_folder:
            os.makedirs(string)
        else:
            raise argparse.ArgumentTypeError(
                f"Invalid folder '{string}'. Folder does not exist."
            )
    return string


def download_files(
    url, extensions, project_folder, create_folder, verbose, overwrite, timeout
):
    """
    Downloads all files with the specified extensions from the given URL and saves them to the project folder.

    Args:
        url (str): The URL of the webpage to scan.
        extensions (list of str): The extensions of the files to download.
        project_folder (str): The folder where the downloaded files will be saved.
        create_folder (bool): If True, create the project folder if it does not exist. If False, raise an error if the project folder does not exist.
        verbose (bool): If True, print detailed download status information. If False, only print a message when the download is complete.
        overwrite (bool): If True, overwrite the files if they already exist. If False, skip the download if the files already exist.
        timeout (int): The number of seconds to wait for the server to send data before giving up.
    """
    # Check if the project folder exists
    if not os.path.exists(project_folder):
        if create_folder:
            os.makedirs(project_folder)
        else:
            raise FileNotFoundError(
                f"The project folder '{project_folder}' does not exist."
            )

    # Get the HTML content of the webpage
    response = requests.get(url, timeout=timeout)
    soup = BeautifulSoup(response.text, "html.parser")

    # Find all links in the webpage
    links = soup.find_all("a")

    # Filter the links to get only the files with the specified extensions
    files_to_download = [
        urljoin(url, link.get("href"))
        for link in links
        if any(link.get("href").endswith(extension) for extension in extensions)
    ]

    # If verbose is set, print a list of all files to be downloaded
    if verbose:
        files_dict = {"files": [os.path.basename(file) for file in files_to_download]}
        print(f"[Files to be downloaded] {json.dumps(files_dict)}")

    # Initialize counters for download results
    download_results = {"success": 0, "overwrite": 0, "skipped": 0, "error": 0}

    # Download the files
    for file in files_to_download:
        filename = os.path.join(project_folder, os.path.basename(file))
        try:
            result = download_file(
                file, filename, verbose, overwrite, timeout
            )  # Added timeout parameter here
            download_results[result] += 1
        except KeyboardInterrupt:
            print(f"\n[Interrupted] Download of {filename} interrupted.")
            if os.path.exists(filename):
                os.remove(filename)
                print(f"[Deleted] Partially downloaded file: {filename}")
            download_results["error"] += 1
            break

    # Print download results
    print(
        f"[Download Results] Success: {download_results['success']}, Overwrite: {download_results['overwrite']}, Skipped: {download_results['skipped']}, Error: {download_results['error']}"
    )


def download_file(url, filename, verbose, overwrite, timeout):
    """
    Downloads a file from the given URL and saves it to the specified filename.

    Args:
        url (str): The URL of the file to download.
        filename (str): The name of the file to save the downloaded content to.
        verbose (bool): If True, print detailed download status information. If False, only print a message when the download is complete.
        overwrite (bool): If True, overwrite the file if it already exists. If False, skip the download if the file already exists.
        timeout (int): The number of seconds to wait for the server to send data before giving up.
    """
    # Check if the file already exists
    if os.path.exists(filename):
        if overwrite:
            print(f"[Overwriting] filename: {filename}")
        else:
            print(f"[Skipping] filename: {filename}")
            return "skipped"

    try:
        response = requests.get(url, stream=True, timeout=timeout)
        total_size = int(response.headers.get("content-length", 0))

        # Download the file
        with open(filename, "wb") as f:
            for data in response.iter_content(1024):
                f.write(data)
                if verbose:
                    # Print the download status information in 'label: value' format
                    print(
                        f"[Downloading] filename: {filename}, downloaded_bytes: {f.tell()}, remaining_bytes: {total_size - f.tell()}"
                    )
        if not verbose:
            print(f"[Downloaded] filename: {filename}")
        return "success"
    except requests.exceptions.RequestException as e:
        print(f"[Error] Failed to download file: {filename}, error: {str(e)}")
        return "error"
    finally:
        # If the file is not fully downloaded, delete it
        if os.path.exists(filename) and os.path.getsize(filename) != total_size:
            os.remove(filename)
            print(f"[Deleted] Partially downloaded file: {filename}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Download files from a URL. If a file already exists, the program will skip it unless the '--overwrite' option is specified.",
        epilog="Example: python file_scraper.py http://example.com -e .pdf .txt -p /path/to/project_folder",
    )
    parser.add_argument("url", type=valid_url, help="The URL of the webpage to scan.")
    parser.add_argument(
        "-e",
        "--extensions",
        type=valid_extension,
        nargs="+",
        default=[".doc"],
        help="The extensions of the files to download.",
    )
    parser.add_argument(
        "-p",
        "--project_folder",
        type=str,
        default="./downloaded_files",
        help="The folder where the downloaded files will be saved.",
    )
    parser.add_argument(
        "-c",
        "--create_folder",
        action="store_true",
        help="Create the project folder if it does not exist.",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Print detailed download status information.",
    )
    parser.add_argument(
        "-o",
        "--overwrite",
        action="store_true",
        help="Overwrite the file if it already exists. If not specified, the program will skip the file.",
    )
    parser.add_argument(
        "-t",
        "--timeout",
        type=int,
        default=30,
        help="The number of seconds to wait for the server to send data before giving up.",
    )

    args = parser.parse_args()
    download_files(
        args.url,
        args.extensions,
        args.project_folder,
        args.create_folder,
        args.verbose,
        args.overwrite,
        args.timeout,
    )
