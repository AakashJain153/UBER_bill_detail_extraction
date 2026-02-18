# Uber Bill Extractor

This script extracts details from Uber PDF receipts and:

- Renames them to YYYYMMDD_{amount}.pdf
- Stores them in a Refined folder
- Generates a summary Excel file

## Installation

1. Install Python 3.9+
2. Install dependencies:

pip install -r requirements.txt

## Usage

Run:

python UBER_Extract_merge_summarize.py

Select the folder containing Uber receipts.

Output will be created in a folder named `Refined`.
