# This script fetches gene data from NCBI Entrez and saves the results to an Excel file.

## Usage

1. Clone this repository.
2. Create a configuration file named `config.ini` with the following content:

	```[NCBI]
	email = your_email@example.com```
3. Replace `your_email@example.com` with your actual email address.
4. Run the script and provide the path to your input Excel file when prompted. The inout excel file should contain a column with the header 'Gene ID' (this column should contain all the gene ids.

## Output

The script will save the output data to an Excel file named `input_file_output.xlsx`, where `input_file` is the name of the input Excel file. The output file will contain the following columns:

* `Gene ID`
* `Accession`
* `Gene Length`
* `Protein Length`

## Dependencies

This script requires the following Python libraries:

* pandas
* Biopython

## Installation

You can install the required libraries using pip:

`pip install pandas biopython`