import configparser
import os
import random
import time

import pandas as pd
from Bio import Entrez, SeqIO

def read_email_from_config(config_file="config.ini"):
    """
    Reads the email address from the configuration file.

    Args:
        config_file: Path to the configuration file.

    Returns:
        The email address as a string.

    Raises:
        FileNotFoundError: If the configuration file is not found.
        KeyError: If the 'email' key is not found in the 'NCBI' section of the config file.
    """
    try:
        config = configparser.ConfigParser()
        config.read(config_file)
        return config['NCBI']['email']
    except FileNotFoundError:
        raise FileNotFoundError(f"Configuration file '{config_file}' not found. "
                                 "Please create a 'config.ini' file with the following content:\n"
                                 "[NCBI]\n"
                                 "email = your_email@example.com")
    except KeyError:
        raise KeyError("The 'email' key is not found in the 'NCBI' section of the config file. "
                       "Please make sure the config file is correctly formatted.")

def fetch_gene_data(gene_id):
    """
    Fetches gene data from NCBI Entrez.

    Args:
        gene_id: The ID of the gene to fetch.

    Returns:
        A tuple containing:
            - protein_accession: The accession number of the protein.
            - gene_length: The length of the gene.
            - protein_length: The length of the protein sequence.
    """
    try:
        # Add a random wait time between 0 and 1 second
        time.sleep(random.uniform(0, 1))

        gene_id = int(gene_id)
        # Fetch gene record
        with Entrez.efetch(db="gene", id=gene_id, retmode="xml") as handle:
            record = Entrez.read(handle)

        # Extract relevant information from the record
        locus = record[0]['Entrezgene_locus'][0]
        product = locus['Gene-commentary_products'][0]['Gene-commentary_products'][0]
        protein_accession = product['Gene-commentary_accession']

        genomic_coords = product['Gene-commentary_genomic-coords'][0]['Seq-loc_mix']['Seq-loc-mix'][0]['Seq-loc_int']['Seq-interval']
        gene_end = int(record[0]['Entrezgene_locus'][0]['Gene-commentary_seqs'][0]['Seq-loc_int']['Seq-interval']['Seq-interval_to'])
        gene_start = int(record[0]['Entrezgene_locus'][0]['Gene-commentary_seqs'][0]['Seq-loc_int']['Seq-interval']['Seq-interval_from'])
        gene_length = abs(gene_end - gene_start)
        print(f"Processing gene ID {gene_id} complete. Inferred gene length was: {gene_length}.")

        # Fetch protein sequence
        with Entrez.efetch(db="protein", id=protein_accession, rettype="fasta", retmode="text") as handle:
            protein_record = SeqIO.read(handle, "fasta")
        protein_length = len(protein_record.seq)
        print(f"Processing accession {protein_accession} complete. Protein length was: {protein_length}.")

        return protein_accession, gene_length, protein_length

    except Exception as e:
        print(f"Error processing gene ID {gene_id}: {e}")
        return None, None, None

def process_gene_list(input_file):
    """
    Processes a list of gene IDs from an Excel file and fetches data for each.

    Args:
        input_file: Path to the input Excel file with gene IDs.
    """
    try:
        email = read_email_from_config()  # Read email from config
        Entrez.email = email  # Set the email for Entrez

        df = pd.read_excel(input_file)
        gene_ids = df['Gene ID'].tolist()  # Assuming the column with gene IDs is named 'Gene ID'

        data = []
        for gene_id in gene_ids:
            accession, gene_len, protein_len = fetch_gene_data(gene_id)
            data.append([gene_id, accession, gene_len, protein_len])

        output_file = os.path.splitext(input_file)[0] + '_output.xlsx'  # Create output file name
        output_df = pd.DataFrame(data, columns=['Gene ID', 'Accession', 'Gene Length', 'Protein Length'])
        output_df.to_excel(output_file, index=False)
        print(f"Output saved to: {output_file}")

    except (FileNotFoundError, KeyError) as e:
        print(e)  # Print the specific error message

if __name__ == "__main__":
    input_excel_file = input("Enter the path to the input Excel file: ")
    process_gene_list(input_excel_file)