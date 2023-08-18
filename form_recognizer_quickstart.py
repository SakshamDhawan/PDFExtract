import os
import numpy as np
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# use your `key` and `endpoint` environment variables
key = os.environ.get('FR_KEY')
endpoint = os.environ.get('FR_ENDPOINT')

def analyze_general_documents(url):

    # create your `DocumentAnalysisClient` instance and `AzureKeyCredential` variable
    document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))

    poller = document_analysis_client.begin_analyze_document_from_url(
        "prebuilt-document", url)
    result = poller.result()

    filename = (url.split("/")[-1]).split(".")[0]  # Get filename from uploaded file URL

    dfs = []  # Create empty list for dataframes to be written

    startrow = 0  # index to output each dataframe on a new row

    initial_df = pd.DataFrame(result.tables)  # Get results into DF

    for k in range(len(result.tables)):
        rows = initial_df[0][k].row_count  # Total number of rows
        columns = initial_df[0][k].column_count  # Total number of columns

        raw_data_df = initial_df[0][k].cells  # Extracted data

        final_df = pd.DataFrame(np.zeros([rows, columns]))  # Create empty dataframe
        # with number of rows and columns defined by data shape

        for j in range(len(raw_data_df)):
            final_df.at[raw_data_df[j].row_index, raw_data_df[j].column_index] = raw_data_df[j].content  # Repopulate
            # dataframe

        dfs.append(final_df)  # Add each dataframe to output list

    with pd.ExcelWriter(filename + ".xlsx") as writer:  # ExcelWriter more powerful than to_csv method
        for df in dfs:
            df.to_excel(writer, engine="xlsxwriter", startrow=startrow, index=False, header=False)
            startrow += (df.shape[0] + 2)


if __name__ == "__main__":
    import sys

    analyze_general_documents(sys.argv[1])
