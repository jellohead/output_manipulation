#!/usr/bin/env python3
import pandas as pd
import pyreadstat

# Import the data set using pyreadstat
data_path = "/Users/patescalona/Library/CloudStorage/OneDrive-Personal/Marketing Reports/CCR Reports/MCCU/Working/MCCU Report Draft and Templat_/9437 Financial Data_rev2.sav"
df, meta = pyreadstat.read_sav(data_path)

# list of variables to use as columns and rows for the crosstabs
row_variables = [
    "Q14_r1_recode",
    "Q14_r2_recode",
    "Q14_r3_recode",
    "Q14_r4_recode",
    "Q14_r5_recode",
    "Q14_r6_recode",
]
column_variables = [
    "Gender",
    "Race_1",
    "Hispanic",
    "Race_2",
    "Asian1",
    "Qual1_recode",
    "Qual7",
    "BankingAccounts",
]
excel_file_name = f"{row_variables[0]}_series.xlsx"
output_path = f"../../../../OneDrive/Marketing Reports/CCR Reports/MCCU/Working/"
print(f"Output file: {output_path + excel_file_name}")

# create the dataframe and excel file to save crosstab outputs to
df1 = pd.DataFrame()
df1.to_excel(f"{output_path + excel_file_name}")


def my_function(row_variables, column_variables):
    for row in row_variables:
        crosstab_result_concat = pd.DataFrame()

        for col in column_variables:
            crosstab_result = (
                pd.crosstab(
                    df[row].map(meta.variable_value_labels[row]),
                    df[col].map(meta.variable_value_labels[col]),
                    dropna=True,
                    normalize="columns",
                )
                .loc[meta.variable_value_labels[row].values()]
                .loc[:, meta.variable_value_labels[col].values()]
            )

            print(f"crosstab result for {col} is: {crosstab_result}")

            # join the result of this loop to the consolidation dataframe
            crosstab_result_concat = pd.concat(
                [crosstab_result_concat, crosstab_result], axis=1
            )
        print(f"crosstab_result_concat for {col} is: \n ", crosstab_result_concat)

        # Append crosstab_result_concat to existing xlsx file here
        reader = pd.read_excel(f"{output_path + excel_file_name}")
        with pd.ExcelWriter(
            f"{output_path + excel_file_name}",
            engine="openpyxl",
            if_sheet_exists="overlay",
            mode="a",
        ) as writer:
            crosstab_result_concat.to_excel(
                writer, sheet_name="Sheet1", startrow=len(reader) + 1
            )
        crosstab_result = []
    print(crosstab_result_concat)


# Call the function
my_function(row_variables, column_variables)
