import pandas as pd
import pyreadstat
import openpyxl

# Import the data set using pyreadstat
data_path = '../../SPSS-Python/spss-datasets/PH_consumer_data_practice.sav'
df, meta = pyreadstat.read_sav(data_path)

# list of variables to use as columns and rows for the crosstabs
row_variables = ['Q3_1', 'Q3_2','Q3_3','Q3_4','Q3_5']
column_variables = ['GENDER', 'WhiteVNonWhite', 'AGEGROUP', 'INCOMEGROUP']
excel_file_name = f'{row_variables[0]}_series.xlsx'
print(excel_file_name)
output_path = f'../output_manipulation/output_folder/'
print(output_path)
print(output_path + excel_file_name)

## This codeblock changes variable value labels                                   ##
## I had to do this because the old label used a '/' and was not python friendly  ##
## Would have been easier to just change the label in SPSS and save as a new dataset ##
# Access the variable metadata
var_name = 'GENDER'
var_metadata = meta.variable_value_labels.get(var_name)

# Check if the variable has value labels
if var_metadata is not None:
    
    # Define the new value labels
    new_value_labels = {
        1.0: 'Male',
        2.0: 'Female',
        3.0: 'Other_non_binary'
    }

    # Update the variable's value labels
    meta.variable_value_labels[var_name] = new_value_labels

    # Save the updated metadata to a new file
    new_meta_path = "../output_manipulation/output_folder/gender_metadata_value_label_updated.sav"
    
else:
    print("Variable does not have value labels.")

# Create a dictionary so that we can selectively display the correct label for a column if necessary.
# I used this line because I saw it in a tutorial but the meta_dict object never gets called in this code
# so i commented it out
#meta_dict = dict(zip(meta.column_names, meta.column_labels))

# create the dataframe and excel file to save crosstab outputs to
# crosstab_result_concat = pd.DataFrame()
df1 = pd.DataFrame()
df1.to_excel(f'{output_path + excel_file_name}')

# generate crosstabs for each combination of rows and columns
for row in row_variables:
    crosstab_result_concat = pd.DataFrame()

    for col in column_variables:
        crosstab_result = pd.crosstab(df[row].\
            #map(meta.variable_value_labels['Q3_1']), \
            map(meta.variable_value_labels[row]), \
            df[col].map(meta.variable_value_labels[col]), \
            dropna=True, normalize='columns'). \
            loc[meta.variable_value_labels[row].values()]. \
            loc[:,meta.variable_value_labels[col].values()]
            
        print(f'crosstab result for {col} is: {crosstab_result}')
        
        # This section is to filter the output to only pass "Checked" responses for row variable
        # This works but output is formatted vertically vice horizontally
        # iloc method will fail if order of row variable value label is changed
        print('\n')
        selected_row = crosstab_result.iloc[-1]
        #selected_row = selected_row.T
        print('selected a row to keep')
        print(selected_row)

        # join the result of this loop to the consolidation dataframe
        crosstab_result_concat = pd.concat([crosstab_result_concat, crosstab_result], axis=1)
    print(f'crosstab_result_concat for {col} is: \n ', crosstab_result_concat)

    # Append crosstab_result_concat to existing xlsx file here
    reader = pd.read_excel(f'{output_path + excel_file_name}')
    with pd.ExcelWriter(f'{output_path + excel_file_name}', engine='openpyxl', if_sheet_exists='overlay', mode='a') as writer:
        crosstab_result_concat.to_excel(writer, sheet_name='Sheet1', startrow=len(reader)+1)
    crosstab_result = []
    # with pd.ExcelWriter('../output_manipulation/output_folder/crosstabs.xlsx', engine='openpyxl', if_sheet_exists='overlay', mode='a') as writer:
    #     crosstab_result_concat.to_excel(writer, sheet_name='Sheet1', startrow=row_num, startcol=col_num)
    # row_num += 3
    # crosstab_result = []
    
    #print(f'crosstab_result_concat for {row} is:')
print(crosstab_result_concat)