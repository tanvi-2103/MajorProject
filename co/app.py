from flask import Flask, render_template, request,render_template_string, jsonify
from prettytable import PrettyTable
import pandas as pd
import xlwings as xw
import numpy as np
import openpyxl

app = Flask(__name__ , static_url_path='/static')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process_data', methods=['POST'])
def process_data():
    # Get input file from HTML form
    # uploaded_file = request.files['IAT1']
    # file_path = 'uploads/input_data.xlsx'
    # uploaded_file.save(file_path)

    IAT1_uploaded = request.files['IAT1']
    IAT1_path = 'uploads/input_data.xlsx'
    IAT1_uploaded.save(IAT1_path)
    IAT1 = pd.read_excel(IAT1_path)

    IAT2_uploaded = request.files['IAT2']
    IAT2_path = 'uploads/input_data.xlsx'
    IAT2_uploaded.save(IAT2_path)
    IAT2 = pd.read_excel(IAT2_path)

    #direct
    uploaded = request.files['Direct']
    file_path = 'uploads/input_data.xlsx'
    uploaded.save(file_path)
    df = pd.read_excel(file_path)

    #indirect
    uploaded1 = request.files['Indirect']
    indirect_path = 'uploads/input_data.xlsx'
    uploaded1.save(indirect_path)
    indirect = pd.read_excel(indirect_path)

    # Read the input Excel file
    # df = pd.read_excel(file_path)

    # The rest of your existing code for data processing
    # ...
    #IAT1
    # Specify the common column you want to process and the number of columns to include
    C_C = 'CO1'
    num_columns_to_process = 4  # You want to process "CO4" and the next 3 columns

    # Find the index of the common column in the DataFrame's columns
    common_column_index = IAT1.columns.get_loc(C_C)
    co_attainment_dict = {}

    iat1_table = PrettyTable()
    # Extract the columns you want to process based on the index

    columns_to_p = IAT1.columns[common_column_index:common_column_index + num_columns_to_process]

    iat1_table.field_names = ["Column", "Total Students Passed (> 60%)", "Total Students Attempted", "CO Attainment", "Attainment Level"]

    for column in columns_to_p:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT1.columns:
         IAT1[column] = pd.to_numeric(IAT1[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT1[column] >= 3).sum()

        # Count filled cells in the column
        filled_cells_count = IAT1[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0

        co_attainment_dict[column] = CO_attainment
        iat1_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])
        # Print or store the results as needed
        # print(f'Column: {column},   Total student passed with > 60%:  {count_greater_than_3},   Total # students attempted the QUESTION:  {filled_cells_count},   CO_attainment: {CO_attainment},    attainment_Level: {attainment_Level}')

    second_column_index = IAT1.columns.get_loc(C_C)+4
    columns_to_next = IAT1.columns[second_column_index:second_column_index + (num_columns_to_process-2) ]
    for column in columns_to_next:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT1.columns:
         IAT1[column] = pd.to_numeric(IAT1[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT1[column] >= 6).sum()

        # Count filled cells in the column
        filled_cells_count = IAT1[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0


        co_attainment_dict[column] = CO_attainment
        iat1_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])

    # Convert PrettyTable to DataFrame
    iat1_table_df = pd.read_html(iat1_table.get_html_string())[0]

    # Transpose the DataFrame
    iat1_transposed_table = iat1_table_df.T

    # Print the transposed table

    new_column_values = ['', 'Total Students Passed (> 60%) ', 'Total Students Attempted', 'CO Attainment ','Attainment Level']  # Replace this with your actual values
    iat1_transposed_table.insert(0, "NewColumn", new_column_values)
    # Set the first row as the header
    iat1_transposed_table.columns = iat1_transposed_table.iloc[0]

    # Drop the first row (which is now the header) to avoid duplicate headers
    transposed_table = iat1_transposed_table[1:]
    print(transposed_table)

    v_variables = {}
    counter = 1
    avg_attainment_list = []

    # Create a dictionary to store average CO attainment values for columns with the same initial three characters
    average_co_attainment_dict = {}

    # Iterate through the columns to calculate average CO attainment
    for column_name, attainment in co_attainment_dict.items():
        initial_3_characters = column_name[:3]
        if initial_3_characters in average_co_attainment_dict:
            average_co_attainment_dict[initial_3_characters].append(attainment)
        else:
            average_co_attainment_dict[initial_3_characters] = [attainment]

    v_variables = {}
    counter = 1
    avg_attainment_list = []
    # Calculate the average CO attainment for each group of columns with the same initial three characters
    for group, attainments in average_co_attainment_dict.items():
        avg_attainment = sum(attainments) / len(attainments)
        avg_attainment_list.append(avg_attainment)



    iat1 = {
        'CO':['CO1','CO2','CO3','CO4','CO5','CO6'],
        'avg_attainment':avg_attainment_list
    }
    while len(avg_attainment_list) < len(iat1['CO']):
        avg_attainment_list.append(0)
    iat1_data = pd.DataFrame(iat1)
    print(iat1_data)

    #IAT2
    # Specify the common column you want to process and the number of columns to include
    C_C = 'CO4'
    num_columns_to_process = 4  # You want to process "CO4" and the next 3 columns

    # Find the index of the common column in the DataFrame's columns
    common_column_index = IAT2.columns.get_loc(C_C)
    co_attainment_dict = {}
    # Extract the columns you want to process based on the index
    iat2_table = PrettyTable()

    columns_to_p = IAT2.columns[common_column_index:common_column_index + num_columns_to_process]

    iat2_table.field_names = ["Column", "Total Students Passed (> 60%)", "Total Students Attempted", "CO Attainment", "Attainment Level"]
    for column in columns_to_p:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT2.columns:
         IAT2[column] = pd.to_numeric(IAT2[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT2[column] >= 3).sum()

        # Count filled cells in the column
        filled_cells_count = IAT2[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0

        co_attainment_dict[column] = CO_attainment
        iat2_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])

        # Print or store the results as needed
        # print(f'Column: {column},   Total student passed with > 60%:  {count_greater_than_3},   Total # students attempted the QUESTION:  {filled_cells_count},   CO_attainment: {CO_attainment},    attainment_Level: {attainment_Level}')

    second_column_index = IAT2.columns.get_loc(C_C)+4
    columns_to_next = IAT2.columns[second_column_index:second_column_index + (num_columns_to_process-2) ]
    for column in columns_to_next:
        # Convert the column to numeric, ignoring non-numeric values
    # Check if the column exists in the DataFrame
        if column in IAT2.columns:
         IAT2[column] = pd.to_numeric(IAT2[column], errors='coerce')

        # Count cells in the column with a value greater than or equal to 3
        count_greater_than_3 = (IAT2[column] >= 6).sum()

        # Count filled cells in the column
        filled_cells_count = IAT2[column].count()

        # Calculate CO_attainment
        CO_attainment = round((count_greater_than_3 / filled_cells_count) * 100)

        # Determine attainment level
        if CO_attainment >= 50 and CO_attainment < 60:
            attainment_Level = 1
        elif CO_attainment >= 60 and CO_attainment < 70:
            attainment_Level = 2
        elif CO_attainment >= 70:
            attainment_Level = 3
        else:
            attainment_Level = 0


        co_attainment_dict[column] = CO_attainment
        iat2_table.add_row([column, count_greater_than_3, filled_cells_count, CO_attainment, attainment_Level])

    # Convert PrettyTable to DataFrame
    iat2_table_df = pd.read_html(iat2_table.get_html_string())[0]

    # Transpose the DataFrame
    iat2_transposed_table = iat2_table_df.T

    # Print the transposed table

    iat2_new_column_values = ['', 'Total Students Passed (> 60%) ', 'Total Students Attempted', 'CO Attainment ','Attainment Level']  # Replace this with your actual values
    iat2_transposed_table.insert(0, "NewColumn", iat2_new_column_values)
    # Set the first row as the header
    iat2_transposed_table.columns = iat2_transposed_table.iloc[0]

    # Drop the first row (which is now the header) to avoid duplicate headers
    iat2_transposed_table = iat2_transposed_table[1:]
    print(iat2_transposed_table)
        # Print or store the results as needed
        # print(f'Column: {column},   Total student passed with > 60%:  {count_greater_than_3},   Total # students attempted the QUESTION:  {filled_cells_count},   CO_attainment: {CO_attainment},    attainment_Level: {attainment_Level}')



    # Create a dictionary to store average CO attainment values for columns with the same initial three characters
    average_co_attainment_dict = {}

    # Iterate through the columns to calculate average CO attainment
    for column_name, attainment in co_attainment_dict.items():
        initial_3_characters = column_name[:3]
        if initial_3_characters in average_co_attainment_dict:
            average_co_attainment_dict[initial_3_characters].append(attainment)
        else:
            average_co_attainment_dict[initial_3_characters] = [attainment]

    # Create a list to store the data for the DataFrame
    table_data = {'CO': [], 'Avg_Attainment': []}

    # Include all CO values in the list, even if they didn't have data
    for co_value in ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6']:
        if co_value in average_co_attainment_dict:
            avg_attainment = sum(average_co_attainment_dict[co_value]) / len(average_co_attainment_dict[co_value])
        else:
            avg_attainment = 0
        table_data['CO'].append(co_value)
        table_data['Avg_Attainment'].append(avg_attainment)

    # Create DataFrame
    iat2_data_df = pd.DataFrame(table_data)

    # Print or use the DataFrame as needed
    print(iat2_data_df)

    iat2_html_table = iat2_data_df.to_html(index=False)




    #Direct
    total_rows = df.shape[0] - 1 if not df.empty else 0
    # print(f'Total number of rows: {total_rows}')



    C_C_IA = 'IA'
    column_sum_IA = df[C_C_IA].sum()

    Average_percentage_IA = (column_sum_IA /(total_rows*20))*100
    # print(Average_percentage_IA)

    count_above_taget_level_IA = (df[C_C_IA] > 10).sum()
    # print(count_above_taget_level_IA)

    percentage_of_direct_assesment_IA = count_above_taget_level_IA*100/total_rows
    # print(percentage_of_direct_assesment_IA)

    if percentage_of_direct_assesment_IA >= 50 and percentage_of_direct_assesment_IA < 60:
            attainment_Level_IA = 1
    elif percentage_of_direct_assesment_IA >= 60 and percentage_of_direct_assesment_IA < 70:
            attainment_Level_IA = 2
    elif percentage_of_direct_assesment_IA >= 70:
            attainment_Level_IA = 3
    else:
            attainment_Level_IA = 0
    # print(attainment_Level_IA)

    Weightage_IA = 20*100/150
    # print(Weightage_IA)

    attainment_level_weightage_IA = Weightage_IA*attainment_Level_IA/100
    # print(attainment_level_weightage_IA)



    C_C_TW = "TW"
    column_sum_TW = df[C_C_TW].sum()

    Average_percentage_TW = (column_sum_TW /(total_rows*25))*100
    # print(Average_percentage_TW)

    count_above_taget_level_TW = (df[C_C_TW] >= 15).sum()
    # print(count_above_taget_level_TW)

    percentage_of_direct_assesment_TW = count_above_taget_level_TW*100/total_rows
    # print(percentage_of_direct_assesment_TW)

    if percentage_of_direct_assesment_TW >= 50 and percentage_of_direct_assesment_TW < 60:
                attainment_Level_TW = 1
    elif percentage_of_direct_assesment_TW >= 60 and percentage_of_direct_assesment_TW < 70:
                attainment_Level_TW = 2
    elif percentage_of_direct_assesment_TW >= 70:
                attainment_Level_TW = 3
    else:
                attainment_Level_TW = 0
    # print(attainment_Level_TW)

    Weightage_TW = 25*100/150
    # print(Weightage_TW)
    attainment_level_weightage_TW = Weightage_TW*attainment_Level_TW/100
    # print(attainment_level_weightage_TW)




    C_C_OP = "O/P"
    column_sum_OP = df[C_C_OP].sum()

    Average_percentage_op = (column_sum_OP /(total_rows*25))*100
    # print(Average_percentage_op)

    count_above_taget_level_op = (df[C_C_OP] >= 14).sum()
    # print(count_above_taget_level_op)

    percentage_of_direct_assesment_op = count_above_taget_level_op*100/total_rows
    # print(percentage_of_direct_assesment_op)

    if percentage_of_direct_assesment_op >= 50 and percentage_of_direct_assesment_op < 60:
                attainment_Level_op = 1
    elif percentage_of_direct_assesment_op >= 60 and percentage_of_direct_assesment_op < 70:
                attainment_Level_op = 2
    elif percentage_of_direct_assesment_op >= 70:
                attainment_Level_op = 3
    else:
                attainment_Level_op = 0
    # print(attainment_Level_op)

    Weightage_op = 25*100/150
    # print(Weightage_op)
    attainment_level_weightage_op = Weightage_TW*attainment_Level_op/100
    # print(attainment_level_weightage_op)




    C_C_ENDSEM = 'END Sem'
    column_sum_ENDSEM = df[C_C_ENDSEM].sum()

    Average_percentage_ENDSem = (column_sum_ENDSEM /(total_rows*80))*100
    # print(Average_percentage_ENDSem)

    count_above_taget_level_ENDSem = (df[C_C_ENDSEM] >= 40).sum()
    # print(count_above_taget_level_ENDSem)

    percentage_of_direct_assesment_ENDSem = count_above_taget_level_ENDSem*100/total_rows
    # print(percentage_of_direct_assesment_ENDSem)

    if percentage_of_direct_assesment_ENDSem >= 50 and percentage_of_direct_assesment_ENDSem < 60:
            attainment_Level_ENDSem = 1
    elif percentage_of_direct_assesment_ENDSem >= 60 and percentage_of_direct_assesment_ENDSem < 70:
            attainment_Level_ENDSem = 2
    elif percentage_of_direct_assesment_ENDSem >= 70:
            attainment_Level_ENDSem = 3
    else:
            attainment_Level_ENDSem = 0
    # print(attainment_Level_ENDSem)

    Weightage_ENDSem = 80*100/150
    # print(Weightage_ENDSem)

    attainment_level_weightage_ENDSem = Weightage_ENDSem*attainment_Level_ENDSem/100
    # print(attainment_level_weightage_ENDSem)




    C_C_IA20 = 'IA20'
    column_sum_IA20 = df[C_C_IA20].sum()
    # print(column_sum_IA20)

    def IA_COs(row_index):
        v = df.loc[row_index, 'IA20']
        CO_IA = v * percentage_of_direct_assesment_IA / column_sum_IA20
        return round(CO_IA,2)  # Return the computed value instead of printing

    c0 = IA_COs(0)
    c1 = IA_COs(1)
    c2 = IA_COs(2)
    c3 = IA_COs(3)
    c4 = IA_COs(4)
    c5 = IA_COs(5)




    C_C_TW25 = 'TW25'
    column_sum_TW25 = df[C_C_TW25].sum()
    # print(column_sum_TW25)

    def TW_COs(row_index):
        v = df.loc[row_index, 'TW25']
        CO_IA = v * percentage_of_direct_assesment_TW/ column_sum_TW25
        return round(CO_IA,2) # Return the computed value instead of printing

    T0 = TW_COs(0)
    T1 = TW_COs(1)
    T2 = TW_COs(2)
    T3 = TW_COs(3)
    T4 = TW_COs(4)
    T5 = TW_COs(5)





    C_C_OP25 = 'O/P25'
    column_sum_OP25 = df[C_C_OP25].sum()
    # print(column_sum_OP25)

    def op_COs(row_index):
        v = df.loc[row_index, 'O/P25']
        CO_IA = v * percentage_of_direct_assesment_op / column_sum_OP25
        return round(CO_IA,2)  # Return the computed value instead of printing

    o0 = op_COs(0)
    o1 = op_COs(1)
    o2 = op_COs(2)
    o3 = op_COs(3)
    o4 = op_COs(4)
    o5 = op_COs(5)





    C_C_END80 = 'END80'
    column_sum_END80 = df[C_C_END80].sum()
    # print(column_sum_END80)

    def E_COs(row_index):
        v = df.loc[row_index, 'END80']
        CO_IA = v * percentage_of_direct_assesment_ENDSem / column_sum_END80
        return round(CO_IA,2)  # Return the computed value instead of printing

    E0 = E_COs(0)
    E1 = E_COs(1)
    E2 = E_COs(2)
    E3 = E_COs(3)
    E4 = E_COs(4)
    E5 = E_COs(5)



    print('Output of Direct Sheet :')
    print()
    direct_data = {
        'CO': ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6'],
        'IA': [c0, c1, c2, c3, c4, c5],
        'TW': [T0,T1,T2,T3,T4,T5],
        'O/P': [o0,o1,o2,o3,o4,o5],
        'END sem': [E0,E1,E2,E3,E4,E5]
    }
    direct_table = pd.DataFrame(direct_data)
    print(direct_table)




    #indirect
    column_name = '%'  # Replace with the actual column name

    # Read the specific column into a DataFrame
    indirect_df = pd.read_excel(indirect_path, usecols=[column_name])
    column_name = '%'  # Replace with the actual column name

    # Read the specific column into a DataFrame
    indirect1 = pd.read_excel(indirect_path, usecols=[column_name])
    transposed_indirect1 = indirect1.T

    transposed_indirect1 = transposed_indirect1.drop(transposed_indirect1.columns[0], axis=1)
    header_names = ["CO1", "CO2", "CO3",'CO4','CO5','CO6']  # Replace with actual header names
    transposed_indirect1.columns = header_names
    print(transposed_indirect1)


#     fcar


    iat1 = pd.DataFrame(iat1)
    transposed_iat1 = iat1.T
    iat2_data_df = pd.DataFrame(table_data)
    iat2_transposed_df = iat2_data_df.T
    direct_table = pd.DataFrame(direct_data)
    transposed_direct = direct_table.T
    print("hello")
    

     # Concatenate the DataFrames vertically
    appended_df = pd.concat([transposed_iat1, iat2_transposed_df, transposed_direct])
    
    # Drop the 'CO' column if it exists
    # if 'CO' in appended_df.columns:
    #     appended_df = appended_df.drop('CO', axis=1)
    
    # Rename columns
    column_names = ['CO1', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6']
    appended_df.columns = column_names

    # Concatenate with transposed_indirect1
    fcar = pd.concat([appended_df, transposed_indirect1])
    final = pd.DataFrame(fcar)
    
    # Filter out rows containing "CO" in any column
    final = final[~final.apply(lambda row: row.astype(str).str.contains('CO').any(), axis=1)]
    new_column_values = ['IAT1', 'IAT2', 'IA', 'TW', 'O/P', 'END Sem','indirect']
    final.insert(0, '', new_column_values)
    fcar_no_zeros = final.replace(0, np.nan)
    # means = {}
    means = {}
    for col in fcar_no_zeros.columns:
        if pd.api.types.is_numeric_dtype(fcar_no_zeros[col]):  # Check if column is numeric
            mean_values = fcar_no_zeros[col].mean()
            means[col] = mean_values
    
    # Concatenate mean values as a new row to the DataFrame
    mean_row = pd.DataFrame(means, index=['Mean'])
    fcar_with_means = pd.concat([fcar_no_zeros, mean_row])
    fcar_with_means = fcar_with_means.round(2)


    # Convert the DataFrame to an HTML table
    fcar_html_table = fcar_with_means.to_html(index=False)

    return render_template('result.html', result=fcar_html_table)

if __name__ == '__main__':
    app.run(debug=True)