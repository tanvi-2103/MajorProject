from flask import Flask, render_template, request,render_template_string
from prettytable import PrettyTable
import pandas as pd
import xlwings as xw
import numpy as np
import openpyxl

directm = Flask(__name__)

@directm.route('/')
def index():
    return render_template('index.html')

@directm.route('/process_data', methods=['POST'])
def process_data():
    uploaded = request.files['Direct']
    file_path = 'uploads/input_data.xlsx'
    uploaded.save(file_path)
    df = pd.read_excel(file_path)

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

    direct_html_table = direct_table.to_html(index=False)

    # Return the results to the HTML template
    return render_template('result.html', result=direct_html_table)
    # return render_template_string('<html><body>{{ html_table }}</body></html>', html_table=iat2_html_table)

if __name__ == '__main__':
    directm.run(debug=True)