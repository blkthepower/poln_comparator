import os
from datetime import datetime as dt

import pandas as pd

RD_POLN_COLUMN_NAME = 'Unnamed: 0'
RD_QTY_COLUMN_NAME = 'Qty Open'
RD_REQUESTED_DATE_COLUMN_NAME = 'Requested Date'
RD_ID_COLUMN_NAME = 'VN ID' 

LD_POLN_COLUMN_NAME = 'Unnamed: 0'
LD_QTY_COLUMN_NAME = 'Balance'
LD_REQUESTED_DATE_COLUMN_NAME = 'FechaRequerida'
LD_ID_COLUMN_NAME = 'IfgNo'
CC_IRREGULAR_REASON_COLUMN_NAME = 'IRREGULARIDAD ENCONTRADA'

IR_WRONG_DATE = 'FECHA DIFERENTE | '
IR_WRONG_POLN = 'POLN DIFERENTE | '
IR_WRONG_QTY = 'CANTIDAD DIFERENTE | '
IR_MULTIPLE_MATCHES = 'MULTIPLES OCURRENCIAS | '
IR_WRONG_ID = 'ID DIFERENTE |'

version_number = '1.0.0'

def welcome_message():
    print('*****POLN COMPARATOR v' + version_number + '*****\n')
    print('[*]Make sure your file follows the correct format. See "format_example_file.xlsx". ')
    print('[*]The file should be an excel file named "target_file.xlsx".')
    print('[*]The file should be located in the same folder as this program. Current path: ' + os.getcwd())
    print('[*]Close any application that is using the file you are going to process.\n')


def get_date_counts(remote_data, local_data):
    
    date_count_rd = remote_data.groupby([RD_REQUESTED_DATE_COLUMN_NAME]).size()
    date_count_ld = local_data.groupby([LD_REQUESTED_DATE_COLUMN_NAME]).size()

    return date_count_ld, date_count_rd

def compare_sheets_data():
    start_comparison = input('Press Enter to start the process: ')
    
    # if str(start_comparison).lower() == 'y':
    print('Processing...')
    try:

        # base_excel_file = pd.read_excel("polivol_test_file.xlsx", index_col=None, header=1)
        base_excel_file = pd.ExcelFile('target_file.xlsx')
        remote_data = pd.read_excel(base_excel_file, sheet_name=base_excel_file.sheet_names[0], index_col=None, header=1)
        local_data = pd.read_excel(base_excel_file, sheet_name=base_excel_file.sheet_names[1], index_col=None, header=0)
        base_excel_file.close()

        # print(remote_data.columns)
        # print(local_data.columns)

        output_file_columns = local_data.columns.append(pd.Index([CC_IRREGULAR_REASON_COLUMN_NAME]))
        irregular_data = pd.DataFrame(columns=output_file_columns)

        subset_remote_data = remote_data.loc[:, [RD_POLN_COLUMN_NAME, RD_ID_COLUMN_NAME, RD_QTY_COLUMN_NAME, RD_REQUESTED_DATE_COLUMN_NAME]]
        subset_remote_data[RD_REQUESTED_DATE_COLUMN_NAME] = pd.to_datetime(subset_remote_data[RD_REQUESTED_DATE_COLUMN_NAME])
        subset_remote_data[RD_REQUESTED_DATE_COLUMN_NAME] = subset_remote_data[RD_REQUESTED_DATE_COLUMN_NAME].dt.strftime('%d-%m-%Y')

        local_data[LD_REQUESTED_DATE_COLUMN_NAME] = pd.to_datetime(local_data[LD_REQUESTED_DATE_COLUMN_NAME])
        local_data[LD_REQUESTED_DATE_COLUMN_NAME] = local_data[LD_REQUESTED_DATE_COLUMN_NAME].dt.strftime('%d-%m-%Y')
        
        foundMatchingPoln = False
        hasMatchingBalance = False
        hasMatchingRequestDate = False
        hasMatchingId = False
        polnHasMatchedBefore = False
        matchingRowsCount = 0
        irregular_reasons = ''
        
        for index_ld, row_ld in local_data.iterrows():
            matchingRowsCount = 0
            irregular_reasons = ''
            polnHasMatchedBefore = False
            row_ld[CC_IRREGULAR_REASON_COLUMN_NAME] = ''
            
            for index_rd, row_rd in subset_remote_data.iterrows():
                foundMatchingPoln = False
                hasMatchingBalance = False
                hasMatchingRequestDate = False
                
                if row_rd[RD_POLN_COLUMN_NAME] == row_ld[LD_POLN_COLUMN_NAME] and row_rd[RD_ID_COLUMN_NAME] == row_ld[LD_ID_COLUMN_NAME]:
                    foundMatchingPoln = True
                    
                    if row_rd[RD_QTY_COLUMN_NAME] == row_ld[LD_QTY_COLUMN_NAME]:
                        hasMatchingBalance = True
                    else:
                        irregular_reasons += IR_WRONG_QTY
                        
                    if row_rd[RD_REQUESTED_DATE_COLUMN_NAME] == row_ld[LD_REQUESTED_DATE_COLUMN_NAME]:
                        hasMatchingRequestDate = True
                    else:
                        irregular_reasons += IR_WRONG_DATE
                
                if foundMatchingPoln:
                    if hasMatchingBalance and hasMatchingRequestDate:
                        matchingRowsCount += 1
                    
                    if polnHasMatchedBefore:                        
                            # irregular_data = pd.concat([irregular_data, row_ld.to_frame().T])               
                            irregular_reasons = ''
                    elif irregular_reasons != '':
                        row_ld[CC_IRREGULAR_REASON_COLUMN_NAME] = irregular_reasons
                    
                    polnHasMatchedBefore = True
                        

            if matchingRowsCount > 1:
                row_ld[CC_IRREGULAR_REASON_COLUMN_NAME] += IR_MULTIPLE_MATCHES
                # irregular_data = pd.concat([irregular_data, matchingRows])

            
            if matchingRowsCount > 0 and row_ld[CC_IRREGULAR_REASON_COLUMN_NAME] != '':
                irregular_data = pd.concat([irregular_data, row_ld.to_frame().T])
                

        print('Found ' + str(len(irregular_data.index)) + ' non-matching rows.')
        
        print('Counting dates...')
        date_count_ld, date_count_rd = get_date_counts(remote_data, local_data)
        
        print('Updating file...')
        with pd.ExcelWriter('target_file.xlsx', mode='a', if_sheet_exists='replace') as excel_writer:         
            if len(irregular_data.index) > 0:            
                irregular_data.to_excel(excel_writer, sheet_name='Irregular data')
            
            date_count_ld.to_excel(excel_writer, sheet_name='Local dates count' )
            date_count_rd.to_excel(excel_writer, sheet_name='Remote dates count')
        
        print("Process completed.")

    except Exception as ex:
        print("Oops! Something wrong happened! " + str(ex))
    # else:
    #     print("Process canceled.")


if __name__ == '__main__':
    welcome_message()
    compare_sheets_data()
    x = input('Press Enter or close this window...')


