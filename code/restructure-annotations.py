'''
The file takes excel sheet, pre-processes the annotations and saves it as a csv
In the pre-processing, we convert the strings of grounding act and CGU ids to lists, 
we add a column for CGUs which are initiated during the utterance and CGUs which are closed during the utterance. 
We also add a final column containing the degree of grounding i.e. low, medium and high.
@author - biswesh.mohapatra@inria.fr
'''

import openpyxl
import csv

if __name__ == "__main__":

    for idx in range(0,1):
        wb = openpyxl.load_workbook("../data/annotated_dialogs/dial_"+str(idx)+".xlsx")
        sheet = wb.active

        csv_list = [] 
        # Iterate through rows
        for row in sheet.iter_rows(values_only=True):
            # Convert tuple to list
            row = list(row)

            grounding_act_column = row[6]

            # Add titles to the three new columns
            if grounding_act_column == 'Grounding Act':
                row.append('Opened CGUs')
                row.append('Closed CGUs')
                row.append('Degree')

            print(row)

            # If grounding act is not empty then split and process them
            if grounding_act_column != 'None' and grounding_act_column != 'Grounding Act' and grounding_act_column is not None:
                current_ids = [int(float(i.strip())) for i in str(row[5]).split(',')]
                current_acts = [a.strip() for a in row[6].split(',')]

                # Keep track of CGUs that started and closed in current utterance
                initiate_id_list = []
                close_id_list = []
                degree_of_grounding = []

                for i in range(len(current_ids)):
                    if current_acts[i] == 'Initiate':
                        initiate_id_list.append(current_ids[i])

                    if current_acts[i] == 'Use':
                        close_id_list.append(current_ids[i]) 
                        degree_of_grounding.append('medium')

                    if current_acts[i] == 'Move':
                        close_id_list.append(current_ids[i]) 
                        degree_of_grounding.append('low')

                    if current_acts[i] == 'Explicit-Ack':
                        close_id_list.append(current_ids[i]) 
                        degree_of_grounding.append('medium')

                row[5] = current_ids
                row[6] = current_acts
                row.append(initiate_id_list)
                row.append(close_id_list)
                row.append(degree_of_grounding)

            print(row)

            # Add row only if time and utterance columns are not empty
            if row[1] is not None and row[2] is not None:
                csv_list.append(row)
        
        with open("../data/final_annotated_dialogs/dial_corrected_"+str(idx)+".csv", "w") as f:
            writer = csv.writer(f)
            writer.writerows(csv_list)
        


    


           





                
