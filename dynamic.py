import pandas as pd
df = pd.read_excel('Pandas Task.xlsx',index_col=None)
df['Result'] = ''
Column_length= len(df.columns)-3
Cell_counter_1=0
Value_list = df['Value'].tolist()
Value_list = ['NA' if x != x else x for x in Value_list]
for Col in range(Column_length,0,-1):
    Level_list=df['LEVEL-{}'.format(Col)].tolist()
    Level_list = ['NA' if x != x else x for x in Level_list]
    if Col==Column_length:
        for Cell in Level_list:
            if Cell != 'NA':
                position_count_1=Cell_counter_1
                Level_result=Value_list[position_count_1]
                df.at[position_count_1,'Result']=Level_result
            Cell_counter_1+=1
    elif Col==Column_length-1:
        Sum_1=Level_sum_1=Cell_counter_2=0
        for Cell in Level_list:
            if Cell != 'NA':
                position_count_2=Cell_counter_2
                Sum_1+=Value_list[position_count_2]
                i=1
                while True:
                    if len(Level_list)-1!='NA':
                        lvl_previous_list.append('NA')
                    if lvl_previous_list[position_count_2+i]!='NA':
                        Sum_1+=Value_list[position_count_2+i]
                        i+=1
                    else:
                        Level_result=Sum_1
                        df.at[position_count_2,'Result']=Level_result
                        Level_sum_1+=Level_result 
                        Sum_1=0              
                        break
            Cell_counter_2+=1
    else:
        Sum_2=Level_sum_2=position_count_3=Bypass=Cell_counter_3=0
        for Cell in Level_list:
            if (Level_list[len(Level_list)-1] != 'NA') and Bypass==0:
                lvl_previous_list.append('NA')
                Bypass=1
            if Cell != 'NA':
                position_count_3=Cell_counter_3
                Sum_2=0
                Sum_2+=Value_list[position_count_3]
                if lvl_previous_list[position_count_3+1]!='NA':
                    Sum_2+=Level_sum_1
                    Level_result=Sum_2
                    df.at[position_count_3,'Result']=Level_result
                    Level_sum_2+=Level_result 
                else:
                    Level_result=Sum_2
                    df.at[position_count_3,'Result']=Level_result
                    Level_sum_2+=Level_result 
            flag1=0           
            Cell_counter_3+=1
        Level_sum_1=Level_sum_2
        Sum_2=0 
    lvl_previous_list=Level_list
(df.style.highlight_max(axis=0, props='background-color:green;', subset=['Result'])
        .highlight_min(axis=0, props='background-color:red;', subset=['Result'])
        .to_excel('pandas_task_output.xlsx', engine='openpyxl'))