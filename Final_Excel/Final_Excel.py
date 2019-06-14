import xlwings as xw
import pandas as pd


"""
Follow the steps:
-----------------
- parse the sheet1 from excel into df1
- parse the sheet2 from excel into df2
- take df1's 1st row as array
- take df2's 1st row as array
- compare(df1, df2)
    + if (equal) ==> same in dfnew
    + if (unequal) ==> keep default value in dfnew
    + if (blank) ==> keep blank in dfnew
"""


#===========================================================INPUTS==============================
sht_name_data = '06122019'
sht_name_default = 'Default'
# sht_data_columns = ["Potato", "Onion", "Pepper", 
#                     "Apple", "Apricot", "Cherry", "Clementine", "Fig", "Grape", "Guava", "Mango", "Melon", "Nectarine", 
#                     "Avacado", "Cucumber", "Tomato", "Carrots", "Asparagus", "Cabbage"]

excel_file_directory = "I:\\github_repos\\My_Freelancing_Projects\\work-4\\data\\06122019.xlsx"


#================================================================MAIN================================================================
def main():
    # wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # for testing

    #------------------------------------------------------------------------------------------------
    # Sheets
    # sht_main = wb.sheets['Main']
    excel_file = pd.ExcelFile(excel_file_directory)


    #------------------------------------------------------------------------------------------------
    # dataframe- 'df_default'
    df_default = excel_file.parse(sht_name_default, skiprows= 0, header= None)      # Fetch without any column header
    df_default.columns = ['sub_category', 'value']
    # sht_main.clear()   # Clear the content and formatting before displaying the data
    # sht_main.range('A1').options(index=False).value = df_default        # show the dataframe (df_default) values into sheet- 'Main'


    #------------------------------------------------------------------------------------------------
    # dataframe- 'df1'
    df1 = excel_file.parse(sht_name_data, skiprows= 1)
    sht_data_columns = df_default['sub_category'].tolist()
    df1 = df1[sht_data_columns]
    # print(df1)    # display 'df1'(float by default)
    df1.fillna(0, inplace=True)     # replace 'NaN' with zero so as to convert all into integer.
    df1 = df1.astype(int)       # convert 'df1' values (float by default) as int.
    print(df1)    # display 'df1'(all converted into int)
    # sht_main.clear()   # Clear the content and formatting before displaying the data
    # sht_main.range('A1').options(index=False).value = df1        # show the dataframe (df1) values into sheet- 'Main'

    #------------------------------------------------------------------------------------------------
    # compare dataframes
    df1_column_headers = df1.columns.tolist()
    df1_indices = df1.index.tolist()
    df_default_subcategory = df_default['sub_category'].tolist()
    df_default_value = df_default['value'].tolist()
    len_subcategory = len(df1_column_headers)   # calculate the len of df1_column_headers list
    # print(len_subcategory)
    len_index = len(df1_indices)
    # print(len_index)
    df1_index0 = df1.iloc[0].tolist()


    if (df1_column_headers == df_default_subcategory) : 
        print("Equal sub-categories in both the sheets!")     # check if both sheets contain the same sub-categories.
        # Fetch list of df1_index0_asint
        # Fetch list of df_default_value
        # Loop in a row of df1
        for i in range(len_index):               # i= [0, 1, 2,...., 293]
            for j in range(len_subcategory):    # j= [0, 1, 2,...., 18]
                if df1.iloc[i,j] != df_default_value[j]:
                    df1.iloc[i,j] = df_default_value[j]
                elif df1.iloc[i,j] == 0:
                    df1.iloc[i,j] = 0

    else:
        print("Sorry! Please correct your sub-categories.")

    #------------------------------------------------------------------------------------------------
    # Display the final dataframe
    # print(df1_column_headers)
    # print(df_default_subcategory)
    # print(df1_index0)
    print(df_default['value'].tolist())
    df1.replace(to_replace= 0, value= 'NaN', inplace=True) 
    print(df1)
    # print(df_default)


# ===============================================================RUN MAIN Function==============================================================
if __name__ == '__main__':
    main()


