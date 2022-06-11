import openpyxl 
import pandas as pd
import string
import os
import logging

from utils.config import directory
from utils.checkreport_month import wantedFileName

logger = logging.getLogger(__name__)

dirr  = directory()
#----------------- Data Directory--------------------------------
rawdatafilename, report_monthyear, reportmonth_datetime = wantedFileName() # Choose latest month

allocation_file = os.path.join(dirr.raw_dir, rawdatafilename)
workbook = openpyxl.load_workbook(allocation_file)

first_day = reportmonth_datetime.replace(day=1)
first_day_ED = first_day.strftime("%d/%m/%Y")
report_month = first_day.strftime("%b'%y")

#--------------Lookup File Directory---------------------
lookupfilename="Lookupfile.xlsx"
lookup_file = os.path.join(dirr.lookup_fileDir, lookupfilename)

#---------lookup reading wb---------
wblookup = pd.ExcelFile(lookup_file)
lookup_typeMat = pd.read_excel(wblookup, sheet_name='Material')
lookup_typeFac = pd.read_excel(wblookup, sheet_name='Factory')
lookup_typeComp = pd.read_excel(wblookup, sheet_name='Vendor')
lookup_typeComp['vendor'] = lookup_typeComp['vendor'].astype('Int64')


def excel_run(visible = True):
    overall = []
    # to find unhide sheet only
    for i in workbook.worksheets:
    
        if i.sheet_state == "visible":
            visible_sheet = i.title
     
            logger.info(f"Visible sheet is {visible_sheet}")
        
            worksheet = workbook[visible_sheet]

            hidden_rows_idx = [
                row - 2 
                for row, dimension in worksheet.row_dimensions.items() 
                if dimension.hidden
            ]

            # Read Excel file as Pandas DataFrame
            df = pd.read_excel(allocation_file,sheet_name=visible_sheet)

            # Open an Excel workbook
            #workbook = openpyxl.load_workbook(allocation_file)

            # List of indices corresponding to all hidden columns
            hidden_cols_idx = [
                string.ascii_uppercase.index(col_name) 
                for col_name in [
                    col 
                    for col, dimension in worksheet.column_dimensions.items() 
                    if dimension.hidden
                ] 
            ]

            # Find names of columns corresponding to hidden column indices
            hidden_cols_name = df.columns[hidden_cols_idx].tolist()

            # Drop the hidden columns and rows
            df.drop(hidden_cols_name, axis=1, inplace=True)
            df.drop(hidden_rows_idx, axis=0, inplace=True)

            # Reset the index
            df.reset_index(drop=True, inplace=True)

            facCol = df['Unnamed: 1']
            factory_list = ['FACTORY 5/23','FACTORY 36','FACTORY 27','FACTORY 33']
            facIndex = facCol[facCol.isin(factory_list)].index
            #print(facIndex)

            for x in range (0,5,1):
                for n, index in enumerate(facIndex[x: x+1]):
           
                    start = facIndex[x]
                    try:
                        end = facIndex[x+1]
                    except IndexError:
                        end = 200
                    
                    
                    logger.info(f"The sheet: {visible_sheet} start from: {start} and end {end}")
                    allocation = df.iloc[start:end, :]
                    Company_col = allocation.where(allocation=='Company Name').dropna(how='all').dropna(axis=1) # Locate the column of the Company Name
                    col=Company_col.columns[0]

                    report_col = allocation.where(allocation == report_month).dropna(how='all').dropna(axis=1) #locate the month of the column 
                    col2 = report_col.columns[0]

                    supply_col = allocation.where(allocation=='Initial Supply Capacity ').dropna(how='all').dropna(axis=1)
                    col3 = supply_col.columns[0]
                    df1 = allocation[[col,col2,col3]]
                    df1 = df1.dropna(subset=[col])
                
                    #print(df1)
                    df1['Factory'] = df1.iloc[0][col]
                    df1 =df1.dropna(subset= ['Unnamed: 5'])
                    df1 = df1[df1['Unnamed: 5'] > 0]
                    df1['Material Description'] = visible_sheet
                    df1['effective_date[dd/mm/yyyy]'] = first_day_ED

                    df1 = pd.merge(df1, lookup_typeMat, left_on='Material Description',
                    right_on='description', how='left')
                    df1 = pd.merge(df1, lookup_typeFac, left_on='Factory',
                                right_on='FACTORY',how='left')
                    df1 = pd.merge(df1, lookup_typeComp, left_on=col,
                    right_on='Company', how='left')

                                        
                    # To break for those vendor that is not mantained 
                    if df1['vendor_name'].isnull().mean() > 0: 
                        Unmantained_names = df1[df1['vendor_name'].isnull()]["Unnamed: 1"].to_list()
                        raise Exception(f"To maintain {Unmantained_names} in the lookup file")
                                    
                    df1 = df1.drop(columns = ['FACTORY','Factory',col,'Company',
                                        'Material Description',
                                        'description']).rename(columns={
                                            col2: 'price',
                                            col3:'avai_capacity',
                                            'Description':'material_desc'
                                            })  

                    df1 =df1.reindex(['plant','factory','vendor','vendor_name',
                    'material','material_desc','effective_date[dd/mm/yyyy]',
                    'avai_capacity','price'], axis=1)

                    overall.append(df1)

            final_file = pd.concat(overall)

            

            final_reportfilename = 'vendor_capacity'+'_'+ report_monthyear +'.csv'
            final_filename = os.path.join(dirr.report_dir, final_reportfilename)
            
            

            logger.info(f"Saving file as csv in {final_filename}")
            final_file.to_csv(final_filename,index=False)
    

if __name__ == "__main__":
    excel_run(visible = True)

