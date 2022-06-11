import itertools
import logging
import pytz
import os
import sys
import time
import pywintypes

from dateutil import tz
from datetime import timedelta, datetime
import numpy as np
import pandas as pd
from PIL import ImageGrab
import win32com.client as win32; win32c = win32.constants

tablecount = itertools.count(1)

# from tg_pm_allocation.util import rgbToInt

# timezone info
my_tz = tz.gettz('Asia/Kuala_Lumpur')

# for logging import

logger = logging.getLogger(__name__)

'''
Library for win32.com Excel 
'''

class Excelbot():
    def __init__(self, filename=None, visible=True):
        try:  # check if excel is running
            self.excel = win32.GetActiveObject("Excel.Application")
        except:
            self.excel = win32.gencache.EnsureDispatch('Excel.Application')
            logger.info("Dispatch new excel application.")
        else:
            logger.info("Use current active excel application.")

        self.excel.Visible = visible
         # to hide display alert, for ease of running script
        self.excel.DisplayAlerts = False
        self.excel.EnableLargeOperationAlert = False
        self.excel.AskToUpdateLinks = False
        # self.excel.ScreenUpdating

        if filename is None:
            logger.info("Adding new workbook")
            try:
                self.wb = self.excel.Workbooks.Add()
            except:
                logger.error(f"Failed to add workbook")
        else:
            logger.info(f"Open workbook file: {filename}")
            try:
                self.wb = self.excel.Workbooks.Open(filename)
                self.wb.Application.Visible = visible
            except Exception as e:
                logger.error(f"Failed to open spreadsheet [{filename}], error : {e}")

    def get_worksheet(self, sheet_name=None):
        if sheet_name is None:
            raise Exception(f"Please enter valid sheetname: {sheet_name}")
        return self.wb.Sheets(sheet_name)

    def refresh_pivot(self, sheet_name=None):
        if sheet_name is None:
            raise Exception(f"Please enter valid sheetname: {sheet_name}")
        logger.info(f"Refresh pivot on sheet: {sheet_name}.")
        ws = self.wb.Sheets(sheet_name)
        pivotCount = ws.PivotTables().Count
        for j in range(1, pivotCount+1):
            ws.PivotTables(j).PivotCache().Refresh()

    def refresh_all(self):
        logger.info("Refresh all.")
        self.wb.RefreshAll()

    def save(self):
        logger.info("Save and close.")
        self.wb.Save()

    def save_as(self, filepath):
        logger.info(f"Save as filepath: {filepath}.")
        # if int(float(excel.Version)) >= 12:
        self.wb.SaveAs(filepath, win32c.xlOpenXMLWorkbook)
        # else:
        #     self.wb.SaveAs('newABCDCatering.xls')

    def close(self):
        logger.info(f"Quit application.")
        self.excel.Application.Quit()

    def compute_range(self, sheet_name, start_range_cell, end_range_cell):
        worksheet = self.wb.Worksheets(sheet_name)
        start_cell = worksheet.Cells(start_range_cell[0], start_range_cell[1])
        end_cell = worksheet.Cells(end_range_cell[0], end_range_cell[1])
        computedRange = worksheet.Range(start_cell, end_cell)
        return computedRange

    def get_used_range(self, sheet_name=None, option='sheet', include_empty_row = True):
        if sheet_name is None:
            raise Exception(f"Please enter valid {sheet_name}")

        ws = self.wb.Sheets(sheet_name)
        ws_used_range = ws.UsedRange

        if option == "sheet":
            row_count = ws_used_range.Rows.Count
            col_count = ws_used_range.Columns.Count

            if include_empty_row:
                row_count += ws_used_range.Row -1
                col_count += ws_used_range.Column - 1
            return (row_count, col_count)

        elif option == "pivot":
            row_count = ws.PivotTables(1).TableRange2.Rows.Count
            col_count = ws.PivotTables(1).TableRange2.Columns.Count

            if include_empty_row:
                row_count += ws.PivotTables(1).TableRange2.Row - 1
                col_count += ws.PivotTables(1).TableRange2.Column - 1

            # return only the first pivot table
            return (row_count, col_count)

    def all_border(self, sheet_name=None, start_cell_range=(1,1), end_cell_range=(2,2)):
        if sheet_name is None:
            raise Exception(f"Please enter valid {sheet_name}")

        selected_range = self.compute_range(sheet_name, start_cell_range, end_cell_range)

        selected_range.Borders(win32c.xlDiagonalDown).LineStyle = win32c.xlNone
        selected_range.Borders(win32c.xlDiagonalUp).LineStyle = win32c.xlNone
        all_edges = [win32c.xlEdgeLeft, win32c.xlEdgeTop, win32c.xlEdgeBottom, win32c.xlEdgeRight, win32c.xlInsideVertical, win32c.xlInsideHorizontal]
        for edges in all_edges:
            selected_range.Borders(edges).LineStyle = win32c.xlContinuous
            selected_range.Borders(edges).ColorIndex = 0
            selected_range.Borders(edges).TintAndShade = 0
            selected_range.Borders(edges).Weight = win32c.xlThin

    def clear_content(self, sheet_name, start_cell_range, end_cell_range, clear_style=True):
        selected_range = self.compute_range(sheet_name, start_cell_range, end_cell_range)
        selected_range.ClearContents()

        if clear_style:
            pass
        #     selected_range.Interior.Color = rgbToInt((255,255,255))

    def auto_filter(self, source_cell_range_start, source_cell_range_end, dest_cell_range_start, dest_cell_range_end, sheet_name=None):
        if sheet_name is None:
            raise Exception(f"Please enter valid {sheet_name}")

        ws = self.wb.Sheets(sheet_name)

        ws.Range(
            ws.Cells(source_cell_range_start[0], source_cell_range_start[1]),
            ws.Cells(source_cell_range_end[0], source_cell_range_end[1])).AutoFill(ws.Range(
                ws.Cells(dest_cell_range_start[0], dest_cell_range_start[1]),
                ws.Cells(dest_cell_range_end[0],dest_cell_range_end[1])), win32c.xlFillDefault)

    def addcell(self, input_data, sheet_name=None, start_i=2, start_j=4):

        if sheet_name is None:
            worksheet = self.excel.worksheet.Add()
        else:
            worksheet = self.wb.Worksheets(sheet_name)

        input_type = type(input_data)
        if input_type == list:
            for i, row in enumerate(input_data):
                for j, item in enumerate(row):
                    worksheet.Cells(i+start_i, j+start_j).Value = item

        elif input_type == pd.core.frame.DataFrame:
            for row in input_data.itertuples():
                i = row[0]
                for j in range(1, len(row)):
                    worksheet.Cells(i+start_i, j+start_j).Value = row[j]

    def append_dataframe_to_sheet(self, dataframe, sheet_name=None, auto_fill_if_not_found=False, start_row=None):
        """ Convert dataframe into excel cell by cell match by header
        Able to work but slow
        
        """        
        logger.info("Appending dataframe to sheet")

        if sheet_name is None:
            raise Exception("Please enter sheet_name")

        start_time = time.time()

        ws = self.wb.Sheets(sheet_name)
        header = ws.UsedRange.Rows(1).Value[0] # get first row as header, [0] as it will return list of list of 1 item
        total_row_used = ws.UsedRange.Rows.Count
        logger.info(f"Original row used: {total_row_used} ")

        n_col = len(header) # from excel file
        n_row_to_insert = len(dataframe) # from SAP
        logger.info(f"To insert: {n_row_to_insert} rows")

        if start_row is None:
            start_row = total_row_used+1 # insert from new row

        for j in range(n_col):
            try:
                current_header = header[j]
                data_col = dataframe[current_header] # get column of selected header
                logger.info(f"Writing {current_header}")
                
                if (pd.core.dtypes.common.is_datetime_or_timedelta_dtype(data_col) 
                    or isinstance(data_col.dtype, pd.core.dtypes.dtypes.DatetimeTZDtype)): # if date object
                    # in_val = data_col.dt.tz_localize('Asia/Kuala_Lumpur') # convert to local time
                    in_val = data_col + pd.Timedelta("08:00:00") # add 8 hours for GMT
                    in_val = in_val.dt.date.values.tolist()
                    in_val = [[datetime(row.year,row.month,row.day,8, tzinfo=my_tz)] if row is not None and not pd.isnull(row) and row != "" else [""] for row in in_val]
                else: # other object
                    data_col = data_col.fillna("") # ignore null value
                    in_val = data_col.values.reshape(-1,1).tolist()

                ws.Range(ws.Cells(start_row, j+1), # start cell
                    ws.Cells(start_row+len(data_col.index)-1,j+1)).Value = in_val

            except KeyError as e:
                logger.info(f"Skip column {e} as not found in excel.")

                if auto_fill_if_not_found:
                    logger.info(f"Auto fill column {e}")

                    # autofill whole column from cell 2 (exclude header)
                    ws.Range(
                        ws.Cells(2, j+1),
                        ws.Cells(2, j+1)).AutoFill(ws.Range(
                            ws.Cells(2, j+1),
                            ws.Cells(start_row+ n_row_to_insert-1,j+1)), win32c.xlFillDefault)

        total_row_used = ws.UsedRange.Rows.Count
        logger.info(f"After insert, row used: {total_row_used} ")

        end_time = time.time()

        logger.info(f"Total time taken: {(end_time-start_time)} seconds")

    def addpivot(self, sourcedata, title, filters=(), columns=(),
                 rows=(), sumvalue=(), sortfield=""):
        """Build a pivot table using the provided source location data
        and specified fields
        sourcedata(type: compute_range return type)
        """
        newsheet = self.wb.Sheets.Add()
        newsheet.Cells(1, 1).Value = title
        newsheet.Cells(1, 1).Font.Size = 16

        # Build the Pivot Table
        tname = "PivotTable%d" % next(tablecount)

        pc = self.wb.PivotCaches().Add(SourceType=win32c.xlDatabase,
                                       SourceData=sourcedata,
                                       Version=win32c.xlPivotTableVersion14)
        pt = pc.CreatePivotTable(TableDestination="%s!R4C1" % newsheet.Name,
                                 TableName=tname,
                                 DefaultVersion=win32c.xlPivotTableVersion14)
        self.wb.Sheets(newsheet.Name).Select()
        self.wb.Sheets(newsheet.Name).Cells(3, 1).Select()
        for fieldlist, fieldc in ((filters, win32c.xlPageField),
                                  (columns, win32c.xlColumnField),
                                  (rows, win32c.xlRowField)):
            for i, val in enumerate(fieldlist):
                self.wb.ActiveSheet.PivotTables(
                    tname).PivotFields(val).Orientation = fieldc
                self.wb.ActiveSheet.PivotTables(
                    tname).PivotFields(val).Position = i+1

        self.wb.ActiveSheet.PivotTables(tname).AddDataField(
            self.wb.ActiveSheet.PivotTables(tname).PivotFields(sumvalue),
            sumvalue,
            win32c.xlSum)
        if len(sortfield) != 0:
            self.wb.ActiveSheet.PivotTables(tname).PivotFields(
                sortfield[0]).AutoSort(sortfield[1], sumvalue)
        newsheet.Name = title

        # Uncomment the next command to limit output file size, but make sure
        # to click Refresh Data on the PivotTable toolbar to update the table
        # newsheet.PivotTables(tname).SaveData = False

        return tname

    def filter_pivot_item(self, pivot_field, sheet_name=None, show_field = []):
        if sheet_name is None:
            raise Exception(f"sheet_name not found, {sheet_name}")

        ws = self.wb.Sheets(sheet_name)

        for pivot_table in ws.PivotTables():
            pivot_table_str= str(pivot_table)
            for pivot_item in ws.PivotTables(pivot_table_str).PivotFields(pivot_field).PivotItems():
                pivot_item_str = str(pivot_item)
                if pivot_item_str not in show_field:
                    ws.PivotTables(pivot_table_str).PivotFields(pivot_field).PivotItems(pivot_item_str).Visible = False
                else:
                    ws.PivotTables(pivot_table_str).PivotFields(pivot_field).PivotItems(pivot_item_str).Visible = True

    def convert_worksheetrange_to_dataframe(self, sheet_name, start_range_cell, end_range_cell):        
        ws_range = self.compute_range(sheet_name, start_range_cell, end_range_cell).Value
        ws_range = np.array(ws_range)
        df = pd.DataFrame(ws_range)
        df.columns = ws_range[0]
        df = df[1:].reset_index(drop=True)
        return df
    
    def delete_unwantedrows_basedon_column(self, sheet_name, column_alph):
        usedrange = self.get_used_range(sheet_name)
        ws = self.get_worksheet(sheet_name)
        delete_row = ws.Cells(ws.Rows.Count, column_alph).End(win32c.xlUp).Row + 1
        try:
            ws.Range(f"{delete_row}:{usedrange[0]}").Delete()
        except Exception as e:
            logger.info(e)
            pass
        
    def save_image(self, sheet_name, cell_range, file_dir, filename) -> None:    
        """Copy the excel cells to clipboard and save as png image to filepath.
        
        Parameters
        ----------
        sheet : sheet name of excel workbook
        cell_range : win32com excel range, eg:'A1:K26' or worksheet.Range(start_cell, end_cell), getrangemethod.
        file_dir : Desired file directory of the save image
        filename : Disired filename of the image file without .png
        """
        ws = self.get_worksheet(sheet_name)
        ws.Range(cell_range).CopyPicture(Format= win32c.xlBitmap)
        time.sleep(0.3)
        img = ImageGrab.grabclipboard()
        img_path = os.path.join(file_dir, filename + '.png')
        img.save(img_path)

if __name__ == "__main__":
    excelbot = Excelbot()