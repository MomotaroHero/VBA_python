import os
import win32com.client as win32
import time
import datetime
import pandas as pd
from pathlib import Path

class DailyReport:
    def __init__(self):
        # Initialize configuration variables
        self.out_folder_name = ""
        self.out_file_name = ""
        self.in_folder_name_dwh = ""
        self.in_file_name_dwh = ""
        self.in_folder_name_pos1 = ""
        self.in_file_name_pos1 = ""
        self.in_folder_name_pos2 = ""
        self.in_file_name_pos2 = ""
        
        self.in_folder_mallpro_name_kix = ""
        self.in_file_mallpro_name_kix = ""
        self.in_folder_mallpro_name_itm = ""
        self.in_file_mallpro_name_itm = ""
        self.in_folder_mallpro_name_kobe = ""
        self.in_file_mallpro_name_kobe = ""
        
        self.md_folder_name = ""
        self.md_folder_name_powerbi = ""
        self.md_file_name_kix = ""
        self.md_file_name_itm = ""
        self.md_file_name_kobe = ""
        
        self.in_folder_name_pax = ""
        self.in_file_name_pax_kix_int = ""
        self.in_file_name_pax_kix_dom = ""
        self.in_file_name_pax_itm_dom = ""
        self.in_file_name_pax_kobe_dom = ""
        
        self.out_folder_name_pax = ""
        self.out_file_name_pax_kix_int = ""
        self.out_file_name_pax_kix_dom = ""
        self.out_file_name_pax_itm_dom = ""
        self.out_file_name_pax_kobe_dom = ""
        
        self.cur_folder_name = ""
        self.start_time = ""
        self.mode = ""
        self.outfile_path = ""
        
        # Initialize Excel application
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        
        # Load configuration
        self.load_config()
    
    def load_config(self):
        """Load configuration from the Excel file"""
        config_file = os.path.join(os.path.dirname(__file__), "VBA02_20220528_SalesReportDaily_ver5.0.xlsm")
        if not os.path.exists(config_file):
            raise FileNotFoundError("Configuration file not found")
        
        # Using pandas to read Excel (alternative to win32com)
        df = pd.read_excel(config_file, sheet_name="開始ボタン", header=None)
        
        # Load settings from the Excel file
        self.in_folder_mallpro_name_kix = df.iloc[25, 3]
        self.in_file_mallpro_name_kix = df.iloc[26, 3]
        self.in_folder_mallpro_name_itm = df.iloc[27, 3]
        self.in_file_mallpro_name_itm = df.iloc[28, 3]
        self.in_folder_mallpro_name_kobe = df.iloc[29, 3]
        self.in_file_mallpro_name_kobe = df.iloc[30, 3]
        
        self.md_folder_name = df.iloc[32, 3]
        self.md_folder_name_powerbi = df.iloc[32, 6]
        self.md_file_name_kix = df.iloc[33, 3]
        self.md_file_name_itm = df.iloc[34, 3]
        self.md_file_name_kobe = df.iloc[35, 3]
        
        self.in_folder_name_pax = df.iloc[37, 3]
        self.in_file_name_pax_kix_int = df.iloc[38, 3]
        self.in_file_name_pax_kix_dom = df.iloc[39, 3]
        
        self.out_folder_name_pax = df.iloc[41, 3]
        self.out_file_name_pax_kix_int = df.iloc[42, 3]
        self.out_file_name_pax_kix_dom = df.iloc[43, 3]
        self.out_file_name_pax_itm_dom = df.iloc[44, 3]
        self.out_file_name_pax_kobe_dom = df.iloc[45, 3]
        
        self.start_time = df.iloc[47, 3]
    
    def control_main(self, mode="solo"):
        """Main control function"""
        self.mode = mode
        
        if mode == "solo":
            self.edit_main()
        elif mode == "sch":
            msg = f"日次レポートの自動送信を開始します。\n翌日 [{self.start_time}] に、自動実行します。"
            print(msg)
            self.time_schedule()
    
    def edit_main(self):
        """Main processing function"""
        # Process mallpro files
        self.edit_mallpro("KIX")
        self.edit_mallpro("ITM")
        self.edit_mallpro("KOBE")
        
        # Process PowerBI files
        self.edit_mallpro_powerbi("KIX")
        self.edit_mallpro_powerbi("ITM")
        self.edit_mallpro_powerbi("KOBE")
        
        # Process PAX data
        self.edit_multiple_pax()
        
        # Schedule next run if needed
        if self.mode == "sch":
            self.time_schedule()
        else:
            self.mode = "sch"
    
    def edit_mallpro(self, ap_code):
        """Process mallpro data"""
        try:
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            
            # Set file paths based on airport code
            if ap_code == "KIX":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_kix, self.in_file_mallpro_name_kix)
                out_path = self.edit_file_path(self.md_folder_name, self.md_file_name_kix)
                sheet_name = "KIX貼付"
            elif ap_code == "ITM":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_itm, self.in_file_mallpro_name_itm)
                out_path = self.edit_file_path(self.md_folder_name, self.md_file_name_itm)
                sheet_name = "ITM貼付"
            elif ap_code == "KOBE":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_kobe, self.in_file_mallpro_name_kobe)
                out_path = self.edit_file_path(self.md_folder_name, self.md_file_name_kobe)
                sheet_name = "UKB貼付"
            
            # Open files
            in_file = self.excel.Workbooks.Open(in_path)
            out_file = self.excel.Workbooks.Open(out_path)
            
            # Process data
            in_sheet = in_file.Sheets(1)
            out_sheet = out_file.Sheets(sheet_name)
            
            # Get used range
            in_range = in_sheet.UsedRange
            out_range = out_sheet.UsedRange
            
            # Delete existing data in output
            out_sheet.Range(out_sheet.Cells(2, 1), out_sheet.Cells(out_range.Rows.Count, out_range.Columns.Count)).EntireRow.Delete()
            
            print("Copying data mallpro")
            # Copy data
            in_sheet.Range(in_sheet.Cells(2, 1), in_sheet.Cells(in_range.Rows.Count, in_range.Columns.Count)).Copy()
            out_sheet.Range("A2").PasteSpecial(Paste=win32.constants.xlPasteValues)
            
            print("Saving")
            # Cleanup
            in_file.Close(False)
            out_file.Save()
            out_file.Close()
            
        finally:
            print("Restoring settings")
            self.excel.DisplayAlerts = True
            self.excel.ScreenUpdating = True
    
    def edit_mallpro_powerbi(self, ap_code):
        """Process mallpro data for PowerBI"""
        try:
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            
            # Set file paths based on airport code
            if ap_code == "KIX":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_kix, self.in_file_mallpro_name_kix)
                out_path = self.edit_file_path(self.md_folder_name_powerbi, self.md_file_name_kix)
                sheet_name = "KIX貼付"
            elif ap_code == "ITM":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_itm, self.in_file_mallpro_name_itm)
                out_path = self.edit_file_path(self.md_folder_name_powerbi, self.md_file_name_itm)
                sheet_name = "ITM貼付"
            elif ap_code == "KOBE":
                in_path = self.edit_file_path(self.in_folder_mallpro_name_kobe, self.in_file_mallpro_name_kobe)
                out_path = self.edit_file_path(self.md_folder_name_powerbi, self.md_file_name_kobe)
                sheet_name = "UKB貼付"
            
            # Open files
            in_file = self.excel.Workbooks.Open(in_path)
            out_file = self.excel.Workbooks.Open(out_path)
            
            # Process data
            in_sheet = in_file.Sheets(1)
            out_sheet = out_file.Sheets(sheet_name)
            
            # Get used range
            in_range = in_sheet.UsedRange
            out_range = out_sheet.UsedRange
            
            # Delete existing data in output
            out_sheet.Range(out_sheet.Cells(2, 1), out_sheet.Cells(out_range.Rows.Count, out_range.Columns.Count)).EntireRow.Delete()
            
            print("Copying data powerBI")
            # Copy data
            in_sheet.Range(in_sheet.Cells(2, 1), in_sheet.Cells(in_range.Rows.Count, in_range.Columns.Count)).Copy()
            out_sheet.Range("A2").PasteSpecial(Paste=win32.constants.xlPasteValues)
            
            print("Saving")
            # Cleanup
            in_file.Close(False)
            out_file.Save()
            out_file.Close()
            
        finally:
            print("Restoring settings")
            self.excel.DisplayAlerts = True
            self.excel.ScreenUpdating = True
    
    def edit_multiple_pax(self):
        """Process multiple PAX files"""
        # Implementation for processing PAX files
        pass
    
    def edit_file_path(self, folder_name, file_name):
        """Create proper file path from folder and file names"""
        return os.path.join(folder_name, file_name)
    
    def time_schedule(self):
        """Schedule the next run"""
        # Calculate next run time
        now = datetime.datetime.now()
        scheduled_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
        scheduled_datetime = datetime.datetime.combine(now.date(), scheduled_time)
        
        # If scheduled time is in the past, schedule for next day
        if scheduled_datetime < now:
            scheduled_datetime += datetime.timedelta(days=1)
        
        # Calculate seconds until next run
        delta = scheduled_datetime - now
        seconds_until_run = delta.total_seconds()
        
        # Schedule the next run
        print(f"Next run scheduled for {scheduled_datetime}")
        time.sleep(seconds_until_run)
        self.edit_main()
    
    def __del__(self):
        """Cleanup when object is destroyed"""
        if hasattr(self, 'excel'):
            self.excel.DisplayAlerts = True
            self.excel.ScreenUpdating = True
            self.excel.Quit()
            del self.excel

