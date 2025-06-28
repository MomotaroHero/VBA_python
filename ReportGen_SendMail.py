import os
import pandas as pd
import win32com.client as win32
import datetime
import time
import pythoncom
from pathlib import Path

class ReportGenerator:
    def __init__(self):
        # Initialize configuration variables
        self.in_folder_name = ""
        self.in_file_name = ""
        self.out_folder_name = ""
        self.out_file_name = ""
        
        self.jp_sheet_name = ""
        self.jp_mail_to = ""
        self.jp_mail_bcc = ""
        self.jp_mail_sub = ""
        self.jp_mail_text = ""
        
        self.en_sheet_name = ""
        self.en_mail_to = ""
        self.en_mail_bcc = ""
        self.en_mail_sub = ""
        self.en_mail_text = ""
        
        self.md_folder_name = ""
        self.md_file_name_kix = ""
        self.md_file_name_itm = ""
        self.md_file_name_kobe = ""
        
        self.in_folder_name_pax = ""
        self.in_file_name_pax_kix_int = ""
        self.in_file_name_pax_kix_dom = ""
        self.in_file_name_pax_itm_dom = ""
        self.in_file_name_pax_kobe_dom = ""
        
        self.start_time = ""
        self.mode = ""
        self.outfile_path = ""

        self.mail_to = ""
        self.mail_bcc = ""
        self.mail_sub = ""
        self.mail_text = ""
        
        # Initialize Excel application
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        
        # Load configuration
        self.load_config()
    
    def load_config(self):
        """Load configuration from the Excel file"""
        config_file = os.path.join(os.path.dirname(__file__), "VBA05_20190709_SalesReport_SendMail_ver2.0.xlsm")
        if not os.path.exists(config_file):
            raise FileNotFoundError("Configuration file not found")
        
        df = pd.read_excel(config_file, sheet_name="開始ボタン", header=None)
        
        # Load settings from the Excel file
        self.in_folder_name = df.iloc[9, 3]
        self.in_file_name = df.iloc[10, 3]
        self.out_folder_name = df.iloc[11, 3]
        self.out_file_name = df.iloc[12, 3]
        
        self.jp_sheet_name = df.iloc[14, 3]
        self.jp_mail_to = df.iloc[15, 3]
        self.jp_mail_bcc = df.iloc[16, 3]
        self.jp_mail_sub = df.iloc[17, 3]
        self.jp_mail_text = df.iloc[18, 3]
        
        self.en_sheet_name = df.iloc[20, 3]
        self.en_mail_to = df.iloc[21, 3]
        self.en_mail_bcc = df.iloc[22, 3]
        self.en_mail_sub = df.iloc[23, 3]
        self.en_mail_text = df.iloc[24, 3]
        
        self.md_folder_name = df.iloc[26, 3]
        self.md_file_name_kix = df.iloc[27, 3]
        self.md_file_name_itm = df.iloc[28, 3]
        self.md_file_name_kobe = df.iloc[29, 3]
        
        self.in_folder_name_pax = df.iloc[31, 3]
        self.in_file_name_pax_kix_int = df.iloc[32, 3]
        self.in_file_name_pax_kix_dom = df.iloc[33, 3]
        self.in_file_name_pax_itm_dom = df.iloc[34, 3]
        self.in_file_name_pax_kobe_dom = df.iloc[35, 3]
        
        self.start_time = df.iloc[37, 3]
    
    def control_main(self, mode="solo"):
        """Main control function"""
        self.mode = mode
        
        if mode == "solo":
            self.proc_main()
        elif mode == "sch":
            msg = f"日次レポートの自動送信を開始します。\n翌日 [{self.start_time}] に、自動実行します。"
            # In a real implementation, you'd use a GUI library to show this message
            print(msg)
            self.time_schedule()
    
    def proc_main(self):
        """Main processing function"""
        # Create report
        self.create_report()
        
        # Create PDFs and send emails
        jp_pdf_path = self.create_pdf("JP")
        en_pdf_path = self.create_pdf("EN")
        
        self.send_mail(jp_pdf_path)
        self.send_mail(en_pdf_path)
        
        # Schedule next run if needed
        self.time_reschedule()
    
    def create_report(self):
        """Create the report by combining data from multiple sources"""
        # Open report file
        rp_file_path = self.edit_file_path(self.in_folder_name, self.in_file_name)
        rp_wb = self.excel.Workbooks.Open(rp_file_path, UpdateLinks=0)
        
        # Open master data files
        md_kix_path = self.edit_file_path(self.md_folder_name, self.md_file_name_kix)
        md_kix_wb = self.excel.Workbooks.Open(md_kix_path)
        
        md_itm_path = self.edit_file_path(self.md_folder_name, self.md_file_name_itm)
        md_itm_wb = self.excel.Workbooks.Open(md_itm_path)
        
        md_kobe_path = self.edit_file_path(self.md_folder_name, self.md_file_name_kobe)
        md_kobe_wb = self.excel.Workbooks.Open(md_kobe_path)
        
        print("start copy sheet data")
        # Process KIX data
        self.copy_sheet_data(md_kix_wb, "レポート", rp_wb, "レポート（KIX)")
        self.copy_sheet_data(md_kix_wb, "オペ", rp_wb, "レポート（ITM)")
        self.copy_sheet_data(md_kix_wb, "上下分離", rp_wb, "レポート（UKB)")

        print("end copy sheet data")
        
        # Process ITM data
        self.copy_sheet_data(md_itm_wb, "レポート", rp_wb, "レポート(ITM)")
        self.copy_sheet_data(md_itm_wb, "オペ", rp_wb, "オペ(ITM)")
        self.copy_sheet_data(md_itm_wb, "上下分離", rp_wb, "上下分離(ITM)")
        
        # Process KOBE data
        self.copy_sheet_data(md_kobe_wb, "レポート", rp_wb, "レポート(UKB)")
        self.copy_sheet_data(md_kobe_wb, "オペ", rp_wb, "オペ(UKB)")
        self.copy_sheet_data(md_kobe_wb, "上下分離", rp_wb, "上下分離(UKB)")
        
        # Calculate and save
        rp_wb.Sheets(self.jp_sheet_name).Calculate()
        rp_wb.Sheets(self.en_sheet_name).Calculate()
        rp_wb.Save()
        
        # Close master files
        md_kix_wb.Close()
        # md_itm_wb.Close()
        # md_kobe_wb.Close()
        
        # Process PAX data
        pax_kix_int_path = self.edit_file_path(self.in_folder_name_pax, self.in_file_name_pax_kix_int)
        pax_kix_int_wb = self.excel.Workbooks.Open(pax_kix_int_path, UpdateLinks=0)
        
        pax_kix_dom_path = self.edit_file_path(self.in_folder_name_pax, self.in_file_name_pax_kix_dom)
        pax_kix_dom_wb = self.excel.Workbooks.Open(pax_kix_dom_path, UpdateLinks=0)
        
        pax_itm_dom_path = self.edit_file_path(self.in_folder_name_pax, self.in_file_name_pax_itm_dom)
        pax_itm_dom_wb = self.excel.Workbooks.Open(pax_itm_dom_path, UpdateLinks=0)
        
        pax_kobe_dom_path = self.edit_file_path(self.in_folder_name_pax, self.in_file_name_pax_kobe_dom)
        pax_kobe_dom_wb = self.excel.Workbooks.Open(pax_kobe_dom_path, UpdateLinks=0)
        
        # Copy PAX data
        self.copy_pax_data(pax_kix_int_wb, "サマリー", rp_wb, self.jp_sheet_name, "C17", "G10")  # KIX Int MTD
        self.copy_pax_data(pax_kix_int_wb, "サマリー", rp_wb, self.jp_sheet_name, "C18", "G11")  # KIX Int YTD
        
        self.copy_pax_data(pax_kix_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "D17", "G12")  # KIX Dom MTD
        self.copy_pax_data(pax_kix_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "D18", "G13")  # KIX Dom YTD
        
        self.copy_pax_data(pax_itm_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "E17", "W10")  # ITM Dom MTD
        self.copy_pax_data(pax_itm_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "E18", "W12")  # ITM Dom YTD
        
        self.copy_pax_data(pax_kobe_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "F17", "W10")  # KOBE Dom MTD
        self.copy_pax_data(pax_kobe_dom_wb, "サマリー", rp_wb, self.jp_sheet_name, "F18", "W12")  # KOBE Dom YTD
        
        # Calculate and save
        rp_wb.Sheets(self.jp_sheet_name).Calculate()
        rp_wb.Sheets(self.en_sheet_name).Calculate()
        rp_wb.Save()
        
        # Close PAX files
        pax_kix_int_wb.Close()
        pax_kix_dom_wb.Close()
        pax_itm_dom_wb.Close()
        pax_kobe_dom_wb.Close()
        rp_wb.Close()
    
    def copy_sheet_data(self, src_wb, src_sheet_name, dest_wb, dest_sheet_name):
        """Copy data from one sheet to another with comprehensive error handling"""
        try:
            # Debug: Print available sheets
            print(f"\nSource workbook sheets: {[sh.Name for sh in src_wb.Sheets]}")
            print(f"Destination workbook sheets: {[sh.Name for sh in dest_wb.Sheets]}")
            
            # Verify source sheet exists
            try:
                src_sheet = src_wb.Sheets(src_sheet_name)
            except Exception as e:
                available_sheets = [sh.Name for sh in src_wb.Sheets]
                raise Exception(
                    f"Source sheet '{src_sheet_name}' not found. Available sheets: {', '.join(available_sheets)}"
                ) from e

            src_sheet.Calculate()
            src_sheet.Activate()  # Sometimes needed for COM operations
            
            # Find used range
            used_range = src_sheet.UsedRange
            last_row = used_range.Find("*", 
                                    SearchOrder=win32.constants.xlByRows,
                                    SearchDirection=win32.constants.xlPrevious).Row
            last_col = used_range.Find("*",
                                    SearchOrder=win32.constants.xlByColumns,
                                    SearchDirection=win32.constants.xlPrevious).Column
            
            # Verify destination sheet exists
            try:
                dest_wb.Activate()  # Activate workbook first
                dest_sheet = dest_wb.Sheets(dest_sheet_name)
                dest_sheet.Activate()
            except Exception as e:
                available_sheets = [sh.Name for sh in dest_wb.Sheets]
                raise Exception(
                    f"Destination sheet '{dest_sheet_name}' not found. Available sheets: {', '.join(available_sheets)}"
                ) from e

            # Copy data
            src_range = src_sheet.Range(src_sheet.Cells(1, 1), src_sheet.Cells(last_row, last_col))
            src_range.Copy()
            
            # Small delay to ensure copy completes
            time.sleep(0.1)
            
            # Paste to destination
            dest_sheet.Range("A1").Select()  # Sometimes needed before paste
            try:
                # Try the simple approach first
                dest_sheet.PasteSpecial(win32.constants.xlPasteValuesAndNumberFormats)
            except:
                # Fallback to more explicit approach if needed
                dest_sheet.PasteSpecial(Format=win32.constants.xlPasteValuesAndNumberFormats,
                                    Link=False,
                                    DisplayAsIcon=False)
            
            # Clear clipboard
            self.excel.CutCopyMode = False
            time.sleep(0.1)  # Small delay
            
            print(f"Successfully copied data from '{src_sheet_name}' to '{dest_sheet_name}'")
            
        except Exception as e:
            print(f"\nERROR in copy_sheet_data:")
            print(f"Source: {src_sheet_name}")
            print(f"Destination: {dest_sheet_name}")
            print(f"Error details: {str(e)}")
            raise
    
    def copy_pax_data(self, src_wb, src_sheet_name, dest_wb, dest_sheet_name, src_cell, dest_cell):
        """Copy specific PAX data from one sheet to another with robust error handling"""
        try:
            # Debug info
            print(f"\nCopying PAX data from {src_sheet_name}[{src_cell}] to {dest_sheet_name}[{dest_cell}]")
            
            # Verify source sheet exists
            try:
                src_sheet = src_wb.Sheets(src_sheet_name)
                src_sheet.Calculate()
            except Exception as e:
                available_sheets = [sh.Name for sh in src_wb.Sheets]
                raise Exception(
                    f"Source sheet '{src_sheet_name}' not found. Available sheets: {', '.join(available_sheets)}"
                ) from e

            # Verify destination sheet exists
            try:
                dest_wb.Activate()  # Activate workbook first
                dest_sheet = dest_wb.Sheets(dest_sheet_name)
            except Exception as e:
                available_sheets = [sh.Name for sh in dest_wb.Sheets]
                raise Exception(
                    f"Destination sheet '{dest_sheet_name}' not found. Available sheets: {', '.join(available_sheets)}"
                ) from e

            # Copy data from source
            try:
                src_sheet.Range(src_cell).Copy()
                time.sleep(0.1)  # Small delay to ensure copy completes
            except Exception as e:
                raise Exception(f"Failed to copy from cell {src_cell} in sheet {src_sheet_name}") from e

            # Paste to destination
            try:
                dest_sheet.Activate()
                dest_range = dest_sheet.Range(dest_cell)
                dest_range.Select()
                
                # Try different PasteSpecial approaches
                try:
                    # First try - simple approach
                    dest_sheet.PasteSpecial(win32.constants.xlPasteValuesAndNumberFormats)
                except:
                    # Fallback - more explicit approach
                    dest_sheet.PasteSpecial(Format=win32.constants.xlPasteValuesAndNumberFormats,
                                        Link=False,
                                        DisplayAsIcon=False)
                
                time.sleep(0.1)  # Small delay
            except Exception as e:
                raise Exception(f"Failed to paste to cell {dest_cell} in sheet {dest_sheet_name}") from e

            # Clear clipboard
            self.excel.CutCopyMode = False
            
            print("PAX data copied successfully")
            
        except Exception as e:
            print(f"\nERROR in copy_pax_data:")
            print(f"Source: {src_sheet_name}[{src_cell}]")
            print(f"Destination: {dest_sheet_name}[{dest_cell}]")
            print(f"Error details: {str(e)}")
            raise
    
    def create_pdf(self, lang_mode):
        """Create PDF from the report"""
        # Set language configuration
        self.set_lang_config(lang_mode)
        
        # Open report file
        report_path = self.edit_file_path(self.out_folder_name, self.out_file_name)
        report_wb = self.excel.Workbooks.Open(report_path)
        print(f"open report file: {report_path}")
        print(f"sheet_name: {self.sheet_name}")
        # Activate the sheet
        sheet = report_wb.Sheets(self.sheet_name)
        sheet.Activate()
        
        print(f"activate sheet: {self.sheet_name}")
        # Create PDF path
        pdf_path = os.path.join(self.out_folder_name, f"{self.mail_sub}_{lang_mode}.pdf")
        
        print(f"pdf_path: {pdf_path}")
        # Export to PDF
        sheet.ExportAsFixedFormat(
            Type=win32.constants.xlTypePDF,
            Filename=pdf_path,
            Quality=win32.constants.xlQualityStandard,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        print(f"export to pdf: {pdf_path}")
        # Close report
        report_wb.Close(SaveChanges=False)
        print(f"close report file: {report_path}")
        return pdf_path
    
    def send_mail(self, pdf_path):
        """Send email with PDF attachment with robust error handling"""
        try:
            # Verify PDF file exists first
            if not os.path.exists(pdf_path):
                raise FileNotFoundError(f"PDF file not found at: {pdf_path}")
            
            # Get absolute path to ensure reliability
            abs_pdf_path = os.path.abspath(pdf_path)
            print(f"Preparing to attach PDF: {abs_pdf_path}")

            # Initialize Outlook
            pythoncom.CoInitialize()
            try:
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                
                # Set email properties
                mail.To = self.mail_to
                mail.BCC = self.mail_bcc
                mail.Subject = self.mail_sub
                mail.Body = self.mail_text + "\n\n"
                
                print(f"mail_to: {self.mail_to}")
                print(f"mail_bcc: {self.mail_bcc}")
                print(f"mail_sub: {self.mail_sub}")
                print(f"mail_text: {self.mail_text}")
                
                # Add attachment with verification
                print(f"Attempting to attach: {abs_pdf_path}")
                if os.path.exists(abs_pdf_path):
                    mail.Attachments.Add(abs_pdf_path)
                    print("Attachment added successfully")
                else:
                    raise FileNotFoundError(f"Attachment file disappeared: {abs_pdf_path}")
                
                # Send email
                print("Sending email...")
                try:
                    mail.Send()
                    print("Email sent successfully")
                except Exception as e:
                    print(f"Outlook error: {str(e)}")
                    raise
                
            except Exception as e:
                print(f"Outlook error: {str(e)}")
                raise
            finally:
                # Cleanup Outlook objects
                try:
                    if 'mail' in locals():
                        del mail
                    if 'outlook' in locals():
                        del outlook
                except:
                    pass
                pythoncom.CoUninitialize()
                
        except Exception as e:
            print(f"Failed to send email: {str(e)}")
            raise
    
    def set_lang_config(self, lang_mode):
        """Set language-specific configuration"""
        if lang_mode == "JP":
            self.sheet_name = self.jp_sheet_name
            self.mail_to = self.jp_mail_to
            self.mail_bcc = self.jp_mail_bcc
            self.mail_sub = self.jp_mail_sub
            self.mail_text = self.jp_mail_text
        elif lang_mode == "EN":
            self.sheet_name = self.en_sheet_name
            self.mail_to = self.en_mail_to
            self.mail_bcc = self.en_mail_bcc
            self.mail_sub = self.en_mail_sub
            self.mail_text = self.en_mail_text
    
    def edit_file_path(self, folder_name, file_name):
        """Create proper file path from folder and file names"""
        if folder_name.endswith("\\"):
            return folder_name + file_name
        else:
            return folder_name + "\\" + file_name
    
    def time_reschedule(self):
        """Schedule next run if needed"""
        if self.mode == "sch":
            self.time_schedule()
        else:
            self.mode = "sch"
    
    def time_schedule(self):
        """Schedule the next run"""
        # This is a simplified version - in a real implementation, you'd use Windows Task Scheduler
        # or a proper scheduling library like APScheduler
        
        # Calculate time until next run
        now = datetime.datetime.now()
        scheduled_time = datetime.datetime.strptime(self.start_time, "%H:%M:%S").time()
        scheduled_datetime = datetime.datetime.combine(now.date(), scheduled_time)
        
        # If scheduled time is in the past, schedule for next day
        if scheduled_datetime < now:
            scheduled_datetime += datetime.timedelta(days=1)
        
        # Calculate seconds until next run
        delta = scheduled_datetime - now
        seconds_until_run = delta.total_seconds()
        
        # Schedule the next run (simplified - in reality use a proper scheduler)
        print(f"Next run scheduled for {scheduled_datetime}")
        time.sleep(seconds_until_run)
        self.proc_main()
    
    def close(self):
        """Explicit cleanup method that should be called when done"""
        if hasattr(self, 'excel') and self.excel is not None:
            try:
                # Try to restore Excel settings
                self.excel.DisplayAlerts = True
                self.excel.ScreenUpdating = True
            except:
                pass  # Excel might already be in a bad state
            
            try:
                # Try to quit Excel
                self.excel.Quit()
            except:
                pass  # Might already be closed
            
            # Release the COM object
            del self.excel
            self.excel = None

    def __del__(self):
        """Backup cleanup in case close() wasn't called"""
        try:
            self.close()
        except:
            pass  # Prevent exceptions during garbage collection


