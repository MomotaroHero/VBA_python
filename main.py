from DailyReport import DailyReport
from ReportGen_SendMail import ReportGenerator

# Example usage
if __name__ == "__main__":
    # report_gen = ReportGenerator()
    # report_gen.control_main(mode="solo")  # Run immediately
    # report_gen.control_main(mode="sch")  # Run on schedule
    
    report_gen = None
    daily_report = None
    try:        
        # daily_report = DailyReport()
        # daily_report.control_main(mode="solo")
        report_gen = ReportGenerator()
        report_gen.control_main(mode="solo")
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        if report_gen is not None:
            report_gen.close()
        if daily_report is not None:
            daily_report.close()

