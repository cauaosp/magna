from utils import read_sheet, schedule_ping_for_printers
import time
import schedule

if __name__ == "__main__":
    excel_list = read_sheet("arquivo.xlsx")
    broke_printers = []
    schedule_ping_for_printers(excel_list, broke_printers)

    while True:
        schedule.run_pending()
        time.sleep(1)