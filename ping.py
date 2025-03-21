from utils import read_sheet, schedule_ping_for_printers
import time
import schedule

if __name__ == "__main__":
    excel_list = read_sheet("arquivo.xlsx")
    broke_printers = []
    schedule_ping_for_printers(excel_list, broke_printers)
    end_time = "23:59"

    while True:
        current_time = time.strftime("%H:%M", time.localtime())

        if current_time >= end_time:
            print(f"Execução encerrada as {end_time}.")
            break

        schedule.run_pending()
        time.sleep(1)