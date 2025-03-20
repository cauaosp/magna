from utils import ping_printers, read_sheet

if __name__ == "__main__":
    excel_list = read_sheet("arquivo.xlsx")
    broke_printers = []

    for printer in excel_list:
        result = ping_printers(printer)
        if result is not None:
            broke_printers.append(result)

    print(broke_printers)