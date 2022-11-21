import pandas as pd
import openpyxl as pxl

def write_in_worksheet(work_sheet_name: str, data: dict):

    df_not_redacted = pd.read_excel("output_not_redacted.xlsx")
    df_data = {
        "Компания": df_not_redacted["Компания"].tolist(),
        "Телефон": df_not_redacted["Телефон"].tolist(),
        "Почта": df_not_redacted["Почта"].tolist(),
        "Адрес": df_not_redacted["Адрес"].tolist()
    }

    df = pd.DataFrame({
        "Компания": df_data["Компания"],
        "Телефон": df_data["Телефон"],
        "Почта": df_data["Почта"],
        "Адрес": df_data["Адрес"]
    })
    
    excel_book = pxl.load_workbook("output.xlsx")
    with pd.ExcelWriter("output.xlsx", "openpyxl") as writer:
        writer.book = excel_book
        df.to_excel(writer, work_sheet_name, index = False)
        writer.save()


if __name__ == "__main__":
    write_in_worksheet("3", {})