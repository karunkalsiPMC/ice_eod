import pandas as pd
import win32com.client as win32
import os
from datetime import date, timedelta, datetime
import calendar
 
outlook = win32.Dispatch('Outlook.Application').GetNamespace("MAPI")
root_folder = outlook.Folders['Karun.Kalsi@plains.com']
inbox = root_folder.Folders["Inbox"]
test_env_inbox = inbox.Folders['Test Enviroment']
messages = test_env_inbox.Items
today = date(2023, 12, 12)
today_datetime = datetime(2023, 12, 12)
 
path = r"H:\LPG\Analyst Pool\Karun\github\localPython\ice_eod\test_outputs"
template_path = os.path.join(path, 'templates')
new_path = os.path.join(path, str(today.year), f"{today.month}. {calendar.month_name[today.month]} {today.year}", str(today.day))
 
if not os.path.exists(new_path):
    os.makedirs(new_path)
    
def save_attachments(subjects):
    for message in messages:
        for attachment in message.Attachments:
            attachment.SaveAsFile(os.path.join(new_path, str(attachment)))
            if message.Unread:
                message.Unread = False
            break

def merge_data(import_file, name, filter_to, merge_col1, merge_col2, naming_dict):
    todays_data = pd.read_excel(os.path.join(new_path, f"{import_file}{today.year}_{today.month}_{today.day}.xlsx"))
    template = pd.read_excel(os.path.join(template_path, f'{name}_Template.xlsx'))
    cols = template.columns
    template[["QuoteFromDate","QuoteToDate"]] = today_datetime
    todays_data = todays_data.loc[todays_data['CONTRACT'].isin(filter_to)]
    todays_data['PriceCurveName'] = todays_data['CONTRACT'].map(naming_dict)
    todays_data['STRIP'] = pd.to_datetime(todays_data['STRIP']).dt.date 
    template['DeliveryPeriod'] = pd.to_datetime(template['DeliveryPeriod']).dt.date
    merged = pd.merge(template, todays_data[['STRIP','PriceCurveName','SETTLEMENT PRICE']], 
                      right_on=["STRIP", "PriceCurveName"], 
                      left_on=[merge_col1, merge_col2], 
                      how='left')
    merged = merged.drop(columns = "Value")
    merged = merged.rename(columns={'SETTLEMENT PRICE':"Value"})
    merged = merged[cols]
    os.remove(os.path.join(new_path, f"{import_file}{today.year}_{today.month}_{today.day}.xlsx"))
    return merged.set_index('PriceCurveName')
 
def export_to_excel(df, name):
    export_loc = os.path.join(new_path, f'{name}_{calendar.month_abbr[today.month]}_{today.day}_{today.year}_test.xlsx')
    writer = pd.ExcelWriter(export_loc, engine="xlsxwriter")
    df = df[["DeliveryPeriod","QuoteFromDate","QuoteToDate","Value","EstimateorActual","PriceType"]].reset_index()
    df.to_excel(writer, sheet_name = "Sheet1", index = False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    format1 = workbook.add_format({'bg_color': "#C4D79B", 
                                    'font_name':"Calibri",
                                    "font_size": 10})
    format_headers = workbook.add_format({'bold': False, 
                                            'font_name':"Calibri",
                                            "font_size": 10})
    format_date = workbook.add_format({'bold': False,
                                    'font_name':"Calibri",
                                    "num_format": "m/d/yyy hh:mm",
                                    "font_size": 10})
    for row_num, row in enumerate(df):
        for col_num, value in enumerate(df[row]):
            if col_num == 0:
                worksheet.write(col_num, row_num, row, format_headers)
            if row == 'Value':
                worksheet.write(col_num+1, row_num, value, format1)
            elif row == 'DeliveryPeriod':
                worksheet.write(col_num+1, row_num, value, format_date)
            elif row == 'QuoteFromDate':
                worksheet.write(col_num+1, row_num, value, format_date)
            elif row == 'QuoteToDate':
                value = value + timedelta(hours = 23, minutes = 59)
                worksheet.write(col_num+1, row_num, value, format_date)
            else: 
                worksheet.write(col_num+1, row_num, value, format_headers)
    worksheet.autofit()
    writer.close()
 
def check_emails_received(subjects):
    count = 0
    for message in messages:
        #TODO DONT FORGET TO CHANGE BACK
        if message.Subject in subjects and message.Senton.date() == today:
            count += 1
    return count == len(subjects)

subjects = [
    "ICE Data: Cleared Canadian Oil Settlement",
    "ICE Data: NGX Gas Settlement",
    "ICE Data: NGX Power Settlement",
    "ICE Data: Cleared Oil Settlement",
    "ICE Data: Cleared NGL Settlement",
    "ICE Data: Cleared Gas Settlement"
]
naming_dict = {
    'XCU': "NGX AESO 7x24 Month Forward",
    'ARV':"ICE Crude Diff - ARV Argus WCS Houston",
    'NGE':"ICE Mt. Belvieu Non TET Natural Gasoline (OPIS) Future",
    'NGL':"ICE Mt. Belvieu TET Natural Gasoline (OPIS) Future",
    "BM2":"ICE - BM2 TETCO - M2 (Receipts)",
    "IBC":"ICE Conway In-Well Normal Butane (OPIS) Future",
    "ISO":"ICE Mt. Belvieu Iso Butane (OPIS) Future",
    "NBI":"ICE Mt. Belvieu Normal Butane (OPIS) Future",
    "NBR":"ICE Mt. Belvieu TET Normal Butane (OPIS) Future",
    "PRC":"ICE Conway Propane (OPIS) Future",
    "PRL":"ICE Mt. Belvieu Propane (OPIS) Future",
    "PRN":"ICE Mt. Belvieu Non TET Propane (OPIS) Future",
    "XW7":"NGX - AB-NIT Same Day Forward Index 5A (CAD)",
    "XUN":"NGX - AB-NIT Same Day Forward Index 5A (USD)",
    "XW6":"NGX - AB-NIT Month Ahead Forward Index 7A (CAD)",
    "XNR":"NGX - AB-NIT Month Ahead Forward Index 7A (USD)",
    "CSH":"ICE Crude Diff - CSH Argus WCS Cushing", 
    "TMF":"ICE - TMX C5+ 1A (TMF)", 
    "TMR":"ICE - TMX SW 1A (TMR)", 
    "TMS":"ICE - TMX SYN 1A (TMS)", 
    "TMU": "ICE - TMX UHC 1A (TMU)",
    "TMW": "ICE - TMX WCS 1A (TMW)"
}


if check_emails_received(subjects):
    
    save_attachments(subjects)
    for subject in subjects:
        merge_col1 = 'DeliveryPeriod'
        merge_col2 = 'PriceCurveName'
        if subject == "ICE Data: Cleared Canadian Oil Settlement":
            name = "AESO 7x24-FWD"
            filter_to = ['XCU']
            import_file = 'ngxcleared_power_'
        elif subject == "ICE Data: Cleared Oil Settlement":
            name = "ICE_CRUDE"
            filter_to = ['ARV', 'NGE', 'NGL']
            import_file = 'icecleared_oil_'
        elif subject == "ICE Data: Cleared Gas Settlement":
            name = "ICE_GAS"
            filter_to = ['BM2']
            import_file = 'icecleared_gas_'
        elif subject == "ICE Data: NGX Power Settlement":
            name = "ICE_SWAPS"
            filter_to = ['IBC',"ISO", "NBI", "NBR", "PRC", "PRL", "PRN"]
            import_file = 'icecleared_ngl_'
        elif subject == "ICE Data: NGX Gas Settlement":
            name = "NGX 5A-7A FWD"
            filter_to = ['XW7', "XUN", "XW6", "XNR"]
            import_file = 'ngxcleared_gas_'
        elif subject == "ICE Data: Cleared NGL Settlement":
            name = "ICE_DIFF"
            filter_to = ["CSH", "TMF", "TMR", "TMS", "TMU","TMW"]
            import_file = 'iceclearedoil_ca_'
        template = merge_data(import_file, name, filter_to, merge_col1, merge_col2, naming_dict)
        export_to_excel(template, name)
else:
    print("Not all required emails have been received.")