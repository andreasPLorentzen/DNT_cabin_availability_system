import json, os.path, requests, numpy
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os.path
import pandas as pd
import datetime as dt
from dateutil.relativedelta import relativedelta

START_TIME = dt.datetime.now()

# functions for settings to be global
def json_to_dict(file) -> dict():
    """Opens a yaml file as a dict"""
    with open(file) as f:
        my_dict = json.load(f)
    return my_dict

def gs(wanted_setting): #gs = Get Setting
    '''Retrives the given setting from the settings file - note that it only reads the file once and never again'''
    if not 'SETTINGS' in globals():
        global SETTINGS
        SETTINGS = json_to_dict(r'C:\Users\andre\OneDrive\CODE\v02 - Google docs\config\settings.json')
    return SETTINGS[wanted_setting]

def logfile( status="", start_time="", number_of_api_requests=0):
    try:
        import csv
        import os
        from datetime import date,  datetime
        file =r"C:\Users\andre\OneDrive\CODE\v02 - Google docs\config\statistics.csv"
        line = ["\n" + str(date.today()),
                str(start_time.strftime("%H:%M:%S")),
                str(datetime.now().strftime("%H:%M:%S")),
                "DNT_cabin_availability_system_v0.2",
                str(os.getlogin()),
                str(status),
                str(number_of_api_requests),
                str(os.path.basename(__file__))]

        with open(file, 'a') as f:
            writer = csv.writer(f, delimiter=";",lineterminator='')
            writer.writerow(line)
        pass
    except:
        print("could not print to logfile")

# This is mostly gathered from https://developers.google.com/sheets/api/quickstart/python
def Connect_With_API():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(gs("Token")):
        creds = Credentials.from_authorized_user_file(gs("Token"), gs("SCOPES"))
    # If there are no (valid) credentials available, let the user log in.

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                # seems like it wont work as well with refreshing. should therfore delete the token.json file and run
                # this function again.
                creds.refresh(Request())
            except:
                # delete TPLEM file
                os.remove(gs("Token"))

                # Run this again
                return Connect_With_API()
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                gs("API_connection"), gs("SCOPES"))
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(gs("Token"), 'w') as token:
            token.write(creds.to_json())

    return creds

# This function creates a list of months to check from todays month
def get_next_months_as_list(numberOfMonths=10) -> list:
    d = dt.date.today()
    list_of_months = [(d.strftime("%Y"),d.strftime("%m"))]
    for i in range(1,numberOfMonths):
        next_month = d + relativedelta(months=i)
        list_of_months.append((next_month.strftime("%Y"), next_month.strftime("%m")))
    return list_of_months

def intF(str):
    if str == None:
        return 0
    else:
        try:
            return int(str)
        except:
            return 1

def printF(level, *string):
    if gs("silent") == "all":
        print(*string)

    elif level <= int(gs("silent")):
        if level == 0:
            print(*string)
        else:
            print("\t"*level, *string)
    return True

def Gather_data_from_API(df, number_of_months_to_check_ahead):
    # gather the data
    #df[gs("presentation_col_last_gathered")] = ""
    number_of_api_requests = 0
    months_to_check = get_next_months_as_list(number_of_months_to_check_ahead)
    print(df)

    for index, row in df.iterrows():
        print(str(row[gs("controller_name_field")]))
        if "STOP" in str(row[gs("controller_name_field")]):
            printF(0,"\tStopping")
            break
        elif row[gs("controller_name_field")] == None or row[gs("controller_name_field")] == "" or type(row[gs("controller_name_field")]) == float:
            printF(1,"\tNo name defined")
            pass
        elif "H#" in row[gs("controller_name_field")]:
            printF(1,f" {row[gs('controller_name_field')]} is an Heading")
            pass
        elif row[gs('controller_store_id')] == None or row[gs('controller_store_id')] == "":
            printF(1,"Empty or text")
            pass

        ## Getting data
        else:
            printF(0, f'Gathering data from API for {row[gs("controller_name_field")]}...')
            # Check if store has defined products in control document, if not, find products trough the store api endpoint.

            if row[gs("controller_product_ids")] == None or row[gs("controller_product_ids")] == "":
                list_of_accommodations = get_products_in_store(row[gs('controller_store_id')])
                number_of_api_requests += 1

            # indicates false product ID. Should be int or string.
            elif type(row[gs("controller_product_ids")]) == float:
                printF(0,"\tType is not recognized as product ID",gs("controller_name_field"),row[gs("controller_product_ids")])
                list_of_accommodations = [str(row[gs("controller_product_ids")]).split(".")]

            else:
                list_of_accommodations = str(row[gs("controller_product_ids")]).split(",")


            df.at[index, gs("presentation_col_link")] = f'https://reservations.visbook.com/{intF(row[gs("controller_store_id")])}'
            df.at[index, gs("presentation_col_last_gathered")] = dt.datetime.now()
            df.at[index, gs("presentation_col_max_products")] = len(list_of_accommodations)

            for month in months_to_check:
                for accommodation in list_of_accommodations:
                    # Creating quarry
                    try:
                        request_url = gs("visbook_base_availability_api_url") + f"{int(row[gs('controller_store_id')])}/availability/{int(accommodation)}/{month[0]}-{month[1]}"
                        printF(3,"\t", request_url)

                        # Getting data
                        resp = requests.get(url=request_url)
                        data = resp.json()
                        number_of_api_requests += 1
                        # Adding data to dataframe
                        for step in data['items']:
                            date = step['date'].replace("T00:00:00","").split("+")[0]  # just to shorten date in final product. no need for time. nor for time difference
                            if date in df:
                                if step['webProducts'][0]['availability']['available']:
                                    if numpy.isnan(df.at[index, date]):
                                        df.at[index, date] = 1
                                    else:
                                        df.at[index, date] = df.at[index, date] + 1
                                else:
                                    df.at[index, date] = 0

                            else:
                                if step['webProducts'][0]['availability']['available']:
                                    df.at[index, date] = 1
                                else:
                                    df.at[index, date] = 0
                    except Exception as e:
                        print("error?")
                        print(e)

    df.drop(gs("controller_store_id"), 1, inplace=True)
    df.drop(gs("controller_product_ids"), 1, inplace=True)
    # # export to excel
    date = dt.datetime.now()
    # output_file = os.path.join(gs("temp_file_location"), date.strftime("%Y-%m-%d") + "_" + gs("temp_file_basename") + ".xlsx")
    # df.to_excel(output_file, sheet_name=gs("temp_file_sheet_name"), startrow=gs("presentation_row_nr_table_heading") - 1, index=False)
    return df, number_of_api_requests

def generate_format_heading(row_nr) -> dict:
    returning = { "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1,
                        "startRowIndex": row_nr - 1,
                        "endRowIndex": row_nr ,
                    },
                    "cell": {
                        "userEnteredFormat": {"textFormat": {"bold": gs("presentation_heading_bold"),
                                                             "fontSize": gs("presentation_heading_font_size")}
                                              }
                    },
                    "fields": "userEnteredFormat.textFormat"
            }
     }
    return returning

def generate_format_date(row_nr, col_nr, bold, border_width, border_style, col_nr_add=1) -> dict:
    returning = { "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startColumnIndex": col_nr,
                        "endColumnIndex": col_nr + col_nr_add,
                        "startRowIndex": row_nr - 1,
                        "endRowIndex": row_nr+400,
                    },
                    "cell": {
                        "userEnteredFormat": {"textFormat": {"bold": bold,
                                                             "fontSize": gs("presentation_weekend_font_size")},
                                              "borders": {
                                                  "right": {
                                                      "style": border_style,
                                                      "width": border_width
                                                  }
                                              }
                                            }
                    },
                    "fields": "userEnteredFormat.textFormat,userEnteredFormat.borders"
            }
     }
    return returning

def get_products_in_store(store_id) -> list:
    list_of_product_ids_to_return = []
    if store_id is None:
        return []
    request_url = gs("visbook_base_store_api_url").replace(gs("visbook_base_store_api_url_id_sign"),str(store_id))

    printF(1,"Gathering products from store-ID api: ", request_url)
    resp = requests.get(url=request_url)
    data = resp.json()  # Check the JSON Response Content documentation below
    for row in data:
        if "webProductId" in row.keys():
            list_of_product_ids_to_return.append(row["webProductId"])

    printF(1,"Found {0} products.".format(len(list_of_product_ids_to_return)))

    return list_of_product_ids_to_return

def generate_url_name(row_nr, name, url, col_nr) -> dict:

    returning = { "repeatCell": {
                    "range": {
                        "sheetId": 0,
                        "startColumnIndex": col_nr,
                        "endColumnIndex": col_nr +1,
                        "startRowIndex": row_nr - 1,
                        "endRowIndex":  row_nr,
                    },
                    "cell": {
                        "userEnteredValue": {"formulaValue": f'=HYPERLINK("{url}";"{name}")'
                                             }
                    },

                    "fields": "*"#"userEnteredFormat.formulaValue"
                    }
                }


    return returning

def main():
    # Cet credentials for Google API
    creds = Connect_With_API()
    service = build('sheets', 'v4', credentials=creds)
    # Get control document
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=gs("controller_sheet_id"), range=gs("controller_data_area")).execute()

    data = result.get('values')
    # since this is an afterthought to add url's to the names, we will only create a look-up table between the name and the url, then drop the url field...
    # I know. This is a prototype. stuff like that is allowed.
    url_lookup_table = [] # cant use dict, as there might be several rows with the same name. gah.
    data[0].pop()
    for row in range(1,len(data)):
        if len(data[row]) >= 4:
            url_lookup_table.append([row,data[row][0], data[row][3]])
            #then dropping it
            data[row].pop()


    control_document_df = pd.DataFrame(data[1:], columns=data[0])

    # Check if control document is empty
    if control_document_df.empty:
        printF(0,'No data found in control document.')
        exit()
    # else:
    #     print(control_document_df)


    ### Gather data from visbook API
    # As pandas dataframe
    Gathered_data, api_requests = Gather_data_from_API(control_document_df, gs("system_months_to_check_ahead"))

    Gathered_data[gs("presentation_col_last_gathered")] = Gathered_data[gs("presentation_col_last_gathered")].dt.strftime('%Y-%m-%d %H:%M:%S')
    data = [Gathered_data.columns.tolist()]
    dataframe = Gathered_data.values.tolist()
    for row in dataframe:
        data.append(row)

    # Fixing data for google sheet
    list_of_rows_with_headings = []

    for row in range(0,len(data)):
        for col in range(0,len(data[row])):
            if pd.isnull(data[row][col]):
                data[row][col] = ""
            elif gs("presentation_heading_sign") in str(data[row][col]):
                list_of_rows_with_headings.append(row)
                data[row][col] = data[row][col][2:]

            # expand the table to remove old values when the months get shorter
        for i in range(0,60):
            data[row].append("")




    ### Write back to Google Sheet
    body = {'values': data}
    #print(body)
    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=gs("presentation_sheet_id"),
            range=gs("presentation_sheet_name") + '!A'+ str(gs("presentation_row_nr_table_heading")) + ':ZZZ500',
            valueInputOption='RAW',
            body=body).execute()
    except Exception as e:
        print(e)
    print(result)

    print("STARTING WITH FORMATTING")


    ### Formatting
    # Remove old Headings
    body = {"requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1,
                    "startRowIndex": gs("presentation_row_nr_table_heading"),
                    "endRowIndex": gs("presentation_row_nr_table_heading") + len(data),
                },
                "cell": {
                    "userEnteredFormat": {"textFormat": {"bold": gs("presentation_normal_bold"),
                                                         "fontSize": gs("presentation_normal_font_size")}
                                          }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        }
    ]
    }
    result = service.spreadsheets().batchUpdate(spreadsheetId=gs("presentation_sheet_id"),
                                                body=body).execute()

    # Add new formating to headings
    format_body = {
        "requests": []
    }
    for row in list_of_rows_with_headings:
        format_body["requests"].append(generate_format_heading(row + gs("presentation_row_nr_table_heading")))

    # adding format to date headings:
    for col in range(4, len(data[0])):
        checkdate = data[0][col].split("-")
        #print(checkdate)
        if len(checkdate) == 3:
            print(checkdate)
            d = dt.datetime(year=int(checkdate[0]),month=int(checkdate[1]),day=int(checkdate[2].split("+")[0])) #had to remove the "+" sign
            print(d)
            #print("\t", d, d.weekday())
            if d.weekday() == 4:
                format_body["requests"].append(generate_format_date(row_nr=gs("presentation_row_nr_table_heading"),
                                                                    col_nr=col,
                                                                    bold=False,
                                                                    border_width=gs(
                                                                        "presentation_weekend_border_left_width"),
                                                                    border_style=gs(
                                                                        "presentation_weekend_border_left_style")))
            elif d.weekday() == 5:
                format_body["requests"].append(generate_format_date(row_nr=gs("presentation_row_nr_table_heading"),
                                                                    col_nr=col,
                                                                    bold=gs("presentation_weekend_bold"),
                                                                    border_width=1,
                                                                    border_style="DASHED"))
            elif d.weekday() == 6:
                format_body["requests"].append(generate_format_date(row_nr=gs("presentation_row_nr_table_heading") ,
                                                                    col_nr=col,
                                                                    bold=gs("presentation_weekend_bold"),
                                                                    border_width=gs("presentation_weekend_border_left_width"),
                                                                    border_style=gs("presentation_weekend_border_left_style")))

            else:
                format_body["requests"].append(generate_format_date(row_nr=gs("presentation_row_nr_table_heading"),
                                                                    col_nr=col,
                                                                    bold=False,
                                                                    border_width=1,
                                                                    border_style="DASHED"))
        else:
            format_body["requests"].append(generate_format_date(row_nr=gs("presentation_row_nr_table_heading"),
                                                                col_nr=col,
                                                                bold=False,
                                                                border_width=1,
                                                                border_style="NONE",
                                                                col_nr_add=len(data[0])-col))
            break
                #print("\t\t", False)

    # format weekends
    for row in url_lookup_table:
        format_body["requests"].append(generate_url_name(row_nr=row[0] + gs("presentation_row_nr_table_heading"),
                                                         name=row[1],
                                                         url=row[2],
                                                         col_nr=0))
    # fix presentation of url to store
    for row in range(1,len(data)):
        if data[row][2] != "" and data[row][2] != None:
            format_body["requests"].append(
                generate_url_name(row_nr=row + gs("presentation_row_nr_table_heading"),
                                  name="Bestillingside",
                                  url=data[row][1],
                                  col_nr=1))

    print("batchupdate")
    result = service.spreadsheets().batchUpdate(spreadsheetId=gs("presentation_sheet_id"),
                                                body=format_body).execute()




    # Gather data and export it to AGOL formater to read. Notice that the upload part will be in another system, using another python env.
    # This is done as i do not want to import arcpy into this env. Also, this is by no means a key feature...

    # Get CABIN document

    printF(0,"Prepering AGOL stuff..")
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=gs("AGOL_CABIN_sheet_id"), range=gs("AGOL_CABIN_data_area")).execute()

    DNT_CABIN_DATA_RAW = result.get('values')

    # convert to dict. a bit lazy today.
    DNT_CABIN_DATA = {}
    for cabin in DNT_CABIN_DATA_RAW:
       DNT_CABIN_DATA[cabin[0]] = {
           "navn": cabin[0],
           "eier": cabin[1],
           "betjeningsgrad": cabin[2],
           "senger": cabin[3],
           "url": cabin[4],
           "long": cabin[5],
           "lat": cabin[6]
       }

    #
    # Combine into single JSON
    list_of_features = {}
    for row in data:
        if row[0] in DNT_CABIN_DATA.keys() and row[0] != data[0][0]:
            list_of_features[row[0]] = []
            for col in range(4,len(row)):
                if row[col] != "":
                    if data[0][col] == "":
                        break
                    new_feature = DNT_CABIN_DATA[row[0]].copy()
                    new_feature["date"] = data[0][col]
                    new_feature["max"] = row[3]
                    new_feature["available"] = row[col]
                    new_feature["last_checked"] = row[2]
                    list_of_features[row[0]].append(new_feature)
                    del new_feature
        else:
            printF(1, row[0], "Not in list of cabin data...")
            if row[0] == "STOP":
                break

    with open(gs("AGOL_CABIN_temp_json_file_path"), "w", encoding='utf-8') as file:
        json.dump(list_of_features, file, ensure_ascii=False)





    printF(0,"Writing log...")
    logfile(status="Sucsess",start_time=START_TIME,number_of_api_requests=api_requests)
if __name__ == '__main__':
    try:
        main()
        print("Success!")
    except Exception as e:
        print(e)
        logfile(status="Failed",start_time=START_TIME)
        print("Failed")


