'''
Note that this script does use the arcpy package and has to have ArcGIS Pro installed on the computer.

ArcGIS Pro does have a large licensing fee and I only use it trough my job.
Therfore this is a seperate script.
'''

import json, os
import datetime as dt

# functions for settings to be global
def printF(level, *string):
    if gs("silent") == "all":
        print(*string)

    elif level <= int(gs("silent")):
        if level == 0:
            print(*string)
        else:
            print("\t"*level, *string)
    return True

def json_to_dict(file) -> dict():
    """Opens a yaml file as a dict"""
    with open(file) as f:
        my_dict = json.load(f)
    return my_dict

def gs(wanted_setting): #gs = Get Setting
    '''Retrives the given setting from the settings file - note that it only reads the file once and never again'''
    if not 'SETTINGS' in globals():
        global SETTINGS
        SETTINGS = json_to_dict('C:/00_GIT/DNT_cabin_availability_system/v02 - Google docs/config/settings.json')
    return SETTINGS[wanted_setting]

def readJson(json_file_path):
    if not os.path.isfile(json_file_path):
        if not gs("silent"): print("Cant find file", json_file_path)
        return False
    with open(json_file_path,"r", encoding='utf-8') as json_file:
        json_data = json.load(json_file)
    return json_data

def replaceData(dataset, targetLayer, cabin_name):
    import arcpy
    describe = arcpy.Describe(targetLayer)
    printF(0, f"Replacing data for: {cabin_name}...")
    printF(1, f"Total of {len(dataset)} new features...")
    printF(3, f"Feature_service: {describe.path}/{describe.name}")
    # 01 DELETE EXISTING DATA
    i = 0
    features_left = True
    while features_left:
        whereClause = """"name" LIKE '%""" + cabin_name + """%'"""

        with arcpy.da.UpdateCursor(in_table=targetLayer, field_names=["name"], where_clause=whereClause) as cur:
            for row in cur:
                i += 1
                if 1 <= int(gs("silent")):
                    print("\t"*1, i, "features deleted.", end="\r")
                cur.deleteRow()

        # Checks if everything is deleted
        # This is necessary as there are sometimes this method of deleting data wont get all features. Normally happens somewhere north of 10k features
        left = 0
        with arcpy.da.UpdateCursor(in_table=targetLayer, field_names=["OBJECTID"], where_clause=whereClause) as cur:
            for row in cur:
                left += 1
        if left == 0:
            features_left = False

    printF(1, i, "features deleted.")

    # 02 APPEND NEW DATA
    i = 0
    sr = arcpy.SpatialReference(gs("AGOL_CABIN_data_epsg"))
    with arcpy.da.InsertCursor(in_table=targetLayer, field_names=["SHAPE"] + gs("AGOL_CABIN_FIELDS")) as pur:
        count = 0
        for feature in dataset:
            count += 1
            if count > gs("AGOL_CABIN_DAYS_AHEAD"):
                break
            row = []
            row.append(
                arcpy.PointGeometry(
                    arcpy.Point(X=feature["long"],Y=feature["lat"]),
                    sr
                )
            )

            row.append(feature["navn"])
            row.append(feature['eier'])
            row.append(feature['betjeningsgrad'])
            row.append(feature['senger'])
            row.append(feature['url'])
            row.append(feature['date'].replace("-","/"))
            row.append(feature['last_checked'].replace("-","/"))
            row.append(feature['available'])
            row.append(feature['max'])
            #pur.insertRow(row)
            try:
                pur.insertRow(row)
                i += 1
                if 3 <= int(gs("silent")):
                    print("\t"*3, "Inserted:", i, end="\r")
            except:
                printF(1, f"Could not import {row}")

            finally:
                del row
    printF(1, i, "features inserted.")

    return i

# This function is just ammended with data.
def logfile( status="", start_time="", number_of_api_requests=0, number_of_features=0):
    try:
        import csv
        import os
        from datetime import date,  datetime
        file =r"C:\00_GIT\DNT_cabin_availability_system\v02 - Google docs\statistics.csv"
        line = ["\n" + str(date.today()),
                str(start_time.strftime("%H:%M:%S")),
                str(datetime.now().strftime("%H:%M:%S")),
                "DNT_cabin_availability_system_v0.2_AGOL",
                str(os.getlogin()),
                str(status),
                str(number_of_api_requests),
                str(os.path.basename(__file__)),
                str(number_of_features),]

        with open(file, 'a') as f:
            writer = csv.writer(f, delimiter=";",lineterminator='')
            writer.writerow(line)
        pass
    except:
        print("could not print to logfile")


def main():
    cabin_data_to_upload = readJson(gs("AGOL_CABIN_temp_json_file_path"))

    # for row in cabin_data_to_upload:
    #     print(row)


    if cabin_data_to_upload == False or cabin_data_to_upload == {}:
        print("NO DATA OR FILE FOUND. EXITING")
        exit()


    import arcpy

    ## Connecting to data.
    # This is an easy way to do it...
    lyrx_file = arcpy.mp.LayerFile(gs("AGOL_CABIN_lyrx_path"))
    cabin_layer = False
    for lyr in lyrx_file.listLayers():
        if lyr.name == gs("AGOL_CABIN_lyrx_lyr_name"):
            cabin_layer = lyr
            printF(0, "Connected to online layer...")

    if cabin_layer == False:
        print("could not connect to AGOL layer. Exiting.")
        exit()


    ## Delete existing data
    # with arcpy.da.UpdateCursor(in_table=cabin_layer, field_names=["SHAPE@"]) as cur:
    #     for row in cur:
    #         cur.deleteRow()

    # Upload new data
    if gs("AGOL_CABIN_DELETE_ALL"):
        printF(0,"Deleting all data")
        replaceData(dataset=[], targetLayer=cabin_layer, cabin_name="")
    number_of_inserted_featrures = 0
    printF(0,"Updating data in ArcGIS Online")
    for cabin in cabin_data_to_upload.keys():
        number_of_inserted_featrures += replaceData(dataset=cabin_data_to_upload[cabin],targetLayer=cabin_layer,cabin_name=cabin)
        printF(2, "total number of features inserted so far:", number_of_inserted_featrures)

    printF(0, "Writing log...")
    logfile(status="Sucsess", start_time=START_TIME, number_of_features=number_of_inserted_featrures)

START_TIME = dt.datetime.now()

if __name__ == '__main__':
    try:
        main()

        print("Success!")
    except Exception as e:
        print(e)
        printF(0, "Writing log...")
        logfile(status="Failed",start_time=START_TIME)
        print("Failed")
