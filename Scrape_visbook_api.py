import requests
import pandas

base_url = r"https://ws.visbook.com/8/api/5602/webproducts/"

colums = ['store_id', 'webProductId','name', 'defaultName', 'unitName', 'sortIndex', 'type', 'maxPeople','minPeople', 'error']
df = pandas.DataFrame(columns=colums)

start = 5000
stop = 10000

for store_id in range(start,stop):

    url = base_url.replace("5602", str(store_id))
    if store_id % 20 == 0:
        print(url)
    try:
        resp = requests.get(url=url)
        data = resp.json() # Check the JSON Response Content documentation below
        if type(data) == list:
            new_row = {}
            for datapoint in data:
                for field in colums:
                    new_row[field] = datapoint.setdefault(field, None)

                new_row["store_id"] = str(store_id)

                df = df.append(new_row, ignore_index=True)


        else:
            new_row = {}
            for field in colums:
                new_row[field] = data.setdefault(field, None)

            new_row["store_id"] = str(store_id)

            df = df.append(new_row, ignore_index=True)

        del data

    except:
        print("FAIL")
        try:
            print("\t", store_id, data)
            del data
        except:
            pass
        new_row = {
                   "store_id": store_id,
                   "error":  "general_failure_with_api"}

        df = df.append(new_row, ignore_index=True)

    del new_row


print(df)
df.to_csv(path_or_buf=f"C:/00_TEMP/crawl_DNT_hytter{start}-{stop}.csv",sep="|",encoding="utf-8")