import requests
from datetime import datetime
import time

URL = "https://httpbin.org/ip"

while True:
    now = datetime.now()
    try:
        r = requests.get(url=URL)
        data = r.json()
        ip = data["origin"]
    except:
        ip = "connection error"

    date = now.strftime("%d-%m-%Y %H:%M:%S")

    with open("plik.txt", "a") as log_file:
        output = "{} \t\t {}".format(date, ip)

        print(output)
        log_file.write(output + "\n")
        time.sleep(30)      # 30 seconds
