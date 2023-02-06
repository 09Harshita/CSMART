import requests
import time

def run(link,return_code,index):

    try:
        r = requests.head(link)
        print(link + ':' + str(r.status_code))
        # prints the int of the status code. Find more at httpstatusrappers.com :)

        return_code[index] = r.status_code

    except requests.ConnectionError:
        print("failed to connect")
        return_code[index] = "failed to connect"

    time.sleep(1)

