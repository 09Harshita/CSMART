__author__ = '720497'

import paramiko
import time
import openpyxl





#wb =openpyxl.load_workbook('Services.xlsx')
client = paramiko.SSHClient()
client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
resp = str()

try:
        client.connect("sapgscl01.mmm.com", username="a9k8jzzad", password=",SR8UEwlI-N(UnlGeEvo")
        channel = client.invoke_shell()
        print("connected")

        command = "dzdo su" + " \n" + "pcs status"
        command2 = (command + " \n")

        time.sleep(1)
        channel.send(command2)
        time.sleep(5)

        text = channel.recv(99999).decode("utf-8")
        print(text)


        client.close()
        time.sleep(5)





except paramiko.AuthenticationException:
    print("1")
    print(":Authentication failed, please verify your credentials")
    client.close()
    time.sleep(1)
except paramiko.BadHostKeyException as badHostKeyException:
    print("2")
    print("Unable to verify server's host key")
    client.close()
    time.sleep(1)
except paramiko.SSHException:
    print("3")
    print(":Unable to establish SSH connection")
    client.close()
    time.sleep(1)
except ValueError as e:
    print("4")
    print(e.args)
    client.close()
    time.sleep(1)
except Exception as e:
    print("5")
    print(e.args)
    client.close()
    time.sleep(1)