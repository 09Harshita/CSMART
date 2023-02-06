__author__ = '720497'

import paramiko
import time
from threading import Semaphore,Thread
threads = []


def run(servername,user,pwd,result,status,i):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    resp = str()
    time.sleep(2)
    try:

        client.connect(servername, username=user, password=pwd)
        channel = client.invoke_shell()
        command = "clustat | grep service \n"
        command2 = "dzdo su - \n"
        time.sleep(5)
        channel.send(command2)
        time.sleep(1)
        channel.send(command)
        time.sleep(5)

        text=channel.recv(99999).decode("utf-8")
        sn=text.splitlines()
        result[i]= sn[23:-1]
        status[i]="Available"
        time.sleep(5)
        client.close()

    except paramiko.AuthenticationException:
        print(servername+":Authentication failed, please verify your credentials")
        status[i]= 'Authentication failed'
        client.close()
        time.sleep(1)

    except paramiko.SSHException:
        print(servername+":Unable to establish SSH connection")
        status[i]="Unavailable"
        client.close()
        time.sleep(1)

    except Exception as e:
        time.sleep(2)
        status[i]="Unusual behavior"
        print(servername,e.args)
        client.close()
