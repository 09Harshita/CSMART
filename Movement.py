__author__ = '720497'


import paramiko
import time

def run1(clustername,system,user,pwd):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    resp = str()
    time.sleep(2)
    try:

        client.connect(clustername, username=user, password=pwd)
        channel = client.invoke_shell()


        command2 = "dzdo su - \n"
        channel.send(command2)
        time.sleep(5)
        index = system.split()
        command = "clusvcadm -r "+index[1]+ " -m "+index[0]+" \n"
        channel.send(command)
        print(system +"Done")
        time.sleep(5)
        client.close()

    except paramiko.AuthenticationException:
        print(clustername+":Authentication failed, please verify your credentials")
        client.close()
        time.sleep(1)

    except paramiko.SSHException:
        print(clustername+":Unable to establish SSH connection")
        client.close()
        time.sleep(1)

    except Exception as e:
        time.sleep(2)
        print(clustername,e.args)
        client.close()


def run2(clustername,system,user,pwd):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    resp = str()
    time.sleep(2)
    try:

        client.connect(clustername, username=user, password=pwd)
        channel = client.invoke_shell()


        command2 = "dzdo su - \n"
        channel.send(command2)
        time.sleep(5)

        for sys in system:

            index = sys.split()
            command = "clusvcadm -r "+index[1]+ " -m "+index[0]+" \n"
            channel.send(command)
            print(sys +"Done")
            time.sleep(50)




        client.close()

    except paramiko.AuthenticationException:
        print(clustername+":Authentication failed, please verify your credentials")
        client.close()
        time.sleep(1)

    except paramiko.SSHException:
        print(clustername+":Unable to establish SSH connection")
        client.close()
        time.sleep(1)

    except Exception as e:
        time.sleep(2)
        print(clustername,e.args)
        client.close()






