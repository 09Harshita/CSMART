__author__ = '720497'

import time
import ServerConnect
from threading import Semaphore,Thread


threads = []

def run(sheet2,x,u,p):
    servername = []
    for i in range(2, sheet2.max_row):
        servername=[]
        if sheet2.cell(i,1).value == x:
            servername.append(sheet2.cell(i,2).value)
            i=i+1
            while(sheet2.cell(i,1).value == None):
                servername.append(sheet2.cell(i,2).value)
                i=i+1
                if (i==sheet2.max_row+1):
                    break
            break

    result = [{} for x in servername]
    OsResult = [{} for x in servername]
    run_service = [{} for x in servername]

    lock = Semaphore(value=1)
    for i in range(len(servername)):
        time.sleep(2)
        process = Thread(target=ServerConnect.run, args=[servername[i]+'.mmm.com',u,p,x,result,OsResult,i,run_service,lock])
        process.start()
        threads.append(process)

    for process in threads:
        process.join()


    j=0

    for i in range(2, sheet2.max_row):
         if sheet2.cell(i,1).value == x:
            sheet2.cell(i,3).value = OsResult[j]
            sheet2.cell(i,4).value = result[j]
            sheet2.cell(i,5).value = run_service[j]
            j=j+1
            i+=1
            while(sheet2.cell(i,1).value == None and sheet2.cell(i,1).value == None ):
                sheet2.cell(i,3).value = OsResult[j]
                sheet2.cell(i,4).value = result[j]
                sheet2.cell(i,5).value = run_service[j]
                j = j+1
                i+=1
                if i==sheet2.max_row+1:
                    break
            break

