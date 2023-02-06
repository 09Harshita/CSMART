'''
import httplib2
h = httplib2.Http()
resp, content = h.request("http://sapgdi1.mmm.com/webdynpro/resources/sap.com/tc~lm~itsam~ui~mainframe~wd/FloorPlanApp?home=true")
if (resp.status == 200):
    print (resp.status)

'''

import requests
try:
    r = requests.head("http://sapgq31.mmm.com/sap/bc/gui/sap/its/webgui/!")
    print(r.status_code)
    # prints the int of the status code. Find more at httpstatusrappers.com :)
except requests.ConnectionError:
    print("failed to connect")

'''


import requests
request = requests.get('http://sapgdi1.mmm.com/startPage')
if request.status_code == 200:
    print('Web site exists')
else:
    print('Web site does not exist')
'''