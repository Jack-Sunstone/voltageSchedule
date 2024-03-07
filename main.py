from datetime import datetime, time
import requests
import json
import win32com.client
outlook = win32com.client.Dispatch("outlook.application")

def callTime():
    timeFormat = '%H:%M:%S'

    timeNow = datetime.now()

    timeStrip = timeNow.strftime(timeFormat)
    timeStripSTR = str(timeStrip)

    timeSchedule = time(hour=4, minute=00, second=00)
    timeScheduleSTR = str(timeSchedule)

    if timeStripSTR == timeScheduleSTR:

        # Defining login details to access Sites
        login_url = 'https://vrmapi.victronenergy.com/v2/auth/login'
        login_string = '{"username":"support@sunstone-systems.com","password":"12Security34!"}'
        # Stores and loads Json request to the login URL
        response = requests.post(login_url, login_string)
        token = json.loads(response.text).get("token")
        headers = {"X-Authorization": 'Bearer ' + token}

        unitIDS = {202554:"ARC0063", 276976:"ARC0102"}

        for i in unitIDS:

            diags_url = "https://vrmapi.victronenergy.com/v2/installations/{}/diagnostics?count=1000".format(i)
            response = requests.get(diags_url, headers=headers)
            data = response.json().get("records")

            batteryVoltage = str([element['rawValue'] for element in data if element['code'] == "bv"][0])

            mail = outlook.CreateItem(0)

            mail.To = "russell@sunstone-systems.com"
            mail.Subject = "{} Battery Volatage".format(unitIDS.get(i))
            mail.Body = "{}".format(float(batteryVoltage))

            mail.Send()


while True:
    callTime()



