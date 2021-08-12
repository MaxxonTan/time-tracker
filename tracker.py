import time
import logging
from datetime import datetime, timedelta
import os.path
import json
try:
    import psutil
    import win32process
    import win32gui
except:
    print("Error importing psutil or win32.")

previousWindow = str()
currentTime = datetime.now().strftime("%I:%M %p")
logging.basicConfig(filename='D:/maxxo/projects/AppTracker/logs/tracker.log',
                    format='%(asctime)s - %(levelname)s  -%(message)s', datefmt='%d-%b-%y %H:%M:%S')

def getCurrentWindow():
    pid = win32process.GetWindowThreadProcessId(
                    win32gui.GetForegroundWindow())[1]
    currentWindow = psutil.Process(pid).name().replace(".exe", "")
    return currentWindow

def writeActivity():
    pass

#log the acitivity history to a json file(acitivityHistory.json)
def writeActivityHistory():
    pass

#make sure activity.json and activityHistory.json exsists
def logActivity(data: dict() = {}):
    global currentDate
    currentDate = datetime.now().strftime("%d-%m-%Y")

    global previousWindow 
    currentWindow = previousWindow = getCurrentWindow()

    if not(os.path.isfile(f"D:/maxxo/projects/AppTracker/logs/activities{currentDate}.json")) :
        with open(f'D:/maxxo/projects/AppTracker/logs/activities{currentDate}.json', 'w+') as file:
            json.dump({"activities": [{"name" : currentWindow, "timeSpent" : [
                {"hours": 0, "minutes": 0, "seconds": 0}], "firstUsed": currentTime, "lastUsed": "0"}]}, file, indent=4)
            file.close()
    else:
        if len(data) != 0:
            with open(f'D:/maxxo/projects/AppTracker/logs/activities{currentDate}.json', 'w+') as file:
                file.seek(0)
                json.dump(data, file, indent=4)
                file.close()
    
def getActivities() :
     with open(f'D:/maxxo/projects/AppTracker/logs/activities{currentDate}.json', 'r') as file:
            activities = json.load(file)
            file.close()
            return activities

def updateActivity(activities: dict(), windowName, timeSpent, endTime):
    index = 0

    for activity in activities["activities"]:
        if activity["name"] == windowName:
            break
        index += 1

    activities["activities"][index]["timeSpent"][0]["seconds"] += timeSpent
    activities["activities"][index]["lastUsed"] = endTime.strftime("%I:%M %p")
    seconds = activities["activities"][index]["timeSpent"][0]["seconds"]
    minutes = activities["activities"][index]["timeSpent"][0]["minutes"]

    # Spliting seconds to minutes and hours (copied from https://github.com/jedi2610/Automatic-Time-Tracker-for-Windows/blob/master/auto.py)
    if seconds >= 60:
        activities["activities"][index]["timeSpent"][0]["minutes"] += seconds // 60
        activities["activities"][index]["timeSpent"][0]["seconds"] = seconds % 60
    if minutes >= 60:
        activities["activities"][index]["timeSpent"][0]["hours"] += minutes // 60
        activities["activities"][index]["timeSpent"][0]["minutes"] = minutes % 60

    return activities

def main():
    global previousWindow
    startTime = datetime.now()

    #initialize files
    logActivity()

    while True:
        try:
            time.sleep(5)
            currentWindow = getCurrentWindow()

            #new lastUsed
            endTime = datetime.now()
            
            if currentWindow == previousWindow:
                #new timespent
                timeSpent = endTime - startTime
                timeSpent = round(timeSpent.total_seconds())

                #updating timeSpent and lastUSed
                activities = getActivities()

                activityHistory = []
                for activity in activities["activities"]:
                    activityHistory.append(activity["name"])
                    
                if currentWindow not in activityHistory:
                    newActivity = {
                        "name" : currentWindow,
                        "timeSpent" : [{"hours": 0, "minutes": 0, "seconds": 0}],
                        "firstUsed" : endTime.strftime("%I:%M %p"),
                        "lastUsed" : ""
                    }
                    activities["activities"].append(newActivity)

                activities = updateActivity(activities, currentWindow, timeSpent, endTime)
                
                print(activities)
                logActivity(activities)
                
            else:
                #Time spent for previous window
                timeSpent = endTime - startTime
                timeSpent = round(timeSpent.total_seconds())

                activities = getActivities()

                #Updating details for previous window
                activities = updateActivity(activities, previousWindow, timeSpent, endTime)

                activityHistory = []
                for activity in activities["activities"]:
                    activityHistory.append(activity["name"])
                
                #Check if current window is already logged into activities.json, if not then create a new entry for a window
                if currentWindow not in activityHistory:
                    newActivity = {
                        "name" : currentWindow,
                        "timeSpent" : [{"hours": 0, "minutes": 0, "seconds": 0}],
                        "firstUsed" : endTime.strftime("%I:%M %p"),
                        "lastUsed" : ""
                    }
                    activities["activities"].append(newActivity)

                print(activities)
                logActivity(activities)

                previousWindow = currentWindow

            startTime = datetime.now()
            
        
        except psutil.NoSuchProcess:
            time.sleep(5)
            continue
        except ProcessLookupError:
            time.sleep(5)
            continue
        except Exception as e:
            logging.error(f'{e} error occured') 

if __name__ == '__main__':
    main()
    