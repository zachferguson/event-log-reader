import win32evtlog
import datetime

server = 'localhost'
#logtype = 'System'
logtype = 'Application'
handle = win32evtlog.OpenEventLog(server,logtype)
flags = win32evtlog.EVENTLOG_SEQUENTIAL_READ | win32evtlog.EVENTLOG_BACKWARDS_READ
total = win32evtlog.GetNumberOfEventLogRecords(handle)
events = win32evtlog.ReadEventLog(handle,flags,0)
print('Records Analyzed: ' + str(total))
curtime = datetime.datetime.now()
print('current time: ' + str(curtime))
print('day: ' + str(curtime.day))
print('hour: ' + str(curtime.hour))

logout = input("In military time (example - 16:23), what time did you log out yesterday?")
print("number of events")
print(len(events))

# loop through events between logout yesterday and start of current day
compevents = []
print(curtime.day -1)
for i in range(len(events)-1, 0, -1):
    if events[i].TimeGenerated.day == curtime.day -1:
        if events[i].TimeGenerated.hour >= logout and events[i].TimeGenerated.hour <= 23:
            compevents.append(events[i])

# loop through time between start of current day and current time
for i in range(len(events)-1, 0, -1):
    if events[i].TimeGenerated.day == curtime.day and events[i].TimeGenerated.hour <= curtime.hour:
        compevents.append(events[i])

print('')
print('There were ' + str(len(compevents)) + ' events between when you logged out and 11:00 PM yesterday.')
for item in compevents:
    print("_________________________________________________________")
    print("Event Source: {}".format(str(item.SourceName)))
    if item.StringInserts != None and len(item.StringInserts) < 1:
        print("{} - {}".format(item.StringInserts[0], item.StringInserts[1]))
    elif item.StringInserts != None:
        print("{}".format(item.StringInserts[0]))
    print("Time: {}:{}".format(str(item.TimeGenerated.hour), str(item.TimeGenerated.minute)))
    print("")
    

