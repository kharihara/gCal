from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime

wb= load_workbook(filename= 'pow3.xlsx')
sheet = wb.worksheets[0]
row_count = sheet.max_row
#print(str(sheet.cell(None, 15,1).value))
index=[]
events=[]
dates=[]
#checks if the column is a date. Assumes that the date is in excell serial format
# puts the index of the date into the list. This shows where to start getting the dates from document
class CalEvent:
    #contains a class for each event 
    def __init__(self, date, time, name, mustloc, execloc, OIC, personnel, uniform):
        self.date = date
        self.time = time
        self.name = name
        self.mustloc = mustloc
        self.execloc = execloc
        self.personnel = personnel
        self.uniform = uniform



for x in range(1,row_count):
    j= (sheet.cell(None,x,1).value)
    if not (j is None) and isinstance(j, float) :
        index.append(x)
        dates.append(j)
        print (j)

for x in range(len(index)):
    i=index[x]
    while sheet.cell(None,i,3).value:
        specDate=dates[x]
        time= sheet.cell(None,i,2).value
        eventName= sheet.cell(None,i,3).value
        mustloc=sheet.cell(None,i,4).value
        execloc= sheet.cell(None,i,5).value
        oic=sheet.cell(None,i,6).value
        personnel= sheet.cell(None,i,7).value
        uniform= sheet.cell(None,i,8).value
        events.append(CalEvent(specDate,time,eventName,mustloc,execloc,oic,personnel,uniform))
        i+=1

    

def convertDatetoExcel(day, month, year):
    offset= 693594
    itime = date(year,month, day)
    n = itime.toordinal()
    return (n - offset)


def convertExceltoDate(serialDate):
    dt = datetime.fromordinal(datetime(1900,1,1).toordinal() + int(serialDate) - 2)
    d= dt.date()
    d=str(d).replace('-','/')
    return d 

def addcolon(time):
    newTime= str(time)
    newTime= newTime[:2] + ':' + newTime[2:]
    return newTime
for x in events:
    print (x.name)

def createNewSheet(eventList):
    newWb= Workbook()
    dest_filename = 'formated_sheet.xlsx'
    ws1 = newWb.active
    ws1.title= "google calender formated sheet"
    colnames = ["Subject",'Start date','Start time','End Date','End Time','All Day Event','Description','Location','Private']
    ws1.append(colnames)
    for x in range(2,len(eventList)):
        descript= "Personnel: " + str(eventList[x].personnel) + "; UOD: " + str(eventList[x].uniform) + "; Execution Location: " + str(eventList[x].execloc)
        ws1.cell(None,x,1).value = eventList[x].name
        ws1.cell(None,x,2).value = convertExceltoDate(eventList[x].date)
        ws1.cell(None,x,7).value = descript
        ws1.cell(None,x,8).value = eventList[x].mustloc
        if not eventList[x].time:
            ws1.cell(None,x,6).value = True
        else:
            ws1.cell(None,x,3).value = addcolon(eventList[x].time)



    newWb.save(filename = dest_filename)

createNewSheet(events)
