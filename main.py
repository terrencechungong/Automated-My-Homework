import math
from random import randrange
import pandas as pd
import openpyxl
from datetime import datetime
from docxtpl import DocxTemplate
from pathlib import Path

base_dir = Path(__file__).parent
word_template_path = base_dir / "ActivityLog.docx"

doc = DocxTemplate(word_template_path)

HEALTH_DATA =  base_dir / 'HealthData.csv'
csvFile = pd.read_csv(HEALTH_DATA)

WORKOUT_DATA =  base_dir / 'workouts.csv'
workCsv = pd.read_csv(WORKOUT_DATA)

oneDay = 86400.0
dayOne = datetime(2023, 1, 27).timestamp()

dayRange = 7
week = 1
all_days = dict()
for index, row in csvFile.iterrows():
    date = row.values[0].split("/")
    date = list(map(lambda n: int(n), date))
    month, day, year = date[0], date[1], 2000 + date[2]
    dateInSeconds = datetime(year, month, day)
    day = round((dateInSeconds.timestamp()-dayOne)/oneDay)
    if dayRange - day < 0:
        dayRange += 7
        week += 1
        all_days[week] = [row.values[0]]
    else:
        if week not in all_days:
            all_days[week] = []
        all_days[week].append(row.values[0])

seen = set()

def additionalContext(weekNum):
    context['stretch1'] = "Pre-workout dynamic stretch routine (upper and lower body)"
    context['time'] = "7 minutes"
    context['weekNumber'] = weekNum

def strengthContext(index, date):
    randonRow = randrange(25)
    context[f"date{index}"] = date
    context[f"strength{index}"] = workCsv.loc[randonRow][0]
    context[f"weight{index}"] = workCsv.loc[randonRow][1]
    context[f"setsReps{index}"] = workCsv.loc[randonRow][2]

def runContext(i, date, time, distance):
    context[f"runDate{i}"] = date
    context[f"cardio{i}"] = "Running"
    context[f"timeDistance{i}"] = f"{distance} miles, {time} minutes"

def duplicateDates(rows, index, runParam, dateRepeats):
    for row in rows:
        for i in range(dateRepeats):
            if row[2] == 'Running':
                runContext(runParam, row[0], row[3], row[1])
                runParam += 1
                break
            else:
                strengthContext(index, row[0])
                index += 1
    return index, runParam

for weekNum, dates in all_days.items():
    context = dict()
    workoutFrequency = math.floor(14/len(dates)) * len(dates)
    i = 1
    for date in dates:
        runCounter = 1
        for j in range(workoutFrequency//len(dates)):
            row = csvFile[csvFile['Date'] == date]
            # choose whether to maek it date or run date
            if len(row.values) > 1:
                if str(row.values) not in seen:
                    seen.add(str(row.values))
                    i, runCounter = \
                        duplicateDates(row.values, i, runCounter, workoutFrequency//len(dates))
            elif row['Workout Type'].values[0] == 'Running':
                runContext(runCounter, date, row[' DURATION'].values[0], row['Walking + Running (mi)'].values[0])
                runCounter += 1
                break
            else:
                strengthContext(i, date)
                i += 1
    additionalContext(weekNum)
    doc.render(context)
    doc.save(base_dir / "documents" / f"ActivityLogWeek_{weekNum}.docx")

csvFile["Week"] = 0
rowIndex = 0
for key, values in all_days.items():
    for date in values:
        csvFile.loc[rowIndex, 'Week'] = key
        rowIndex += 1
csvFile.to_csv(HEALTH_DATA, index=False)

for index, row in csvFile.iterrows():
    if row.values[2] == 'Running':
        runTime = str(row.values[3])
        runTime = int(runTime[0:len(runTime)-1])
        if runTime <= 6:
            csvFile.loc[index, 'Walking + Running (mi)'] = 1.00
        else:
            csvFile.loc[index, 'Walking + Running (mi)'] = round(runTime*60/390, 2)
csvFile.to_csv(HEALTH_DATA, index=False)

for index, row in csvFile.iterrows():
    csvFile.loc[index, 'Walking + Running (mi)'] = 0.0
csvFile.to_csv(HEALTH_DATA, index=False)

for index, row in csvFile.iterrows():
    if str(row.values[1]) == 'nan':
        csvFile.loc[index, 'Walking + Running (mi)'] = 0.0
csvFile.to_csv(HEALTH_DATA, index=False)

for index, row in csvFile.iterrows():
    if str(row.values[3]) == 'nan':
        csvFile.drop(index, axis=0, inplace=True)
csvFile.to_csv(HEALTH_DATA, index=False)