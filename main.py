import math
from random import randrange
import pandas as pd
import openpyxl
from datetime import datetime
from docxtpl import DocxTemplate
from pathlib import Path

# Set the path to the Word document template
base_dir = Path(__file__).parent
word_template_path = base_dir / "ActivityLog.docx"

# Load the Word document template into memory
doc = DocxTemplate(word_template_path)

# Load the health data from a CSV file
HEALTH_DATA = base_dir / 'HealthData.csv'
csvFile = pd.read_csv(HEALTH_DATA)

# Load the workout data from a CSV file
WORKOUT_DATA = base_dir / 'workouts.csv'
workCsv = pd.read_csv(WORKOUT_DATA)

# Define constants for time calculations
oneDay = 86400.0
dayOne = datetime(2023, 1, 27).timestamp()

# Create a dictionary to store all the dates by week
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

# Set up an empty set to track whether certain dates have been seen before
seen = set()

# Define a function to add additional context to the Word document template
def additionalContext(weekNum):
    context['stretch1'] = "Pre-workout dynamic stretch routine (upper and lower body)"
    context['time'] = "7 minutes"
    context['weekNumber'] = weekNum

# Define a function to add strength training context to the Word document template
def strengthContext(index, date):
    randonRow = randrange(25)
    context[f"date{index}"] = date
    context[f"strength{index}"] = workCsv.loc[randonRow][0]
    context[f"weight{index}"] = workCsv.loc[randonRow][1]
    context[f"setsReps{index}"] = workCsv.loc[randonRow][2]

# Define a function to add running context to the Word document template
def runContext(i, date, time, distance):
    context[f"runDate{i}"] = date
    context[f"cardio{i}"] = "Running"
    context[f"timeDistance{i}"] = f"{distance} miles, {time} minutes"

# Define a function to duplicate dates in the data
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

# Loop through each week in the dictionary of dates
# Iterate through all_days dictionary containing week numbers and dates
for weekNum, dates in all_days.items():
    context = dict()
    # Calculate workout frequency for the week based on number of dates in the week
    workoutFrequency = math.floor(14/len(dates)) * len(dates)
    i = 1
    # Iterate through each date in the week
    for date in dates:
        runCounter = 1
        # Repeat for each workoutFrequency/len(dates) number of times
        for j in range(workoutFrequency//len(dates)):
            # Retrieve row from csvFile corresponding to current date
            row = csvFile[csvFile['Date'] == date]
            # Check if there are duplicate entries for the same date
            if len(row.values) > 1:
                # If so, check if the row has already been seen before
                if str(row.values) not in seen:
                    # If not, add row to the document and update i and runCounter
                    seen.add(str(row.values))
                    i, runCounter = \
                        duplicateDates(row.values, i, runCounter, workoutFrequency//len(dates))
            # If there are no duplicates, check if the workout type is running
            elif row['Workout Type'].values[0] == 'Running':
                # If so, add running context to the document and update runCounter
                runContext(runCounter, date, row[' DURATION'].values[0], row['Walking + Running (mi)'].values[0])
                runCounter += 1
                break
            # If workout type is not running, add strength training context to the document and update i
            else:
                strengthContext(i, date)
                i += 1
    # Add additional context to the document for the week and render and save the document
    additionalContext(weekNum)
    doc.render(context)
    doc.save(base_dir / "documents" / f"ActivityLogWeek_{weekNum}.docx")

# Add week numbers to csvFile
csvFile["Week"] = 0
rowIndex = 0
for key, values in all_days.items():
    for date in values:
        csvFile.loc[rowIndex, 'Week'] = key
        rowIndex += 1
csvFile.to_csv(HEALTH_DATA, index=False)

# Update Walking + Running (mi) values in csvFile based on duration of running workouts
for index, row in csvFile.iterrows():
    if row.values[2] == 'Running':
        runTime = str(row.values[3])
        runTime = int(runTime[0:len(runTime)-1])
        if runTime <= 6:
            csvFile.loc[index, 'Walking + Running (mi)'] = 1.00
        else:
            csvFile.loc[index, 'Walking + Running (mi)'] = round(runTime*60/390, 2)
csvFile.to_csv(HEALTH_DATA, index=False)

# Reset Walking + Running (mi) values in csvFile to 0
for index, row in csvFile.iterrows():
    csvFile.loc[index, 'Walking + Running (mi)'] = 0.0
csvFile.to_csv(HEALTH_DATA, index=False)

# Update Walking + Running (mi) values in csvFile to 0 if they are NaN
for index, row in csvFile.iterrows():
    if str(row.values[1]) == 'nan':
        csvFile.loc[index, 'Walking + Running (mi)'] = 0.0
csvFile.to_csv(HEALTH_DATA, index=False)

#
