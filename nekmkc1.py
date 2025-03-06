#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# Export the Hungarian Opera's "munkaelrendelés" Excel (converted to csv by LibreOffice) to a iCalendar that can be fed to Google Calendar (or other calendar app).
# The second column contains the musician's names in uppercase - the leftmost one is relevant, X in this column marks an appearance - rows describing the play.
# Not all rows have dates, as in the xlsx/PDF versions the first two columns are merged for multiple rows - this is the date of the play.
# The first parameter is the musician's name (UTF-8) and the second parameter is the csv input file's name.
# @author Dr. Nagy Elemér Károly
# @license Unlicense or public domain or CC Zero - whichever you choose ;)
# version 1.0 rc1 (works with the first test file received)

# Check Python version
import sys
if sys.version_info[0]<3:
	print("This script requries Python 3, detected Python version is '%s'." % sys.version_info)
	exit(-3)
else:
	if (sys.version_info[0]==3) and (sys.version_info[1]<11):
		print("ERROR! Tested only on Python 3.11 / Debian 12 - please upgrade (@see venv!) test+fix the code or contact the author.");
		exit(-4)

import locale, os.path, csv, re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# Check arguments
if len(sys.argv)<2:
	print("Usage: %s INPUT_FILE MUSICIAN_NAME" % sys.argv[0]);
	print("INPUT_FILE is a csv, second row contains musician names in uppercase, 'x' or 'X' in that column denotes an assignment.")
	print("All other rows with empty first column should be ignored until a non-empty first column (date) is found.")
	print("Column 0 is date (if empty, use previous value), column 3/4 is from/to, 5-7 are details.")
	exit(-1)

if not os.path.isfile(sys.argv[1]):
	print("ERROR! '%s' is not an INPUT_FILE." % sys.argv[1])
	exit(-2)

# Parse arguments
input_file=sys.argv[1].strip()
musician_name=sys.argv[2].upper().strip()
print("INPUT_FILE is '%s'." % input_file)
print("MUSICIAN_NAME is '%s'." %  musician_name)
output_file="%s.ics" % input_file
print("OUPUT_FILE is '%s'." % output_file)

# Internal variables
target_column=-1
skip_rows=True
locale.setlocale(locale.LC_ALL, 'hu_HU.UTF-8')

# Process the input CSV line by line, print output .icsv on the way
with open(input_file, mode='r') as infile:
	input=csv.reader(infile)
	with open(output_file, mode='w') as outfile:
		outfile.write("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:nekmkc1.py\r\n") # Header, see iCalendar spec.
		event_count=0 # Count found events, print at the end, user could compare this with Google Calendar import stats.

		for i, line in enumerate(input):
			if i==1: # Second line, this should containt the musicians' names - first occurence is visible and relevant, the rest is internal statistics?
				for j, item in enumerate(line):
					if item==musician_name:
						target_column=j
						break
				if target_column==-1:
					print("Musician '%s' not found in second row, aborting." % musician_name);
					exit -3;
				else:
					print("Musician '%s' found in second row in column %d ;)" % (musician_name, target_column));
			else: # All other lines either have a month+day in the first cell or are empty and the previous date should be used (before the first date, ignore all).
				item=line[0].strip()
				if item!='': # We found a non-empty (hopefully date) field, so we are through the header, we don't need to skip any more rows.
					skip_rows=False
					# Date does not contain year, so it can be this year or next year - we guess whichever is closer in time and for that, we need to parse as both.
					# This is a quite solid strategy as this file is generated for 2-3-4 month periods and often changed, hence we don't know when we run this
					# We could also use the second column, which is the day of the week - TODO/FIXME/consider this
					now=datetime.utcnow()
					thisyear=now.strftime('%Y')
					nextyear=(now+timedelta(days = 365)).strftime('%Y')
					dateCY=datetime.strptime("%s %s" % (thisyear, item), "%Y %B %d.") # Date with current year added to month+day.
					dateNY=datetime.strptime("%s %s" % (nextyear, item), "%Y %B %d.") # Date with next year added to month+day.
					targetdate=dateCY if abs(dateCY-now)<abs(dateNY-now) else dateNY # How I miss my precious ?: operator...
			if not skip_rows:
				# If we have a valid date and an X in the column with the musician's name, this line is a play in which the musician performs, so we'll export it
				if line[target_column].strip()=='':
					pass # No performance in this play
				elif line[target_column].upper().strip()=='X':
					# The Hungarian Opera is located in Budapest and Hungary is a single-timezone country, so times are of course local times.
					fromTS=datetime.strptime("%s %s" % (targetdate.strftime("%Y/%B/%d"), line[3]), "%Y/%B/%d %H:%M").replace(tzinfo=ZoneInfo("Europe/Budapest"))
					toTS=datetime.strptime("%s %s" % (targetdate.strftime("%Y/%B/%d"), line[4]), "%Y/%B/%d %H:%M").replace(tzinfo=ZoneInfo("Europe/Budapest"))
					# Due to daylight savings, we have to check if the target local date is in CET or in CEST - the iCalendar format prefers Zulu time / UTC
					fromTS_UTC=datetime.strptime("%s %s" % (targetdate.strftime("%Y/%B/%d"), line[3]), "%Y/%B/%d %H:%M").replace(tzinfo=ZoneInfo("UTC"))
					# We'll get the offset of the two timezones - but due to precision issues, this has values like "-1 day +22:59:9999995", so...
					timezone_offset=round((fromTS-fromTS_UTC).total_seconds()/60) # ... we round and convert it to minutes, and now has values like '-60' or '-120'
					fromTS+=timedelta(minutes=timezone_offset) # TS: timestamp
					toTS+=timedelta(minutes=timezone_offset)
					fromZ=fromTS.strftime("%Y%m%dT%H%M00Z"); # Z: Zulu time (military lingo for UTC)
					toZ=toTS.strftime("%Y%m%dT%H%M00Z");
					exportZ=(datetime.now()+timedelta(minutes=timezone_offset)).strftime("%Y%m%dT%H%M00Z");
					# We should have an UID for each event, and that is totally missing from this Excel sheet. The same sheet is used in a rolling fashion, so no RowID.
					# We are really struck with this one, as removing/replacing a musician will cause the event to disappear from the iCalendar export...
					# ... and thus re-importing it for refresh will leave the old event in the calendar, in lack of real UIDs.
					# TODO:FIXME: add zero-length events with valid UIDs when there is no X in targetcolumn to make refreshes work better?
					# Anyway, we are trying to build and usable UID form the date, time, type and name of the play, removing non-ascii alphanumeric characters
					uid=re.sub('[^0-9a-zA-Z]+', '', "%s%s%s%s%s" % (targetdate, line[3], line[4], line[6], line[7]))
					# For some practices, only the room is used for 'Location' but Google Maps can't find that building, so add it
					location="Eiffel Műhelyház, %s" % line[5] if "terem" in line[5] else line[5]

					# Finally, write the event block with all the collected data.
					outfile.write("BEGIN:VEVENT\r\nUID:%s\r\nDTSTAMP:%s\r\nDTSTART:%s\r\nDTEND:%s\r\nSUMMARY:%s: %s\r\nLOCATION:%s\r\nEND:VEVENT\r\n" %
						(uid, exportZ, fromZ, toZ, line[7], line[6], location))
					event_count+=1
				else:
					print("FATAL! Item expected to be either '' or 'X' but is '%s', aborting..." % item)
					exit(-5)
		outfile.write("END:VCALENDAR\r\n")
		print("Written %d events to '%s'." % (event_count, output_file))
