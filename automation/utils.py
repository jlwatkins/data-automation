import time
import datetime
import calendar

def convert_datetime_string_to_unix_utc(time_string):
    time_format_string = '%Y-%m-%dT%H:%M:%S.%fZ'
    unix_seconds_utc = None
    try:
        date = datetime.datetime.strptime(time_string, time_format_string)
        unix_seconds_utc = calendar.timegm(date.timetuple())
    except ValueError as e:
        print("Error: Could not parse date: " + str(e))
        return None

    return unix_seconds_utc

def convert_datetime_string_to_unix_local(time_string):
    time_format_string = '%Y-%m-%dT%H:%M:%S.%fZ'
    unix_seconds_utc = None
    try:
        date = datetime.datetime.strptime(time_string, time_format_string)
        unix_seconds_utc = time.mktime(date.timetuple())
    except ValueError as e:
        print("Error: Could not parse date: " + str(e))
        return None

    return unix_seconds_utc

def get_time_diff_local_utc_seconds():
    return calendar.timegm(time.localtime()) - calendar.timegm(time.gmtime())