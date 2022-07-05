import datetime
from json.decoder import JSONDecodeError
import time
from datetime import timedelta
import json
import os


user = os.getlogin()
data_folder = r"C:\\Users\\%s\\AppData\\Roaming\\PKaser" %user
tracker_file =data_folder + "\\nt-json-files\\ttimes_stored.json"
total_times = data_folder + "\\nt-json-files\\total_times.json"

def get_today_date():
    today_date = str(datetime.datetime.now().date()).replace("-","_")
    return today_date

today_date = get_today_date()

def load_times():
    if os.path.exists(tracker_file):
        with open(tracker_file) as json_file:
            d_time_tracking =  json.load(json_file)
        return d_time_tracking
    else:
        d_time_tracking = {today_date:{}}
        with open(tracker_file, "w") as json_file:
            json.dump(d_time_tracking, json_file)
        return d_time_tracking

def load_total_times():
    if os.path.exists(total_times):
        with open(total_times) as json_file:
            d_time_tracking =  json.load(json_file)
        json_file.close()
        return d_time_tracking
    else:
        d_time_tracking = {today_date:{}}
        with open(total_times, "w") as json_file:
            json.dump(d_time_tracking, json_file)
        json_file.close()
        return d_time_tracking


def keys_exists(element, *keys):
    '''
    Check if *keys (nested) exists in `element` (dict).
    '''
    if not isinstance(element, dict):
        raise AttributeError('keys_exists() expects dict as first argument.')
    if len(keys) == 0:
        raise AttributeError('keys_exists() expects at least two arguments, one given.')

    _element = element
    for key in keys:
        try:
            _element = _element[key]
        except KeyError:
            return False
    return True

def get_now():
    now = datetime.datetime.now()
    now = datetime.datetime.strftime(now, '%H:%M:%S')
    return str(now)


def update_time(currently_selected, last_selected):

    times = load_times()
    try:
        d_time_tracking = times[today_date]
    except KeyError: # Error happens when today_date is not in dictionary.
        new_day = {today_date:{}}
        times.update(new_day)
        d_time_tracking = times[today_date]

    #print("BEFORE: %s"%d_time_tracking)
    if keys_exists(d_time_tracking,last_selected):
        print("if keys_exists(d_time_tracking,last_selected): - ran - timetracker.py")
        for time_lists in d_time_tracking[last_selected]:
            print("for time_lists in d_time_tracking[last_selected]: - ran - time_lists: %s-%s - timetracker.py"%(last_selected,time_lists))
            for end_time in time_lists:
                print("for end_time in time_lists: %s - timetracker.py"%(end_time))
                if end_time == "00:00:00" and last_selected != currently_selected:
                    print("if end_time == '00:00:00': %s"%(end_time))
                    i_end_time = time_lists.index(end_time)
                    time_lists[i_end_time] = get_now()
                    try:
                        print(" try - d_time_tracking[currently_selected].append([get_now(),'00:00:00'])")
                        d_time_tracking[currently_selected].append([get_now(),"00:00:00"])
                    except KeyError:
                        print("except KeyError - d_time_tracking[currently_selected]= [[get_now(),'00:00:00']]")
                        d_time_tracking[currently_selected]= [[get_now(),"00:00:00"]]
    else:
        print("else - d_time_tracking[currently_selected]= [[get_now(),'00:00:00']]")
        d_time_tracking[currently_selected]= [[get_now(),"00:00:00"]]


    with open(tracker_file, 'w') as fp:
        json.dump(times, fp)

    #print("AFTER: %s"%d_time_tracking) 


def calculate_times(d_time_tracking):
    # dictionary variable keeps track of time as users switches through cases.
    dictionary = load_times()
    # dictionary_two variable keeps track of total time user has spent of each case for the day.
    dictionary_two = {today_date:{}}
    d_total_times = dictionary_two[today_date]

    d_time_tracking = dictionary[today_date]
    for each_case in d_time_tracking:
        sum_time = datetime.timedelta()
        if keys_exists(d_total_times,each_case):
            pass
        else:
            d_total_times[each_case] = ""
            len_time_lists = len(d_time_tracking[each_case])
            for a_list_of_times in d_time_tracking[each_case]:
                if not "00:00:00" in a_list_of_times:
                    time_1 = datetime.datetime.strptime(a_list_of_times[0],"%H:%M:%S")
                    time_2 = datetime.datetime.strptime(a_list_of_times[1],"%H:%M:%S")
                    time_difference = time_2 - time_1
                    total_seconds = time_difference.total_seconds()
                    sum_time += datetime.timedelta(seconds=total_seconds)
                    
        dictionary_two[today_date][each_case] = str(sum_time)
    print("Total Times: %s"%d_total_times)
    with open(total_times, 'w') as fp:
        json.dump(dictionary_two, fp)

def calculate_selected_times(list_of_selected_items):
    sum_time = datetime.timedelta()
    for i in list_of_selected_items:
        (h,m,s) = i.split(':')
        d = datetime.timedelta(hours=int(h), minutes=int(m), seconds=int(s))
        sum_time += d

    return sum_time

#listoftimes = ['0:11:19', '0:16:13', '0:00:08', '0:09:21']
#calculate_selected_times(listoftimes)

# What will go in on_select Function.
#last_selected = "P123456_RAMI"
#currently_selected =  "P123456_TUCKERONE"
#update_time(currently_selected,last_selected)
#d_time_tracking = load_times()
#calculate_times(d_time_tracking)




