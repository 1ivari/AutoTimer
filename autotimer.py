from __future__ import print_function
import time
from os import system
from activity import *
import json
import datetime
import sys
if sys.platform in ['Windows', 'win32', 'cygwin']:
    import win32gui
    from pywinauto import Application
elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
    from AppKit import NSWorkspace
    from Foundation import *
elif sys.platform in ['linux', 'linux2']:
    import linux as l

active_window_name = ""
activity_name = ""
start_time = datetime.datetime.now()
activeList = AcitivyList([])
first_time = True

json_name = "log.json"

# split URL returned by Pywin and take the first part e.g. youtube.com from youtube.com/video/292929


def url_to_name(url):
    string_list = url.split('/')
    return string_list[0]

# cleanup common active windows to minimize the variety of activities


def getCategory(activityName):
    # set identifiers for Outlook
    retval = activityName
    outlookIdentifiers = ['Outlook', 'RE:', 'Inbox']
    teamsIdentifiers = ['Teams']
    webIdentifiers = ["Chrome", "Edge"]
    excelIdentifiers = ["Exce"]
    wordIdentifiers = ["Word"]
    pptIdentifiers = ["Powerpoint"]
    pdfIdentifiers = ["Adobe", "pdf"]

    if any(substring in activityName for substring in outlookIdentifiers):
        retval = "Outlook"

    elif any(substring in activityName for substring in webIdentifiers):
        retval = "Web"

    elif any(substring in activityName for substring in excelIdentifiers):
        retval = "Excel"
        
    elif any(substring in activityName for substring in teamsIdentifiers):
        retval = "Teams"
    
    elif any(substring in activityName for substring in wordIdentifiers):
        retval = "Word"
    
    elif any(substring in activityName for substring in pptIdentifiers):
        retval = "Powerpoint"
        
    elif any(substring in activityName for substring in pdfIdentifiers):
        retval = "PDF"

    return retval


"""     if activityName in outlookIdentifiers:
        retval = "Outlook:" + activityName

    elif activityName in webIdentifiers:
        retval = "Web:" + activityName """


# get the window that is active at a given time


def get_active_window():
    _active_window_name = None
    # Windows used
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        window = win32gui.GetForegroundWindow()
        _active_window_name = win32gui.GetWindowText(window)
    # Mac used
    elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
        _active_window_name = (NSWorkspace.sharedWorkspace()
                               .activeApplication()['NSApplicationName'])
    # Other platforms not supported
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)
    return _active_window_name

# get the url from chrome address bar - consider adding Edge support here


def get_chrome_url():

   # windows platform
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        app = Application(backend='uia')
        app.connect(title_re=".*Chrome.*")
        dlg = app.top_window()
        url = dlg.child_window(
            title="Address and search bar", control_type="Edit").get_value()
        return url

    # MacOs
    elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
        textOfMyScript = """tell app "google chrome" to get the url of the active tab of window 1"""
        s = NSAppleScript.initWithSource_(
            NSAppleScript.alloc(), textOfMyScript)
        results, err = s.executeAndReturnError_(None)
        return results.stringValue()

    # Unsupported platforms
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)
    return _active_window_name


def get_edge_url():
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        app = Application(backend='uia')
        app.connect(title_re=".*Edge.*", found_index=0)
        dlg = app.top_window()
        wrapper = dlg.child_window(title="App bar", control_type="ToolBar")
        url = wrapper.descendants(control_type='Edit')[0]
        retval = url.get_value().split('/')
        return retval[2]


try:
    activeList.initialize_me()
except Exception:
    print('No json')


try:
    while True:
        previous_site = ""
        if sys.platform not in ['linux', 'linux2']:
            new_window_name = get_active_window()
            if 'Google Chrome' in new_window_name:
                try:
                    new_window_name = url_to_name(get_chrome_url())
                    new_window_name = "Chrome - " + new_window_name
                except:
                    new_window_name = "Chrome - " + new_window_name
            elif 'Edge' in new_window_name:
                try:
                    new_window_name = get_edge_url()
                    new_window_name = "Edge - " + new_window_name
                except:
                    new_window_name = "Edge - " + new_window_name
        if sys.platform in ['linux', 'linux2']:
            new_window_name = l.get_active_window_x()
            if 'Google Chrome' in new_window_name:
                new_window_name = l.get_chrome_url_x()

        if active_window_name != new_window_name:
            activity_name = active_window_name
            print(activity_name)

            if not first_time:
                end_time = datetime.datetime.now()
                time_entry = TimeEntry(start_time, end_time, 0, 0, 0, 0)
                time_entry._get_specific_times()

                exists = False
                for activity in activeList.activities:
                    if activity.name == activity_name:
                        exists = True
                        activity.time_entries.append(time_entry)

                if not exists:
                    category = getCategory(active_window_name)
                    print(category)
                    activity = Activity(category, activity_name, [time_entry])
                    activeList.activities.append(activity)
                with open(json_name, 'w') as json_file:
                    json.dump(activeList.serialize(), json_file,
                              indent=4, sort_keys=True)
                    start_time = datetime.datetime.now()
            first_time = False
            active_window_name = new_window_name

        # sleep for 1 second
        time.sleep(1)

except KeyboardInterrupt:
    with open(json_name, 'w') as json_file:
        json.dump(activeList.serialize(), json_file, indent=4, sort_keys=True)
