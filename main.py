import yaml
import os
import requests
import json
import time
import atexit
import logging
from datetime import datetime, date, timedelta
import pytz
import msal

from graph import Graph


def parse_cal_date(dateStr):
    d = datetime.strptime(dateStr, "%Y-%m-%dT%H:%M:%S.0000000")
    exchangeTz = pytz.utc
    localTz = pytz.timezone('Asia/Kolkata')
    return exchangeTz.localize(d).astimezone(localTz);

def format_orgmode_date(dateObj):
  return dateObj.strftime("%Y-%m-%d %H:%M")

def format_orgmode_time(dateObj):
    return dateObj.strftime("%H:%M")

def parse_entry(appt):
    apptstart = parse_cal_date(appt['start']['dateTime'])
    apptend = parse_cal_date(appt['end']['dateTime'])
    tags = ":meeting:work:"
    if "Out of office" in appt['subject']:
        return
    if appt['categories']:
        tags = tags + ":" + ":".join([t.lower() for t in appt['categories']]) + ":"

    if apptstart.date() == apptend.date():
        dateStr = "<" +  format_orgmode_date(apptstart) + "-" + format_orgmode_time(apptend) + ">"
    else:
        dateStr = "<" +  format_orgmode_date(apptstart) + ">--<" + format_orgmode_date(apptend) + ">"
    body = appt['bodyPreview'].translate({ord('\r'): None})

    print(f"* {dateStr} {appt['subject']} {tags}")

    print(":PROPERTIES:")
    if appt['location']['displayName'] is not None:
        print(f":LOCATION: {appt['location']['displayName']}")

    if appt['onlineMeeting'] is not None:
        print(f":JOINURL: {appt['onlineMeeting']['joinUrl']}")

    print(f":RESPONSE: {appt['responseStatus']['response']}")
    print(":END:")

    print(f"{body}")

    print("")

def dump_to_org(entries):
    for appt in entries:
        parse_entry(appt)

def main():
    with open("config.yaml", "r") as stream:
        try:
            config = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
            return

    azure_settings = config["azure"]
    token_cache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.bin"):
        token_cache.deserialize(open("token_cache.bin", "r").read())
        atexit.register(lambda:
            open("token_cache.bin", "w").write(token_cache.serialize())
            # Hint: The following optional line persists only when state changed
            if token_cache.has_state_changed else None
        )
    graph = Graph(azure_settings, token_cache)
    entries = graph.get_calendar_entries()
    dump_to_org(entries["value"])
    # print(json.dumps(entries, indent=4))

main()
