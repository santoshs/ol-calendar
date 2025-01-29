from orgparse import node, load
from pathlib import Path
from platformdirs import user_config_dir
import yaml
import msal
import os
import json
from datetime import datetime
import pytz
from typing import Optional, List, Dict
import re
import sys

from graph import Graph
from orgnode import OrgNode, OrgFile


def parse_cal_date(dateStr):
    d = datetime.strptime(dateStr, "%Y-%m-%dT%H:%M:%S.0000000")
    exchangeTz = pytz.utc
    localTz = pytz.timezone('Asia/Kolkata')
    return exchangeTz.localize(d).astimezone(localTz)


def format_orgmode_date(dateObj):
    return dateObj.strftime("%Y-%m-%d %H:%M")


def format_orgmode_time(dateObj):
    return dateObj.strftime("%H:%M")


def get_calendar(azure_settings):
    config_path = Path(user_config_dir("ol-calendar"))
    token_cache = msal.SerializableTokenCache()
    token_cache_file = config_path.joinpath("token_cache.bin")
    if os.path.exists(token_cache_file):
        token_cache.deserialize(open(token_cache_file, "r").read())
        atexit.register(lambda:
            open(token_cache_file, "w").write(token_cache.serialize())
            if token_cache.has_state_changed else None
        )
    graph = Graph(azure_settings, token_cache)
    calendar = graph.get_calendar_entries()

    skip_pattern = r"(^Canceled:|PTO|out of office|OOO)"
    cal_entries = {}

    entries = {}
    for e in calendar.get('value'):
        if re.search(skip_pattern, e['subject'], re.IGNORECASE) or \
           re.search(skip_pattern, e['bodyPreview'], re.IGNORECASE):
            print(f"Skipping {e['subject']}")
            continue

        entries[e['subject']] = build_entry(e)

    return entries


def build_entry(appt: Dict) -> OrgNode:
    """
    Build an OrgNode from a calendar appointment.
    :param appt: Dictionary containing calendar appointment details.
    :return: An OrgNode instance.
    """
    apptstart = parse_cal_date(appt['start']['dateTime'])
    apptend = parse_cal_date(appt['end']['dateTime'])

    tags = ["meeting", "work"]
    if appt.get('categories'):
        tags.extend(tag.lower() for tag in appt['categories'])

    properties = {}
    if appt['location'].get('displayName'):
        properties['LOCATION'] = appt['location']['displayName']
    if appt.get('onlineMeeting') and appt['onlineMeeting'].get('joinUrl'):
        properties['JOINURL'] = appt['onlineMeeting']['joinUrl']
    if appt.get('id'):
        properties['MEETING_ID'] = appt['id']
    if appt.get('responseStatus') and appt['responseStatus'].get("response"):
        properties["RESPONSE_STATUS"] = appt['responseStatus']["response"]

    timestamp = apptstart

    return OrgNode(
        heading=appt['subject'],
        todo="",
        tags=tags,
        properties=properties,
        body=appt.get('bodyPreview', ""),
        timestamp=timestamp
    )

def update(old, new):
    old.heading = new.heading
    old.body = new.body

    if not old.timestamp or not new.timestamp:
        return

    if old.timestamp < new.timestamp:
        print(f"Updating {old.heading}")
        old.update_timestamp(new.timestamp)

        if old.todo == 'DONE' or old.todo == 'CANCELLED':
            old.change_todo_state("")

    return


def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <org-file-path>")
        return

    org_file = sys.argv[1]

    config_path = Path(user_config_dir("ol-calendar"))
    config_file = config_path.joinpath("config.yaml")
    if not config_path.exists() or not config_file.exists():
        print(f"Config file not present: {config_file}")
        return

    with open(config_file, "r") as stream:
        try:
            config = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
            return

    calendar = get_calendar(config["azure"])
    # calendar = {}
    orgfile = OrgFile.from_file(org_file)

    existing_entries = {}
    for c in orgfile.children:
        if c.heading in existing_entries.keys():
            if existing_entries[c.heading].clocks:
                existing_entries[c.heading].clocks.extend(c.clocks if c.clocks else [])
            else:
                existing_entries[c.heading].clocks = c.clocks
        else:
            existing_entries[c.heading] = c

    orgfile.children = []
    for heading in calendar:
        if calendar[heading].timestamp is None:
            print(f"Missing Timestamp in \n: {calendar[id].to_org_string()}")
            continue
        if heading in existing_entries.keys():
            update(existing_entries[heading], calendar[heading])
            orgfile.children.append(existing_entries.pop(heading))
        else:
            print(f"Adding new entry {heading}")
            orgfile.children.append(calendar[heading])


    existing_keys = list(existing_entries.keys())
    for e in existing_keys:
        if not existing_entries[e].todo:
            print("Removing cancelled/stale event")
            existing_entries.pop(e)

    children = list(existing_entries.values())
    children.extend(orgfile.children)
    orgfile.children = children
    orgfile.to_file(org_file)
    return

if __name__ == "__main__":
    main()
