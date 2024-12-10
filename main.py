import atexit
import os
import sys
from datetime import datetime
from pathlib import Path
import re

import msal
import pytz
import yaml
from platformdirs import user_config_dir

from graph import Graph


def parse_cal_date(dateStr):
    d = datetime.strptime(dateStr, "%Y-%m-%dT%H:%M:%S.0000000")
    exchangeTz = pytz.utc
    localTz = pytz.timezone('Asia/Kolkata')
    return exchangeTz.localize(d).astimezone(localTz)

def format_orgmode_date(dateObj):
    return dateObj.strftime("%Y-%m-%d %H:%M")

def format_orgmode_time(dateObj):
    return dateObj.strftime("%H:%M")

def get_week_heading(dateObj):
    year, week_num = dateObj.isocalendar()[:2]
    return f"{year}-w{week_num:02}"

def parse_entry(appt):
    apptstart = parse_cal_date(appt['start']['dateTime'])
    apptend = parse_cal_date(appt['end']['dateTime'])
    tags = ":meeting:work:"

    skip_pattern = r"(^Canceled:|PTO|Out of office)"
    if re.match(skip_pattern, appt['subject']):
        return None

    if appt['responseStatus']['response'] == "notResponded":
        return None

    if appt['categories']:
        tags = tags + ":".join([t.lower() for t in appt['categories']]) + ":"

    if apptstart.date() == apptend.date():
        dateStr = "<" + format_orgmode_date(apptstart) + "-" + format_orgmode_time(apptend) + ">"
    else:
        dateStr = "<" + format_orgmode_date(apptstart) + ">--<" + format_orgmode_date(apptend) + ">"
    body = appt['bodyPreview'].translate({ord('\r'): None})

    entry = []
    entry.append(f"*** {dateStr} {appt['subject']} {tags}")
    entry.append(":PROPERTIES:")
    if appt['location']['displayName']:
        entry.append(f":LOCATION: {appt['location']['displayName']}")
    if appt['onlineMeeting']:
        entry.append(f":JOINURL: {appt['onlineMeeting']['joinUrl']}")
    entry.append(f":RESPONSE: {appt['responseStatus']['response']}")
    entry.append(":END:")
    entry.append(body)
    entry.append("")
    return "\n".join(entry)

def dump_to_org(entries, org_file):
    weeks = {}
    for appt in entries:
        parsed_entry = parse_entry(appt)
        if parsed_entry:
            week_heading = get_week_heading(parse_cal_date(appt['start']['dateTime']))
            if week_heading not in weeks:
                weeks[week_heading] = []
            weeks[week_heading].append(parsed_entry)

    # Read the existing org file content
    if os.path.exists(org_file):
        with open(org_file, "r") as f:
            org_content = f.readlines()
    else:
        org_content = []

    # Update org content with new entries
    updated_content = []
    for week_heading, entries in weeks.items():
        heading_found = False
        for line in org_content:
            updated_content.append(line)
            if line.strip() == f"* {week_heading}":
                heading_found = True
                updated_content.extend(entries)
                org_content = org_content[len(updated_content):]
                break
        if not heading_found:
            updated_content.append(f"* {week_heading}\n")
            updated_content.extend(entries)

    # Write updated content back to file
    with open(org_file, "w") as f:
        f.writelines(updated_content)

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <org-file-path>")
        return

    org_file = sys.argv[1]

    config_path = Path(user_config_dir("ol-calendar"))
    config_file = config_path.joinpath("config.yaml")
    if not config_path.exists() or not config_file.exists():
        print("Config file not present: %s" % config_file)
        return

    with open(config_file, "r") as stream:
        try:
            config = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
            return

    azure_settings = config["azure"]
    token_cache = msal.SerializableTokenCache()
    token_cache_file = config_path.joinpath("token_cache.bin")
    if os.path.exists(token_cache_file):
        token_cache.deserialize(open(token_cache_file, "r").read())
        atexit.register(lambda:
            open(token_cache_file, "w").write(token_cache.serialize())
            if token_cache.has_state_changed else None
        )
    graph = Graph(azure_settings, token_cache)
    entries = graph.get_calendar_entries()

    # Update org file
    dump_to_org(entries["value"], org_file)

main()
