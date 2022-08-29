from __future__ import print_function
import datetime
import pytz
import json
import sys
from google_api import get_calendar_service


def prep_work_ranges(config):
    tz = pytz.timezone(config["timezone"])
    today = datetime.datetime.now(tz).date()
    min_time = datetime.datetime.now(tz) + datetime.timedelta(
        hours=config["hours_till_first_meeting"]
    )
    min_time = min_time.replace(minute=0, second=0, microsecond=0)
    ranges = []
    for i in range(config["days_forward"]):
        day = today + datetime.timedelta(days=i)
        dow = day.strftime("%a")
        if dow not in config["days"]:
            continue
        for avail_start, avail_end in config["days"][dow]:
            start = tz.localize(
                datetime.datetime.combine(
                    day, datetime.time.fromisoformat(avail_start)
                ),
                is_dst=None,
            )
            end = tz.localize(
                datetime.datetime.combine(day, datetime.time.fromisoformat(avail_end)),
                is_dst=None,
            )
            if datetime.time.fromisoformat(avail_end) < datetime.time.fromisoformat(
                avail_start
            ):
                end = end + datetime.timedelta(days=1)
            if end < min_time:
                continue
            if start < min_time:
                start = min_time
            assert start < end
            ranges.append([start, end])
    return ranges


def get_busy_ranges(config, service, calendar_id):
    tz = pytz.timezone(config["timezone"])

    # Call the Calendar API
    now = datetime.datetime.now(tz).isoformat()
    end = (
        datetime.datetime.now(tz) + datetime.timedelta(days=config["days_forward"])
    ).isoformat()
    body = dict(
        timeMin=now,
        timeMax=end,
        items=[{"id": calendar_id}],
        timeZone=config["timezone"],
    )
    freebusy = service.freebusy().query(body=body).execute()
    ranges = []
    for r in freebusy["calendars"][calendar_id]["busy"]:
        start = datetime.datetime.fromisoformat(r["start"])
        end = datetime.datetime.fromisoformat(r["end"]) + datetime.timedelta(
            minutes=config["meeting_spare_after"]
        )
        ranges.append([start, end])
    return ranges


def ceil_dt(dt, base, minutes):
    return dt + (base - dt) % datetime.timedelta(minutes=minutes)


def combine_ranges(config, free, busy):
    ranges = []
    free_idx = 0
    busy_idx = 0

    while free_idx < len(free):
        while busy_idx < len(busy) and busy[busy_idx][1] <= free[free_idx][0]:
            busy_idx += 1
        if busy_idx == len(busy) or busy[busy_idx][0] >= free[free_idx][1]:
            # no busy or busy starts after end of this free - take all free
            start = free[free_idx][0]
            end = free[free_idx][1]
            free_idx += 1
        elif busy[busy_idx][0] <= free[free_idx][0]:
            if busy[busy_idx][1] >= free[free_idx][1]:
                # busy over all this free range, skip it
                free_idx += 1
                continue
            else:
                # busy starts before this free range, and ends in middle of it
                # update start of this free range, rounded up to nearest boundary
                new_free_start = ceil_dt(
                    busy[busy_idx][1],
                    free[free_idx][0],
                    config["meeting_length_minutes"],
                )
                if new_free_start < free[free_idx][1]:
                    free[free_idx][0] = new_free_start
                else:
                    free_idx += 1
                busy_idx += 1
                continue
        elif busy[busy_idx][1] >= free[free_idx][1]:
            # busy starts in this free range, and ends after it
            start = free[free_idx][0]
            end = busy[busy_idx][0]
            free_idx += 1
        else:
            # busy starts and ends within this free range
            start = free[free_idx][0]
            end = busy[busy_idx][0]
            # update start of this free range, rounded up to nearest boundary
            new_free_start = ceil_dt(
                busy[busy_idx][1], free[free_idx][0], config["meeting_length_minutes"]
            )
            if new_free_start < free[free_idx][1]:
                free[free_idx][0] = new_free_start
            else:
                free_idx += 1
            busy_idx += 1

        length = (end - start).total_seconds() // 60
        if length >= config["meeting_length_minutes"]:
            ranges.append([start, end])

    return ranges


def print_ranges(config, ranges):
    tz = pytz.timezone(config["show_timezone"])

    def format_time(t):
        if config["show_24hr"]:
            return t.strftime("%-H:%M")
        else:
            if t.minute != 0:
                return t.strftime("%-I:%M%p").lower()
            else:
                return t.strftime("%-I%p").lower()

    # https://stackoverflow.com/questions/5891555/display-the-date-like-may-5th-using-pythons-strftime
    def suffix(d):
        return "th" if 11 <= d <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(d % 10, "th")

    def custom_strftime(t, format):
        return t.strftime(format).replace("{S}", str(t.day) + suffix(t.day))

    day_ranges = []
    day = None
    for r0, r1 in ranges:
        r0 = r0.astimezone(tz)
        r1 = r1.astimezone(tz)

        if day == r0.weekday():
            day_ranges[-1].append([r0, r1])
        else:
            day_ranges.append([[r0, r1]])
            day = r0.weekday()
    for day_list in day_ranges:
        range_str_list = [
            f"{format_time(r0)} - {format_time(r1)}" for r0, r1 in day_list
        ]
        day = custom_strftime(day_list[0][0], "%a (%b {S}):")
        print(f" * {day:14s} {', '.join(range_str_list).lower()}")


def main():
    service = get_calendar_service()

    calendar_list = service.calendarList().list().execute()
    calendar = calendar_list["items"][0]

    config = json.load(open(sys.argv[1], "r"))

    timezone_str = (
        ""
        if config["show_timezone_name"] is None
        else f" (all {config['show_timezone_name']})"
    )
    print(f"Availability for next few days{timezone_str}:")

    free = prep_work_ranges(config)
    busy = get_busy_ranges(config, service, calendar["id"])
    ranges = combine_ranges(config, free, busy)
    print_ranges(config, ranges)


if __name__ == "__main__":
    main()
