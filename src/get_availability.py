from __future__ import print_function
import datetime
from multiprocessing.dummy import Value
import pytz
import json
from tabulate import tabulate
from google_api import get_calendar_service
from optparse import OptionParser
from dateutil.parser import parser

_CONFIG_DEFAULT_KEYS = {
    "days_forward": 14,
    "hours_till_first_meeting": 3,
    "meeting_length_minutes": 30,
    "meeting_spare_before": 0,
    "meeting_spare_after": 0,
    "show_timezone": "America/Los_Angeles",
    "show_24hr": False,
    "show_timezone_name": "PT",
    "week_starts_on_sunday": False,
    "local_timezone": "America/Los_Angeles",
    "days": {
        "Mon": [["9am", "5pm"]],
        "Tue": [["9am", "5pm"]],
        "Wed": [["9am", "5pm"]],
        "Thu": [["9am", "5pm"]],
        "Fri": [["9am", "2pm"]],
        "Sat": [],
        "Sun": [],
    },
}


def _parse_timestr(timestr):
    # Alternative version for just ISO: datetime.time.fromisoformat(timestr)
    return parser().parse(timestr).time()


def prep_work_ranges(config):
    for key in config.keys():
        assert key in _CONFIG_DEFAULT_KEYS, f"Unknown key in configuration: {key}"
    for key, val in _CONFIG_DEFAULT_KEYS.items():
        config.setdefault(key, val)

    tz = pytz.timezone(config["local_timezone"])
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
                datetime.datetime.combine(day, _parse_timestr(avail_start)),
                is_dst=None,
            )
            end = tz.localize(
                datetime.datetime.combine(day, _parse_timestr(avail_end)),
                is_dst=None,
            )
            if _parse_timestr(avail_end) < _parse_timestr(avail_start):
                end = end + datetime.timedelta(days=1)
            if end < min_time:
                continue
            if start < min_time:
                start = min_time
            assert start <= end
            ranges.append([start, end])
    return ranges


def get_busy_ranges(config, service, calendar_id):
    tz = pytz.timezone(config["local_timezone"])

    # Call the Calendar API
    now = datetime.datetime.now(tz).isoformat()
    end = (
        datetime.datetime.now(tz) + datetime.timedelta(days=config["days_forward"])
    ).isoformat()
    body = dict(
        timeMin=now,
        timeMax=end,
        items=[{"id": calendar_id}],
        timeZone=config["local_timezone"],
    )
    freebusy = service.freebusy().query(body=body).execute()
    ranges = []
    for r in freebusy["calendars"][calendar_id]["busy"]:
        start = datetime.datetime.fromisoformat(r["start"]) - datetime.timedelta(
            minutes=config["meeting_spare_before"]
        )
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


def _to_weekday(t, config):
    return (t.weekday() + int(config["week_starts_on_sunday"])) % 7


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
    last_weekday = _to_weekday(day_ranges[0][0][0], config)
    for day_list in day_ranges:
        range_str_list = [
            f"{format_time(r0)} - {format_time(r1)}" for r0, r1 in day_list
        ]
        if _to_weekday(day_list[0][0], config) < last_weekday:
            print("Next week:")
        last_weekday = _to_weekday(day_list[0][0], config)
        day = custom_strftime(day_list[0][0], "%a (%b {S}):")
        print(f" * {day:14s} {', '.join(range_str_list).lower()}")


def get_args():
    parser = OptionParser()
    parser.add_option(
        "-l",
        "--list",
        dest="list",
        help="list calendars",
        action="store_true",
        default=False,
    )
    parser.add_option(
        "-c",
        "--calendar",
        action="append",
        dest="cal",
        help="choose a calendar for busy times (multiple allowed)",
    )
    parser.add_option(
        "-t",
        "--time_config",
        metavar="FILE",
        dest="conf",
        help="choose a configuration JSON file",
    )
    parser.add_option(
        "-o",
        "--opt",
        action="append",
        dest="opt",
        metavar="OPTNAME=VALUE",
        help="override a configuration option (multiple allowed), use OPT=VAL format. See -O for possible options",
    )
    parser.add_option(
        "-O",
        "--list-config-options",
        dest="list_conf_options",
        help="list possible configuration options",
        action="store_true",
        default=False,
    )

    (options, args) = parser.parse_args()

    if options.list_conf_options:
        print(
            tabulate(
                [[key, str(val)] for key, val in _CONFIG_DEFAULT_KEYS.items()],
                headers=["Option name", "Default value"],
                maxcolwidths=[None, 50],
            )
        )
        return None

    conf_override = {}
    if options.opt:
        for opt_entry in options.opt:
            if "=" not in opt_entry:
                parser.error(
                    "Any configuration option should be provided as OPTNAME=VALUE"
                )
            name, val_str = opt_entry.split("=")
            if name not in _CONFIG_DEFAULT_KEYS:
                parser.error(f"Unknown option {name}, use -O to see possible options")
            try:
                if isinstance(_CONFIG_DEFAULT_KEYS[name], (dict, list)):
                    val = json.loads(val_str)
                elif isinstance(_CONFIG_DEFAULT_KEYS[name], bool):
                    try:
                        val = int(val_str)
                    except ValueError:
                        try:
                            val = {"false": False, "true": True}[val_str.lower()]
                        except KeyError:
                            parser.error(
                                f"Bad type for option {name}, should be like a boolean, but is '{val_str}'"
                            )
                    val = bool(val)
                else:
                    val = type(_CONFIG_DEFAULT_KEYS[name])(val_str)
                conf_override[name] = val
            except ValueError as e:
                parser.error(
                    f"Bad type for option {name}, should be like '{_CONFIG_DEFAULT_KEYS[name]}' but is '{val_str}'. Error is: {e}"
                )

    if options.list and options.cal:
        parser.error("options -l and -c are mutually exclusive")
    # if options.conf is None:
    #    parser.error("must set time configuration with -t")
    if options.cal is None and not options.list:
        parser.error("must set either -c or -l")
    return options, conf_override


def _order_cal_list(cal):
    return (cal["accessRole"] != "owner", -len(cal["defaultReminders"]), cal["id"])


def main():
    opts, conf_override = get_args()
    if opts is None:
        return

    service = get_calendar_service()
    calendar_list = service.calendarList().list().execute()

    chosen_cals = []
    not_found = []
    if opts.cal is not None:
        calendar_by_ids = {cal["id"]: cal for cal in calendar_list["items"]}
        for cal in opts.cal:
            if cal not in calendar_by_ids:
                not_found.append(cal)
            else:
                chosen_cals.append(calendar_by_ids[cal])
        if len(not_found) > 0:
            print(f"Calendars {not_found} not found! Possible calendars are:")

    if len(not_found) > 0 or opts.list:
        table = [
            (cal["id"], cal["summary"])
            for cal in sorted(calendar_list["items"], key=_order_cal_list)
        ]
        print(tabulate(table, headers=["Id", "Name"]))
        return

    if opts.conf:
        config = json.load(open(opts.conf, "r"))
    else:
        config = _CONFIG_DEFAULT_KEYS.copy()
    config.update(conf_override)

    timezone_str = (
        ""
        if config["show_timezone_name"] is None
        else f" (all {config['show_timezone_name']})"
    )

    free = prep_work_ranges(config)
    for chosen_cal in chosen_cals:
        busy = get_busy_ranges(config, service, chosen_cal["id"])
        free = combine_ranges(config, free, busy)

    print(f"Availability for next few days{timezone_str}:")
    print_ranges(config, free)


if __name__ == "__main__":
    main()
