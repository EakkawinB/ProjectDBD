from datetime import datetime, timedelta
from config import RUN_START_TIME, RUN_END_TIME

def format_elapsed(start_dt, now_dt):
    elapsed = now_dt - start_dt
    total_sec = int(elapsed.total_seconds())

    days = total_sec // 86400
    hours = (total_sec % 86400) // 3600
    minutes = (total_sec % 3600) // 60

    return f"{days} วัน {hours} ชม {minutes} นาที", elapsed

def estimate_avg_time_per_company(start_dt, now_dt, processed):
    if processed <= 0:
        return None
    return (now_dt - start_dt).total_seconds() / processed

def estimate_eta(now_dt, remaining_seconds):
    cur = now_dt
    remaining = remaining_seconds

    while remaining > 0:
        cur_time = cur.time()

        if cur_time < RUN_START_TIME:
            cur = datetime.combine(cur.date(), RUN_START_TIME)
        elif cur_time > RUN_END_TIME:
            cur = datetime.combine(cur.date() + timedelta(days=1), RUN_START_TIME)

        end_today = datetime.combine(cur.date(), RUN_END_TIME)
        usable_today = (end_today - cur).total_seconds()

        if remaining <= usable_today:
            return cur + timedelta(seconds=remaining)

        remaining -= usable_today
        cur = datetime.combine(cur.date() + timedelta(days=1), RUN_START_TIME)

    return cur