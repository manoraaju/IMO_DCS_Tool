import datetime
from datetime import datetime
import time
first_time = datetime.strptime('2019-05-17 15:18:00', '%Y-%m-%d %H:%M:%S')
first_time = datetime.strftime(first_time, "%d-%b-%Y %H:%M")
time.sleep(2)
later_time = datetime.strptime('9/12/2019  12:00:00 AM', '%m/%d/%Y %I:%M:%S %p')
duration = later_time - first_time
duration_in_s = duration.total_seconds()
days  = divmod(duration_in_s, 86400)[0]
hours = divmod(duration_in_s, 3600)[0]
minutes = divmod(duration_in_s, 60)[0]
print(days)
print(hours)
print(minutes)

import math
test = -math.inf

print(test)


missing_duration_in_s = -63000

missing_hours = (missing_duration_in_s // 3600)
missing_minutes = (missing_duration_in_s // 60)
missing_minutes = (missing_duration_in_s // 60) % (60*missing_minutes/abs(missing_minutes))
print(1)