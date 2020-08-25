import datetime
from datetime import datetime
import time
first_time = datetime.strptime('5/5/2019 12:00:00 PM', '%m/%d/%Y %I:%M:%S %p')
time.sleep(2)
later_time = datetime.strptime('5/4/2019 11:59:00 AM', '%m/%d/%Y %I:%M:%S %p')
duration = later_time - first_time
duration_in_s = duration.total_seconds()
days  = divmod(duration_in_s, 86400)[0]
hours = divmod(duration_in_s, 3600)[0]
minutes = divmod(duration_in_s, 60)[0]
print(days)
print(hours)
print(minutes)