import datetime
from datetime import datetime
import time
first_time = datetime.strptime('9/15/2019  5:30:00 PM', '%m/%d/%Y %I:%M:%S %p')
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
