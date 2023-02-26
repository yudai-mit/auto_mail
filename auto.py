import schedule
import time
import sys
import datetime

import winsound
from sound import sound1
import os





winsound.PlaySound("「起動します」.wav",winsound.SND_FILENAME)
#試験用
#schedule.every().minutes.at(":00").do(sound1)
schedule.every().hours.at(":00").do(sound1)

while True:
    schedule.run_pending()
    time.sleep(1)
    

