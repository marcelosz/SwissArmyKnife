#! /usr/bin/env python
#
# urlRequestTimes.py - simple script that displays the total time to perform
#                      a network request. It also displays DNS and connection 
#                      times separately. Useful to test HTTP web servers.
# 

import sys
import re
import pycurl
from io import BytesIO

if sys.argv[1:]:
	# only http or https URLs allowed for now
    protocol = re.match("https*:\/\/", sys.argv[1])
else:
    sys.exit('Please provide a HTTP or HTTPS URL as an argument!')

url = sys.argv[1]

buffer = BytesIO()
c = pycurl.Curl()
c.setopt(c.URL, url)
c.setopt(c.WRITEDATA, buffer)
c.perform()

# HTTP response code
print('Status: %d' % c.getinfo(c.RESPONSE_CODE))
# Elapsed time for DNS query
print('DNS: %f sec' % c.getinfo(c.NAMELOOKUP_TIME))
# Elapsed time for connection
print('Connect: %f sec' % c.getinfo(c.CONNECT_TIME))
# Total elapsed time
print('Total time: %f sec' % c.getinfo(c.TOTAL_TIME))

c.close()