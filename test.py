import re

hahaDate = 'May-84'

pattern1 = re.compile("\w\w\w-\d\d")

print(pattern1.match(hahaDate))