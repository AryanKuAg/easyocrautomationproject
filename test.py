import re
string1 = "498results 34should get"
nothing = int(re.search(r'\d+', string1).group())
print(nothing)
