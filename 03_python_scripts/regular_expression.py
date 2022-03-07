import re
phoneNumRegex = re.compile(r'\d\d\d-\d\d\d-\d\d\d\d')
mo = phoneNumRegex.search('My number is 445-448-7755.')
mo.group()
print('Phone number found: ' + mo.group())

#Grouping with Parentheses

phoneNumRegex = re.compile(r'(\d\d\d)-(\d\d\d-\d\d\d\d)')
mo = phoneNumRegex.search('My number is 415-555-4242.')
mo.group(1)
mo.group(2)
mo.group(0)
mo.group()

areaCode, mainNumber = mo.groups()
print(areaCode)

print(mainNumber)
