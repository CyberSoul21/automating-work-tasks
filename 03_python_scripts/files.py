import os
import re

path =  os.path.join('usr','bin','spam')
print(path)

os.path.basename(path)

os.path.dirname(path)

os.path.split(path)

#Current directory#
path = os.getcwd()

#Change Directory#

#os.chdir('C:\\windows\\system34')

##os.makedirs('.\myFirstFolder') #Create directory

##helloFile = open('hello.txt','w')#Create file

os.makedirs('.\myFirstFolder')
os.chdir('.\myFirstFolder')
helloFile = open('hi.txt','w')
helloFile.close()
print('I am here: ' + os.getcwd())
os.chdir(path)
print('Now, I am here: ' + os.getcwd())
