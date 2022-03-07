import os
import re
import math

fileName = input('Type name of file: ')
fileName = fileName + '.txt'

file = open(fileName)

##inFile = file.read()

lines = input('how many lines do you wish in each file? ')
lines = int(lines)
##l = len(inFile)
c = 1
r=''
f = True
#[this is a line \n ]
a = 1

while f :
        r = file.readline()
        if c == 1 and len(r) !=0:
                newFile = 'part' + str(a) + '.txt'
                file2 = open(newFile,'a')
        if c <= lines and len(r) !=0:
                
##                print(r)
                file2.write(r)

                if c == lines:
                        c = 1
                        a += 1
                        file2.close()
                else:
                        c += 1
        if len(r) == 0:
                f = False
##                print('Entro if')



file.close()
file2.close()

print(c)





