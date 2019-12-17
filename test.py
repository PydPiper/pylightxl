
from _src.readxl import readxl
from time import time


t = time()

db = readxl('_test/testbook.xlsx')
db.ws('Sheet1').row(1)
print('process time,',time()-t)
print(db.worksheet('Sheet1').size)
t = time()
print(db.worksheet('Sheet1').rows)
print(time()-t)



