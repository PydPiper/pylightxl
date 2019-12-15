
from _src.readxl import readxl
from time import time


t = time()

db = readxl('book22.xlsx')
db.worksheet('Sheet1').row(1)
print(time()-t)
t = time()
print(db.worksheet('Sheet1').size)
print(db.worksheet('Sheet1').rows)
print(time()-t)



