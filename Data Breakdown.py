import xlwings as xw

TEXT_WIDTH = 40 #How many characters in Each Line
DIGITS = 3 # how many digits to be shown while calculating averages and %age


sheet = xw.Book('test.xlsx').sheets[0]  #Name of the excel sheet having the data

arr = sheet.range('A2:A99').value     # The range of the data column

pure = []
Sum = 0
count = 0


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


    
for i in arr:
    if i != None:
        if is_number(i):
            Sum+=float(i)
            count+=1
            pure.append(i)
pure.sort()
perc = {}
for i in pure:

    
    if i in perc.keys():
        perc[i]+=1
    else:
        perc[i] = 1
data = []
keys = list(perc.keys())

Max = keys[0]
Min = keys[0]
print("DETAILED BREAKDOWN")
print()

for i in perc:
    if perc[i]>perc[Max]:
        Max = i
    if perc[i]<perc[Min]:
        Min = i
    data.append((str(i),str(perc[i]),str((perc[i]/count)*100)[:1+DIGITS]))
    


    

data1 = [['VALUE'],['No. of entries'],['% of total entries']]
columns = []

for i in data1:
    columns.append(i[0])
width = [0,0,0]
data.insert(0,tuple(columns))

for row in data:
    t = 0
    for column in row:
        if width[t] < len(str(column)):
            width[t] = len(str(column))
        t+=1
        
for row in data:
    t=0
    
    for column in row:
        x = width[t]
        
        for i in str(column):
            print (i,end = '')
            x-=1
            
        while x != 0:
           print(' ',end = '')
           x-=1
        t+=1
        print(' | ',end ='')        
    print()
print()
print('Average :',str((Sum/count))[:2+DIGITS])
print()

mid = len(pure) // 2
res = (pure[mid] + pure[~mid]) / 2
print("Median : " + str(res))
print()

print('Maximum %age at ',Max,' With %age',str((perc[Max]/count)*100)[:1+DIGITS])
print('Minimum %age at ',Min,' With %age',str((perc[Min]/count)*100)[:1+DIGITS])

