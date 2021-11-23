t_hour = ''
utc_offset='0'
correct_hour = str(int(t_hour)-int(utc_offset))
if int(correct_hour)==24:
    correct_hour=str(0)
elif int(correct_hour)>24:
    correct_hour=str(int(correct_hour)-24)
elif int(correct_hour)<0:
    correct_hour=str(int(correct_hour)+24)    
print (correct_hour)