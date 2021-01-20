today = Date
snd = "1/1/3000"
week = Weekday(Date)
hr = Hour(Time)
min = Minute(Time)
tm = hr&":"&min
nicetime = TimeValue(tm)
ft = tm&":"&week

ssnd = DateDiff("d",today,snd)

msgbox ft
