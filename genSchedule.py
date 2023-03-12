import pandas as pd
import xlsxwriter
import datetime, calendar

def genSchedule(df, month,monthStr, year):

    workbook = xlsxwriter.Workbook('ตารางเวร' + monthStr + str(year+543) + '.xlsx')
    worksheet = workbook.add_worksheet()


    # variables
    try:
        holiday = df.pivot(index='เดือนที่หยุด', columns='วันที่หยุด', values='วันหยุด')              
        drivers = df['คนขับเย็น'].tolist()
        drivers = [x for x in drivers if str(x) != 'nan']
        replace_driver = df['คนขับทดแทน'].tolist()[0]
        hDrivers = df['คนขับวันหยุด'].tolist()
        hDrivers = [x for x in hDrivers if str(x) != 'nan']
        dayShifters =list(df[['หัวหน้าเวรกลางวัน', 'ลูกเวรกลางวัน']].itertuples(index=False, name=None))
        dayShifters = [(x,y) for x,y in dayShifters if str(x) != 'nan' and str(y) != 'nan']
        nightShifters = df['เวรกลางคืน'].tolist()
        nightShifters = [x for x in nightShifters if str(x) != 'nan']
    except:
        return 2

    numDay = pd.Timestamp(year, 12, 31).dayofyear
   
    
    days = []
    
    for m in range(1,12+1):
        for i in range(1,calendar.monthrange(year, m)[1] +1):
            days.append(datetime.date(year, m, i)) 

        
    dayNameInYear = [day.strftime('%A') for day in days]
    dayToThai = {'Monday': 'จันทร์', 'Tuesday': 'อังคาร', 'Wednesday': 'พุธ', 'Thursday': 'พฤหัสบดี', 'Friday': 'ศุกร์', 'Saturday': 'เสาร์', 'Sunday': 'อาทิตย์'}
    dayNameInYear = [dayToThai[day] for day in dayNameInYear]
    days = [day.strftime('%d') for day in days]

    for d, m in holiday.items():
        # month is dataframe
        for i, v in m.items():
            if str(v) != 'nan':
                dayInYear = pd.Timestamp(year, int(i), int(d)).dayofyear
                dayNameInYear[dayInYear-1] = dayNameInYear[dayInYear-1] + ' (' + v + ')'

    # day of holiday
    yearHolidayIndex = [i for i, day in enumerate(dayNameInYear) if '(' in day or day == 'อาทิตย์' or day == 'เสาร์']

    # Only keep the day in the month
    numMonth = calendar.monthrange(year, month)[1]
    numStartMonth = pd.Timestamp(year, month, 1).dayofyear
    numEndMonth = pd.Timestamp(year, month, numMonth).dayofyear
    dayNameInMonth = dayNameInYear[numStartMonth-1:numEndMonth]




    ## CRITICAL PART ##

    # create schedule
    dayShift = [""] * numMonth
    nightShift = [""] * numMonth
    holidayDriverShift = [""] * numMonth
    normDriverShift = [""] * numMonth
    bookShift = [""] * numMonth

    yearBookShift = [""] * numDay
    yearNormDriverShift = [""] * numDay
    yearHolidayDriverShift = [""] * numDay
    yearDayShift = [""] * numDay
    yearNightShift = [""] * numDay

    # assign day shift
    j = 0
    for i in yearHolidayIndex:
        yearDayShift[i] = dayShifters[j]
        j = (j+1) % len(dayShifters)
    
    dayShift = yearDayShift[numStartMonth-1:numEndMonth]

    # assign night shift
    j = 0
    for i in range(numDay):
        yearNightShift[i] = nightShifters[j]
        j = (j+1) % len(nightShifters)

    nightShift = yearNightShift[numStartMonth-1:numEndMonth]

    # assign holiday driver shift
    j = 0
    for i in yearHolidayIndex:
        yearHolidayDriverShift[i] = hDrivers[j]
        j = (j+1) % len(hDrivers)
    
    holidayDriverShift = yearHolidayDriverShift[numStartMonth-1:numEndMonth]


    # assign driver shift

    j = 0
    for i in range(0,numDay):
        if i not in yearHolidayIndex:
            yearNormDriverShift[i] = drivers[j]
            j = (j+1) % len(drivers)

    # assign book shift
    # swap driver each week by once

    ## if there're need to continue the sequence of replacing drivers from last year ##
        # if year != 2023:
        #     last_replaced = drivers.index(df['แทนล่าสุด'].tolist()[0])
        #     drivers = drivers[last_replaced+1:] + drivers[:last_replaced+1]


    j = 0
    pointer = 0
    if year != 2023:
        last_replaced = drivers.index(df['แทนล่าสุด'].tolist()[0])
        pointer = (last_replaced+1) % len(drivers)
        
    while j < numDay:
        d = drivers[pointer] # need to be replaced
        pointer = (pointer+1) % len(drivers)
        while j < numDay and j in yearHolidayIndex:
            j += 1

        
        while j < numDay and yearNormDriverShift[j] != d:
            j += 1
           
        if j >= numDay:
            break
        
        yearBookShift[j] = yearNormDriverShift[j]
        yearNormDriverShift[j] = replace_driver
        while j < numDay and  'เสาร์' not in dayNameInYear[j]:
            j += 1

    # choose only in selected month
    bookShift = yearBookShift[numStartMonth-1:numEndMonth]
    normDriverShift = yearNormDriverShift[numStartMonth-1:numEndMonth]

    # create output dataframe
    output = pd.DataFrame({'วันที่': days[numStartMonth-1:numEndMonth], 'วัน': dayNameInMonth,'ส่งหนังสือ': bookShift, 'เวรขับรถเย็น': normDriverShift, 'เวรขับรถวันหยุด': holidayDriverShift, 'เวรกลางวัน': dayShift, 'เวรกลางคืน': nightShift})

    # print(output)
    # save to excel
    header_format = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D9D9'})
    col_format = workbook.add_format({ 'border': 1, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#80aaff'})
    subcell_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    holiday_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#ffc34d'})
    weekend_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#88cc00'})

    row = 0
    col = 0
    # write title
    worksheet.merge_range(row,col,row,col+7, "ตารางเวร ประจำเดือน " + monthStr + " ปี พ.ศ."+ str(year+543), header_format)
    row += 1
    # write header
    worksheet.write(row,col, "วันที่", col_format)
    worksheet.write(row,col+1, "วัน", col_format)
    worksheet.write(row,col+2, "ส่งหนังสือ", col_format)
    worksheet.write(row,col+3, "เวรขับรถเย็น", col_format)
    worksheet.write(row,col+4, "เวรขับรถวันหยุด", col_format)
    worksheet.merge_range(row,col+5,row,col+6, "เวรกลางวัน", col_format)
    worksheet.write(row,col+7, "เวรกลางคืน", col_format)
    row += 1

    for index, r in output.iterrows():
        if r['วัน'] == 'อาทิตย์' or r['วัน'] == 'เสาร์':
            cell_format = weekend_format
        elif '(' in r['วัน']:
            cell_format = holiday_format
        else:
            cell_format = subcell_format
        worksheet.write(row,col, r['วันที่'], cell_format)
        worksheet.write(row,col+1, r['วัน'], cell_format)
        worksheet.write(row,col+2, r['ส่งหนังสือ'], cell_format)
        worksheet.write(row,col+3, r['เวรขับรถเย็น'], cell_format)
        worksheet.write(row,col+4, r['เวรขับรถวันหยุด'], cell_format)
        if r['เวรกลางวัน'] != "":
            worksheet.write(row,col+5, r['เวรกลางวัน'][0], cell_format)
            worksheet.write(row,col+6, r['เวรกลางวัน'][1], cell_format)
        else:
            worksheet.write(row,col+5, "", cell_format)
            worksheet.write(row,col+6, "", cell_format)
        worksheet.write(row,col+7, r['เวรกลางคืน'], cell_format)
        row += 1
    try:
        workbook.close()
        return 0
    except:
        return 1
    
    
    
