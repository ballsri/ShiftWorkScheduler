import pandas as pd
import xlsxwriter
import datetime, calendar
import numpy as np

def genSchedule(df, month,monthStr, year):

    workbook = xlsxwriter.Workbook('ตารางเวร' + monthStr + str(year+543) + '.xlsx')
    worksheet = workbook.add_worksheet()


    # variables
    holiday = pd.Series(df['วันที่หยุด'].values, index=df['วันหยุด'].values).to_dict()
    holiday = {k: int(v) for k, v in holiday.items() if str(k) != 'nan'}   
    drivers = df['คนขับ'].tolist()
    drivers = [x for x in drivers if str(x) != 'nan']
    replace_driver = df['คนขับทดแทน'].tolist()[0]
    dayShifters =list(df[['หัวหน้าเวรกลางวัน', 'ลูกเวรกลางวัน']].itertuples(index=False, name=None))
    nightShifters = df['เวรกลางคืน'].tolist()
    nightShifters = [x for x in nightShifters if str(x) != 'nan']

    # day of week
    numDay = calendar.monthrange(year, month)[1]
    days = [datetime.date(year, month, day) for day in range(1, numDay+1)]
    dayOfWeek = [day.strftime('%A') for day in days]
    dayToThai = {'Monday': 'จันทร์', 'Tuesday': 'อังคาร', 'Wednesday': 'พุธ', 'Thursday': 'พฤหัสบดี', 'Friday': 'ศุกร์', 'Saturday': 'เสาร์', 'Sunday': 'อาทิตย์'}
    dayOfWeek = [dayToThai[day] for day in dayOfWeek]
    days = [day.strftime('%d') for day in days]

    for k, v in holiday.items():
        dayOfWeek[v-1] = dayOfWeek[v-1] + ' (' + k + ')'

    # day of holiday
    holidayIndex = [i for i, day in enumerate(dayOfWeek) if '(' in day or day == 'อาทิตย์' or day == 'เสาร์']

    # shuffle shifters
    np.random.shuffle(drivers)
    np.random.shuffle(dayShifters)
    np.random.shuffle(nightShifters)

    # create schedule
    dayShift = [""] * numDay
    nightShift = [""] * numDay
    holidayDriverShift = [""] * numDay
    normDriverShift = [""] * numDay
    bookShift = [""] * numDay

    # assign day shift
    j = 0
    for i in holidayIndex:
        dayShift[i] = dayShifters[j]
        j = (j+1) % len(dayShifters)

    # assign night shift
    j = 0
    for i in range(numDay):
        nightShift[i] = nightShifters[j]
        j = (j+1) % len(nightShifters)

    # assign holiday driver shift
    j = 0
    for i in holidayIndex:
        holidayDriverShift[i] = drivers[j]
        j = (j+1) % len(drivers)

    # assign normal driver shift
    # remove replace driver from drivers
    drivers.remove(replace_driver)
    
    j = 0
    for i in range(numDay):
        if dayShift[i] == "":
            normDriverShift[i] = drivers[j]
            j = (j+1) % len(drivers)

    
    # swap one of the normak driver with replace driver to do another task
    # one job per week
    # 3 -> 11 -> 19 -> 27 -> 30 ( start new week)
    # find the first non holiday day
    for i in range(numDay):
        if dayShift[i] == "":
            break
    # swap driver in the next 7 days and iterate it
    while i <= numDay:
        bookShift[i] = normDriverShift[i]
        normDriverShift[i] = replace_driver
        if dayOfWeek[i] == "ศุกร์":
            i = (i + 3)
        else:
            i = (i + 8)

    # create output dataframe
    output = pd.DataFrame({'วันที่': days, 'วัน': dayOfWeek,'ส่งหนังสือ': bookShift, 'เวรขับรถเย็น': normDriverShift, 'เวรขับรถวันหยุด': holidayDriverShift, 'เวรกลางวัน': dayShift, 'เวรกลางคืน': nightShift})

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

    workbook.close()
    
    
