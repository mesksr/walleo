import openpyxl
import os
import shutil
import sys
import cv2
import numpy as np
import ctypes

sym2cat = {} # maps symbols to category assigned
symbols = [] # all symbols used
today_score = 0
today_data = []
font = cv2.FONT_ITALIC

def copy_file():
    # copies 'Today 2018-01.xlsx' to 'data.xlsx'
    try:
        source = r"C:\Users\mesksr\Dropbox\Today 2018-01.xlsx"
        destination = os.getcwd()+"\data.xlsx"
        shutil.copy2(source, destination)
        print ("Success: data.xlsx retrived\n")
    except:
        print ("Error: can't retrive data.xlsx")
        sys.exit()

def read_categories():
    # read categories.txt
    try:
        source = r"C:\Users\mesksr\Desktop\Walleo\categories.txt"
        file = open(source, 'r')
        for line in file:
            line = line.strip()            
            category, symbol = line.split(' - ')
            symbl = symbol.lower()
            sym2cat[symbol] = category
            symbols.append(symbol)
        print ("Success: categories.txt read\n")
    except:
        print ("Error: can't read categories.txt")
        sys.exit()


def get_records(days):
    # read records.txt and return data of last few days
    last_few_days = []
    try:
        source = r"C:\Users\mesksr\Desktop\Walleo\records.txt"
        file = open(source, 'r')
        for line in file:
            if (':' not in line):
                break
            line = line.strip()
            date, score = line.split(' : ')
            if (len(last_few_days) >= days):
                if (days == 1):
                    last_few_days = []
                else:
                    last_few_days = last_few_days[1-days:]
            last_few_days.append((date, score))
        file.close()        
        for date, score in last_few_days:
            print ("\tOld data:", date, '--->', score)
        print ("Success: records.txt read\n")
    except:
        print ("Error: can't read records.txt")
        sys.exit()
    return last_few_days

def calc_score(data):
    # based on the data, calculate and return score
    none_present = False
    if ('None' in data):
        i = 0
        temp = ''
        while (i < len(data)):
            if (data[i] == 'N'):
                i += 4
            else:
                temp += data[i]
                i += 1
        none_present = True
        data = temp
    # if none is present then day's data is not complete
    # date_data_complete == not none_present
    score = 0
    time = {}
    productive = 0
    waste = 0
    for symbol in symbols:
        hours = data.count(symbol)//6
        minutes = (data.count(symbol)%6)*10
        time[symbol] = data.count(symbol)/6
        temp = 0
        
        if (symbol == '@'): # Sleep : 5-7 hour
            if (time[symbol] < 5):
                temp -= (10 - time[symbol])
            elif (time[symbol] <= 7):
                temp += 10
            else:
                temp -= (time[symbol] + 3)
                
        elif (symbol == '1' or symbol == '4' or symbol == 'g'): # Bath, College, House Work : any
            temp += (time[symbol])
                
        elif (symbol == '2'): # Eat/Cook : 0.75-2 hour
            # TODO - check if ate thrice?
            if (time[symbol] < 0.75 or time[symbol] > 2):
                temp -= 5
            else:
                temp += 5
                
        elif (symbol == '3'): # Exercise : any
            if (time[symbol] == 0):
                temp -= 5
            else:
                temp += 5
                                
        elif (symbol == '5' or symbol == '6' or symbol == '7' or symbol == '8'): # Study, MOOCs, Self Project, C. Coding : 1+ hour
            if (time[symbol] <= 1):
                temp -= ((time[symbol]-1)**2)*10
            else:
                temp += ((time[symbol]-1)**2)*10
            productive += time[symbol]
                
        elif (symbol == '9'): # Language : 0.5+ hour
            if (time[symbol] <= 1): 
                temp -= ((time[symbol]-0.5)**2)*(15/0.25)
            else:
                temp += ((time[symbol]-0.5)**2)*(15/0.25)
            productive += time[symbol]
                
        elif (symbol == 'a' or symbol == 'b'): # Her, Family: 0.5+ hour
            if (time[symbol] <= 0.5):
                temp -= ((time[symbol]-0.5)**2)*20
            else:
                temp += (time[symbol]-0.5)
                
        elif (symbol == 'c' or symbol == 'd' or symbol == 'e' or symbol == 'h'): # Novel, Games, Movie/TV : 0 - 1 hour
            if (time[symbol] <= 1):
                temp += ((time[symbol]-1)**2)*15
            else:
                temp -= ((time[symbol]-1)**2)*15
            waste += time[symbol]
                
        elif (symbol == 'f'): # Internet/Other : 0 - 1 hour
            if (time[symbol] <= 1):
                temp += ((time[symbol]-1)**2)*15
            else:
                temp -= ((time[symbol]-1)**2)*15
            waste += time[symbol]
                                
        elif (symbol == 'h'): # Shopping : 0 - 1 hour
            if (time[symbol] <= 1):
                temp += ((time[symbol]-1)**2)*10
            else:
                temp -= ((time[symbol]-1)**2)*10
                
        elif (symbol == 'i'): # Travel : 0 - 1 hour
            if (time[symbol] <= 1):
                temp += ((time[symbol]-1)**2)*10
            else:
                temp -= ((time[symbol]-1)**2)*10
                
        print ('\t'+sym2cat[symbol], str(hours)+':'+str(minutes), 'gives', temp)
        today_data.append([sym2cat[symbol], str(hours)+':'+str(minutes), temp])
        score += temp

    temp = 0
    if (productive <= 3):
        temp -= ((productive-3)**2)*(20/9)
    else:
        temp += ((productive-3)**2)*(20/9)
    print ('\t**Productive**', 'gives', temp)
    temp = 0
    if (waste <= 5):
        temp += ((waste-5)**2)*(20/25)
    else:
        temp -= ((waste-5)**2)*(20/25)
    print ('\t**Waste**', 'gives', temp)
        
    return score, not none_present 

def update_records():
    # reads data.xlsx and updates records.txt
    try:
        wb = openpyxl.load_workbook(filename = 'data.xlsx')
        if (len(wb.get_sheet_names())!=1):
            print ("Error: more than one sheets")
            sys.exit()
        sheet = wb[wb.get_sheet_names()[0]]
        pos = 2
        last_date = get_records(1)[-1][0]
        print ('\tLast date:', last_date)
        while (sheet['B'+str(pos)].value is not None):
            global today_data
            today_data = []
            date = str(sheet['B'+str(pos)].value)[:10]
            if (date > last_date):
                # read data
                date_data = '' # saves data for the current date
                date_data += ''.join(list(map(str, map(lambda x: x.value, sheet[str(pos)][2:74]))))
                date_data += ''.join(list(map(str, map(lambda x: x.value, sheet[str(pos+1)][2:74]))))
                #print (date, date_data)
                # calculate score
                date_score, date_data_complete = calc_score(date_data)
                #print (date_score, date_data_complete)
                if (date_data_complete):
                    # write in file
                    try:
                        print ('\tNew data:', date, '--->', date_score, '\n')
                        source = r"C:\Users\mesksr\Desktop\Walleo\records.txt"
                        file = open(source, 'a')
                        file.write(str(date)+' : '+str(date_score)+'\n')
                        file.close()
                    except:
                        print ("Error: can't read records.txt")
                        sys.exit()

                else:
                    print ('\tNew data:', date, '--->', date_score, '---> incomplete')
                    global today_score
                    today_score = date_score
            pos += 2
        print ("Success: data.xlsx read\n")
    except:
        print ("Error: can't read data.xlsx")
        sys.exit()

def normalize(old_scores, mx_new):
    mx_old = max(max(list(map(lambda x: abs(int(float(x))), old_scores))), int(today_score))
    new_scores = list(map(lambda x: int(int(float(x))*(mx_new/mx_old)), old_scores))
    for i in range(len(old_scores)):
        print ('\t', old_scores[i], 'to', new_scores[i])
    print ("Success: score normalized\n")
    return new_scores

def dotted_line(canvas, start, end, color):
    if (start[1] > end[1]):
        start, end = end, start
    y = start[1] 
    while (y <= end[1]):
        cv2.line(canvas, (start[0], y), (start[0], min(y+5, end[1])), color)
        y += 10
        
def put_text(canvas, start, end, color, text):
    x, y = end
    x -= 15
    if (start[1] <= end[1]):
        y += 15
    else:
        y -= 5
    cv2.putText(canvas, text, (x, y), font, 0.5, color, 1)
        
def draw(days):
    global today_score
    height = 786
    width = 1366
    border_h = 200
    border_wl = 200
    border_wr = 350
    taskbar_gutter = 40
    extra_length = 30
    window_height = height - 2*border_h - taskbar_gutter
    window_width = width - border_wl - border_wr
    segments = days
    segment_size = window_width//segments
    c_radius = 6
    green = (0, 255, 0)
    blue = (255, 127, 127)
    white = (255, 255, 255)
    dark_gray = (50, 50, 50)
    red = (0, 0, 255)
    records =  get_records(days)
    dates = list(map(lambda x: x[0][-2:]+'/'+x[0][-5:-3], records))+['Today']
    scores = list(map(lambda x: x[1], records))+[today_score]
    normalized_scores = normalize(scores, window_height//2)


    # drawing main axis and background lines
    canvas = np.zeros((height, width, 3), dtype = "uint8")
    for i in range(-200, 200, 20):
        cv2.line(canvas, (border_wl - extra_length, border_h + window_height//2 + i),
                 (border_wl + extra_length + window_width, border_h+window_height//2 + i), dark_gray)

    cv2.line(canvas, (border_wl - extra_length, border_h + window_height//2),
             (border_wl + extra_length + window_width, border_h+window_height//2), white)

    # drawing lines, circles
    for i in range(days):
        start = (border_wl + i*segment_size, border_h+(window_height//2)-normalized_scores[i])
        if (i != days-1):
            end = (border_wl + (i+1)*segment_size, border_h+(window_height//2)-normalized_scores[i+1])
            cv2.line(canvas, start, end, white, thickness=2)
        cv2.circle(canvas, start, c_radius, white)
        dotted_end = (border_wl + i*segment_size, border_h+(window_height//2))
        dotted_line(canvas, start, dotted_end, white)
        
    # drawing lines, circles - today's part
    prv_start = start
    start = (border_wl + (i+1)*segment_size, border_h+(window_height//2)-normalized_scores[-1])
    dotted_end = (border_wl + (i+1)*segment_size, border_h+(window_height//2))
    
    cv2.line(canvas, prv_start, start, red, thickness=2)
    dotted_line(canvas, start, dotted_end, white) 
    cv2.circle(canvas, start, c_radius, red, -1)

    # putting text
    for i in range(days):
        start = (border_wl + i*segment_size, border_h+(window_height//2)-normalized_scores[i])
        dotted_end = (border_wl + i*segment_size, border_h+(window_height//2))
        put_text(canvas, start, dotted_end, blue, dates[i])

    # putting text - today's  part    
    start = (border_wl + (i+1)*segment_size, border_h+(window_height//2)-normalized_scores[-1])
    dotted_end = (border_wl + (i+1)*segment_size, border_h+(window_height//2))
    put_text(canvas, start, dotted_end, blue, 'Today')

    today_score = str(today_score)
    if ('.' in today_score):
            today_score = today_score[:today_score.find('.')+2]
    if (today_score[0] == '-'):
        cv2.putText(canvas, today_score, (start[0]-15, start[1]+20), font, 0.5, red, 1)
    else:        
        cv2.putText(canvas, today_score, (start[0]-15, start[1]-10), font, 0.5, red, 1)

    # putting today's data 
    height_available = height - taskbar_gutter
    parts = len(today_data)
    part_height = height_available//parts
    y = part_height//2
    for cat, time, score in today_data:
        x = width - 205
        score = str(score)
        if ('.' in score):
            score = score[:score.find('.')+2]
        pos = time.find(':')
        if (pos == 1):
            time = '0'+time
        if (len(time) == 4):
            time += '0'
        
        if (float(score) < 0):
            cv2.putText(canvas, time + ' ' + cat+' .. '+score, (x, y), font, 0.5, red, 1)
        elif (float(score) == 0):
            cv2.putText(canvas, time + ' ' + cat+' .. '+score, (x, y), font, 0.5, white, 1)
        else:
            cv2.putText(canvas, time + ' ' + cat+' .. '+score, (x, y), font, 0.5, green, 1)
        y += part_height
        
    cv2.imwrite("wallpaper.jpeg", canvas)
    ctypes.windll.user32.SystemParametersInfoW(20, 0, os.getcwd()+"/wallpaper.jpeg" , 0)
    print ("Success: wallpaper set\n")

    # TODO - Draw a pie chart of total time spent

    
    
    
copy_file()
read_categories()
update_records()
draw(7)

