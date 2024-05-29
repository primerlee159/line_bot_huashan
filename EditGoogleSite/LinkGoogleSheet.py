import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials

sheet = -1
sheet_name = ''

def open_sheet(sheet_name):
    auth_json_path = 'pro-variety-424600-m6-ff590f99699c.json'
    gss_scopes = ['https://spreadsheets.google.com/feeds']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(auth_json_path,gss_scopes)
    gss_client = gspread.authorize(credentials)

    spreadsheet_key = '1KioGawgpql5GGBQj6_Q62Bbw73HNW1LHw4ojGMqtcuw'
    sheet = gss_client.open_by_key(spreadsheet_key).worksheet(sheet_name)

    return sheet

def de_code(loc,date,time):
    loc_code = -1
    date_code = date.split('/')
    time_code = -100

    if loc == '華山':
        loc_code = 5
    elif loc == '龍山':
        loc_code = 8
    elif loc == '龍山二':
        loc_code = 8
    
    if time == '早':
        if loc == '華山':       
            time_code = 0
        elif loc == '龍山':     
            time_code = 0
        elif loc == '龍山二':
            time_code = 1
    elif time == '午':
        if loc == '華山':       
            time_code = 1
        elif loc == '龍山':     
            time_code = 2
        elif loc == '龍山二':
            time_code = 3
    elif time == '晚':
        if loc == '華山':       
            time_code = 2
        elif loc == '龍山':     
            time_code = 4
        elif loc == '龍山二':
            time_code = 5
    
    return int(date_code[1])+4,loc_code+time_code



def del_work(name,loc,date,time):
    global sheet
    err_message = 'OK!'
    row,col = de_code(loc,date,time)
    get_name = sheet.cell(row,col).value
    
    if get_name is None:
        if loc == '華山' or loc == '龍山二':
            err_message = f'{date}{time}請假 Error : 該日期應為空班，請確認請假日期，謝謝!'
        elif loc == '龍山':
            err_message = del_work(name,'龍山二',date,time)

    elif name in get_name:
        sheet.update_cell(row,col,'')
        err_message = f'{date}{time}請假 OK!'
    else:
        if loc == '華山' or loc == '龍山二':
            err_message = f'{date}{time}請假 Error : 名字錯誤，請確認該日期是否有排班，謝謝!'
        elif loc == '龍山':
            err_message = del_work(name,'龍山二',date,time)
    
    return err_message

def del_works(name,loc,date,time):
    ### 請假 (刪除該cell，是否更改為請假另紀錄在備註) ###
    err_messages = []
    if len(time) == 1:
        err_messages.append(del_work(name,loc,date,time))
    else:
        for t in time:
            err_messages.append(del_work(name,loc,date,t))
    return err_messages

def add_work(name,loc,date,time):
    global sheet
    err_message = 'OK!'
    row,col = de_code(loc,date,time)
    get_name = sheet.cell(row,col).value
    
    if get_name is None:
        sheet.update_cell(row,col,name)
        err_message = f'{date}{time}增班 OK!'
    else:
        if loc == '華山' or loc == '龍山二':
            err_message = f'{date}{time}增班 Error : 該日期已有人排班，請確認想增班的日期，謝謝!'
        elif loc == '龍山':
            err_message = add_work(name,'龍山二',date,time)
    return err_message

def add_works(name,loc,date,time):
    ### 請假 (刪除該cell，是否更改為請假另紀錄在備註) ###
    err_messages = []
    if len(time) == 1:
        err_messages.append(add_work(name,loc,date,time))
    else:
        for t in time:
            err_messages.append(add_work(name,loc,date,t))
    return err_messages

def EditGoogleSheet(new_comment):
    global sheet
    month = datetime.date.today().month
    sheet_name = f'{month}月'
    decode_comment = new_comment.split(' ')
    remark = None
    
    if len(decode_comment) == 5:
        work,name,loc,date,time = decode_comment
    elif len(decode_comment) == 6:
        work,name,loc,date,time,remark = decode_comment
    else:
        work = -1
    

    err_messages = []
    check_month = int(date.split('/')[0])

    if work == -1:
        err_messages.append('Error : 格式錯誤，請重新確認格式，謝謝!')
    elif month != check_month:
        err_messages.append('Error : 月份錯誤，請重新確認增減班日期，謝謝!')
    elif loc != '華山' and loc != '龍山':
        err_messages.append('Error : 地點錯誤，分隊只能填華山或龍山，謝謝!')  
    else:
        sheet = open_sheet(sheet_name = sheet_name)
        if work == '增班':
            err_messages = add_works(name,loc,date,time)
        elif work == '請假':
            err_messages = del_works(name,loc,date,time)
        else:
            err_messages.append('Error : 請輸入確認格式，第一格只能填寫增班或請假，謝謝!')

    if sheet != -1:
        set_remark_row = len(sheet.get_all_values())
    else:
        set_remark_row = 0

    if set_remark_row < 40:
        set_remark_row = 40
    else:
        set_remark_row += 1
    
    push_cell = 0
    for message in err_messages:
        if not 'Error' in message:
            sheet.update_cell(set_remark_row+push_cell,3,f'{name} {loc}{message}')
            if remark is not None:
                sheet.update_cell(set_remark_row+push_cell,4,remark)
            push_cell += 1
                


    return err_messages

if __name__ == '__main__':
    
    for i in EditGoogleSheet('增班 李相承 華山 5/30 早午晚 10-18'):
        print(i)
    



