import pandas as pd
import re, operator, datetime, os
from xlsxwriter.utility import xl_rowcol_to_cell
from pytz import timezone
def get_last_updated_comments(data):   
    data=str(data)
    #Compile a pettern of 2020-12-14 14:51:31 - Amit ABHANG (Work-Notes) [code]Requested chain is Activated.[/code]
    pattern = re.compile("(\d{4}.\d{2}.\d{2}.\d{2}.\d{2}.\d{2})\W{3}(\S*[\w ]+) (.*)\n(.+)")    
    parsed=re.finditer(pattern,data)
    sorted_data=sorted(parsed,key=operator.itemgetter(0))
    tz = timezone('Europe/Paris')
    now_time = datetime.datetime.now(tz)
    map_data={'who':None,'updated_date':None,'before':None}    
    for i in sorted_data:  
        date=i.groups()[0]
        name=i.groups()[1]
        item=i.groups()[2]
        data=i.groups()[3]  
        if not name in ['HCL MoogSOFT']:
            map_data['who'] = name 
            map_data['updated_date']=date
        import dateutil.relativedelta
        elapsedTime=dateutil.relativedelta.relativedelta(now_time.replace(tzinfo=None), datetime.datetime.strptime(date,"%Y-%m-%d %H:%M:%S"))         
        map_data['before']=', '.join(f'{v} {k}' for k,v in {'days':elapsedTime.days, 'hours':elapsedTime.hours, 'min':elapsedTime.minutes}.items() if v)    
 
    return map_data['updated_date'], map_data['who'], map_data['before']

def setHeader(sheet,row,col,h):  
    global LastCol
    for hCol, hVal in enumerate(h):
        sheet.write(row,hCol+col,hVal,header_format)
    LastCol = col + hCol
            
def create_data(df_name,sheet_name=None,sheet_header=None):
    global LastRow
    df_name.to_excel(writer, sheet_name=sheet_name, startrow=start_row+1, startcol=start_col, index=False, header=False,encoding='latin1')
    LastRow = start_row + df_name.shape[0] -1
    worksheet = writer.sheets[sheet_name]
    setHeader(worksheet,start_row,start_col,sheet_header)
    return worksheet 

def generate_excel(data_frame,op_file):
    global start_row, start_col, LastCol, LastRow, header_format, writer
    start_row= 1
    start_col = 0
    LastRow = 0
    LastCol = 0
    writer = pd.ExcelWriter(op_file, engine='xlsxwriter')
    workbook  = writer.book
    #EXCEL FORMATTING
    red_format = workbook.add_format({'bg_color':'#F74A4A','border':1,'font_name':'calibri light','font_size':'9','align':"center"})
    green_format = workbook.add_format({'bg_color':'#50DE89','border':1,'font_name':'calibri light','font_size':'9','align':"center"})
    header_format = workbook.add_format({'bold':True,'bg_color':"#808080",'border':1, 'font_name':"calibri",'font_size':10,'align':"center", 'color':"#FFFFFF"})
    data_format = workbook.add_format({'border':1,'font_name':'calibri light','font_size':'9','align':"center"})
    cell_format = workbook.add_format({'align': 'center','valign': 'vcenter','border':1})
    merge_format = workbook.add_format({'bold':True, 'color':"#FFFFFF",'font_size':14,'align': 'center','bg_color':"#404244",'border': 1})
    date_format = workbook.add_format({'border':1, 'font_name':'calibri light','font_size':'9','num_format':'yyyy-mm-dd hh:mm:ss'})    
    
    for k,v in data_frame.items():
        header=v.columns.values
        worksheet_dc=create_data(v,k,header)         
        print('createing worksheet',k)
        '''Arrange The Row Col to set Data'''
        col_range=xl_rowcol_to_cell(start_row + 1, start_col)+":"+xl_rowcol_to_cell(LastRow + 1, LastCol)
        FirstCell = xl_rowcol_to_cell(start_row - 1, start_col)
        LastCell = xl_rowcol_to_cell(start_row - 1, LastCol)
        merge_range = FirstCell + ":" + LastCell
        TopLeftCell = xl_rowcol_to_cell(start_row, start_col)
        BottomRightCell = xl_rowcol_to_cell(LastRow, LastCol)
        FilterRange = TopLeftCell+":"+BottomRightCell
        
        now_date = datetime.datetime.strptime('1900-01-01 12:00:00', "%Y-%m-%d %H:%M:%S")                   
        worksheet_dc.merge_range(merge_range, "Last Notes Audit Report", merge_format)    
        worksheet_dc.set_column(col_range, None, data_format)
        worksheet_dc.conditional_format(col_range, {'type': 'date','criteria': 'greater than','value': now_date,'format': date_format})              
        worksheet_dc.conditional_format(col_range, {'type': 'blanks','format': date_format})
        worksheet_dc.set_column(start_col,LastCol,15)
        worksheet_dc.autofilter(FilterRange)        
    writer.save()
    
def process_file(xls):
    global df
    try:          
        xl = pd.ExcelFile(xls)
        df = xl.parse(xl.sheet_names[0],skiprows=0,index_col=None)
        sheet_name='RITMs'
        if 'INC' in df.Number.iloc[0]:
            sheet_name='Incidents'
            #df['Opened']=df['Opened'].astype('str')
            
        if 'PTASK' in df.Number.iloc[0]:
            sheet_name='Problems'
        if 'CTASK' in df.Number.iloc[0]:
            sheet_name='Changes'               
        
        df['Merged'] = df['Work-Notes'].fillna('').str.cat(df['Additional comments'].fillna(''))
        df['Last Updated Time'],df['Updated By'],df['Before']=zip(*df['Merged'].map(get_last_updated_comments))
        df.dropna(subset=['Last Updated Time','Updated By'], how='all',inplace=True)
        df['Last Updated Time']=df['Last Updated Time'].apply(pd.to_datetime)
        df.drop(['Merged','Work-Notes','Additional comments'],axis=1, inplace=True)        
        return df, sheet_name
    except Exception as e:
        print(str(e))
		
if __name__=='__main__':
    df_list={}
    op_file='last_work_notes_audit.xlsx'
    for i in os.listdir('./data'):
        df_list[process_file(os.path.abspath('./data/'+i))[1]]=process_file(os.path.abspath('./data/'+i))[0]
    generate_excel(df_list,op_file) 
    