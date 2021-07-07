# -*- coding: utf-8 -*-
"""
Created on Thu Mar 19 16:34:01 2020

@author: lockchi
"""


from re import findall
from easygui import multenterbox
# import easygui
from xlsxwriter import Workbook
from string import ascii_lowercase
import socket
import plxscripting
from pandas import DataFrame
# import pandas as pd
from getpass import getuser
from urllib.request import urlretrieve

#Plaxis script import
from plxscripting.easy import *


phase_name = []
'''-------------------------------------------------------------------------------------------------------------------'''
#Extract Result
def plaxis_result():
    #Initial Variable (not for editing)
    stage_number = 0
    max_strut_list = []
    print('----------------------')
    print('Total no. of stage in this model:',len(list(g_o.Phases)))
    print('----------------------')
    print('Please wait')
    print('----------------------')
    
    #Empty Data Frame
    data = {'stage':[],'x': [],'y': [],'m':[]}
    df= DataFrame(data,columns = ['stage','x','y','m'])
    #print(df)
    
    
    for i in g_o.Phases[0:]:
        #Intial Variables
        stage_number = stage_number + 1
        #Print Stage Name
        # print(findall(r'"(.*?)"', g_o.echo(g_o.Phases[stage_number-1].Identification))[0])
        stage_list = []
        try:
            print('Working on Stage',stage_number,': ',findall(r'"(.*?)"', g_o.echo(g_o.Phases[stage_number-1].Identification))[0])
            phase_name.append(findall(r'"(.*?)"', g_o.echo(g_o.Phases[stage_number-1].Identification))[0])
            # print(phase_name)
            #test if have result
            n2nX = g_o.getresults(i, g_o.ResultTypes.NodeToNodeAnchor.X, 'node')
            #Have result
            n2nX = g_o.getresults(i, g_o.ResultTypes.NodeToNodeAnchor.X, 'node')
            n2nY = g_o.getresults(i, g_o.ResultTypes.NodeToNodeAnchor.Y, 'node')
            n2nID = g_o.getresults(i, g_o.ResultTypes.NodeToNodeAnchor.ElementID, 'node')
            n2nM = g_o.getresults(i, g_o.ResultTypes.NodeToNodeAnchor.AnchorForce2D, 'node')
            #Filtered Result
            fn2nX = list(n2nX)[::2]
            fn2nY = list(n2nY)[::2]
            fn2nID = list(n2nID)[::2]
            fn2nM = list(n2nM)[::2]
            max_strut_list.append(len(list(fn2nX)))
            #Stage List
            for a in range(0,len(list(fn2nX))):
                stage_list.append(stage_number)
                #print(stage_number)
            #print('stage list',stage_list)
            #print(fn2nM)
            
            data1 = {'stage':stage_list,'x': fn2nX,'y': fn2nY,'ID': fn2nID,'m':fn2nM}
            df1= DataFrame(data1,columns = ['stage','x','y','ID','m'])
            df = df.append(df1)
             #print(df)
        #If no result    
        except plxscripting.plx_scripting_exceptions.PlxScriptingError:
            print ('No Active Strut in Stage %s' %stage_number)
        
    #print('done')
    #selected_row = df.loc[df['stage'] == 5]
    #print(selected_row)
    #print (selected_row['m'])
    #print(max(max_strut_list))
    #print(stage_number)
    return(df,max(max_strut_list),stage_number)
    #print(strut_result()[0])
    
#Find Max Strut Number in choice/range    
def max_strut_in_range():       
    max_strut = 0
    max_strut_in_range_list = []
    temp_slv_ls = []
    for i in range(0,stage_number):
        #print('i value',i)
        #print(df)
        temp = df.loc[(df['stage'] == i)&(df['x'] >= strut_range_min)&(df['x'] <= strut_range_max)]
        temp_slv_ls = df['y'].values.tolist()
        temp_slv_ls = list(set(temp_slv_ls))
        temp_slv_ls.sort(reverse = True)
        # temp_slv_ls = temp_slv_ls.drop_duplicates(subset=['brand'])
        print(temp_slv_ls)
        max_strut_in_range_list.append(len(list(temp['m'])))
        
    max_strut = max(max_strut_in_range_list)
    print('----------------------')
    print('No. of Strut in User Selected Range:',max_strut)
    return max_strut,temp_slv_ls

#get user password and port from pop up box
def check_user_passnport():
    #Read File
    try: 
        username = getuser()
        file1 = open("%s.txt"%username, 'r') 
        Lines = file1.readlines()
        password = Lines[0].strip()
        port = Lines[1].strip()
        print('Recorded User Password:',password)
        print('Recorded User Port:',port)
        file1.close()
        
    except (OSError, IOError,IndexError):
        print('No Data Found, Need User input')    
        msg = "No User Info Found, Please Input the Password and Port Number of Plaxis Output (Only for Once)"
        title = "Plaxis Strut Auto Result Extraction"
        fieldNames = ["  Plaxis Output Password","  Plaxis Output Port"]
        fieldValues = []  # we start with blanks for the values
        fieldValues = multenterbox(msg,title, fieldNames)
        
        # make sure that none of the fields was left blank
        while 1:
            if fieldValues == None: break
            errmsg = ""
            for i in range(len(fieldNames)):
              if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
            if errmsg == "": break # no problems found
            fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)
        print ("Reply was:", fieldValues)
        #write pass
        username = getuser()
        file1 = open("%s.txt"%username,"w")
        L = [fieldValues[0]," \n",fieldValues[1]," \n"]
        file1.writelines(L)
        file1.close()
        return (fieldValues)

def get_user_passnport():
    #read pass
    username = getuser()
    file1 = open("%s.txt"%username, 'r') 
    Lines = file1.readlines()
    password = Lines[0].strip()
    port = Lines[1].strip()
    file1.close()
    return(password,port)
    
def get_strut_x():
    msg = "Please Enter the X Coordinate Range                                                                 (Please Make Sure the X Boundry Includes the Left Point of the Struts)"
    title = "Plaxis Strut Auto Result Extraction"
    fieldNames = ["Left X Coordinate","Right X Coordinate","Save File Location"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multenterbox(msg,title, fieldNames)
    
    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
          if fieldValues[i].strip() == "":
            errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if errmsg == "": break # no problems found
        fieldValues = multenterbox(errmsg, title, fieldNames, fieldValues)
    print ("Output Address:", fieldValues[2])
    return (fieldValues)

        

'''-------------------------------------------------------------------------------------------------------------------'''
#testing file location
#filedir = r'C:\Users\lockchi\Desktop\New folder (9)\New folder'


    
#Run Function for port and password
check_user_passnport()
userpassword,userport = get_user_passnport()

print('Please Input Strut Boundry and Output Path into Pop Up Box')

#User Select Strut input
#strut_range_min,strut_range_max = get_strut_x()

s_min,s_max,s_address= get_strut_x()
strut_range_min = int(s_min)
strut_range_max = int(s_max)
filedir = '%s'%str(s_address)
#print (strut_range_min,strut_range_max)



#start connection
s_o, g_o = new_server(socket.gethostbyname(socket.gethostname()), userport, password=userpassword)

#Run Function for result
df,temp,stage_number = plaxis_result()

#Excel Setting

#Variables
job_name = 'Job Name'
section_name = 'Section No.'
model_revision = 'Plaxis Model Revision: R%d' % 0
strut_layer_num,st_lv_ls = max_strut_in_range()
stage_num = stage_number

#Excel startup
workbook = Workbook('%s\StrutResult.xlsx' % (filedir) )
worksheet = workbook.add_worksheet('Output')

#Format Selection
#Bold Result
cell_formatRT = workbook.add_format()
cell_formatRT.set_num_format('0')
cell_formatRT.set_align('center')
cell_formatRT.set_border(2)
cell_formatRT.set_bold(True)
#Title Left Alin
cell_formatTI = workbook.add_format()
cell_formatTI.set_num_format('0')
cell_formatTI.set_align('left')
cell_formatTI.set_border(2)
cell_formatTI.set_bold(True)
#Text
cell_formatTX = workbook.add_format()
cell_formatTX.set_num_format('0')
cell_formatTX.set_align('center')
cell_formatTX.set_border(1)
#Value 2dp format
cell_format2DP = workbook.add_format()
cell_format2DP.set_num_format('0.00')
cell_format2DP.set_align('center')
#Value 4dp format      
cell_format4DP = workbook.add_format()
cell_format4DP.set_num_format('0.0000')
cell_format4DP.set_align('center')
#Value Date format   
cell_formatD = workbook.add_format()
cell_formatD.set_num_format('dd/mm/yyyy')
cell_formatD.set_align('center')
#Table with light border format   
cell_formatLB = workbook.add_format()
cell_formatLB.set_num_format('0.00')
cell_formatLB.set_align('center')
cell_formatLB.set_border(1)
#Merge Title Box Format *worksheet.merge_range(from row,from col,to row,to col, input text, merge_format)
merge_format = workbook.add_format({'align': 'center','valign': 'vcenter','border': 2,'bold': False})

#Logo Input
# data = urlretrieve("https://i.ibb.co/KWCW9cS/Capture.png", '%s\lambethlogo.png' % (filedir))
# worksheet.insert_image('A1', '%s\lambethlogo.png' % (filedir), {'x_scale': 3, 'y_scale': 1})

#Title Input
worksheet.write_column(0,1,(job_name,section_name,model_revision))
worksheet.merge_range(4,0,6,0,'Stage', merge_format)
worksheet.merge_range(4,1,6,1,'Description', merge_format)
worksheet.merge_range(2,2,2,strut_layer_num+1,'Strut Level', merge_format)
worksheet.merge_range(4,2,4,strut_layer_num+1,'Strut Forces', merge_format)
#worksheet.merge_range(4,strut_layer_num+2,4,strut_layer_num+3,'Left Wall', merge_format)
#worksheet.merge_range(4,strut_layer_num+4,4,strut_layer_num+5,'Right Wall', merge_format)
#worksheet.write_row(5,strut_layer_num+2,('V','M','V','M'),cell_formatLB)
worksheet.write_column(stage_num+7,1,('Envelopes','Spacing (m)','Force per m run (kN/m)'),cell_formatTI)


#Stage and Strut Layer Input
for i in range(2,strut_layer_num+2):
   worksheet.write(5,i, 'S%d' %(i-1) , cell_formatLB)
   worksheet.write(6,i, 'kN' , cell_formatLB)
   worksheet.write(stage_num+8,i,'1',cell_formatRT)
   worksheet.write_formula(stage_num+9,i,'=%s%d/%s%d'%(ascii_lowercase[i],stage_num+8,ascii_lowercase[i],stage_num+9),cell_formatRT)
   
for i in range(7,stage_num+7):
   worksheet.write(i,1, '%s' %(phase_name[i-7]) , cell_formatTX)
   worksheet.write(i,0, i-7 , cell_formatTX)
   #print('stage num',stage_num)
   
for i in range(7,stage_num+8):   
   #write data          
   selected_row = df.loc[(df['stage'] == i-7)&(df['x'] > strut_range_min)&(df['x'] < strut_range_max)]
   #print ('i for stage',i-8)
   #print (df.loc[(df['stage'] == i-7)])
   #print (list(selected_row['m']))
   # worksheet.write_row(i-1,2,list(selected_row['m']),cell_format2DP)
   worksheet.write_row(3,2,list(st_lv_ls),cell_format2DP)
   print(selected_row)
   print(st_lv_ls)
   print(len(selected_row.index))
   
   k = 0
   ii = 0
   while k<len(selected_row.index):
       for ia in st_lv_ls:
            if selected_row.iloc[k]['y'] == ia:
                   print('write',selected_row.iloc[k]['m'],2+k,ii)
                   # workbook = Workbook('%s\StrutResult.xlsx' % (filedir) )
                   # worksheet = workbook.add_worksheet('Output')
                   worksheet.write_row(i-1,2+ii,[selected_row.iloc[k]['m']],cell_format2DP)
                   # workbook.close()
                   k=k+1
                   if k>=len(selected_row.index):
                       break
            else: 
                    print('Ignore')
            ii = ii+1
            
       # print(row['y'], row['m'])

   
   
#Env Formula Input
for i in range(2,strut_layer_num+2):
   worksheet.write_formula(stage_num+7,i,'{=Max(Abs(%s8:%s%d))}'%(ascii_lowercase[i],ascii_lowercase[i],stage_num+7),cell_formatRT)
   
#Column Size
worksheet.set_column('A:A', 23.78)
worksheet.set_column('B:B', 53.22)
worksheet.set_column('C:Z', 14.89)


workbook.close()
print('----------------------')
print('Script Complete')
print('----------------------')










