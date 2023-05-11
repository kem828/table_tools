# -*- coding: utf-8 -*-
"""
Created on Fri Oct 27 10:25:59 2017

@author: 388560
Keinan Marks
keinan@keinanmarks.com
"""
import pandas as pd
from appJar import gui
import gc
import difflib,os
app = gui()



#add filepath to list when selected from file entry in the append tool
def addlistitem(opass):
    app.addListItem('selected',app.getEntry('appendsel'))

#Gets all spreadsheets in a folder
#Used in the append tools "join contents of folder" setting
def get_files(workpathin):
    sheetlist = []
    for file in os.listdir(workpathin):
        if file.endswith(".xls") or file.endswith('.xlsx') or file.endswith('.csv'):
            sheetlist.append(workpathin +'/' + str(file))
    return sheetlist

#Append Tool Functions
#load and concat tables
#This is a tool to append a number of tables together, by using matching headers
#any header that is not matched in a subsequent table leads to the creation of a new\
#header. Also includes option to append together all sheets in a folder, as well
#as to designate header row (default row 1)
#as of 6/5/18 GUI could use some redesign (but it might be as good as appJar gets)
def append(opass):
    folder_status = app.getCheckBox('Join Contents of Folder')
    outpath = app.getEntry('outpatha')
    outname = app.getEntry('outnamea')
    if folder_status == False:
        intab = app.getEntry('primtab')
    else:
        folder = app.getEntry('direntry')
        tabs = get_files(folder)
        intab = tabs[0]
    try:
        merges = tabs[1:]
    except:
        pass
    try:
        header_row = int(app.getEntry('Header Row')) - 1
    except:
        header_row = 0
    try:
        if header_row > 0:
            pass
        else:
            header_row = 0
    except:
        header_row = 0
    print('outputting to ' + outpath + outname)
    print('Reading Primary Table')
    try:
        if '.csv' in intab:
            prim = pd.read_csv(intab,memory_map=True,skiprows = header_row,keep_default_na=False,na_values=['#N/A'])
        elif '.xls' in intab or '.xlsx' in intab:
            prim = pd.read_excel(intab,0,skiprows = header_row,keep_default_na=False,na_values=['#N/A'])
        else:
            app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
        primdf = pd.DataFrame(data=prim)

    except Exception as e:
        app.warningBox('exceptionfuncjoin',e,parent=None)
    print(app.getRadioButton('select'))
    if folder_status == False:
        if app.getRadioButton('select') == 'Highlighted Tables':
            print(' Joining selected tables only!')
            merges = app.getListBox('selected')
        else:
            print('Joining all tables!')
        merges = app.getAllListItems('selected')
    try:
        merged = primdf
        for tabs in merges:
            print(tabs)
            if '.csv' in tabs:
                m1 = pd.read_csv(tabs,skiprows = header_row,keep_default_na=False,na_values=['#N/A'])
            elif '.xls' in intab or '.xlsx' in intab:
                m1 = pd.read_excel(tabs,0,skiprows = header_row,keep_default_na=False,na_values=['#N/A'])
            else:
                app.warningBox('badfile','The files must  be csv, xls or xlsx',parent=None)
            m2 = pd.DataFrame(data = m1)
            print('joining table ' + tabs)
            if app.getCheckBox('Shared Columns only') == True:
                merged = pd.concat([merged,m2],join = 'inner',ignore_index=True)
            else:
                merged = pd.concat([merged,m2],ignore_index=True)
        joined = merged
        print('Converting back to Excel')
        writer = pd.ExcelWriter(outpath+'/'+outname+'.xlsx')
        joined.to_excel(writer)
        print('Writing File')
        writer.save()
        app.infoBox(':-)','Merge Complete',parent=None)
    except Exception as e:
        app.warningBox('exceptionfuncjoin',e,parent=None)


#This is a very simple function to create a list of potential matches from difflib get closests matches
#It then uses the first return in that list (the highest scoring) to attempt a join
#Currently takes a long time. Matches are pretty decent though

def lambda_diffy(word, possibilities):
    match_list = difflib.get_close_matches(str(word), possibilities)
    if len(match_list) == 0:
        match = 'No Close Match'
    else:
        match = match_list[0]
    return match

def fuzzy_matcher(intab,jointab,leftkey,rightkey,jointype):
    if '.csv' in intab:
        ware = pd.read_csv(intab,memory_map=True)
        wbill = pd.read_csv(jointab,memory_map=True)#.merge(wbill,left_on = leftkey,right_on = rightkey,how = jointype)
    elif '.xls' in intab or '.xlsx' in intab:
        ware = pd.read_excel(intab,0)
        wbill = pd.read_excel(jointab,0)
    else:
        app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
    
    wbill['Original Key Field Value'] = wbill[rightkey]
    wbill[rightkey] = wbill[rightkey].apply(lambda x: lambda_diffy(x, ware[leftkey]))


    joined = ware.merge(wbill,left_on = leftkey,right_on = rightkey,how = jointype)
    return joined
        



#Join Tool Functions  
#Join Function
#This function performs the type of join selected (default left) on the two
#selected tables and selected key fields. Has an option for an experimental
#fuzzy logic matcher
def join(take):
    intab=app.getEntry('intab')
    jointab = app.getEntry('jointab')
    leftkey = app.getEntry('leftjoin')
    rightkey = app.getEntry('rightjoin')
    outname = app.getEntry('outnamej')
    outpath = app.getEntry('outpathj')
    jointype = app.getOptionBox("Join Type:")
    leftkeydrop = app.getOptionBox('incol')
    rightkeydrop = app.getOptionBox('incol2')
    if len(leftkeydrop)>len(leftkey):
        leftkey = leftkeydrop
    if len(rightkeydrop)>len(rightkey):
        rightkey = rightkeydrop
    print(leftkey)
    #If the "Fuzzy Logic" checkbox is marked off, run through the fuzzy matching
    #function instead of the normal pandas merge function
    if app.getCheckBox('Fuzzy Logic Matching (experimental)'):
        try:
            joined = fuzzy_matcher(intab,jointab,leftkey,rightkey,jointype)
            writer = pd.ExcelWriter(outpath+'/'+outname+'.xlsx')
            joined.to_excel(writer)
            writer.save()
            app.infoBox(':-)','Join Complete',parent=None)
        except Exception as e:
            app.warningBox('exceptionfuncjoin',e,parent=None)
    else:
        try:
            if '.csv' in intab:
                ware = pd.read_csv(intab,memory_map=True,keep_default_na=False,na_values=['#N/A'])
                wbill = pd.read_csv(jointab,memory_map=True,keep_default_na=False,na_values=['#N/A'])
                joined = ware.merge(wbill,left_on = leftkey,right_on = rightkey,how = jointype)
                joined.to_csv(outpath+'/'+outname+'.csv',sep=',')
            elif '.xls' in intab or '.xlsx' in intab:
                ware = pd.read_excel(intab,0,keep_default_na=False,na_values=['#N/A'])
                wbill = pd.read_excel(jointab,0,keep_default_na=False,na_values=['#N/A'])
                joined = ware.merge(wbill,left_on = leftkey,right_on = rightkey,how = jointype)
                writer = pd.ExcelWriter(outpath+'/'+outname+'.xlsx')
                joined.to_excel(writer)
                writer.save()
            else:
                app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
            app.infoBox(':-)','Join Complete',parent=None)
        except Exception as e:
            app.warningBox('exceptionfuncjoin',e,parent=None)

#simple function to get and return all uniques in a chosen field in a dataframe
def get_uniques(frame,field):
    try:
        vals = frame[field].unique()
        print(vals)
        return sorted(vals)
    except Exception as e:
        app.warningBox('exceptionfuncjoin','Error during Unique value Collection\n'+str(e),parent=None)
        
#this is the split function. It collects all unique values in the chosen field
#and iterates through them, copying values that match the unique into a new worksheet
#for each one
def split(opass):
    field = app.getOptionBox('fields_to_split')
    table =app.getEntry('split_table')
    outpath = app.getEntry('outpathf')
    outname = app.getEntry('outnamef')
    if '.csv' in table:
        intable = pd.read_csv(table)
    elif '.xls' in table or '.xlsx' in table:
        intable = pd.read_excel(table,0)
    uniques = get_uniques(intable,field)
    for unique in uniques:
        filtered = intable[field] == unique
        filtered = intable[filtered]
        writer = pd.ExcelWriter(outpath + '/'+ outname + str(unique) + '.xlsx')
        filtered.to_excel(writer)
        writer.save()
        gc.collect
    app.infoBox(':-)','Split Complete',parent=None)
#Update left table field titles 
def pickjoinin(opass):
    app.openSubWindow('Join Tool')
    print(opass)
    jin = app.getEntry('intab')
    if '.csv' in jin:
        intable = pd.read_csv(jin)
        #wbill = pd.read_csv(jointab)
    elif '.xls' in jin or '.xlsx' in jin:
        intable = pd.read_excel(jin,0)
        #wbill = pd.read_excel(jointab,0)
    else:
        app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
    cols = list(intable.columns.values)
    try:
        app.addOptionBox('incol',cols,6,0)
    except:
        app.changeOptionBox('incol',cols)

#update right table field titles
def pickjoinon(opass2):
    app.openSubWindow('Join Tool')
    print(opass2)
    jin2 = app.getEntry('jointab')
    if '.csv' in jin2:
        intable = pd.read_csv(jin2)
        #wbill = pd.read_csv(jointab)
    elif '.xls' in jin2 or '.xlsx' in jin2:
        intable = pd.read_excel(jin2,0)
        #wbill = pd.read_excel(jointab,0)
    else:
        app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
    cols2 = list(intable.columns.values)
    try:
        app.addOptionBox('incol2',cols2,8,0)
    except:
        app.changeOptionBox('incol2',cols2)

def fill_column_values_split(opass):
    app.openSubWindow('Field Split Tool')
    cols = ['Fields To Split On']
    jin = app.getEntry('split_table')
    if '.csv' in jin:
        intable = pd.read_csv(jin)
    elif '.xls' in jin or '.xlsx' in jin:
        intable = pd.read_excel(jin,0)
    else:
        app.warningBox('badfile','The files must both be csv, xls or xlsx',parent=None)
    for values in list(intable.columns.values):
        cols.append(values)
    try:
        app.changeOptionBox('fields_to_split',cols)

    except:
        app.addOptionBox('fields_to_split',cols,4,0)

def externalDrop(data):
    print("Data dropped:", data)

#App selector GUI (launch window)
def launch(win):
    app.showSubWindow(win)

def main():
#    import pandas as pd
#    from appJar import gui
#    import gc
#    import difflib,os
#    #Build GUI
#    app = gui()
    
#This main function builds all the GUI Elements
#Each toolset has its own subwindow that is called from
#a main window
#Typically, selecting a file will load that field in order to populate a 
#column list. This isn't too bad in pandas
    
    #Join GUI (Subwindow 1)
    app.startSubWindow('Join Tool')
    app.addLabel('Join Tool','Select Tables to Join:\n')
    app.addLabel('intablab', 'Table to Join to (left)',1,0)
    app.addFileEntry('intab',2,0)
    app.addLabel('jointab', 'Table to Join in (right)',3,0)
    app.addFileEntry('jointab',4,0)
    app.addLabel('leftjoin', 'Left Join Field',5,0)
    app.addEntry('leftjoin',6,0)
    app.addLabel('rightjoin', 'Right Join Field',7,0)
    app.addEntry('rightjoin',8,0)
    app.addLabel('outpathj', 'Folder to save to')
    app.addDirectoryEntry('outpathj')
    app.addLabel('outnamej', 'Output file name')
    app.addEntry('outnamej')
    app.addLabelOptionBox("Join Type:", ['left','right','outer','inner'])
    app.addCheckBox('Fuzzy Logic Matching (experimental)')
    app.addButton('Join',join)
    app.setEntryDefault('leftjoin','must be exact name of column')
    app.setEntryDefault('rightjoin', 'must be exact name of column')
    app.setEntryChangeFunction('intab',pickjoinin)
    app.setEntryChangeFunction('jointab',pickjoinon)
    app.stopSubWindow()
    
    
    #Append GUI(Subwindow 2)
    app.startSubWindow('Append Tool')
    app.addLabel('append','Select Tables to Combine:\n')
    app.addLabel('primtablab', 'Primary Table to Append to',1,0,2)
    app.addFileEntry('primtab',2,0,1)
    app.addDirectoryEntry('direntry',2,1,1)
    app.addLabel('appendlab', 'Select tables to append',3,0,2)
    app.addFileEntry('appendsel',4,0,2)
    app.addLabel('selectlab','Tables for merge:',9,0,2)
    app.addLabel('checklab','Select whether all tables below will be \npart of the merge, or only a selection',5,0,2)
    app.addListBox('selected',(''),10,0,2)
    app.addRadioButton('select','All Tables',6,0)
    app.addRadioButton('select', 'Highlighted Tables',6,1)
    app.addCheckBox('Shared Columns only',7,0)
    app.addCheckBox('Join Contents of Folder',7,1)
    app.addLabelEntry('Header Row',8,0,2)
    app.setEntryDefault('Header Row','1')
    app.setListBoxMulti('selected',multi=True)
    app.setEntryChangeFunction('appendsel',addlistitem)
    app.addLabel('outpatha', 'Folder to save to',11,0,2)
    app.addDirectoryEntry('outpatha',12,0,2)
    app.addLabel('outnamea', 'Output file name',13,0,2)
    app.addEntry('outnamea',14,0,2)
    app.addButton('Merge Tables',append,15,0,2)
    app.setEntryDropTarget('appendsel')
    app.stopSubWindow()
    
    #Split GUI (Subwindow 3)
    app.startSubWindow('Field Split Tool')
    app.addLabel('split_table', 'Table to Split',1,0,2)
    app.addFileEntry('split_table',2,0,2)
    app.addOptionBox('fields_to_split',['Fields To Split On'],4)
    app.setEntryChangeFunction('split_table',fill_column_values_split)
    app.addLabel('outpathf', 'Folder to save to',5,0)
    app.addDirectoryEntry('outpathf',6,0)
    app.addLabel('outnamef', 'Output file name (+Unique Value)',7,0)
    app.addEntry('outnamef',8,0)
    app.addButton('Split Table',split,9,0)
    app.stopSubWindow()
    
    
    #Main Screen GUI
    app.setTitle('Table Tools v.45 - 9/28/18 - Keinan Marks')
    app.addButton('Join Tool',launch,1)
    app.addButton('Append Tool',launch,2)
    app.addButton('Field Split Tool',launch,3)
    
    #Launch GUI
    app.go()

            
            
            
main()

