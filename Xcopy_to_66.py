import pyodbc
import sys
import time
import datetime 
#import shutil
#import codecs
#import zipfile
#import logging
#import logging.config
#import threading
import sqlite3
#import pdb;pdb.set_trace() # h-help
#***********************************************************************************************
#Init start Time
startTm=time.time()
#***********************************************************************************************
#declare variables

#Remote server
#str_connect_R='DRIVER={SQL Server};SERVER=10.100.100.61;DATABASE=USTANOVKI;UID=as-test;PWD=test'
str_connect_R='DRIVER={SQL Server};SERVER=10.100.100.61;DATABASE=tep_mb;UID=sa;PWD=google3519google'

tagtblname = 'Table1'	

current_datetime=time.strftime("%d/%m/%y",time.localtime())

#define thread procedure
def update_table(tbl_x, db_x):
	try:

    #open remote database
		conR = pyodbc.connect(str_connect_R)		
		curR = conR.cursor()

		#open local database
		conL = sqlite3.connect(db_x)
		curL = conL.cursor()
	
		str_exec_l="select * from %s " % (tagtblname)
		curL.execute(str_exec_l)
		rows=curL.fetchall()
		countRows=len(rows)

		curR.execute('delete from {}'.format(tbl_x))
		print ('delete records from 44...')
		conR.commit()
		print ('records deleted...')

		i=0
		                    
		for row in rows:
			i+=1

			print ('[{}]->'.format(i),row[0]) #,type(row[1])
			str_exec_r="insert into {} values{}".format(tbl_x, row)

			curR.execute(str_exec_r)
			conR.commit()

		print ('update {}  [{}]rec'.format(tbl_x,countRows))
	except:
		print(sys.exc_info()[0])
	finally:
		curR.close()
		conR.close()
		curL.close()
		conL.close()
#**************************************************************************************************
#main programm
if __name__ == "__main__":
	try:
		arr_table=['mb_upv', 'tep_upv']
		dbTag = [r'd:\SQL\MB_UPV_PYTHON\db\mb_upv.db', r'd:\SQL\TEP_PYTHON\db\tep_upv.db']

		for tab, db in zip(arr_table, dbTag):
			update_table(tab, db)

	finally:
		print(time.time()-startTm)