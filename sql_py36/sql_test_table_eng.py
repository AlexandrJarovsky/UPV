import pyodbc
import os
import time
import sqlite3
from datetime import datetime, timedelta

dat=datetime.now() + timedelta(days=-28)
m_n=dat.strftime("%m")
year_n=dat.strftime("%Y")
print (dat,m_n,year_n)

def find_arch (base, serv):

  try:

#    cnn = pyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (serv, base, id_user, passwor_id))
    print(base)
    cnn = sqlite3.connect(str(base))
    cursor = cnn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    flag=0
  
    for row in cursor.fetchall():
      if ('_'+year_n+'_'+'%02d'% (int(m_n))) in row[0]:
        print('Server: '+serv+'; base: '+os.path.split(base)[1]+';  archiv:  ' +row[0] + ' suschestvuet ', file=file_op)
        flag=1

    if flag==0:

      NewFileName = "Archiv" + '_' + year_n + '_' + m_n
#      SQLStr = 'SELECT * INTO ' + NewFileName + ' FROM Table1'
      SQLStr = "CREATE TABLE " + NewFileName + " AS SELECT * FROM Table1"
      print ('3', NewFileName, SQLStr)
      cursor1 = cnn.cursor()
      cursor1.execute(SQLStr)
      cnn.commit()
      print('Server: '+serv+'; base: '+base+';  archiv:  ' + NewFileName + ' sozdan ', file=file_op)

    cnn.close()

  except Exception as Er:
    open(os.curdir+'\\report_serv_tab.txt' , 'a' ).write( 'Error: '+serv+' ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

if os.path.exists(os.curdir+'\\list_server_table.txt'): 

  try:

    file_op=open( os.curdir+r'\otchet_table_arch.txt' , 'w')
    #file_op.close()
    print (str(time.asctime())+'\n',file=file_op)

    for line in open( os.curdir +'\\list_server_table.txt' , 'r' ) : #забираем из файла конфигурации данные для функции перезапуска
      if 'rem' not in line:
        if len(line) > 20:
          find_arch(line.split(' ')[0],line.split(' ')[1].rstrip('\n'))
        else:
          print (line.split(' ')[0].rstrip('\n'), file=file_op)

#      print('', file=file_op)

    file_op.close()

  except Exception as Er:
    open(os.curdir+'\\report_serv_tab.txt' , 'a' ).write( 'Error: ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

else:
  open(os.curdir+'\\report._serv_tab.txt' , 'a' ).write( 'Not found file configure: ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
