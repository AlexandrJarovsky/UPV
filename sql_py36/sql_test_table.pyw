import pyodbc
import os
import time
from datetime import datetime, timedelta

dat=datetime.now() + timedelta(days=-28)
m_n=dat.strftime("%m")
year_n=dat.strftime("%Y")
print (dat,m_n,year_n)

def find_arch (serv, base, id_user, passwor_id):

  try:

    cnn = pyodbc.connect('DRIVER={SQL Server};SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (serv, base, id_user, passwor_id))
    cursor = cnn.cursor()
    cursor.execute('SELECT * FROM information_schema.tables')
    flag=0
  
    for row in cursor.fetchall():
      if ('_'+year_n+'_'+'%02d'% (int(m_n))) in row[2]:
        print('Сервер: '+serv+'; база: '+base+';  архив:  ' +row[2] + ' suschestvuet ', file=file_op)
        flag=1

    if flag==0:

      NewFileName = "Archiv" + '_' + year_n + '_' + m_n
      SQLStr = 'SELECT * INTO ' + NewFileName + ' FROM Table1'
      print ('3', NewFileName, SQLStr)
      cursor1 = cnn.cursor()
      cursor1.execute(SQLStr)
      cnn.commit()
      print('Сервер: '+serv+'; база: '+base+';  архив:  ' + NewFileName + ' sozdan ', file=file_op)

    cnn.close()

  except Exception as Er:
    open(os.curdir+'\\report_serv_tab.txt' , 'a' ).write( 'Ошибка: '+serv+' ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

if os.path.exists(os.curdir+'\\list_server_table.txt'): 

  try:

    file_op=open( os.curdir+r'\otchet_table_arch.txt' , 'w')
    #file_op.close()
    print (str(time.asctime())+'\n',file=file_op)

    for line in open( os.curdir +'\\list_server_table.txt' , 'r' ) : #забираем из файла конфигурации данные для функции перезапуска
      if len(line) > 10:
        print (line.split(' ')[0] + '  ' + line.split(' ')[4].rstrip('\n'), file=file_op)
        for l in line.split(' ')[1].split(','):
          find_arch(line.split(' ')[0],l,line.split(' ')[2].rstrip('\n'),line.split(' ')[3].rstrip('\n'))
      print('', file=file_op)

    file_op.close()

  except Exception as Er:
    open(os.curdir+'\\report_serv_tab.txt' , 'a' ).write( 'Ошибка: ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

else:
  open(os.curdir+'\\report._serv_tab.txt' , 'a' ).write( 'Не найден файл конфигурации: ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
