from datetime import datetime
import os

def reset_po_time(filename,taskname,exename): #функция перезапуска програмы при отсутсвии обновления более 2 минут

  today = datetime.today()
  try:

    if os.path.exists(filename):

      if os.path.exists(exename): 

        line=open( filename , 'r' ).read() 

        if not line:
          line=str(today)

        m=line.rstrip('\n').split()[1].split(':')[1]
        h=line.rstrip('\n').split()[1].split(':')[0]

        if int(today.hour) != int(h) or int(today.minute)-int(m) > 2:
          os.system("taskkill /f /im " + taskname)
          os.startfile(exename)
          print( taskname + '\n', 'Last restart     : '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n', 'Last modification: '+ line, file=open( os.path.abspath('.')+r'\restart_count.txt' , 'a' ))
  
      else:
        open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл для запуска: ' + exename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

    else:
      open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл обновлений: ' + filename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Ошибка модуля res_test: ' + str(Er) +': '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

if os.path.exists(os.path.abspath('.')+'\\config.txt'): 

  try:
  
    for line in open( os.path.abspath('.')+'\\config.txt' , 'r' ) : #забираем из файла конфигурации данные для функции перезапуска
      if len(line) > 10:
        s=line.rstrip('\n').split()
        reset_po_time(s[0],s[1],s[2])
  
  except FileNotFoundError as fn:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл : ' + str(fn) +': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
  
  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Ошибка: ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

else:
  open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл конфигурации: ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
