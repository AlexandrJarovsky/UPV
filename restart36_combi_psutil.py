from datetime import datetime
import os, subprocess, time
import psutil

def reset_po_time(filename, taskname, exename='', exename1=''): #функция перезапуска програмы при отсутсвии обновления более 2 минут

  today = datetime.today()
  try:

    if os.path.exists(filename):

      if os.path.exists(exename): 

        line=open( filename , 'r' ).read() 

        if not line:
          line=str(today)
          open( filename , 'w' ).write(today.strftime("%d.%m.%Y %H:%M:%S"))

        m=line.rstrip('\n').split()[1].split(':')[1]
        h=line.rstrip('\n').split()[1].split(':')[0]

        if int(today.hour) != int(h) or int(today.minute)-int(m) > 2:

          if '.exe' in taskname:
            os.system("taskkill /f /im " + taskname)
            os.startfile(exename)
          else:
            for proc in psutil.process_iter(['name', 'cmdline']):
              if "python" in proc.info['name'] and exename in str(proc.info['cmdline']):
#              if "python" in proc.info['name']:
                print(proc.info)
                proc.terminate()
                time.sleep(1)

#            if len(subprocess.check_output('tasklist /FI "WINDOWTITLE eq ' + taskname + '" /nh').splitlines())>1:
#              subprocess.check_output('taskkill /F /FI "WINDOWTITLE eq ' + taskname + '"')
            if os.path.exists(exename1): 
              subprocess.Popen(exename1, cwd=os.path.split(exename1)[0])

          print( taskname + '\n', 'Last restart     : '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n', 'Last modification: '+ line, file=open( os.path.abspath('.')+r'\restart_count.txt' , 'a' ))
  
      else:
        open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл для запуска: ' + exename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

    else:
      open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл обновлений: ' + filename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

      if '.exe' in taskname:
        os.system("taskkill /f /im " + taskname)
        os.startfile(exename)
      else:

        for proc in psutil.process_iter(['name', 'cmdline']):
          if "python" in proc.info['name'] and exename in str(proc.info['cmdline']):
            print(proc.info)
            proc.terminate()
            time.sleep(1)

#        if len(subprocess.check_output('tasklist /FI "WINDOWTITLE eq ' + taskname + '" /nh').splitlines())>1:
#          subprocess.check_output('taskkill /F /FI "WINDOWTITLE eq ' + taskname + '"')

        if os.path.exists(exename1): 
          subprocess.Popen(exename1, cwd=os.path.split(exename1)[0])

      print( taskname + '\n', 'Last restart     : '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n', file=open( os.path.abspath('.')+r'\restart_count.txt' , 'a' ))

  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Ошибка модуля res_test: ' + str(Er) +': '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

if os.path.exists(os.path.abspath('.')+'\\config_.txt'): 

  try:
  
    for line in open( os.path.abspath('.')+'\\config_.txt' , 'r' ) : #забираем из файла конфигурации данные для функции перезапуска
      if len(line) > 10 and 'rem' not in line:
        s=line.rstrip('\n').split()
        if len(s)==4:
          reset_po_time(s[0],s[1],s[2],s[3])
        elif len(s)==3:
          reset_po_time(s[0],s[1],s[2])
        else:
          open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не корректные значения в строке конфигурации: ' 
                + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )
  
 
  except FileNotFoundError as fn:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл : ' + str(fn) +': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
  
  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Ошибка: ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

else:
  open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Не найден файл конфигурации: ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
