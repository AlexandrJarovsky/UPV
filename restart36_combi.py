from datetime import datetime
import os, subprocess

def reset_po_time(filename,taskname,exename): #функция перезапуска програмы при отсутсвии обновления более 2 минут

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
            if len(subprocess.check_output('tasklist /FI "WINDOWTITLE eq ' + taskname + '" /nh').splitlines())>1:
              subprocess.check_output('taskkill /F /FI "WINDOWTITLE eq ' + taskname + '"')

            subprocess.Popen(exename, cwd=os.path.split(exename)[0])
          print( taskname + '\n', 'Last restart     : '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n', 'Last modification: '+ line, file=open( os.path.abspath('.')+r'\restart_count.txt' , 'a' ))
  
      else:
        open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Not found startfile: ' + exename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

    else:
      open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Not found updatefile: ' + filename + ' ' + today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

      if '.exe' in taskname:
        os.system("taskkill /f /im " + taskname)
        os.startfile(exename)
      else:
        if len(subprocess.check_output('tasklist /FI "WINDOWTITLE eq ' + taskname + '" /nh').splitlines())>1:
          subprocess.check_output('taskkill /F /FI "WINDOWTITLE eq ' + taskname + '"')

        if os.path.exists(exename): 
          subprocess.Popen(exename, cwd=os.path.split(exename)[0])

      print( taskname + '\n', 'Last restart     : '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n', file=open( os.path.abspath('.')+r'\restart_count.txt' , 'a' ))

  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Error module res_test: ' + str(Er) +': '+ today.strftime("%d.%m.%Y %H:%M:%S") + '\n' )

if os.path.exists(os.path.abspath('.')+'\\config.txt'): 

  try:
  
    for line in open( os.path.abspath('.')+'\\config.txt' , 'r' ) : #забираем из файла конфигурации данные для функции перезапуска
      if len(line) > 10 and 'rem' not in line:
        s=line.rstrip('\n').split()
        reset_po_time(s[0],s[1],s[2])
  
  except FileNotFoundError as fn:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Not found file : ' + str(fn) +': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
  
  except Exception as Er:
    open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Error: ' + str(Er) +': '+ datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )

else:
  open( os.path.abspath('.')+'\\report.txt' , 'a' ).write( 'Not found configurationfile: ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n' )
