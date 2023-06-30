# coding: utf8

from datetime import datetime, timedelta
import sqlite3
import os, threading, shutil
import mb_print
import fast_update
from multiprocessing import Process
from time import sleep


with open(os.curdir + '\\config.ini') as conf:
    conf_list = []
    for c in conf:
        if c != '\n':
            conf_list.append(c.strip().split('='))
    conf_dict = dict(conf_list)

LastDate = datetime.now()
Flag_prn = False
HFlag = False

def opros(proc_rasch=1, lab_density=1): # 15 sec...
    try:

        global LastDate, Flag_prn, HFlag
        HourFlag = False
        TDate = datetime.now()

        try:
            cnn_mydb = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
            cursor = cnn_mydb.cursor()

            d_c=dict()
            d_c = {rows[1]:rows[0] for rows in cnn_mydb.execute("pragma table_info(Table1)").fetchall()}

        except sqlite3.DatabaseError as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
		         'Error database connect: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
            return

        try:
            cnn_rem = sqlite3.connect(conf_dict['DATABASE_REM'])
            cur_rem = cnn_rem.cursor()
        except sqlite3.DatabaseError as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error remuv database connect: ' + str(Er) + ': ' + datetime.today().strftime(
                    "%d.%m.%Y %H:%M:%S") + '\n')

        # ===================== обработка перехода часа==========================================================
        try:

            if LastDate.hour != TDate.hour:
                HourFlag = True
                cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))

                for st in cursor.fetchall():

                    if st[d_c['Sz_hour']] >= 0:

                        MinusDay = 0
                        if 1 <= TDate.hour <= 8:
                            v_num = 1
                        elif 9 <= TDate.hour <= 16:
                            v_num = 2
                        elif 17 <= TDate.hour <= 23:
                            v_num = 3
                        elif TDate.hour == 0:
                            v_num = 3
                            MinusDay = -1

                        if st[d_c['Prizn']] != "":
                            if st[d_c['Prizn']] == "UROV":
                                if st[d_c['Tzn']] > 0:
                                    cursor.execute(
                                    "UPDATE %s SET Sz_N"% (conf_dict['Table_work']) + str(
                                        (TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V= %s WHERE Shifr='%s'" % (st[d_c['Sz_hour']], st[d_c['Shifr']]))
                                else:
                                    cursor.execute(
                                    "UPDATE %s SET Sz_N"% (conf_dict['Table_work']) + str(
                                        (TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                            else:
                                sql = "SELECT Sz_N" + str((TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                    v_num) + "V from %s WHERE Shifr='%s'" % (conf_dict['Table_work'], st[d_c['Shifr']])
                                sql1 = "SELECT M_N" + str((TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                    v_num) + "V from %s WHERE Shifr='%s'" % (conf_dict['Table_work'], st[d_c['Shifr']])
                                cursor.execute("UPDATE %s SET Sz_N" % (conf_dict['Table_work']) + str(
                                    (TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                    v_num) + "V= %s WHERE Shifr='%s'" % (
                                               cursor.execute(sql).fetchone()[0] + st[d_c['Sz_hour']], st[d_c['Shifr']]))
                                cursor.execute("UPDATE %s SET M_N" % (conf_dict['Table_work']) + str(
                                    (TDate + timedelta(days=MinusDay)).day) + "_" + str(
                                    v_num) + "V= %s WHERE Shifr='%s'" % (
                                               cursor.execute(sql1).fetchone()[0] + st[d_c['Sz_hour_m']], st[d_c['Shifr']]))

                            cursor.execute(
                                "UPDATE %s SET n= %s WHERE Shifr='%s'" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                            cursor.execute(
                                "UPDATE %s SET Sz_hour= %s WHERE Shifr='%s'" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                            cursor.execute(
                                "UPDATE %s SET Sz_hour_m= %s WHERE Shifr='%s'" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))

                if len(cursor.execute("SELECT lastupdate FROM updatetime").fetchall()) != 0:
                    sql = "UPDATE updatetime SET LastUpdate= '%s'" % TDate.strftime(conf_dict['str_datatype'])
                else:
                    sql = "INSERT INTO updatetime VALUES ('%s')" % TDate.strftime(conf_dict['str_datatype'])
                cursor.execute(sql)
                cnn_mydb.commit()
                open(os.curdir + '\\report_error.txt', 'a').write(
                    datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ' : Refresh of chenge hour' + '\n')

        except Exception as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error on refresh of chenge hour: ' + str(Er) + ': ' + datetime.today().strftime(
                    "%d.%m.%Y %H:%M:%S") + '\n')
        # =================================================================================================

        # =============================обработка каждые 15 сек====================================
        try:

            ''''up_proc = Process(target=fastupdate, args=(conf_dict, d_c), daemon=True)
            up_proc.start()
            time_out = 0
            while up_proc.is_alive():
                sleep(0.5)
                time_out += 1
                if time_out == 20:
#                    up_proc.terminate()
                    break
            print(time_out)'''

            fast_update.fastupdate(cursor, cur_rem, conf_dict, d_c)
            cnn_mydb.commit()

            '''if lab_density==1:

                fast_update.lab_densiti(cursor, conf_dict, d_c)
                cnn_mydb.commit()'''

        except Exception as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error refresh now : ' + str(Er) + ': ' + datetime.today().strftime(
                    "%d.%m.%Y %H:%M:%S") + '\n')
        # ===========================================================================================================

        if HourFlag:
#        if not HourFlag:
            # ============= Пересчет процентов за сутки при переходе часа =================================
            try:

                if proc_rasch==1:

                    fast_update.rasch_proc(cursor, conf_dict, d_c)
                    cnn_mydb.commit()

                    open(os.curdir + '\\report_error.txt', 'a').write(
                        datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ' : raschet procentov' + '\n')

            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error on raschet procentov(opros) : ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')
            # ===========================================================================================================
            # ============= Данные из лаборатории по плотностям =================================
            try:

                if lab_density==1:

                    fast_update.lab_densiti(cursor, conf_dict, d_c)
                    cnn_mydb.commit()

                    open(os.curdir + '\\report_error.txt', 'a').write(
                        datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ' : Obnovlenie plotnosti' + '\n')

            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error on refresh lab analis(opros) : ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')
            # ===========================================================================================================

        if Flag_prn:

            # ============================ Ежедневный бэкап базы =======================================
            try:

                db_name=os.path.split(cursor.execute("PRAGMA database_list;").fetchall()[0][2])[1]
                shutil.copyfile(os.path.abspath(os.curdir) + conf_dict['DATABASE'], 
                                os.path.abspath(os.curdir) + '\\BACKUP\\' + str(datetime.now().date()) +  '_' + db_name)
#                shutil.copyfile(os.path.abspath(os.curdir) + conf_dict['DATABASE'], os.path.abspath(os.curdir) + '\\BACKUP\\' + db_name)

                open(os.curdir + '\\report_error.txt', 'a').write(
                    datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Backup created \n')

            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error on backup created: ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')

            Flag_prn = False

            set_print = cursor.execute("SELECT print_check FROM seting").fetchone()
            if len(cursor.execute("SELECT print_check FROM seting").fetchall()) != 0:
#                mb_print.prin_pr(flag_prn=int(set_print[0]))
                threading.Thread(target=lambda: mb_print.prin_pr(flag_prn=int(set_print[0])),args=()).start()
            else:
#                mb_print.prin_pr()
                threading.Thread(target=lambda: mb_print.prin_pr(),args=()).start()

        # ====================================================================================

        # ============================ Создание архива и очистка базы в начале месяца===================================
        if HFlag and TDate.day == 1 and TDate.hour == 0 and TDate.minute == 45:
            #    if not HFlag :

            try:

                NewFileName = 'Archiv_' + (datetime.today() + timedelta(days=-28)).strftime("%Y_%m")
                cursor.execute("CREATE TABLE " + NewFileName + " AS SELECT * FROM %s" % (conf_dict['Table_work']))
                cnn_mydb.commit()
                open(os.curdir + '\\report_error.txt', 'a').write(
                    datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Archiv created ' + NewFileName + '\n')

            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error on archiv created ' + NewFileName + ': ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')

            try:
                cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
                for st in cursor.fetchall():
                    for i in range(1, 32):
                        cursor.execute("UPDATE %s SET Sz_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "1V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET Sz_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "2V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET Sz_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "3V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET M_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "1V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET M_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "2V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET M_N" % (conf_dict['Table_work']) + str(
                            i) + "_" + "3V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))

                cnn_mydb.commit()
                open(os.curdir + '\\report_error.txt', 'a').write(
                    datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ' : Clear base complit' + '\n')
            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error on clear base : ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')

            HFlag = False
        # ====================================================================================

        # ======================== Выставление флагов ============================================================

        if HourFlag:

            if TDate.hour == 0:

                Flag_prn = True
                if TDate.day == 1:
                    HFlag = True

                if os.path.exists(os.curdir + '\\report_error.txt'):
                    fil1 = os.path.getsize(os.curdir + '\\report_error.txt')
                    if fil1 > 1000000:
                        os.remove(os.curdir + '\\report_error.txt')

        # ====================================================================================

        LastDate = TDate
        cursor.close()
        cur_rem.close()
        cnn_mydb.close()
        cnn_rem.close()

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error in opros : ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')


if __name__ == '__main__':
    opros()