# coding: utf8
from datetime import datetime, timedelta
import pyodbc
import sqlite3
import os, shutil
from fast_update import rasch_proc

def restoredata(proc_rasch=1):
    LastDate = datetime.now()

    with open(os.curdir + '\\config.ini') as conf:
        conf_list = []
        for c in conf:
            if c != '\n':
                conf_list.append(c.strip().split('='))
        conf_dict = dict(conf_list)

        # ===================== Restore data ==========================================================
    try:

        cnn_mydb = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
        cursor = cnn_mydb.cursor()

        d_c=dict()
        d_c = {rows[1]:rows[0] for rows in cnn_mydb.execute("pragma table_info(Table1)").fetchall()}

        if len(cursor.execute("SELECT lastupdate FROM updatetime").fetchall()) != 0:

            Temp_update = datetime.strptime(cursor.execute("SELECT lastupdate FROM updatetime").fetchone()[0],conf_dict['str_datatype'])

            while (LastDate - Temp_update).days * 86400 + (
                    LastDate - Temp_update).seconds >= 0:  # and (LastDate - Temp_update).seconds/3600 > 0.5:

                if (LastDate - Temp_update).days * 86400 + (LastDate - Temp_update).seconds < 3600 and (
                        LastDate.hour == Temp_update.hour):

                    cnn_mydb.commit()
                    cursor.close()
                    cnn_mydb.close()
                    return ()

                Temp_update += timedelta(hours=1)
                print(Temp_update)

                # ===================== Chenge hour ==========================================================
                try:

                    cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))

                    for st in cursor.fetchall():

                        if st[d_c['Sz_hour']] >= 0:

                            MinusDay = 0
                            if 1 <= Temp_update.hour <= 8:
                                v_num = 1
                            elif 9 <= Temp_update.hour <= 16:
                                v_num = 2
                            elif 17 <= Temp_update.hour <= 23:
                                v_num = 3
                            elif Temp_update.hour == 0:
                                v_num = 3
                                MinusDay = -1

                            if st[d_c['Prizn']] != "":
                                if st[d_c['Prizn']] == "UROV":
                                    if st[d_c['Tzn']] > 0:
                                        cursor.execute(
                                        "UPDATE %s SET Sz_N"% (conf_dict['Table_work']) + str(
                                            (Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                            v_num) + "V= %s WHERE Shifr='%s'" % (st[d_c['Sz_hour']], st[d_c['Shifr']]))
                                    else:
                                        cursor.execute(
                                        "UPDATE %s SET Sz_N"% (conf_dict['Table_work']) + str(
                                            (Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                            v_num) + "V= %s WHERE Shifr='%s'" % (0, st[d_c['Shifr']]))
                                else:
                                    sql = "SELECT Sz_N" + str((Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V from %s WHERE Shifr='%s'" % (conf_dict['Table_work'], st[d_c['Shifr']])
                                    sql1 = "SELECT M_N" + str((Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V from %s WHERE Shifr='%s'" % (conf_dict['Table_work'], st[d_c['Shifr']])
                                    cursor.execute("UPDATE %s SET Sz_N" % (conf_dict['Table_work']) + str(
                                        (Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V= %s WHERE Shifr='%s'" % (
                                                   cursor.execute(sql).fetchone()[0] + st[d_c['Sz_hour']], st[d_c['Shifr']]))
                                    cursor.execute("UPDATE %s SET M_N" % (conf_dict['Table_work']) + str(
                                        (Temp_update + timedelta(days=MinusDay)).day) + "_" + str(
                                        v_num) + "V= %s WHERE Shifr='%s'" % (
                                                   cursor.execute(sql1).fetchone()[0] + st[d_c['Sz_hour_m']], st[d_c['Shifr']]))

                    if len(cursor.execute("SELECT lastupdate FROM updatetime").fetchall()) != 0:
                        sql = "UPDATE updatetime SET LastUpdate= '%s'" % Temp_update.strftime(conf_dict['str_datatype'])
                    else:
                        sql = "INSERT INTO updatetime VALUES ('%s')" % Temp_update.strftime(conf_dict['str_datatype'])
                    cursor.execute(sql)
                    cnn_mydb.commit()

                except Exception as Er:
                    open(os.curdir + '\\report_error.txt', 'a').write(
                        'Error perechoda chasa (restore): ' + str(Er) + ': ' + datetime.today().strftime(
                            "%d.%m.%Y %H:%M:%S") + '\n')
                    # =================================================================================================

                open(os.curdir + '\\report_error.txt', 'a').write('Restor dete :  ' + str(Temp_update) + '\n')

                if Temp_update.hour == 0:

                    if Temp_update.day == 1:

                        try:

                            if proc_rasch==1:

                                rasch_proc(cursor, conf_dict, d_c)
                                cnn_mydb.commit()

                        except Exception as Er:
                            open(os.curdir + '\\report_error.txt', 'a').write(
                                'Error rascheta procentov (restore nachalo meciyca): ' + str(
                                    Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')

                        # =============       Создать архив в начале месяца         =================================
                        try:

                            NewFileName = 'Archiv_' + (datetime.today() + timedelta(days=-28)).strftime("%Y_%m")
                            cursor.execute("CREATE TABLE " + NewFileName + " AS SELECT * FROM %s" % (conf_dict['Table_work']))
                            cnn_mydb.commit()

                            open(os.curdir + '\\report_error.txt', 'a').write(
                                 datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Archiv created (restore): ' + NewFileName + '\n')
	
                        except Exception as Er:
                            open(os.curdir + '\\report_error.txt', 'a').write(
                                'Error sozdaniy archiva (restore): ' + NewFileName + ': ' + str(
                                    Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')

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

                            open(os.curdir + '\\report_error.txt', 'a').write(Temp_update.strftime(
                                "%d.%m.%Y %H:%M:%S") + ' Ochistka basy (restore) ' + '\n')
                        except Exception as Er:
                            open(os.curdir + '\\report_error.txt', 'a').write(
                                ' Error ochistki basy (restore) : ' + str(
                                    Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
                            # ====================================================================================

            # =============       Расчет процентов        =================================
            try:

                if proc_rasch==1:

                    rasch_proc(cursor, conf_dict, d_c)

            except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error rascheta procentov (restore): ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')
                # ===========================================================================================================

        cnn_mydb.commit()
        cursor.close()
        cnn_mydb.close()
        print('Данные востановлены')

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write \
	    ('Error pri vostanovlenii dannych: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')

        # =====================================================================================================


if __name__ == '__main__':
    restoredata()
