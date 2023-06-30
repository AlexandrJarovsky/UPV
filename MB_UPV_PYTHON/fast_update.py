# coding: utf8

#from curses.ascii import NUL
from datetime import datetime, timedelta
import pyodbc
import os, time
import sqlite3

def lab_densiti(cursor, conf_dict, d_c):

    # ============= Данные из лаборатории по плотностям =================================
    try:

        str_cnn = 'DRIVER=%s;SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (
        conf_dict['DRIVER'], conf_dict['SERVER_LAB'], conf_dict['DATABASE_LAB'], conf_dict['UID_LAB'],
        conf_dict['PWD_LAB'])
        cnn_lab = pyodbc.connect(str_cnn, readonly=True)
        cur_lab = cnn_lab.cursor()

        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
        for st in cursor.fetchall():

            if (st[d_c['kod']] != 0) and (st[d_c['nom']] != ''):

                sql = "SELECT top 1 * FROM analyse_day where (k_res=" + str(st[d_c['kod']]) + ") and (k_obkt=" + str(
                    st[d_c['kod_ob']]) + ") and (par_cod=" + str(st[d_c['nom']]) + ") order by dt desc"
                st_lab = cur_lab.execute(sql).fetchone()

                if st_lab != None:
                    if st_lab.PAR_ZN != 'None' and str(st_lab.PAR_ZN).strip('') != '' and float(
                            st_lab.PAR_ZN) != 0:

                        cursor.execute("UPDATE %s SET P_nom= %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], float(str(st_lab.PAR_ZN)) if in_numer(str(st_lab.PAR_ZN)) else st[d_c['P_nom']], st[d_c['Shifr']]))
                        cursor.execute("UPDATE %s SET ed_p= %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], int(st_lab.K_EDI) if str(st_lab.K_EDI).isnumeric() else 0,
                        st[d_c['Shifr']]))


        cur_lab.close()
        cnn_lab.close()

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error updating densities(fast) : ' + str(Er) + ': ' + datetime.today().strftime(
                "%d.%m.%Y %H:%M:%S") + '\n')

def rasch_proc(cursor, conf_dict, d_c):

    try:

        a_sum = [0]* 31

        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
        for st in cursor.fetchall():

            if st[d_c['Prizn']][:7] == "NEFT_IN" and st[d_c['Prizn']] not in ['NEFT_IN', 'NEFT_IN7']:
                for i in range(1,32):
                    sql = "SELECT M_N" + str(i) + "_" + "1V from %s WHERE Shifr='%s'" % (
                    conf_dict['Table_work'], st[d_c['Shifr']])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V from %s WHERE Shifr='%s'" % (
                    conf_dict['Table_work'], st[d_c['Shifr']])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V from %s WHERE Shifr='%s'" % (
                    conf_dict['Table_work'], st[d_c['Shifr']])
                    a_sum[i-1] += (cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] +
                                cursor.execute(sql2).fetchone()[0])

            cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
            for st in cursor.fetchall():
#                if "NEFT_POC" not in st[d_c['Prizn']]:
                for i in range(1, 32):
                    if a_sum[i - 1] != 0:
                        sql = "SELECT M_N" + str(i) + "_" + "1V from %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], st[d_c['Shifr']])
                        sql1 = "SELECT M_N" + str(i) + "_" + "2V from %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], st[d_c['Shifr']])
                        sql2 = "SELECT M_N" + str(i) + "_" + "3V from %s WHERE Shifr='%s'" % (
                        conf_dict['Table_work'], st[d_c['Shifr']])
                        s_tmp = cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                                cursor.execute(sql2).fetchone()[0]
                        cursor.execute(
                            "UPDATE %s SET P_" % (conf_dict['Table_work']) + str(i) + "= %s WHERE Shifr='%s'" % (
                            (s_tmp / a_sum[i - 1]) * 100, st[d_c['Shifr']]))

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Raschet procentov (Fast_update) : ' + str(Er) + ': ' + datetime.today().strftime(
                "%d.%m.%Y %H:%M:%S") + '\n')
    # ===========================================================================================================

def in_numer(num_str=''):
    try:
        float(num_str)
        return True
    except Exception:
        return False

def opros_shifra(curs, Shifr='', db='', zn=0, prizn='', conf_dict={}):
#def opros_shifra(curs, Shifr='', db='', zn=0):
    try:
        if Shifr.strip() and prizn != 'NEFT_OUT_8':

            sql = """SELECT tzn FROM %s WHERE Shifr='%s'""" % (db, Shifr)
            if curs.execute(sql).fetchone():
                return curs.execute(sql).fetchone()[0]
            else:
                return -1

        elif prizn == 'NEFT_OUT_8' and Shifr == 'FIRCA3018':

            try:
                str_cnn = 'DRIVER=%s;SERVER=%s;DATABASE=%s;UID=%s;PWD=%s' % (
                conf_dict['DRIVER'], conf_dict['SERVER_REM1'], conf_dict['DATABASE_REM1'], conf_dict['UID_REM1'],
                conf_dict['PWD_REM1'])
                cnn_rem = pyodbc.connect(str_cnn, readonly=True, timeout=1)
                curs = cnn_rem.cursor()
            except pyodbc.DatabaseError as Er:
                #    except Exception as Er:
                open(os.curdir + '\\report_error.txt', 'a').write(
                    'Error database connect for SERVER_REM1: ' + str(Er) + ': ' + datetime.today().strftime(
                        "%d.%m.%Y %H:%M:%S") + '\n')
                return -1

            sql = """SELECT tzn FROM %s WHERE Shifr='%s'""" % (conf_dict['Table_rem1'], Shifr)
            if curs.execute(sql).fetchone():

                tzn = curs.execute(sql).fetchone().tzn
                curs.close()
                cnn_rem.close()
                return tzn

            else:

                curs.close()
                cnn_rem.close()
                return -1

        else:
            return zn


    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error remov shifr: ' + Shifr + ' ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
        return zn

# =============================обработка каждые 15 сек====================================
def fastupdate(cursor, cur_rem, conf_dict, d_c):

    try:

        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))

        try:
            cnn_mydb = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
            cursor_p = cnn_mydb.cursor()
            potoc = cursor_p.execute('SELECT * FROM potoki').fetchone()

            d_potoc=dict()
            d_potoc = {rows[1]:rows[0] for rows in cnn_mydb.execute("pragma table_info(potoki)").fetchall()}

        except sqlite3.DatabaseError as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
		         'Error database connect: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')

        for st in cursor.fetchall():

            result=Gt=Gt1=Gt2=Gt3=Gt4=Gt5=T=T1=T2=P=P1=P2=0
            list_zn=[]
            for sh in str(conf_dict['shifrs']).split(','):
                Zn = 'Zn_' + sh.split('_')[1]
                Gt_temp = opros_shifra(cur_rem, st[d_c[sh]], conf_dict['Table_rem'], st[d_c[Zn]], st[d_c['Prizn']], conf_dict)
                if Gt_temp:
                    Gt = Gt_temp if Gt_temp < 0 or Gt_temp > 0.000001 else 0
                else:
                    Gt = 0
                cursor.execute("UPDATE %s SET " % (conf_dict['Table_work']) + Zn + "= %s WHERE Shifr='%s'" % (
                                Gt, st[d_c['Shifr']]))
                list_zn.append(Gt)

            if -1 in list_zn:
                result = -1
            elif -99999 not in list_zn:

                if st[d_c['Prizn']] not in ['NEFT_IN4', 'NEFT_IN2']:
                    result = list_zn[0]    
                elif st[d_c['Prizn']] in ['NEFT_OUT_8']:
                    result = list_zn[0]*2
                else:
                    result = 0

            else:
                result = -99999

            cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))

            if result >= 0 and st[d_c['Prizn']] != 'sep':

                if st[d_c['Prizn']] in ['NEFT_IN9']:
                    if potoc[d_potoc['uvg100_200']]:
                        result = result/2

                    elif potoc[d_potoc['uvg100']]:
                        pass

                    else:

                        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        continue

                    cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))
                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result * st[d_c['P_nom']] / 1000  - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)

                    cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['P_nom']], st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

                if st[d_c['Prizn']] in ['NEFT_IN8']:
                    if potoc[d_potoc['uvg100_200']]:
                        result = result/2

                    elif potoc[d_potoc['uvg200']]:
                        pass

                    else:

                        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        continue

                    cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))
                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result * st[d_c['P_nom']] / 1000  - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)

                    cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['P_nom']], st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

                if st[d_c['Prizn']] in ['NEFT_IN10']:
                    if potoc[d_potoc['vsg100_200']]:
                        result = result/2

                    elif potoc[d_potoc['vsg100']]:
                        pass

                    else:

                        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        continue

                    cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))
                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result * st[d_c['P_nom']] / 1000  - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)

                    cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['P_nom']], st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

                if st[d_c['Prizn']] in ['NEFT_IN11']:
                    if potoc[d_potoc['vsg100_200']]:
                        result = result/2

                    elif potoc[d_potoc['vsg200']]:
                        pass

                    else:

                        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                            conf_dict['Table_work'], 0, st[d_c['Shifr']]))
                        continue

                    cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))
                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result * st[d_c['P_nom']] / 1000  - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)

                    cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['P_nom']], st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))



                if st[d_c['Prizn']] in ['NEFT_IN1', 'NEFT_IN3', 'NEFT_OUT_3', 'NEFT_OUT_4', 'NEFT_OUT_5']:

                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result / 1000 - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

                elif st[d_c['Prizn']] in ['NEFT_IN4', 'NEFT_IN2', 'NEFT_IN7', 'NEFT_IN']:
                    pass

                else:

                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                    Sz_hour_m = st[d_c['Sz_hour_m']] + (result * st[d_c['P_nom']] / 1000  - st[d_c['Sz_hour_m']]) / (st[d_c['n']] + 1)
                    Density = st[d_c['P_nom']]

                    cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Sz_hour_m, st[d_c['Shifr']]))
                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Density, st[d_c['Shifr']]))

                    cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

        # ===================================Расчеты вне цикла========================================================================
                  # ===================================Потери на ХОВ========================================================================
        gt_in2 = 0
        gt_in4 = 0

        temp = 0
        temp = cursor.execute("""SELECT Zn_Gt FROM %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 'FIRCA1006')).fetchone()[0]
        gt_in2 = temp if temp>=0 else 0

        temp = 0
        temp = cursor.execute("""SELECT Zn_Gt FROM %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], 'FIRCA2006')).fetchone()[0]
        gt_in4 = temp if temp>=0 else 0

        m_in = 0
        m_out = 0
        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
        for st in cursor.fetchall():

            if st[d_c['Prizn']] in ['NEFT_OUT', 'NEFT_OUT_7']:
                m_out += st[d_c['Sz_hour_m']]

            if st[d_c['Prizn']] in ['NEFT_IN1', 'NEFT_IN9', 'NEFT_IN5', 'NEFT_IN10', 'NEFT_IN3', 'NEFT_IN8', 'NEFT_IN6', 'NEFT_IN11']:
                m_in += st[d_c['Sz_hour_m']]

        if (gt_in2 == 0) and (gt_in4 == 0):
            result_m_in2 = 0
            result_m_in4 = 0
        else:
            result_m_in2 = (gt_in2 / (gt_in2 + gt_in4)) * (m_out * 100.006 / 100 - m_in)
            result_m_in4 = (gt_in4 / (gt_in2 + gt_in4)) * (m_out * 100.006 / 100 - m_in)

        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result_m_in2, 'FIRCA1006'))
        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result_m_in2, 'FIRCA1006'))

        cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result_m_in4, 'FIRCA2006'))
        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result_m_in4, 'FIRCA2006'))

                  # ===================================Итого по секциям========================================================================
        m_c100 = 0
        m_c200 = 0
        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
        for st in cursor.fetchall():

            if st[d_c['Prizn']] in ['NEFT_IN1', 'NEFT_IN2', 'NEFT_IN9', 'NEFT_IN5', 'NEFT_IN10']:
                m_c100 += st[d_c['Sz_hour_m']]

            if st[d_c['Prizn']] in ['NEFT_IN3', 'NEFT_IN4', 'NEFT_IN8', 'NEFT_IN6', 'NEFT_IN11']:
                m_c200 += st[d_c['Sz_hour_m']]

        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Prizn='%s'""" % (conf_dict['Table_work'], m_c100, 'NEFT_IN7'))
        cursor.execute("""UPDATE %s SET Sz_hour_m = %s WHERE Prizn='%s'""" % (conf_dict['Table_work'], m_c200, 'NEFT_IN'))

        # ===========================================================================================================

        open(os.curdir + '\\test.txt', 'w').write(datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
        print('Last update ----- ', str(datetime.now()).split('.')[0])

        cursor_p.close()
#        cur_rem.close()
        cnn_mydb.close()
#        cnn_rem.close()

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write('Error obrabotki tekucshich znacheniy : ' + st[d_c['Prizn']] + str(Er) + ': '
            + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
    # ===========================================================================================================


if __name__ == '__main__':
    pass
