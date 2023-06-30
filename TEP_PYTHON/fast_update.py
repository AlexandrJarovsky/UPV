# coding: utf8

from datetime import datetime, timedelta
import time
import pyodbc
import os

def in_numer(num_str=''):
    try:
        float(num_str)
        return True
    except Exception:
        return False

def opros_shifra(curs, Shifr='', db='', zn=0):
    try:
        if Shifr.strip():

            sql = """SELECT tzn FROM %s WHERE Shifr='%s'""" % (db,Shifr)
            if curs.execute(sql).fetchone():
                return curs.execute(sql).fetchone()[0]
            else:
                return -1

        else:
            return zn
    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error remov shifr: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
        return zn

def lab_densiti(cursor, conf_dict,d_c):

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
                    st[d_c['a']]) + ") and (par_cod=" + str(st[d_c['nom']]) + ") order by dt desc"
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
            'Refresh lab analis (fast) : ' + str(Er) + ': ' + datetime.today().strftime(
                "%d.%m.%Y %H:%M:%S") + '\n')

def rasch_proc(cursor, conf_dict,d_c):

    try:

        a_sum = [0]* 31

        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))
        for st in cursor.fetchall():

            if "NEFT_IN" in st[d_c['Prizn']]:
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

# =============================обработка каждые 15 сек====================================
def fastupdate(cursor, cur_rem, conf_dict, d_c):
    try:

        cursor.execute('SELECT * FROM %s order by sortorder' % (conf_dict['Table_work']))

        for st in cursor.fetchall():

            result=Gt=0
            list_zn=[]
            for sh in str(conf_dict['shifrs']).split(','):
                Zn = 'Zn_' + sh.split('_')[1]
                Gt = opros_shifra(cur_rem, st[d_c[sh]], conf_dict['Table_rem'], st[d_c[Zn]])
                cursor.execute("UPDATE %s SET " % (conf_dict['Table_work']) + Zn + "= %s WHERE Shifr='%s'" % (
                    Gt, st[d_c['Shifr']]))
                list_zn.append(Gt)

            if -1 in list_zn:
                result = -1
            elif -99999 not in list_zn:
                if st[d_c['Prizn']] == "PAR1":
#                    result = list_zn[0]/1000+(list_zn[1]+list_zn[2])
                    result = list_zn[0]/1000+list_zn[1] #03.11.2022 по просьбе Бутковского
                else:
                    result = list_zn[0]
            else:
                result = -99999

            cursor.execute("""UPDATE %s SET Tzn = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], result, st[d_c['Shifr']]))

            if result >= 0:

                '''if st[d_c['Prizn']] == "PAR":

                    Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)

                else:

                    Sz_hour = st[d_c['Sz_hour']] + (result * (
                        st[d_c['P_nom']] / 1000 - (0.001828 - 0.00132 * st[d_c['P_nom']] / 1000) * (
                        list_zn[3] - 20))- st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)

                    Density = (st[d_c['P_nom']] /1000 - (0.001828 - 0.00132 * st[d_c['P_nom']]/1000) * (
                        list_zn[3] - 20)) * 1000

                    cursor.execute("""UPDATE %s SET Density = %s WHERE Shifr='%s'""" % (
                        conf_dict['Table_work'], Density, st[d_c['Shifr']]))'''

                Sz_hour = st[d_c['Sz_hour']] + (result - st[d_c['Sz_hour']]) / (st[d_c['n']] + 1)
                cursor.execute("""UPDATE %s SET Sz_hour = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], Sz_hour, st[d_c['Shifr']]))
                cursor.execute("""UPDATE %s SET n = %s WHERE Shifr='%s'""" % (conf_dict['Table_work'], st[d_c['n']] + 1, st[d_c['Shifr']]))

        print('Last update ----- ', str(datetime.now()).split('.')[0])

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write('Error obrabotki tekucshich znacheniy : ' + str(Er) + ': '
            + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
    # ===========================================================================================================


if __name__ == '__main__':
    pass
