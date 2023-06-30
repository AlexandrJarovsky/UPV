# coding: utf8

from datetime import datetime, timedelta
import time
import sqlite3
import os, sys, glob
import print_exel
import win32com.client
import pythoncom


def prin_pr(TableName_prn='', day_prn='', mount_prn='', yaer_prn='', flag='', flag_prn=True, flag_prn_hour=True):
    try:

        with open(os.curdir + '\\config.ini') as conf:
            conf_list = []
            for c in conf:
                if c != '\n':
                    conf_list.append(c.strip().split('='))
            conf_dict = dict(conf_list)

        if not flag:
            TableName = conf_dict['Table_work']
        else:
            TableName = TableName_prn

        try:

            cnn_mydb = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
            cursor = cnn_mydb.cursor()
            cnn_count = sqlite3.connect(os.curdir + conf_dict['DATABASE'])
            cur_count = cnn_count.cursor()

            d_c=dict()
            i=0
            for rows in cnn_mydb.execute("pragma table_info(Table1)").fetchall():
                d_c[rows[1]]=i
                i+=1
        except sqlite3.DatabaseError as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error database connect(print) : ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
            return

        pythoncom.CoInitialize()
        Excel1 = win32com.client.DispatchEx("Excel.Application")

        if not glob.glob('obolochka.xls*'):
            con_in = cursor.execute("SELECT count(*) FROM {} WHERE Prizn LIKE '%_IN%'".format(TableName)).fetchone()[0]
            con_out = cursor.execute("SELECT count(*) FROM {} WHERE Prizn LIKE '%_OUT%'".format(TableName)).fetchone()[0]
            con_urov = cursor.execute("SELECT count(*) FROM {} WHERE Prizn LIKE '%POC%'".format(TableName)).fetchone()[0]
            print_exel.pr_exel(con_in, con_out+1, con_urov, conf_dict['excel_format'], conf_dict['titul_list'])

        tip_file = glob.glob('obolochka.xls*')[0].split('.')[1]
        f_obol = os.path.join(os.curdir, glob.glob('obolochka.xls*')[0])

        wb1 = Excel1.Workbooks.Open(os.path.abspath(f_obol))
        sheet1 = wb1.Worksheets(1)

        if not day_prn:
            TempDate = (datetime.now() + timedelta(days=-1)).day
        else:
            TempDate = day_prn

        SQLStr = "SELECT Ima, Shifr, Ed, Prizn, P_nom, "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_1V as v1_V, "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_2V as v2_V, "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_3V as v3_V, "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_1V + "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_2V + "
        SQLStr = SQLStr + "Sz_N" + str(TempDate) + "_3V as day_V, "
        SQLStr = SQLStr + "Sz_N1_1V + Sz_N1_2V + Sz_N1_3V + "
        SQLStr = SQLStr + "Sz_N2_1V + Sz_N2_2V + Sz_N2_3V + "
        SQLStr = SQLStr + "Sz_N3_1V + Sz_N3_2V + Sz_N3_3V + "
        SQLStr = SQLStr + "Sz_N4_1V + Sz_N4_2V + Sz_N4_3V + "
        SQLStr = SQLStr + "Sz_N5_1V + Sz_N5_2V + Sz_N5_3V + "
        SQLStr = SQLStr + "Sz_N6_1V + Sz_N6_2V + Sz_N6_3V + "
        SQLStr = SQLStr + "Sz_N7_1V + Sz_N7_2V + Sz_N7_3V + "
        SQLStr = SQLStr + "Sz_N8_1V + Sz_N8_2V + Sz_N8_3V + "
        SQLStr = SQLStr + "Sz_N9_1V + Sz_N9_2V + Sz_N9_3V + "
        SQLStr = SQLStr + "Sz_N10_1V + Sz_N10_2V + Sz_N10_3V + "
        SQLStr = SQLStr + "Sz_N11_1V + Sz_N11_2V + Sz_N11_3V + "
        SQLStr = SQLStr + "Sz_N12_1V + Sz_N12_2V + Sz_N12_3V + "
        SQLStr = SQLStr + "Sz_N13_1V + Sz_N13_2V + Sz_N13_3V + "
        SQLStr = SQLStr + "Sz_N14_1V + Sz_N14_2V + Sz_N14_3V + "
        SQLStr = SQLStr + "Sz_N15_1V + Sz_N15_2V + Sz_N15_3V + "
        SQLStr = SQLStr + "Sz_N16_1V + Sz_N16_2V + Sz_N16_3V + "
        SQLStr = SQLStr + "Sz_N17_1V + Sz_N17_2V + Sz_N17_3V + "
        SQLStr = SQLStr + "Sz_N18_1V + Sz_N18_2V + Sz_N18_3V + "
        SQLStr = SQLStr + "Sz_N19_1V + Sz_N19_2V + Sz_N19_3V + "
        SQLStr = SQLStr + "Sz_N20_1V + Sz_N20_2V + Sz_N20_3V + "
        SQLStr = SQLStr + "Sz_N21_1V + Sz_N21_2V + Sz_N21_3V + "
        SQLStr = SQLStr + "Sz_N22_1V + Sz_N22_2V + Sz_N22_3V + "
        SQLStr = SQLStr + "Sz_N23_1V + Sz_N23_2V + Sz_N23_3V + "
        SQLStr = SQLStr + "Sz_N24_1V + Sz_N24_2V + Sz_N24_3V + "
        SQLStr = SQLStr + "Sz_N25_1V + Sz_N25_2V + Sz_N25_3V + "
        SQLStr = SQLStr + "Sz_N26_1V + Sz_N26_2V + Sz_N26_3V + "
        SQLStr = SQLStr + "Sz_N27_1V + Sz_N27_2V + Sz_N27_3V + "
        SQLStr = SQLStr + "Sz_N28_1V + Sz_N28_2V + Sz_N28_3V + "
        SQLStr = SQLStr + "Sz_N29_1V + Sz_N29_2V + Sz_N29_3V + "
        SQLStr = SQLStr + "Sz_N30_1V + Sz_N30_2V + Sz_N30_3V + "
        SQLStr = SQLStr + "Sz_N31_1V + Sz_N31_2V + Sz_N31_3V AS Total_month, "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_1V as v1_M, "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_2V as v2_M, "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_3V as v3_M, "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_1V + "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_2V + "
        SQLStr = SQLStr + "M_N" + str(TempDate) + "_3V as day_M, "
        SQLStr = SQLStr + "M_N1_1V + M_N1_2V + M_N1_3V + "
        SQLStr = SQLStr + "M_N2_1V + M_N2_2V + M_N2_3V + "
        SQLStr = SQLStr + "M_N3_1V + M_N3_2V + M_N3_3V + "
        SQLStr = SQLStr + "M_N4_1V + M_N4_2V + M_N4_3V + "
        SQLStr = SQLStr + "M_N5_1V + M_N5_2V + M_N5_3V + "
        SQLStr = SQLStr + "M_N6_1V + M_N6_2V + M_N6_3V + "
        SQLStr = SQLStr + "M_N7_1V + M_N7_2V + M_N7_3V + "
        SQLStr = SQLStr + "M_N8_1V + M_N8_2V + M_N8_3V + "
        SQLStr = SQLStr + "M_N9_1V + M_N9_2V + M_N9_3V + "
        SQLStr = SQLStr + "M_N10_1V + M_N10_2V + M_N10_3V + "
        SQLStr = SQLStr + "M_N11_1V + M_N11_2V + M_N11_3V + "
        SQLStr = SQLStr + "M_N12_1V + M_N12_2V + M_N12_3V + "
        SQLStr = SQLStr + "M_N13_1V + M_N13_2V + M_N13_3V + "
        SQLStr = SQLStr + "M_N14_1V + M_N14_2V + M_N14_3V + "
        SQLStr = SQLStr + "M_N15_1V + M_N15_2V + M_N15_3V + "
        SQLStr = SQLStr + "M_N16_1V + M_N16_2V + M_N16_3V + "
        SQLStr = SQLStr + "M_N17_1V + M_N17_2V + M_N17_3V + "
        SQLStr = SQLStr + "M_N18_1V + M_N18_2V + M_N18_3V + "
        SQLStr = SQLStr + "M_N19_1V + M_N19_2V + M_N19_3V + "
        SQLStr = SQLStr + "M_N20_1V + M_N20_2V + M_N20_3V + "
        SQLStr = SQLStr + "M_N21_1V + M_N21_2V + M_N21_3V + "
        SQLStr = SQLStr + "M_N22_1V + M_N22_2V + M_N22_3V + "
        SQLStr = SQLStr + "M_N23_1V + M_N23_2V + M_N23_3V + "
        SQLStr = SQLStr + "M_N24_1V + M_N24_2V + M_N24_3V + "
        SQLStr = SQLStr + "M_N25_1V + M_N25_2V + M_N25_3V + "
        SQLStr = SQLStr + "M_N26_1V + M_N26_2V + M_N26_3V + "
        SQLStr = SQLStr + "M_N27_1V + M_N27_2V + M_N27_3V + "
        SQLStr = SQLStr + "M_N28_1V + M_N28_2V + M_N28_3V + "
        SQLStr = SQLStr + "M_N29_1V + M_N29_2V + M_N29_3V + "
        SQLStr = SQLStr + "M_N30_1V + M_N30_2V + M_N30_3V + "
        SQLStr = SQLStr + "M_N31_1V + M_N31_2V + M_N31_3V AS Total_month_mass, "
        SQLStr = SQLStr + "P_" + str(TempDate) + " as P_D "
        SQLStr = SQLStr + "FROM " + TableName + " order by SortOrder"

        s_s1v = 0
        s_s2v = 0
        s_s3v = 0
        s_s1m = 0
        s_s2m = 0
        s_s3m = 0
        s_sDv = 0
        s_sDm = 0
        s_sMv = 0
        s_sMm = 0
        s_pD = 0
        s_pM = 0

        p_s1v = 0
        p_s2v = 0
        p_s3v = 0
        p_s1m = 0
        p_s2m = 0
        p_s3m = 0
        p_sDv = 0
        p_sDm = 0
        p_sMv = 0
        p_sMm = 0
        p_pD = 0
        p_pM = 0

        for rs in cur_count.execute(SQLStr):

            if rs[3][:7] == "NEFT_IN" and rs[3] not in ['NEFT_IN', 'NEFT_IN7']:
                s_s1v += rs[5]
                s_s2v += rs[6]
                s_s3v += rs[7]
                s_s1m += rs[10]
                s_s2m += rs[11]
                s_s3m += rs[12]
                s_sDv += rs[8]
                s_sDm += rs[13]
                s_pD += rs[15]

                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT Sz_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT Sz_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT Sz_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_sMv += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                             cursor.execute(sql2).fetchone()[0]

                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_sMm += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                             cursor.execute(sql2).fetchone()[0]

            if rs[3] in ["NEFT_OUT", "NEFT_OUT_7"]:
                p_s1v += rs[5]
                p_s2v += rs[6]
                p_s3v += rs[7]
                p_s1m += rs[10]
                p_s2m += rs[11]
                p_s3m += rs[12]
                p_sDv += rs[8]
                p_sDm += rs[13]
                p_pD += rs[15]

                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT Sz_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT Sz_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT Sz_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    p_sMv += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                             cursor.execute(sql2).fetchone()[0]

                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    p_sMm += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                             cursor.execute(sql2).fetchone()[0]

        Mount_name = {1: " января ", 2: " февраля ", 3: " марта ", 4: " апреля ", 5: " мая ", 6: " июня ", \
                      7: " июля ", 8: " августа ", 9: " сентября ", 10: " октября ", 11: " ноября ", 12: " декабря "}

        if not flag:
 
           if not day_prn:
                RepDat = str(TempDate) + str(Mount_name[(datetime.now() + timedelta(days=-1)).month]) + str(
     			    (datetime.now() + timedelta(days=-1)).year) + "г."
           else:
                RepDat = str(TempDate) + str(Mount_name[datetime.now().month]) + str(datetime.now().year) + "г."

        else:
            RepDat = str(TempDate) + str(Mount_name[int(TableName.split('_')[2])]) + str(TableName.split('_')[1]) + "г."

        s_month_all = 0
        cursor.execute('SELECT * FROM %s order by sortorder' % (TableName))
        for st in cursor.fetchall():

            if st[d_c['Prizn']][:7] == "NEFT_IN" and st[d_c['Prizn']] not in ['NEFT_IN', 'NEFT_IN7']:
                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, st[d_c['Shifr']])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, st[d_c['Shifr']])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, st[d_c['Shifr']])
                    s_month_all += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                        cursor.execute(sql2).fetchone()[0]

        s_pM = 0
        q = 0

        for rs in cur_count.execute(SQLStr):
            if rs[3][:7] == "NEFT_IN":

                s_month = 0
                s_month_m = 0
                p_m = 0

                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT Sz_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT Sz_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT Sz_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                               cursor.execute(sql2).fetchone()[0]

                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month_m += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                                 cursor.execute(sql2).fetchone()[0]

                if s_month_all != 0:
                    p_m = (s_month_m / s_month_all) * 100

                sheet1.Range("A" + str(q + 6)).value = rs[0]
                sheet1.Range("B" + str(q + 6)).value = rs[4]
                sheet1.Range("C" + str(q + 6)).value = rs[5]
                sheet1.Range("D" + str(q + 6)).value = rs[10]
                sheet1.Range("E" + str(q + 6)).value = rs[6]
                sheet1.Range("F" + str(q + 6)).value = rs[11]
                sheet1.Range("G" + str(q + 6)).value = rs[7]
                sheet1.Range("H" + str(q + 6)).value = rs[12]
                sheet1.Range("I" + str(q + 6)).value = rs[8]
                sheet1.Range("J" + str(q + 6)).value = rs[13]
                sheet1.Range("K" + str(q + 6)).value = rs[15]
                sheet1.Range("L" + str(q + 6)).value = s_month
                sheet1.Range("M" + str(q + 6)).value = s_month_m
                sheet1.Range("N" + str(q + 6)).value = p_m
                q = q + 1

        if s_month_all != 0:
            s_pM = (s_sMm / s_month_all) * 100

        sheet1.Range("A" + str(q + 6)).value = "ВСЕГО"
        sheet1.Range("C" + str(q + 6)).value = s_s1v
        sheet1.Range("D" + str(q + 6)).value = s_s1m
        sheet1.Range("E" + str(q + 6)).value = s_s2v
        sheet1.Range("F" + str(q + 6)).value = s_s2m
        sheet1.Range("G" + str(q + 6)).value = s_s3v
        sheet1.Range("H" + str(q + 6)).value = s_s3m
        sheet1.Range("I" + str(q + 6)).value = s_sDv
        sheet1.Range("J" + str(q + 6)).value = s_sDm
        sheet1.Range("K" + str(q + 6)).value = s_pD
        sheet1.Range("L" + str(q + 6)).value = s_sMv
        sheet1.Range("M" + str(q + 6)).value = s_sMm
        sheet1.Range("N" + str(q + 6)).value = s_pM

        p_pM = 0

        for rs in cur_count.execute(SQLStr):
            if rs[3][:8] == "NEFT_OUT":

                s_month = 0
                s_month_m = 0
                p_m = 0

                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT Sz_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT Sz_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT Sz_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                               cursor.execute(sql2).fetchone()[0]

                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month_m += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                                 cursor.execute(sql2).fetchone()[0]

                if s_month_all != 0:
                    p_m = (s_month_m / s_month_all) * 100

                if rs[3] =="NEFT_OUT_1":
                    sheet1.Range("A" + str(q + 8)).value = "Из них"
                    q += 1
                sheet1.Range("A" + str(q + 8)).value = rs[0]    
                sheet1.Range("B" + str(q + 8)).value = rs[4]    
                sheet1.Range("C" + str(q + 8)).value = rs[5]    
                sheet1.Range("D" + str(q + 8)).value = rs[10]   
                sheet1.Range("E" + str(q + 8)).value = rs[6]    
                sheet1.Range("F" + str(q + 8)).value = rs[11]   
                sheet1.Range("G" + str(q + 8)).value = rs[7]    
                sheet1.Range("H" + str(q + 8)).value = rs[12]   
                sheet1.Range("I" + str(q + 8)).value = rs[8]    
                sheet1.Range("J" + str(q + 8)).value = rs[13]   
                sheet1.Range("K" + str(q + 8)).value = rs[15]   
                sheet1.Range("L" + str(q + 8)).value = s_month  
                sheet1.Range("M" + str(q + 8)).value = s_month_m
                sheet1.Range("N" + str(q + 8)).value = p_m                  
                q = q + 1                                                    
                                                                             
        if s_month_all != 0:                                                
            p_pM = (p_sMm / s_month_all) * 100                               
                                                                            
        sheet1.Range("A" + str(q + 8)).value = "ИТОГО"                       
        sheet1.Range("C" + str(q + 8)).value = p_s1v                        
        sheet1.Range("D" + str(q + 8)).value = p_s1m                         
        sheet1.Range("E" + str(q + 8)).value = p_s2v                        
        sheet1.Range("F" + str(q + 8)).value = p_s2m                        
        sheet1.Range("G" + str(q + 8)).value = p_s3v                       
        sheet1.Range("H" + str(q + 8)).value = p_s3m                     
        sheet1.Range("I" + str(q + 8)).value = p_sDv                           
        sheet1.Range("J" + str(q + 8)).value = p_sDm
        sheet1.Range("K" + str(q + 8)).value = p_pD
        sheet1.Range("L" + str(q + 8)).value = p_sMv
        sheet1.Range("M" + str(q + 8)).value = p_sMm
        sheet1.Range("N" + str(q + 8)).value = p_pM

        sheet1.Range("A" + str(q + 9)).value = "ПОТЕРИ"
        sheet1.Range("D" + str(q + 9)).value = s_s1m - p_s1m
        sheet1.Range("F" + str(q + 9)).value = s_s2m - p_s2m
        sheet1.Range("H" + str(q + 9)).value = s_s3m - p_s3m
        sheet1.Range("J" + str(q + 9)).value = s_sDm - p_sDm
        sheet1.Range("K" + str(q + 9)).value = s_pD - p_pD
        sheet1.Range("M" + str(q + 9)).value = s_sMm - p_sMm
        sheet1.Range("N" + str(q + 9)).value = s_pM - p_pM

        q += 1
        for rs in cur_count.execute(SQLStr):
            if rs[3] == "POC":

                s_month = 0
                s_month_m = 0
                p_m = 0

                for i in range(1, int(TempDate) + 1):
                    sql = "SELECT Sz_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT Sz_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT Sz_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                               cursor.execute(sql2).fetchone()[0]

                    sql = "SELECT M_N" + str(i) + "_" + "1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql1 = "SELECT M_N" + str(i) + "_" + "2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    sql2 = "SELECT M_N" + str(i) + "_" + "3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                    s_month_m += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                                 cursor.execute(sql2).fetchone()[0]

                if s_month_all != 0:
                    p_m = (s_month_m / s_month_all) * 100

                sheet1.Range("A" + str(q + 10)).value = rs[0]    
                sheet1.Range("B" + str(q + 10)).value = rs[4]    
                sheet1.Range("C" + str(q + 10)).value = rs[5]    
                sheet1.Range("D" + str(q + 10)).value = rs[10]   
                sheet1.Range("E" + str(q + 10)).value = rs[6]    
                sheet1.Range("F" + str(q + 10)).value = rs[11]   
                sheet1.Range("G" + str(q + 10)).value = rs[7]    
                sheet1.Range("H" + str(q + 10)).value = rs[12]   
                sheet1.Range("I" + str(q + 10)).value = rs[8]    
                sheet1.Range("J" + str(q + 10)).value = rs[13]   
                sheet1.Range("K" + str(q + 10)).value = rs[15]   
                sheet1.Range("L" + str(q + 10)).value = s_month  
                sheet1.Range("M" + str(q + 10)).value = s_month_m
                sheet1.Range("N" + str(q + 10)).value = p_m                  
                q += 1    

        sheet1.Range("A" + str(q + 12)).value = RepDat

        sheet1.Range("K" + str(q + 11)).value = "ОАО МНПЗ. НХП. УПВ"
        sheet1.Range("K" + str(q + 12)).value = "Начальник установки _____________"

        if flag_prn:
            sheet1.PageSetup.Zoom = 100
            for i in range(int(conf_dict['count_print'])):
                sheet1.PrintOut()

        if glob.glob(os.curdir + '\\PRINTER\\' + RepDat + tip_file):
            os.remove(os.path.join(os.curdir, glob.glob(os.curdir + '\\PRINTER\\' + RepDat + tip_file)[0]))

        wb1.SaveAs(os.path.abspath(os.curdir) + '\\PRINTER\\' + RepDat + tip_file)

        wb1.Close()
        Excel1.Quit()

        cursor.close()
        cur_count.close()
        cnn_mydb.close()

        open(os.curdir + '\\report_error.txt', 'a').write(
            datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Otchet sozdan\n')

        pythoncom.CoUninitialize()

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error pri pechati otcheta : ' + str(Er) + ': ' + datetime.today().strftime(
                "%d.%m.%Y %H:%M:%S") + '\n')
        pythoncom.CoUninitialize()
        wb1.Close(SaveChanges=0)
        Excel1.Quit()


if __name__ == '__main__':
    prin_pr()
