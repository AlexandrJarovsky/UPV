# coding: utf8

from datetime import datetime, timedelta
import pyodbc
import sqlite3
import os, sys, glob
import print_exel
import win32com.client
import win32print


def prin_pr(TableName_prn='', day_prn='', mount_prn='', yaer_prn='', flag='', flag_prn=1):
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
            d_c = {rows[1]:rows[0] for rows in cnn_mydb.execute("pragma table_info(Table1)").fetchall()}

        except pyodbc.DatabaseError as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error database connect(print) : ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
            return

        Excel1 = win32com.client.DispatchEx("Excel.Application")

        if not glob.glob('obolochka.xls*'):
            con_in = cursor.execute("SELECT count(*) FROM {}".format(TableName)).fetchone()[0]
#            con_out = cursor.execute("SELECT count(*) FROM {} WHERE Prizn LIKE '%_OUT%'".format(TableName)).fetchone()[0]
            print_exel.pr_exel(con_in, conf_dict['excel_format'], conf_dict['titul_list'])

        tip_file = glob.glob('obolochka.xls*')[0].split('.')[1]
        f_obol = os.path.join(os.curdir, glob.glob('obolochka.xls*')[0])

        wb1 = Excel1.Workbooks.Open(os.path.abspath(f_obol))
        sheet1 = wb1.Worksheets(1)

        if not day_prn:
            TempDate = (datetime.now() + timedelta(days=-1)).day
        else:
            TempDate = day_prn

        SQLStr = "SELECT Ima, Shifr, Ed, Prizn, "
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
        SQLStr = SQLStr + "M_N31_1V + M_N31_2V + M_N31_3V AS Total_month_mass "
        SQLStr = SQLStr + "FROM " + TableName + " order by SortOrder"

        Mount_name = {1: " января ", 2: " февраля ", 3: " марта ", 4: " апреля ", 5: " мая ", 6: " июня ", \
                      7: " июля ", 8: " августа ", 9: " сентября ", 10: " октября ", 11: " ноября ", 12: " декабря "}

        Ed_izm = {1: "м3/ч", 2: "%", 3: "см", 4: "нм3/ч", 5: "т/ч", 6: "ГКАЛ", 10: "кг/ч"}

        if not flag:
 
           if not day_prn:
                RepDat = str(TempDate) + str(Mount_name[(datetime.now() + timedelta(days=-1)).month]) + str(
     			    (datetime.now() + timedelta(days=-1)).year) + "г."
           else:
                RepDat = str(TempDate) + str(Mount_name[datetime.now().month]) + str(datetime.now().year) + "г."

        else:
            RepDat = str(TempDate) + str(Mount_name[int(TableName.split('_')[2])]) + str(TableName.split('_')[1]) + "г."

        q = 0

        for rs in cur_count.execute(SQLStr):

            s_month = 0

            for i in range(1, int(TempDate) + 1):
                sql = "SELECT Sz_N" + str(i) + "_1V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                sql1 = "SELECT Sz_N" + str(i) + "_2V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                sql2 = "SELECT Sz_N" + str(i) + "_3V FROM %s WHERE Shifr='%s'" % (TableName, rs[1])
                s_month += cursor.execute(sql).fetchone()[0] + cursor.execute(sql1).fetchone()[0] + \
                        cursor.execute(sql2).fetchone()[0]

            sheet1.Range("B" + str(q + 5)).value = rs[0]

            if rs[3] != 'sep':

                sheet1.Range("A" + str(q + 5)).value = rs[1]
                sheet1.Range("C" + str(q + 5)).value = Ed_izm[int(rs[2])] 
                sheet1.Range("D" + str(q + 5)).value = rs[7]/24
                sheet1.Range("E" + str(q + 5)).value = rs[7]
                sheet1.Range("F" + str(q + 5)).value = s_month
                if rs[1] in ["FIC2710", "FI1710"]:
                    sheet1.Range("G" + str(q + 5)).value = s_month * 0.752

            q = q + 1

        sheet1.Range("A" + str(q + 8)).value = RepDat

        sheet1.Range("D" + str(q + 7)).value = "ОАО МНПЗ. НХП. УПВ"
        sheet1.Range("D" + str(q + 8)).value = "Начальник установки ___________________"

        if glob.glob(os.curdir + '\\PRINTER\\' + RepDat + tip_file):
            os.remove(os.path.join(os.curdir, glob.glob(os.curdir + '\\PRINTER\\' + RepDat + tip_file)[0]))

        wb1.SaveAs(os.path.abspath(os.curdir) + '\\PRINTER\\' + RepDat + tip_file)

        if flag_prn:

            sheet1.PageSetup.Zoom = 100
            for i in range(int(conf_dict['count_print'])):
                sheet1.PrintOut()

            open(os.curdir + '\\report_error.txt', 'a').write(
                datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Otchet printed\n')
#            win32print.SetDefaultPrinter(printDef)

        wb1.Close()
        Excel1.Quit()

        cursor.close()
        cur_count.close()
        cnn_mydb.close()

        open(os.curdir + '\\report_error.txt', 'a').write(
            datetime.today().strftime("%d.%m.%Y %H:%M:%S") + ': Otchet sozdan\n')

    except Exception as Er:
        open(os.curdir + '\\report_error.txt', 'a').write(
            'Error pri pechati otcheta : ' + str(Er) + ': ' + datetime.today().strftime(
                "%d.%m.%Y %H:%M:%S") + '\n')
        wb1.Close(SaveChanges=0)
        Excel1.Quit()


if __name__ == '__main__':
    prin_pr()
