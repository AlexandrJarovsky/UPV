# coding: utf8

from datetime import datetime, timedelta
import win32com.client
import os, glob


def pr_exel(col_in=0, format_data='0,00', titul_list=''):
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12
    xlContinuous = 1
    xlGeneral = 1
    xlThin = 2
    xlAutomatic = -4105
    xlBottom = -4107
    xlNone = -4142
    xlLeft = -4131
    xlRight = -4152
    xlCenter = -4108
    xlUnderlineStyleNone = -4142

    if True:
        try:
            Excel1 = win32com.client.DispatchEx("Excel.Application")

            if glob.glob(os.curdir + '\\obolochka.xls*'):
                os.remove(os.path.join(os.curdir, glob.glob('obolochka.xls*')[0]))
            wb1 = Excel1.Workbooks.Add()

            sheet1 = wb1.Worksheets(1)

            # ==================формирование оболочки==========================================================

            # ==================НЕИЗМЕНЯЕМАЯ ЧАСТЬ==========================================================
            sheet1.Range("A1:G4").HorizontalAlignment = xlCenter
            sheet1.Range("A1:G4").VerticalAlignment = xlCenter
            sheet1.Range("A1:G4").Font.Name = "Arial CYR"
            sheet1.Range("A1:G4").Font.FontStyle = "полужирный"
            sheet1.Range("A3:G4").ShrinkToFit = True
            sheet1.Range("A3:G4").WrapText = True

            sheet1.Range("A1:G2").MergeCells = True
            sheet1.Range("A1:G2").Font.Size = 14
            sheet1.Range("A1:G2").value = titul_list

            sheet1.Range("A3:G4").Font.Size = 10

            sheet1.Range("A3:A4").MergeCells = True
            sheet1.Range("A3:A4").FormulaR1C1 = "№ позиции"
            sheet1.Columns("A").ColumnWidth = 10

            sheet1.Range("B3:B4").MergeCells = True
            sheet1.Range("B3:B4").FormulaR1C1 = "Наим-ние продукта"
            sheet1.Columns("B").ColumnWidth = 37

            sheet1.Range("C3:C4").MergeCells = True
            sheet1.Range("C3:C4").FormulaR1C1 = "Ед. изм"
            sheet1.Columns("C").ColumnWidth = 5

            sheet1.Range("D3:D4").MergeCells = True
            sheet1.Range("D3:D4").FormulaR1C1 = "Среднечас. расход"
            sheet1.Columns("D").ColumnWidth = 13

            sheet1.Range("E3:E4").MergeCells = True
            sheet1.Range("E3:E4").FormulaR1C1 = "Суточный расход"
            sheet1.Columns("E").ColumnWidth = 13

            sheet1.Range("F3:F4").MergeCells = True
            sheet1.Range("F3:F4").FormulaR1C1 = "Расход за месяц"
            sheet1.Columns("F").ColumnWidth = 13

            sheet1.Range("G3:G4").MergeCells = True
            sheet1.Range("G3:G4").FormulaR1C1 = "Гкал"
            sheet1.Columns("G").ColumnWidth = 13

            sheet1.Range("A3:G4").Borders(xlEdgeLeft).Weight = xlThin
            sheet1.Range("A3:G4").Borders(xlEdgeTop).Weight = xlThin
            sheet1.Range("A3:G4").Borders(xlEdgeBottom).Weight = xlThin
            sheet1.Range("A3:G4").Borders(xlEdgeRight).Weight = xlThin
            sheet1.Range("A3:G4").Borders(xlInsideVertical).Weight = xlThin
            sheet1.Range("A3:G4").Borders(xlInsideHorizontal).Weight = xlThin

            # ============================================================================


            sheet1.Range("B5:B" + str(col_in + 4)).Font.Bold = True

            sheet1.Range("A3:G" + str(col_in + 4)).Borders(xlEdgeLeft).Weight = xlThin
            sheet1.Range("A3:G" + str(col_in + 4)).Borders(xlEdgeTop).Weight = xlThin
            sheet1.Range("A3:G" + str(col_in + 4)).Borders(xlEdgeBottom).Weight = xlThin
            sheet1.Range("A3:G" + str(col_in + 4)).Borders(xlEdgeRight).Weight = xlThin
            sheet1.Range("A3:G" + str(col_in + 4)).Borders(xlInsideVertical).Weight = xlThin

            sheet1.Range("A5:B" + str(col_in + 4)).HorizontalAlignment = xlLeft
            sheet1.Range("A5:B" + str(col_in + 4)).VerticalAlignment = xlBottom
            sheet1.Range("A5:B" + str(col_in + 4)).IndentLevel = 0
            sheet1.Range("A5:B" + str(col_in + 4)).ShrinkToFit = True

            sheet1.Range("C5:G" + str(col_in + 4)).HorizontalAlignment = xlRight
            sheet1.Range("C5:G" + str(col_in + 4)).VerticalAlignment = xlBottom
            sheet1.Range("C5:G" + str(col_in + 4)).ShrinkToFit = True

            sheet1.Range("D5:G" + str(col_in + 4)).NumberFormat = format_data
            sheet1.Range("A5:C" + str(col_in + 4)).NumberFormat = "@"
            sheet1.Rows("5:" + str(col_in + 4)).RowHeight = 16.5

            sheet1.PageSetup.LeftMargin = 30
            sheet1.PageSetup.RightMargin = 1
            sheet1.PageSetup.BottomMargin = 1.5
            sheet1.PageSetup.TopMargin = 1.5
            sheet1.PageSetup.Orientation = 2

            # ============================================================================

            wb1.SaveAs(os.path.abspath(os.curdir) + '\\obolochka')

            wb1.Close()
            Excel1.Quit()

        except Exception as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Error pechati obolochki: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
            wb1.Close(SaveChanges=0)
            Excel1.Quit()


if __name__ == '__main__':
    pr_exel()
