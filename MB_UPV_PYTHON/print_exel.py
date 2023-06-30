# coding: utf8

from datetime import datetime, timedelta
import win32com.client
import os, glob


def pr_exel(col_in=2, col_out=5, col_urov=-1, format_data='0,00', titul_list=''):
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
            sheet1.Range("A1:N2").HorizontalAlignment = xlCenter
            sheet1.Range("A1:N2").VerticalAlignment = xlCenter
            sheet1.Range("A1:N2").MergeCells = True
            sheet1.Range("A1:N2").Font.Name = "Arial CYR"
            sheet1.Range("A1:N2").Font.FontStyle = "полужирный"
            sheet1.Range("A1:N2").Font.Size = 14

            sheet1.Range("A3:A4").HorizontalAlignment = xlCenter
            sheet1.Range("A3:A4").VerticalAlignment = xlCenter
            sheet1.Range("A3:A4").ShrinkToFit = True
            sheet1.Range("A3:A4").MergeCells = True

            sheet1.Range("B3:B4").HorizontalAlignment = xlCenter
            sheet1.Range("B3:B4").VerticalAlignment = xlCenter
            sheet1.Range("B3:B4").WrapText = True
            sheet1.Range("B3:B4").MergeCells = True

            sheet1.Range("C3:D3").HorizontalAlignment = xlCenter
            sheet1.Range("C3:D3").VerticalAlignment = xlBottom
            sheet1.Range("C3:D3").MergeCells = True

            sheet1.Range("E3:F3").HorizontalAlignment = xlCenter
            sheet1.Range("E3:F3").VerticalAlignment = xlBottom
            sheet1.Range("E3:F3").MergeCells = True

            sheet1.Range("G3:H3").HorizontalAlignment = xlCenter
            sheet1.Range("G3:H3").VerticalAlignment = xlBottom
            sheet1.Range("G3:H3").MergeCells = True

            sheet1.Range("I3:K3").HorizontalAlignment = xlCenter
            sheet1.Range("I3:K3").VerticalAlignment = xlBottom
            sheet1.Range("I3:K3").MergeCells = True

            sheet1.Range("L3:N3").HorizontalAlignment = xlCenter
            sheet1.Range("L3:N3").VerticalAlignment = xlBottom
            sheet1.Range("L3:N3").MergeCells = True

            sheet1.Range("A3:N4").Borders(xlEdgeLeft).Weight = xlThin
            sheet1.Range("A3:N4").Borders(xlEdgeTop).Weight = xlThin
            sheet1.Range("A3:N4").Borders(xlEdgeBottom).Weight = xlThin
            sheet1.Range("A3:N4").Borders(xlEdgeRight).Weight = xlThin
            sheet1.Range("A3:N4").Borders(xlInsideVertical).Weight = xlThin
            sheet1.Range("A3:N4").Borders(xlInsideHorizontal).Weight = xlThin

            sheet1.Range("A3:N4").Font.Name = "Arial CYR"
            sheet1.Range("A3:N4").Font.FontStyle = "полужирный"
            sheet1.Range("A3:N4").Font.Size = 10

            sheet1.Range("A1:N2").FormulaR1C1 = titul_list
            sheet1.Range("A3:A4").FormulaR1C1 = "Наим-ние сырья и продуктов"
            sheet1.Columns("A:A").ColumnWidth = 26

            sheet1.Range("B3:B4").FormulaR1C1 = "Плотн. Кг/м3"

            sheet1.Range("B3:B4").Font.Name = "Arial CYR"
            sheet1.Range("B3:B4").Font.FontStyle = "полужирный"
            sheet1.Range("B3:B4").Font.Size = 10

            sheet1.Range("C3:N3").NumberFormat = "@"
            sheet1.Range("C3:D3").FormulaR1C1 = "0-8"
            sheet1.Range("E3:F3").FormulaR1C1 = "8-16"
            sheet1.Range("G3:H3").FormulaR1C1 = "16-0"
            sheet1.Range("I3:K3").FormulaR1C1 = "Сутки"
            sheet1.Range("L3:N3").FormulaR1C1 = "Месяц"

            sheet1.Range("C4:N4").HorizontalAlignment = xlCenter
            sheet1.Range("C4:N4").VerticalAlignment = xlBottom

            sheet1.Range("A5").FormulaR1C1 = "Взято"
            sheet1.Range("C4").FormulaR1C1 = "м3"
            sheet1.Range("D4").FormulaR1C1 = "тн"
            sheet1.Range("E4").FormulaR1C1 = "м3"
            sheet1.Range("F4").FormulaR1C1 = "тн"
            sheet1.Range("G4").FormulaR1C1 = "м3"
            sheet1.Range("H4").FormulaR1C1 = "тн"
            sheet1.Range("I4").FormulaR1C1 = "м3"
            sheet1.Range("J4").FormulaR1C1 = "тн"
            sheet1.Range("K4").FormulaR1C1 = "%"
            sheet1.Range("L4").FormulaR1C1 = "м3"
            sheet1.Range("M4").FormulaR1C1 = "тн"
            sheet1.Range("N4").FormulaR1C1 = "%"
            sheet1.Columns("B:B").ColumnWidth = 6.43
            sheet1.Columns("C:C").ColumnWidth = 9.14
            sheet1.Columns("D:D").ColumnWidth = 7
            sheet1.Columns("E:E").ColumnWidth = 9.14
            sheet1.Columns("F:F").ColumnWidth = 7
            sheet1.Columns("G:G").ColumnWidth = 9.14
            sheet1.Columns("H:H").ColumnWidth = 7
            sheet1.Columns("I:I").ColumnWidth = 10.43
            sheet1.Columns("J:J").ColumnWidth = 9
            sheet1.Columns("K:K").ColumnWidth = 6.14
            sheet1.Columns("L:L").ColumnWidth = 10.14
            sheet1.Columns("M:M").ColumnWidth = 9.29
            sheet1.Columns("N:N").ColumnWidth = 6.14
            # ============================================================================

            sheet1.Range("A" + str(col_in + 7)).FormulaR1C1 = "Получено"

            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).Font.Bold = True

            sheet1.Range("A3:N" + str(col_in + col_out + col_urov + 10)).Borders(xlEdgeLeft).Weight = xlThin
            sheet1.Range("A3:N" + str(col_in + col_out + col_urov + 10)).Borders(xlEdgeTop).Weight = xlThin
            sheet1.Range("A3:N" + str(col_in + col_out + col_urov + 10)).Borders(xlEdgeBottom).Weight = xlThin
            sheet1.Range("A3:N" + str(col_in + col_out + col_urov + 10)).Borders(xlEdgeRight).Weight = xlThin
            sheet1.Range("A3:N" + str(col_in + col_out + col_urov + 10)).Borders(xlInsideVertical).Weight = xlThin

            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).HorizontalAlignment = xlLeft
            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).VerticalAlignment = xlBottom
            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).IndentLevel = 0
            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).ShrinkToFit = True

            sheet1.Range("B5:N" + str(col_in + col_out + col_urov + 10)).HorizontalAlignment = xlRight
            sheet1.Range("B5:N" + str(col_in + col_out + col_urov + 10)).VerticalAlignment = xlBottom
            sheet1.Range("B5:N" + str(col_in + col_out + col_urov + 10)).ShrinkToFit = True

            sheet1.Range("B5:N" + str(col_in + col_out + col_urov + 10)).NumberFormat = format_data
            sheet1.Range("C" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).NumberFormat = "0.000"
            sheet1.Range("A5:A" + str(col_in + col_out + col_urov + 10)).NumberFormat = "@"
            sheet1.Rows("5:" + str(col_in + col_out + col_urov + 10)).RowHeight = 16.5

            sheet1.Range("A" + str(col_in + 6) + ":N" + str(col_in + 6)).Borders(xlEdgeLeft).Weight = xlThin
            sheet1.Range("A" + str(col_in + 6) + ":N" + str(col_in + 6)).Borders(xlEdgeTop).Weight = xlThin
            sheet1.Range("A" + str(col_in + 6) + ":N" + str(col_in + 6)).Borders(xlEdgeBottom).Weight = xlThin
            sheet1.Range("A" + str(col_in + 6) + ":N" + str(col_in + 6)).Borders(xlEdgeRight).Weight = xlThin
            sheet1.Range("A" + str(col_in + 6) + ":N" + str(col_in + 6)).Borders(xlInsideVertical).Weight = xlThin

            sheet1.Range("A5").HorizontalAlignment = xlCenter
            sheet1.Range("A5").VerticalAlignment = xlBottom

            sheet1.Range("A" + str(col_in + 6) + ":A" + str(col_in + 7)).HorizontalAlignment = xlCenter
            sheet1.Range("A" + str(col_in + 6) + ":A" + str(col_in + 7)).VerticalAlignment = xlBottom

            sheet1.Range("A" + str(col_in + col_out + 8) + ":A" + str(col_in + col_out + 9)).HorizontalAlignment = xlCenter
            sheet1.Range("A" + str(col_in + col_out + 8) + ":A" + str(col_in + col_out + 9)).VerticalAlignment = xlBottom

            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlEdgeLeft).Weight = xlThin
            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlEdgeTop).Weight = xlThin
            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlEdgeBottom).Weight = xlThin
            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlEdgeRight).Weight = xlThin
            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlInsideVertical).Weight = xlThin
            sheet1.Range("A" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Borders(
                    xlInsideHorizontal).Weight = xlThin

            sheet1.Range("B" + str(col_in + 6) + ":N" + str(col_in + 6)).Font.Bold = True
            sheet1.Range("B" + str(col_in + col_out + 8) + ":N" + str(col_in + col_out + 9)).Font.Bold = True
            sheet1.Range("B11:N11").Font.Bold = True
            sheet1.Range("B17:N17").Font.Bold = True

            sheet1.PageSetup.LeftMargin = 0.5
            sheet1.PageSetup.RightMargin = 0.5
            sheet1.PageSetup.BottomMargin = 0.5
            sheet1.PageSetup.TopMargin = 0.5
            sheet1.PageSetup.Orientation = 2

            # ============================================================================

            wb1.SaveAs(os.path.abspath(os.curdir) + '\\obolochka')

            wb1.Close()
            Excel1.Quit()

        except Exception as Er:
            open(os.curdir + '\\report_error.txt', 'a').write(
                'Ошибка при печати оболочки: ' + str(Er) + ': ' + datetime.today().strftime("%d.%m.%Y %H:%M:%S") + '\n')
            wb1.Close(SaveChanges=0)
            Excel1.Quit()


if __name__ == '__main__':
    pr_exel()
