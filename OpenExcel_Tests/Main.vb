Imports OpenExcel
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class Main

    Public Shared Sub Main()

        Using xls = XLWorkbook.Create("Excel.xlsx")

            Dim sheet1 = xls.NewSheet("Sheet1")
            Dim sheet2 = xls.NewSheet("Sheet2")
            Dim sheet3 = xls.NewSheet("Sheet3", "Sheet2")
            sheet1.Cell("B1") = "B1"
            sheet1.Cell("A1") = "A1"
            sheet1.Cell("C1") = "C1"

            For i = 100 To 10 Step -1

                sheet1.Cell("D", i) = "D" + i.ToString
            Next
            sheet1.DeleteLine(15, 2)
            sheet1.InsertBeforeColumn("B", 2)
            For i = 10 To 100

                sheet1.Cell("B", i) = "B" + i.ToString
            Next

            sheet1.Rows(20, 60).Hidden = True
            sheet1.CopyInsertBeforeLine(1, 2)
            sheet1.CopyInsertBeforeMultiLine(50, 52, 2, 4)
            sheet1.CopyInsertBeforeMultiColumn("B:D", "A", 2)

            sheet1.Cell("B17") = "B17"
            Dim cell = sheet1.CellValue("B17")
            cell.Style.TopBorder.Style = BorderStyleValues.Thin
            cell.Style.PatternFill.PatternType = PatternValues.Solid
            cell.Style.PatternFill.ForegroundColor = New ForegroundColor() With {.Rgb = HexBinaryValue.FromString("ffff00")}
            cell.UpdateStyle()
        End Using

        For Each x In New ITestExcel() {
                New ColumnsTest
            }

            x.Exec(x.GetType.Name + ".xlsx")

        Next

    End Sub
End Class
