Imports OpenExcel


Public Class ColumnsTest
    Implements ITestExcel


    Public Overridable Sub Exec(ByVal path As String) Implements ITestExcel.Exec

        Using xls = Excel.Create(path)

            Dim sheet1 = xls.NewSheet("Sheet1")

            sheet1.Cell("A1") = "A1"
            sheet1.Cell("B2") = "B2"
            sheet1.Cell("C3") = "C3"
            sheet1.Cell("D4") = "D4"
            sheet1.Cell("E5") = "E5"

            ' ABCDE →　ABCBCDE
            sheet1.CopyInsertBeforeMultiColumn("B:C", "B", 1)

            ' ABCBCDE → ABCDBCDE
            sheet1.CopyInsertBeforeMultiColumn("F:F", "D", 1)

            ' ABCDBCDE → ABCDBCDABCDBCDE
            sheet1.CopyInsertBeforeMultiColumn("A:G", "A", 1)

            Dim sheet2 = xls.NewSheet("Sheet2")

            sheet2.Cell("A1") = "A1"
            sheet2.Cell("B2") = "B2"
            sheet2.Cell("C3") = "C3"
            sheet2.Cell("D4") = "D4"
            sheet2.Cell("E5") = "E5"

            sheet2.Columns("B").Hidden = True

            sheet2.Columns("D:E").Hidden = True

            sheet2.Columns("A:D").Hidden = False
            sheet2.Columns("A:D").Width = 5
        End Using
    End Sub

End Class
