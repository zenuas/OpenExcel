Imports OpenExcel


Public Class Main

    Public Shared Sub Main()


        Using xls = Excel.Create("dummy.xlsx")

            xls.NewSheet("Sheet1")
        End Using
        Using xls = Excel.Create("Excel.xlsx")

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
            'sheet1.InsertLineBefore(25, 2)

            'sheet1.VisibleLine(20, False, 2)
        End Using

        Using xls = Excel.Open("Excel.xlsx")

            Dim sheet2 = xls.WorkSheets("Sheet2")
            sheet2.Cell("A1") = "test2"

            For i = 2 To 100

                sheet2.Cell(i, i) = CellIndex.ConvertColumnName(i) + i.ToString
            Next
            
            xls.SaveAs("Excel2.xlsx")

        End Using

        'Using doc = SpreadsheetDocument.Create("test.xlsx", SpreadsheetDocumentType.Workbook)

        '    Dim book_part = doc.AddWorkbookPart
        '    Dim sheet_part = book_part.AddNewPart(Of WorksheetPart)()

        '    book_part.Workbook = New Workbook(
        '        New Sheets(
        '            New Sheet With
        '            {
        '                .Id = book_part.GetIdOfPart(sheet_part),
        '                .SheetId = 1,
        '                .Name = "test_sheet"
        '            })
        '        )
        '    sheet_part.Worksheet = New Worksheet(
        '        New SheetData(
        '            New Row(
        '                New Cell With
        '                {
        '                    .DataType = CellValues.String,
        '                    .CellReference = "A1",
        '                    .CellValue = New CellValue("test aa")
        '                })
        '            )
        '        )
        'End Using
    End Sub
End Class
