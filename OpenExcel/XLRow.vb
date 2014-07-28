Imports System
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLRow

    Private sheet_ As XLWorksheet
    Private row_index_ As UInteger

    Public Sub New(ByVal sheet As XLWorksheet, ByVal row_index As UInteger)

        Me.sheet_ = sheet
        Me.row_index_ = row_index
    End Sub

    Public Overridable ReadOnly Property Worksheet As XLWorksheet
        Get
            Return Me.sheet_
        End Get
    End Property

    Public Overridable ReadOnly Property RowIndex As UInteger
        Get
            Return Me.row_index_
        End Get
    End Property

    Public Overridable ReadOnly Property Row As Row
        Get
            Return Me.Worksheet.GetRow(Me.RowIndex)
        End Get
    End Property

End Class
