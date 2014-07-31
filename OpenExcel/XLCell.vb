Imports System
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLCell

    Private worksheet_ As XLWorksheet
    Private index_ As CellIndex

    Public Sub New(ByVal worksheet As XLWorksheet, ByVal index As CellIndex)

        Me.worksheet_ = worksheet
        Me.index_ = index
    End Sub

    Public Overridable ReadOnly Property Worksheet As XLWorksheet
        Get
            Return Me.worksheet_
        End Get
    End Property

    Public Overridable ReadOnly Property Index As CellIndex
        Get
            Return Me.index_
        End Get
    End Property

    Private cell_cache_ As Cell = Nothing
    Public Overridable ReadOnly Property Value As Cell
        Get
            If Me.cell_cache_ Is Nothing Then Me.cell_cache_ = Me.Worksheet.GetCell(Me.Index.Column, Me.Index.Row)
            Return Me.cell_cache_
        End Get
    End Property

#Region "style"

    Private style_cache_ As XLStyle = Nothing
    Public Overridable ReadOnly Property Style As XLStyle
        Get
            If Me.style_cache_ IsNot Nothing Then Return Me.style_cache_
            If Me.Value.StyleIndex Is Nothing OrElse Not Me.Value.StyleIndex.HasValue Then Me.Value.StyleIndex = 0
            Me.style_cache_ = Me.Worksheet.Workbook.Stylesheet.Styles(CInt(Me.Value.StyleIndex.Value))
            Return Me.style_cache_
        End Get
    End Property

    Public Overridable Sub UpdateStyle()

        If Me.style_cache_ Is Nothing Then Return

        Me.Value.StyleIndex = Me.style_cache_.SaveCellFormat
        Me.style_cache_ = Nothing
    End Sub

#End Region

End Class
