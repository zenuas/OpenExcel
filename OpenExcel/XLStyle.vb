Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLStyle

    Private stylesheet_ As XLStylesheet

    Public Sub New(ByVal stylesheet As XLStylesheet)

        Me.stylesheet_ = stylesheet
    End Sub

    Public Sub New(ByVal stylesheet As XLStylesheet, ByVal index As Integer)
        Me.New(stylesheet)

        Me.LoadCellFormat(index)
    End Sub

    Public Overridable ReadOnly Property Stylesheet As XLStylesheet
        Get
            Return Me.stylesheet_
        End Get
    End Property

    Public Overridable Sub LoadCellFormat(ByVal index As Integer)

        Dim format = CType(Me.Stylesheet.CellFormats(index), CellFormat)

        If format.BorderId.HasValue Then

            Dim x = CType(Me.Stylesheet.Borders(CInt(format.BorderId.Value)), Border)
            Me.top_border_ = x.TopBorder
            Me.bottom_border_ = x.BottomBorder
            Me.left_border_ = x.LeftBorder
            Me.right_border_ = x.RightBorder
            Me.horizontal_border_ = x.HorizontalBorder
            Me.vertical_border_ = x.VerticalBorder
            Me.diagonal_border_ = x.DiagonalBorder
        End If
    End Sub

    Public Overridable Function SaveCellFormat() As UInteger

        Dim new_format As New CellFormat

        If Me.top_border_ IsNot Nothing OrElse
            Me.bottom_border_ IsNot Nothing OrElse
            Me.left_border_ IsNot Nothing OrElse
            Me.right_border_ IsNot Nothing OrElse
            Me.horizontal_border_ IsNot Nothing OrElse
            Me.vertical_border_ IsNot Nothing OrElse
            Me.diagonal_border_ IsNot Nothing Then

            Dim border As New Border
            border.TopBorder = Me.top_border_
            border.BottomBorder = Me.bottom_border_
            border.LeftBorder = Me.left_border_
            border.RightBorder = Me.right_border_
            border.HorizontalBorder = Me.horizontal_border_
            border.VerticalBorder = Me.vertical_border_
            border.DiagonalBorder = Me.diagonal_border_

            Me.Stylesheet.Borders.Append(border)
            new_format.BorderId = CUInt(Me.Stylesheet.Borders.ChildElements.Count - 1)
        End If

        Me.Stylesheet.CellFormats.Append(new_format)
        Return CUInt(Me.Stylesheet.CellFormats.ChildElements.Count - 1)
    End Function

#Region "border"

    Private top_border_ As TopBorder = Nothing
    Private bottom_border_ As BottomBorder = Nothing
    Private left_border_ As LeftBorder = Nothing
    Private right_border_ As RightBorder = Nothing
    Private horizontal_border_ As HorizontalBorder = Nothing
    Private vertical_border_ As VerticalBorder = Nothing
    Private diagonal_border_ As DiagonalBorder = Nothing

    Public Overridable ReadOnly Property TopBorder As TopBorder
        Get
            If Me.top_border_ Is Nothing Then Me.top_border_ = New TopBorder
            Return Me.top_border_
        End Get
    End Property

    Public Overridable ReadOnly Property BottomBorder As BottomBorder
        Get
            If Me.bottom_border_ Is Nothing Then Me.bottom_border_ = New BottomBorder
            Return Me.bottom_border_
        End Get
    End Property

    Public Overridable ReadOnly Property LeftBorder As LeftBorder
        Get
            If Me.left_border_ Is Nothing Then Me.left_border_ = New LeftBorder
            Return Me.left_border_
        End Get
    End Property

    Public Overridable ReadOnly Property RightBorder As RightBorder
        Get
            If Me.right_border_ Is Nothing Then Me.right_border_ = New RightBorder
            Return Me.right_border_
        End Get
    End Property

    Public Overridable ReadOnly Property HorizontalBorder As HorizontalBorder
        Get
            If Me.horizontal_border_ Is Nothing Then Me.horizontal_border_ = New HorizontalBorder
            Return Me.horizontal_border_
        End Get
    End Property

    Public Overridable ReadOnly Property VerticalBorder As VerticalBorder
        Get
            If Me.vertical_border_ Is Nothing Then Me.vertical_border_ = New VerticalBorder
            Return Me.vertical_border_
        End Get
    End Property

    Public Overridable ReadOnly Property DiagonalBorder As DiagonalBorder
        Get
            If Me.diagonal_border_ Is Nothing Then Me.diagonal_border_ = New DiagonalBorder
            Return Me.diagonal_border_
        End Get
    End Property

#End Region

End Class
