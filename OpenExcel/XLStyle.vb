Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLStyle

    Private stylesheet_ As XLStylesheet
    Private cell_format_ As CellFormat = Nothing

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

        If format.BorderId IsNot Nothing AndAlso format.BorderId.HasValue Then

            Dim x = CType(Me.Stylesheet.Borders(CInt(format.BorderId.Value)), Border)
            Me.top_border_ = x.TopBorder
            Me.bottom_border_ = x.BottomBorder
            Me.left_border_ = x.LeftBorder
            Me.right_border_ = x.RightBorder
            Me.horizontal_border_ = x.HorizontalBorder
            Me.vertical_border_ = x.VerticalBorder
            Me.diagonal_border_ = x.DiagonalBorder
        End If

        If format.FillId IsNot Nothing AndAlso format.FillId.HasValue Then

            Dim x = CType(Me.Stylesheet.Fills(CInt(format.FillId.Value)), Fill)
            Me.pattern_fill_ = x.PatternFill
            Me.gradient_fill_ = x.GradientFill
        End If

        Me.cell_format_ = format
    End Sub

    Public Overridable Function SaveCellFormat() As UInteger

        Dim new_format As CellFormat
        If Me.cell_format_ IsNot Nothing Then

            new_format = CType(Me.cell_format_.Clone, CellFormat)
        Else
            new_format = New CellFormat
        End If

        If Me.top_border_ IsNot Nothing OrElse
            Me.bottom_border_ IsNot Nothing OrElse
            Me.left_border_ IsNot Nothing OrElse
            Me.right_border_ IsNot Nothing OrElse
            Me.horizontal_border_ IsNot Nothing OrElse
            Me.vertical_border_ IsNot Nothing OrElse
            Me.diagonal_border_ IsNot Nothing Then

            Me.Stylesheet.Borders.Append(New Border With
                {
                    .TopBorder = Me.top_border_,
                    .BottomBorder = Me.bottom_border_,
                    .LeftBorder = Me.left_border_,
                    .RightBorder = Me.right_border_,
                    .HorizontalBorder = Me.horizontal_border_,
                    .VerticalBorder = Me.vertical_border_,
                    .DiagonalBorder = Me.diagonal_border_
                })
            new_format.BorderId = CUInt(Me.Stylesheet.Borders.ChildElements.Count - 1)
        End If

        If Me.pattern_fill_ IsNot Nothing OrElse
            Me.gradient_fill_ IsNot Nothing Then

            Dim x = New Fill 
            If Me.pattern_fill_ IsNot Nothing Then x.PatternFill = Me.pattern_fill_
            If Me.gradient_fill_ IsNot Nothing Then x.GradientFill = Me.gradient_fill_
            Me.Stylesheet.Fills.Append(x)
            new_format.FillId = CUInt(Me.Stylesheet.Fills.ChildElements.Count - 1)
        End If

        Me.Stylesheet.CellFormats.Append(new_format)
        Me.cell_format_ = new_format
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

#Region "fill"

    Private pattern_fill_ As PatternFill
    Private gradient_fill_ As GradientFill

    Public Overridable ReadOnly Property PatternFill As PatternFill
        Get
            If Me.pattern_fill_ Is Nothing Then Me.pattern_fill_ = New PatternFill
            Return Me.pattern_fill_
        End Get
    End Property

    Public Overridable ReadOnly Property GradientFill As GradientFill
        Get
            If Me.gradient_fill_ Is Nothing Then Me.gradient_fill_ = New GradientFill
            Return Me.gradient_fill_
        End Get
    End Property

#End Region

End Class
