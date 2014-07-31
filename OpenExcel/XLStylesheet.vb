Imports System
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml


Public Class XLStylesheet

    Private xls_ As XLWorkbook
    Private style_ As Stylesheet

    Public Sub New(ByVal xls As XLWorkbook, ByVal style As Stylesheet)

        Me.xls_ = xls
        Me.style_ = style
        Me.Init()
    End Sub

    ''' <summary>
    ''' スタイル初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub Init()

        If Me.CellFormats.ChildElements.Count > 0 Then Return

        ' CellFormatsの1要素目(StyleIndex=0)はFont、Fill、Borderを設定しないとフォーマットエラーとなる
        ' Microsoft Excel 2010、2013で現象確認
        Me.CellFormats.Append(New CellFormat() With {.FontId = 0, .FillId = 0, .BorderId = 0})
        Me.Fonts.Append(New Font())
        Me.Fills.Append(New Fill(), New Fill())
        Me.Borders.Append(New Border())
    End Sub

#Region "property"

    Public Overridable ReadOnly Property Styles(ByVal index As Integer) As XLStyle
        Get
            Return New XLStyle(Me, index)
        End Get
    End Property

    Public Overridable ReadOnly Property Stylesheet As Stylesheet
        Get
            Return Me.style_
        End Get
    End Property

    Public Overridable ReadOnly Property Fills As Fills
        Get
            Return Me.AppendStyleElement(Function() New Fills)
        End Get
    End Property

    Public Overridable ReadOnly Property NumberingFormats As NumberingFormats
        Get
            Return Me.AppendStyleElement(Function() New NumberingFormats)
        End Get
    End Property

    Public Overridable ReadOnly Property Fonts As Fonts
        Get
            Return Me.AppendStyleElement(Function() New Fonts)
        End Get
    End Property

    Public Overridable ReadOnly Property Borders As Borders
        Get
            Return Me.AppendStyleElement(Function() New Borders)
        End Get
    End Property

    Public Overridable ReadOnly Property CellStyleFormats As CellStyleFormats
        Get
            Return Me.AppendStyleElement(Function() New CellStyleFormats)
        End Get
    End Property

    Public Overridable ReadOnly Property CellFormats As CellFormats
        Get
            Return Me.AppendStyleElement(Function() New CellFormats)
        End Get
    End Property

    Public Overridable ReadOnly Property CellStyles As CellStyles
        Get
            Return Me.AppendStyleElement(Function() New CellStyles)
        End Get
    End Property

    Public Overridable ReadOnly Property DifferentialFormats As DifferentialFormats
        Get
            Return Me.AppendStyleElement(Function() New DifferentialFormats)
        End Get
    End Property

    Public Overridable Function AppendStyleElement(Of T As OpenXmlElement)(ByVal f As Func(Of T)) As T

        Dim x = Me.Stylesheet.Where(Function(e) TypeOf e Is T).FirstOrDefault
        If x IsNot Nothing Then Return CType(x, T)
        Return CType(Me.AppendStyleElement(f()), T)
    End Function

    Public Overridable Function AppendStyleElement(ByVal child As OpenXmlElement) As OpenXmlElement

        Dim xs = New Type() {
                GetType(Fonts),
                GetType(Fills),
                GetType(Borders),
                GetType(CellStyleFormats),
                GetType(CellStyles),
                GetType(DifferentialFormats),
                GetType(TableStyles),
                GetType(StylesheetExtensionList)
            }
        Dim find = False
        For i As Integer = xs.Length - 1 To 0 Step -1

            Dim x = xs(i)
            If find Then

                Dim before = Me.Stylesheet.Where(Function(c) c.GetType Is x).FirstOrDefault
                If before IsNot Nothing Then

                    Me.Stylesheet.InsertAfter(child, before)
                    Return child
                End If
            Else

                If child.GetType Is x Then find = True
            End If
        Next

        Me.Stylesheet.InsertAfter(child, Nothing)
        Return child
    End Function

#End Region

End Class
