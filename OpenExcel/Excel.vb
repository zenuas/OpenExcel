Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


<Assembly: System.Reflection.AssemblyVersion("0.0.*")> 
Public Class Excel
    Implements IDisposable

#Region "constructor"

    Protected Sub New()

    End Sub

    ''' <summary>
    ''' 空のシートを作成する
    ''' </summary>
    ''' <param name="path">保存先</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(ByVal path As String) As Excel

        Dim self = New Excel
        self.iscreate_ = True
        self.SavePath = path
        self.original_path_ = path
        self.xls_ = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)
        self.Init()
        Return self
    End Function

    ''' <summary>
    ''' 既存のブックを開く
    ''' </summary>
    ''' <param name="path">対象</param>
    ''' <param name="auto_save">自動保存指定</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Open(ByVal path As String, Optional ByVal auto_save As Boolean = False) As Excel


        Dim self = New Excel
        self.iscreate_ = False
        If Not auto_save Then

            Dim temp = IO.Path.GetTempFileName
            IO.File.Copy(path, temp, True)
            path = temp
            self.SavePath = ""
        Else
            self.SavePath = path
        End If
        self.original_path_ = path
        self.xls_ = SpreadsheetDocument.Open(path, True)
        self.UpdateSheets()
        Return self
    End Function

    ''' <summary>
    ''' ブックの初期化を行い空のシートを作成する
    ''' </summary>
    ''' <param name="sheet_name">初期作成シート名</param>
    ''' <remarks></remarks>
    Public Overridable Sub Init(Optional ByVal sheet_name As String = "Sheet1")

        If Me.Document.WorkbookPart Is Nothing Then

            Dim book_part = Me.Document.AddWorkbookPart
            book_part.Workbook = New Workbook
            book_part.Workbook.Append(New Sheets)
        End If
    End Sub

    Public Overridable Sub UpdateSheets()

        Me.sheets_.Clear()
        For Each sheet_part In Me.Document.WorkbookPart.WorksheetParts

            Me.sheets_.Add(sheet_part.Worksheet, New XLWorksheet(sheet_part.Worksheet))
        Next
    End Sub

#End Region

#Region "property"

    Private iscreate_ As Boolean
    Private original_path_ As String
    Private xls_ As SpreadsheetDocument
    Private sheets_ As New Dictionary(Of Worksheet, XLWorksheet)

    Public Overridable ReadOnly Property OriginalPath As String
        Get
            Return Me.original_path_
        End Get
    End Property

    Public Overridable Property SavePath As String

    Public Overridable ReadOnly Property Document As SpreadsheetDocument
        Get
            Return Me.xls_
        End Get
    End Property

    Public Overridable Function WorksheetToXLWorkSheet(ByVal sheet As Worksheet) As XLWorksheet

        Return Me.sheets_(sheet)
    End Function

#End Region

#Region "save"

    ''' <summary>
    ''' 保存
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub Save()

        Me.SaveAs(Me.SavePath)
    End Sub

    ''' <summary>
    ''' 名前を付けて保存
    ''' </summary>
    ''' <param name="path">保存先</param>
    ''' <remarks>
    ''' 別名保存ができないので一旦上書きをさせてから、終了時にファイル移動を行っている
    ''' そのため最後に保存したファイル名のみが残る
    '''   Create("a.xlsx") -> Save() -> SaveAs("b.xlsx") -> Dispose() と行うと b.xlsx のみが残る
    '''   Create("a.xlsx") -> Save() -> SaveAs("b.xlsx") -> SaveAs("c.xlsx") -> Dispose() と行うと c.xlsx のみが残る
    '''   Open("a.xlsx")   -> Save() -> SaveAs("b.xlsx") -> SaveAs("c.xlsx") -> Dispose() と行うと c.xlsx のみが残る
    ''' </remarks>
    Public Overridable Sub SaveAs(ByVal path As String)

        If String.IsNullOrEmpty(path) Then Throw New ArgumentException("path")

        Me.SavePath = path
        Me.Document.WorkbookPart.Workbook.Save()
    End Sub

#End Region

#Region "sheet operation"

    Public Overridable ReadOnly Property WorkSheets(ByVal sheet_name As String) As XLWorksheet
        Get
            Return Me.GetSheetByName(sheet_name)
        End Get
    End Property

    ''' <summary>
    ''' シートを名前で参照する
    ''' </summary>
    ''' <param name="sheet_name">シート名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetSheetByName(ByVal sheet_name As String) As XLWorksheet

        Return Me.GetSheetById(Me.Document.WorkbookPart.Workbook.Sheets.Descendants(Of Sheet)().Where(Function(s) s.Name = sheet_name).First.Id)
    End Function

    ''' <summary>
    ''' シートをIDで指定する
    ''' </summary>
    ''' <param name="sheet_id">ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetSheetById(ByVal sheet_id As String) As XLWorksheet

        Dim book_part = Me.Document.WorkbookPart
        Return Me.WorksheetToXLWorkSheet(book_part.WorksheetParts.Where(Function(s) book_part.GetIdOfPart(s).Equals(sheet_id)).First.Worksheet)
    End Function

    ''' <summary>
    ''' シートをインデックスで指定する
    ''' </summary>
    ''' <param name="sheet_index">インデックス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetSheetByIndex(ByVal sheet_index As Integer) As XLWorksheet

        Return Me.WorksheetToXLWorkSheet(Me.Document.WorkbookPart.WorksheetParts(sheet_index).Worksheet)
    End Function

    ''' <summary>
    ''' シートをコピーする
    ''' </summary>
    ''' <param name="from">コピー元シート名</param>
    ''' <param name="to_">コピー先シート名</param>
    ''' <param name="insert_position">挿入位置、指定シート名の前に追加する、空文字指定時はブックの末尾に追加する</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function CopySheet(ByVal from As String, ByVal to_ As String, Optional ByVal insert_position As String = "") As XLWorksheet

        Dim from_sheet = Me.GetSheetByName(from)
        Return Nothing
    End Function

    ''' <summary>
    ''' シートを追加する
    ''' </summary>
    ''' <param name="sheet_name">シート名</param>
    ''' <param name="insert_position">挿入位置、指定シート名の前に追加する、空文字指定時はブックの末尾に追加する</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function NewSheet(ByVal sheet_name As String, Optional ByVal insert_position As String = "") As XLWorksheet

        Dim sheets = Me.Document.WorkbookPart.Workbook.GetFirstChild(Of Sheets)()

        Dim sheet_id = 1UI
        If sheets.Elements(Of Sheet).Count > 0 Then

            sheet_id = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max + 1UI
        End If

        Dim sheet_part = Me.Document.WorkbookPart.AddNewPart(Of WorksheetPart)()
        Dim sheet_data = New SheetData
        sheet_part.Worksheet = New Worksheet(sheet_data)
        Dim sheet = New Sheet With
            {
                .Id = Me.Document.WorkbookPart.GetIdOfPart(sheet_part),
                .SheetId = sheet_id,
                .Name = sheet_name
            }

        Dim before = Me.Document.WorkbookPart.Workbook.Sheets.Elements(Of Sheet).Where(Function(s) s.Name = insert_position).FirstOrDefault
        If before Is Nothing Then

            Me.Document.WorkbookPart.Workbook.Sheets.Append(sheet)
        Else

            Me.Document.WorkbookPart.Workbook.Sheets.InsertBefore(sheet, before)
        End If

        Dim x = New XLWorksheet(sheet_part.Worksheet)
        Me.sheets_.Add(sheet_part.Worksheet, x)
        Return x
    End Function

#End Region

#Region "IDisposable implements"

    Public Sub Dispose() Implements IDisposable.Dispose

        If Me.Document IsNot Nothing Then

            Me.Document.Dispose()

            If Not String.IsNullOrEmpty(Me.SavePath) AndAlso Not Me.SavePath.Equals(Me.OriginalPath) Then

                If IO.File.Exists(Me.SavePath) Then IO.File.Delete(Me.SavePath)
                IO.File.Move(Me.OriginalPath, Me.SavePath)

            ElseIf Not Me.iscreate_ Then

                If IO.File.Exists(Me.OriginalPath) Then IO.File.Delete(Me.OriginalPath)
            End If
        End If
        Me.xls_ = Nothing
    End Sub
#End Region

End Class
