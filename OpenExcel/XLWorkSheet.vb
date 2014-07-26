Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLWorksheet

    Private worksheet_ As Worksheet
    Private sheet_ As SheetData
    Private columns_ As Columns = Nothing

    Public Sub New(ByVal worksheet As Worksheet)
        MyBase.New()

        Me.worksheet_ = worksheet
        Me.sheet_ = worksheet.Descendants(Of SheetData).First
    End Sub

    Public ReadOnly Property Worksheet As Worksheet
        Get
            Return Me.worksheet_
        End Get
    End Property

    Public ReadOnly Property SheetData As SheetData
        Get
            Return Me.sheet_
        End Get
    End Property

    Public ReadOnly Property Columns As Columns
        Get
            If Me.columns_ Is Nothing Then

                Me.columns_ = Me.Worksheet.Descendants(Of Columns).FirstOrDefault
                If Me.columns_ Is Nothing Then

                    Me.columns_ = New Columns
                    Me.Worksheet.Append(Me.columns_)
                End If
            End If
            Return Me.columns_
        End Get
    End Property

#Region "cell operation"

    ''' <summary>
    ''' セルプロパティ
    ''' </summary>
    ''' <param name="name">A1形式のセル名</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Cell(ByVal name As String) As String
        Get
            Return Me.GetCell(name).CellValue.InnerText
        End Get
        Set(ByVal value As String)

            Dim x = Me.GetCell(name)
            x.DataType = CellValues.String
            x.CellValue = New CellValue(value)
        End Set
    End Property

    ''' <summary>
    ''' セルプロパティ
    ''' </summary>
    ''' <param name="col">列名</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Cell(ByVal col As String, ByVal row As Integer) As String
        Get
            Return Me.Cell(col, CUInt(row))
        End Get
        Set(ByVal value As String)

            Me.Cell(col, CUInt(row)) = value
        End Set
    End Property

    ''' <summary>
    ''' セルプロパティ
    ''' </summary>
    ''' <param name="col">列名</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Cell(ByVal col As String, ByVal row As UInteger) As String
        Get
            Return Me.GetCell(col, row).CellValue.InnerText
        End Get
        Set(ByVal value As String)

            Dim x = Me.GetCell(col, row)
            x.DataType = CellValues.String
            x.CellValue = New CellValue(value)
        End Set
    End Property

    ''' <summary>
    ''' セルプロパティ
    ''' </summary>
    ''' <param name="col">列番号(1列目から開始)</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Cell(ByVal col As Integer, ByVal row As Integer) As String
        Get
            Return Me.Cell(CellIndex.ConvertColumnName(CUInt(col)), CUInt(row))
        End Get
        Set(ByVal value As String)

            Me.Cell(CellIndex.ConvertColumnName(CUInt(col)), CUInt(row)) = value
        End Set
    End Property

    ''' <summary>
    ''' セルプロパティ
    ''' </summary>
    ''' <param name="col">列番号(1列目から開始)</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property Cell(ByVal col As UInteger, ByVal row As UInteger) As String
        Get
            Return Me.Cell(CellIndex.ConvertColumnName(col), row)
        End Get
        Set(ByVal value As String)

            Me.Cell(CellIndex.ConvertColumnName(col), row) = value
        End Set
    End Property

    ''' <summary>
    ''' 行取得
    ''' </summary>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 行データがない場合はブランク行を追加する
    ''' </remarks>
    Public Overridable Function GetRow(ByVal row As UInteger) As Row

        Dim x = Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = row).FirstOrDefault()
        If x Is Nothing Then

            x = New Row
            x.RowIndex = row

            Dim before = Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value < row).LastOrDefault
            Me.SheetData.InsertAfter(x, before)
        End If

        Return x
    End Function

    ''' <summary>
    ''' セル取得
    ''' </summary>
    ''' <param name="name">A1形式のセル名</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' セルデータがない場合はセルを追加する
    ''' </remarks>
    Public Overridable Function GetCell(ByVal name As String) As Cell

        Dim x = CellIndex.ConvertCellIndex(name)
        Return Me.GetCell(Me.GetRow(x.Row), name)
    End Function

    ''' <summary>
    ''' セル取得
    ''' </summary>
    ''' <param name="col">列番号(1列目から開始)</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' セルデータがない場合はセルを追加する
    ''' </remarks>
    Public Overridable Function GetCell(ByVal col As UInteger, ByVal row As UInteger) As Cell

        Return Me.GetCell(Me.GetRow(row), CellIndex.ToAddress(col, row))
    End Function

    ''' <summary>
    ''' セル取得
    ''' </summary>
    ''' <param name="col">列名</param>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' セルデータがない場合はセルを追加する
    ''' </remarks>
    Public Overridable Function GetCell(ByVal col As String, ByVal row As UInteger) As Cell

        Return Me.GetCell(Me.GetRow(row), CellIndex.ToAddress(col, row))
    End Function

    ''' <summary>
    ''' セル取得
    ''' </summary>
    ''' <param name="row">行データ</param>
    ''' <param name="name">A1形式のセル名</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' セルデータがない場合はセルを追加する
    ''' </remarks>
    Public Overridable Function GetCell(ByVal row As Row, ByVal name As String) As Cell

        Dim x = row.Elements(Of Cell).Where(Function(c) name.Equals(c.CellReference.Value)).FirstOrDefault
        If x Is Nothing Then

            x = New Cell
            x.CellReference = name

            Dim before = row.Elements(Of Cell).Where(Function(c) String.Compare(c.CellReference.Value, name) < 0).LastOrDefault
            row.InsertAfter(x, before)
        End If

        Return x
    End Function

    'Public Overridable Sub DeleteCell(ByVal col As UInteger, ByVal row As UInteger)

    'End Sub

    'Public Overridable Sub DeleteCell(ByVal col As String, ByVal row As UInteger)

    'End Sub

    'Public Overridable Sub DeleteCell(ByVal name As String)

    'End Sub

#End Region

#Region "line operation"

    ''' <summary>
    ''' 行削除
    ''' </summary>
    ''' <param name="row">削除対象行</param>
    ''' <param name="count">削除行数</param>
    ''' <remarks>
    ''' 行削除してもExcelのように式の範囲が自動再設定されない
    ''' </remarks>
    Public Overridable Sub DeleteLine(ByVal row As UInteger, Optional ByVal count As UInteger = 1)

        For Each x In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= row AndAlso r.RowIndex.Value <= row + count - 1).Reverse

            x.Remove()
        Next

        For Each x In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= row + count)

            x.RowIndex.Value -= count
            For Each c In x.Elements(Of Cell)()

                Dim ref = CellIndex.ConvertCellIndex(c.CellReference)
                c.CellReference = CellIndex.ToAddress(ref.Column, x.RowIndex.Value)
            Next
        Next
    End Sub

    ''' <summary>
    ''' 前に行追加
    ''' </summary>
    ''' <param name="row">追加位置</param>
    ''' <param name="count">追加行数</param>
    ''' <returns>追加した最初の行データ</returns>
    ''' <remarks>
    ''' 行追加してもExcelのように式の範囲が自動再設定されない
    ''' </remarks>
    Public Overridable Function InsertBeforeLine(ByVal row As UInteger, Optional ByVal count As UInteger = 1) As Row

        For Each x In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= row).Reverse

            x.RowIndex.Value += count
            For Each c In x.Elements(Of Cell)()

                Dim ref = CellIndex.ConvertCellIndex(c.CellReference)
                c.CellReference = CellIndex.ToAddress(ref.Column, x.RowIndex.Value)
            Next
        Next

        Dim before = Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value < row).LastOrDefault
        Dim new_row As Row = Nothing
        For i = 1UI To count

            new_row = New Row
            new_row.RowIndex = row + count - i
            Me.SheetData.InsertAfter(new_row, before)
        Next

        Return new_row
    End Function

    ''' <summary>
    ''' 前に行コピー追加
    ''' </summary>
    ''' <param name="from">コピー元</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">追加回数</param>
    ''' <remarks>
    ''' 行追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeLine(ByVal from As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Me.CopyInsertBeforeMultiLine(from, from, to_, count)
    End Sub

    ''' <summary>
    ''' 行の表示設定
    ''' </summary>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <param name="visible">表示フラグ</param>
    ''' <param name="count">対象行数</param>
    ''' <remarks></remarks>
    Public Overridable Sub VisibleLine(ByVal row As UInteger, ByVal visible As Boolean, Optional ByVal count As UInteger = 1)

        'For Each x In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= row AndAlso r.RowIndex.Value <= row + count - 1)

        '    x.Hidden = Not visible
        'Next

        For i = 0UI To count - 1UI

            Me.GetRow(row + i).Hidden = Not visible
        Next
    End Sub

#End Region

#Region "multi-line operation"

    Public Overridable Sub CopyInsertBeforeMultiLine(ByVal from_start As UInteger, ByVal from_end As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Me.GetRow(from_start)
    End Sub

    Public Overridable Sub CopyInsertBeforeMultiLine(ByVal from As String, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Dim x = CellIndex.ConvertRange(from)
        Me.CopyInsertBeforeMultiLine(x.Item1, x.Item2, to_, count)
    End Sub

#End Region

#Region "column operation"

    ''' <summary>
    ''' 列削除
    ''' </summary>
    ''' <param name="col">列位置(1列目から開始)</param>
    ''' <param name="count">対象列数</param>
    ''' <remarks></remarks>
    Public Overridable Sub DeleteColumn(ByVal col As UInteger, Optional ByVal count As UInteger = 1)

        For Each r In Me.SheetData.Elements(Of Row)()

            For Each c In r.Elements(Of Cell).Where(
                Function(x)
                    Dim index = CellIndex.ConvertCellIndex(x.CellReference)
                    Return index.Column >= col AndAlso index.Column <= col + count - 1
                End Function).Reverse

                c.Remove()
            Next

            For Each c In r.Elements(Of Cell).Where(Function(x) CellIndex.ConvertCellIndex(x.CellReference).Column >= col + count).Reverse

                Dim index = CellIndex.ConvertCellIndex(c.CellReference)
                c.CellReference = CellIndex.ToAddress(index.Column - count, index.Row)
            Next
        Next
    End Sub

    ''' <summary>
    ''' 列削除
    ''' </summary>
    ''' <param name="col">列名</param>
    ''' <param name="count">対象列数</param>
    ''' <remarks></remarks>
    Public Overridable Sub DeleteColumn(ByVal col As String, Optional ByVal count As UInteger = 1)

        Me.DeleteColumn(CellIndex.ConvertColumnIndex(col), count)
    End Sub

    Public Overridable Sub InsertBeforeColumn(ByVal col As UInteger, Optional ByVal count As UInteger = 1)

    End Sub

    Public Overridable Sub InsertBeforeColumn(ByVal col As String, Optional ByVal count As UInteger = 1)

        Me.InsertBeforeColumn(CellIndex.ConvertColumnIndex(col), count)
    End Sub

    Public Overridable Sub CopyInsertBeforeColumn(ByVal from As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

    End Sub

    Public Overridable Sub CopyInsertBeforeColumn(ByVal from As String, ByVal to_ As String, Optional ByVal count As UInteger = 1)

    End Sub

    Public Overridable Sub VisibleColumn(ByVal col As UInteger, ByVal visible As Boolean, Optional ByVal count As UInteger = 1)

        ' Columnがなければ作る
        ' Columnがあるが A:C (Column {Min=1, Max=3})になっており、 B列のみ非表示する場合割らないといけない？
        For Each x In Me.Columns.Elements(Of Column).Where(Function(c) c.Min.Value >= col AndAlso c.Max.Value <= col + count - 1)

            x.Hidden = Not visible
        Next
    End Sub

    Public Overridable Sub VisibleColumn(ByVal col As String, ByVal visible As Boolean, Optional ByVal count As UInteger = 1)

        Me.VisibleColumn(CellIndex.ConvertColumnIndex(col), visible, count)
    End Sub

#End Region

End Class
