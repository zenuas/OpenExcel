﻿Imports System.Linq
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

    Public ReadOnly Property ColumnsData As Columns
        Get
            If Me.columns_ Is Nothing Then

                Me.columns_ = Me.Worksheet.Descendants(Of Columns).FirstOrDefault
                If Me.columns_ Is Nothing Then

                    Me.columns_ = New Columns
                    Me.Worksheet.InsertBefore(Me.columns_, Me.SheetData)
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
            
            Dim after = Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value > row).FirstOrDefault
            Me.SheetData.InsertBefore(x, after)
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
    ''' 行セット取得
    ''' </summary>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Rows(ByVal row As UInteger) As XLRows
        Get
            Return New XLRows(Me, row, row)
        End Get
    End Property

    ''' <summary>
    ''' 行セット取得
    ''' </summary>
    ''' <param name="from">行範囲開始(1行目から開始)</param>
    ''' <param name="to_">行範囲終了(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Rows(ByVal from As UInteger, ByVal to_ As UInteger) As XLRows
        Get
            Return New XLRows(Me, from, to_)
        End Get
    End Property

    ''' <summary>
    ''' 行セット取得
    ''' </summary>
    ''' <param name="row">行番号(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Rows(ByVal row As Integer) As XLRows
        Get
            Return New XLRows(Me, CUInt(row), CUInt(row))
        End Get
    End Property

    ''' <summary>
    ''' 行セット取得
    ''' </summary>
    ''' <param name="from">行範囲開始(1行目から開始)</param>
    ''' <param name="to_">行範囲終了(1行目から開始)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Rows(ByVal from As Integer, ByVal to_ As Integer) As XLRows
        Get
            Return New XLRows(Me, CUInt(from), CUInt(to_))
        End Get
    End Property

    ''' <summary>
    ''' 行セット取得
    ''' </summary>
    ''' <param name="row">行番号(A:A形式)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Rows(ByVal row As String) As XLRows
        Get
            Dim x = CellIndex.ConvertRange(row)
            Return New XLRows(Me, x.Item1, x.Item2)
        End Get
    End Property

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
    ''' <remarks>
    ''' 行追加してもExcelのように式の範囲が自動再設定されない
    ''' </remarks>
    Public Overridable Sub InsertBeforeLine(ByVal row As UInteger, Optional ByVal count As UInteger = 1)

        For Each x In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= row).Reverse

            x.RowIndex.Value += count
            For Each c In x.Elements(Of Cell)()

                Dim ref = CellIndex.ConvertCellIndex(c.CellReference)
                c.CellReference = CellIndex.ToAddress(ref.Column, x.RowIndex.Value)
            Next
        Next
    End Sub

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

#End Region

#Region "multi-line operation"

    ''' <summary>
    ''' 前に複数行コピー追加
    ''' </summary>
    ''' <param name="from_start">コピー元開始</param>
    ''' <param name="from_end">コピー元終了</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 行追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' コピー元の範囲内に追加位置を設定してはいけない
    '''   from_start &lt; to_ &amp;&amp; to_ &lt; from_end の場合エラー
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeMultiLine(ByVal from_start As UInteger, ByVal from_end As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Dim length = count * (from_end - from_start + 1UI)
        Me.InsertBeforeLine(to_, length)
        If to_ <= from_start Then

            from_start += length
            from_end += length
        End If

        Dim after = Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value > to_).FirstOrDefault

        For i = 0UI To count - 1UI

            For Each row In Me.SheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value >= from_start AndAlso r.RowIndex.Value <= from_end)

                Dim copy_row = CType(row.Clone, Row)
                copy_row.RowIndex = to_ + (from_end - from_start + 1UI) * i + copy_row.RowIndex.Value - from_start
                For Each c In copy_row.Elements(Of Cell)()

                    Dim index = CellIndex.ConvertCellIndex(c.CellReference)
                    c.CellReference = CellIndex.ToAddress(index.Column, copy_row.RowIndex.Value)
                Next
                Me.SheetData.InsertBefore(copy_row, after)
            Next
        Next
    End Sub

    ''' <summary>
    ''' 前に複数行コピー追加
    ''' </summary>
    ''' <param name="from">コピー元範囲</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 行追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' コピー元の範囲内に追加位置を設定してはいけない
    '''   from開始 &lt; to_ &amp;&amp; to_ &lt; from終了 の場合エラー
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeMultiLine(ByVal from As String, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Dim x = CellIndex.ConvertRange(from)
        Me.CopyInsertBeforeMultiLine(x.Item1, x.Item2, to_, count)
    End Sub

#End Region

#Region "column operation"

    ''' <summary>
    ''' カラムセット取得
    ''' </summary>
    ''' <param name="col">列名(A:A形式)</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Columns(ByVal col As String) As XLColumns
        Get
            Dim x = CellIndex.ConvertRange(col)
            Return New XLColumns(Me, x.Item1, x.Item2)
        End Get
    End Property

    ''' <summary>
    ''' 列削除
    ''' </summary>
    ''' <param name="col">列位置(1列目から開始)</param>
    ''' <param name="count">対象列数</param>
    ''' <remarks>
    ''' 列削除してもExcelのように式の範囲が自動再設定されない
    ''' </remarks>
    Public Overridable Sub DeleteColumn(ByVal col As UInteger, ByVal count As UInteger)

        For Each r In Me.SheetData.Elements(Of Row)()

            For Each c In r.Elements(Of Cell).Where(
                Function(x)
                    Dim index = CellIndex.ConvertCellIndex(x.CellReference)
                    Return index.Column >= col AndAlso index.Column < col + count
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
    Public Overridable Sub DeleteColumn(ByVal col As String, ByVal count As UInteger)

        Me.DeleteColumn(CellIndex.ConvertColumnIndex(col), count)
    End Sub

    ''' <summary>
    ''' 列削除
    ''' </summary>
    ''' <param name="col">列名(A:A形式)</param>
    ''' <remarks></remarks>
    Public Overridable Sub DeleteColumn(ByVal col As String)

        Dim x = CellIndex.ConvertRange(col)
        Me.DeleteColumn(x.Item1, x.Item2 - x.Item1 + 1UI)
    End Sub

    ''' <summary>
    ''' 前に列追加
    ''' </summary>
    ''' <param name="col">列位置(1列目から開始)</param>
    ''' <param name="count">追加列数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub InsertBeforeColumn(ByVal col As UInteger, ByVal count As UInteger)

        For Each r In Me.SheetData.Elements(Of Row)()

            For Each c In r.Elements(Of Cell).Where(Function(x) CellIndex.ConvertCellIndex(x.CellReference).Column >= col).Reverse

                Dim index = CellIndex.ConvertCellIndex(c.CellReference)
                c.CellReference = CellIndex.ToAddress(index.Column + count, index.Row)
            Next
        Next
    End Sub

    ''' <summary>
    ''' 前に列追加
    ''' </summary>
    ''' <param name="col">列名(A:A形式)</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub InsertBeforeColumn(ByVal col As String)

        Dim x = CellIndex.ConvertRange(col)
        Me.InsertBeforeColumn(x.Item1, x.Item2 - x.Item1 + 1UI)
    End Sub

    ''' <summary>
    ''' 前に列追加
    ''' </summary>
    ''' <param name="col">列名</param>
    ''' <param name="count">追加列数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub InsertBeforeColumn(ByVal col As String, Optional ByVal count As UInteger = 1)

        Me.InsertBeforeColumn(CellIndex.ConvertColumnIndex(col), count)
    End Sub


    ''' <summary>
    ''' 前に列コピー追加
    ''' </summary>
    ''' <param name="from">コピー元</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeColumn(ByVal from As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Me.CopyInsertBeforeMultiColumn(from, from, to_, count)
    End Sub

    ''' <summary>
    ''' 前に列コピー追加
    ''' </summary>
    ''' <param name="from">コピー元</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeColumn(ByVal from As String, ByVal to_ As String, Optional ByVal count As UInteger = 1)

        Me.CopyInsertBeforeColumn(CellIndex.ConvertColumnIndex(from), CellIndex.ConvertColumnIndex(to_), count)
    End Sub

#End Region

#Region "multi-column operation"

    ''' <summary>
    ''' 前に複数列コピー追加
    ''' </summary>
    ''' <param name="from_start">コピー元開始</param>
    ''' <param name="from_end">コピー元終了</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' コピー元の範囲内に追加位置を設定してはいけない
    '''   from_start &lt; to_ &amp;&amp; to_ &lt; from_end の場合エラー
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeMultiColumn(ByVal from_start As UInteger, ByVal from_end As UInteger, ByVal to_ As UInteger, Optional ByVal count As UInteger = 1)

        Dim length = count * (from_end - from_start + 1UI)
        Me.InsertBeforeColumn(to_, length)
        If to_ <= from_start Then

            from_start += length
            from_end += length
        End If

        For i = 0UI To count - 1UI

            For Each row In Me.SheetData.Elements(Of Row)()

                Dim insert_to = to_ + (from_end - from_start + 1UI) * i
                Dim after = row.Elements(Of Cell).Where(Function(c) CellIndex.ConvertCellIndex(c.CellReference).Column > insert_to).FirstOrDefault

                For Each c In row.Elements(Of Cell).Where(
                    Function(x)
                        Dim ref = CellIndex.ConvertCellIndex(x.CellReference)
                        Return ref.Column >= from_start AndAlso ref.Column <= from_end
                    End Function)

                    Dim copy_col = CType(c.Clone, Cell)
                    Dim index = CellIndex.ConvertCellIndex(copy_col.CellReference)
                    copy_col.CellReference = CellIndex.ToAddress(insert_to + index.Column - from_start, index.Row)

                    row.InsertBefore(copy_col, after)
                Next
            Next
        Next
    End Sub

    ''' <summary>
    ''' 前に複数列コピー追加
    ''' </summary>
    ''' <param name="from_start">コピー元開始</param>
    ''' <param name="from_end">コピー元終了</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' コピー元の範囲内に追加位置を設定してはいけない
    '''   from_start &lt; to_ &amp;&amp; to_ &lt; from_end の場合エラー
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeMultiColumn(ByVal from_start As String, ByVal from_end As String, ByVal to_ As String, Optional ByVal count As UInteger = 1)

        Me.CopyInsertBeforeMultiColumn(CellIndex.ConvertColumnIndex(from_start), CellIndex.ConvertColumnIndex(from_end), CellIndex.ConvertColumnIndex(to_), count)
    End Sub

    ''' <summary>
    ''' 前に複数列コピー追加
    ''' </summary>
    ''' <param name="from">コピー元範囲</param>
    ''' <param name="to_">追加位置</param>
    ''' <param name="count">コピー回数</param>
    ''' <remarks>
    ''' 列追加してもExcelのように式の範囲が自動再設定されない、式は再計算されない
    ''' コピー元の範囲内に追加位置を設定してはいけない
    '''   from開始 &lt; to_ &amp;&amp; to_ &lt; from終了 の場合エラー
    ''' </remarks>
    Public Overridable Sub CopyInsertBeforeMultiColumn(ByVal from As String, ByVal to_ As String, Optional ByVal count As UInteger = 1)

        Dim x = CellIndex.ConvertRange(from)
        Me.CopyInsertBeforeMultiColumn(x.Item1, x.Item2, CellIndex.ConvertColumnIndex(to_), count)
    End Sub

#End Region

End Class
