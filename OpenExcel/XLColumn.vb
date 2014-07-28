Imports System
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLColumn

    Private sheet_ As XLWorksheet
    Private from_ As UInteger
    Private to_ As UInteger

    Public Sub New(ByVal sheet As XLWorksheet, ByVal from As UInteger, ByVal to_ As UInteger)

        Me.sheet_ = sheet
        Me.from_ = from
        Me.to_ = to_
    End Sub

    Public Overridable ReadOnly Property Worksheet As XLWorksheet
        Get
            Return Me.sheet_
        End Get
    End Property

    Public Overridable ReadOnly Property From As UInteger
        Get
            Return Me.from_
        End Get
    End Property

    Public Overridable ReadOnly Property [To] As UInteger
        Get
            Return Me.to_
        End Get
    End Property

#Region "column property"

    Private Enum CacheType

        NoRead
        Cached
        Readed
    End Enum

    Private cache_type_ As CacheType = CacheType.NoRead
    Private column_cache_ As Column = Nothing

    Public Overridable ReadOnly Property Column() As Column
        Get
            Return Me.GetColumn
        End Get
    End Property

    Public Overridable ReadOnly Property ColumnCache() As Column
        Get
            Return Me.GetCache
        End Get
    End Property

    Public Overridable Function GetCache() As Column

        If Me.cache_type_ = CacheType.NoRead Then

            Me.column_cache_ = Me.Worksheet.ColumnsData.Elements(Of Column).Where(Function(c) c.Min.Value <= Me.To AndAlso c.Max.Value >= Me.From).FirstOrDefault
            Me.cache_type_ = CacheType.Cached
        End If

        Return Me.column_cache_
    End Function

    Public Overridable Function GetColumn() As Column

        If Me.cache_type_ <> CacheType.Readed Then Me.GetCache()

        If Me.column_cache_ Is Nothing Then

            Me.column_cache_ = New Column()
            Me.column_cache_.Min = Me.From
            Me.column_cache_.Max = Me.To

            Me.Worksheet.ColumnsData.InsertBefore(Me.column_cache_, Me.Worksheet.ColumnsData.Elements(Of Column).Where(Function(c) c.Min.Value >= Me.To).FirstOrDefault)

        ElseIf Me.column_cache_.Min.Value < Me.From AndAlso Me.column_cache_.Max.Value > Me.To Then

            Dim before = CType(Me.column_cache_.Clone, Column)
            Dim after = CType(Me.column_cache_.Clone, Column)
            before.Max = Me.From - 1UI
            Me.column_cache_.Min = Me.From
            Me.column_cache_.Max = Me.To
            after.Min = Me.To + 1UI

            Me.column_cache_ = column_cache_
            Me.Worksheet.ColumnsData.InsertBefore(before, Me.column_cache_)
            Me.Worksheet.ColumnsData.InsertAfter(after, Me.column_cache_)

        ElseIf Me.column_cache_.Min.Value < Me.From Then

            Dim after = CType(Me.column_cache_.Clone, Column)
            Me.column_cache_.Max = Me.From - 1UI
            after.Min = Me.From

            Me.Worksheet.ColumnsData.InsertAfter(after, Me.column_cache_)
            Me.column_cache_ = after

        ElseIf Me.column_cache_.Max.Value > Me.To Then

            Dim before = CType(Me.column_cache_.Clone, Column)
            Me.column_cache_.Min = Me.To + 1UI
            before.Max = Me.To

            Me.Worksheet.ColumnsData.InsertBefore(before, Me.column_cache_)
            Me.column_cache_ = before
        End If

        Return Me.column_cache_
    End Function

#End Region

    Public Overridable Property Width As Double?
        Get
            If Me.ColumnCache Is Nothing OrElse Not Me.ColumnCache.Width.HasValue Then Return Nothing
            Return Me.ColumnCache.Width.Value
        End Get
        Set(ByVal value As Double?)

            Me.Column.Width = XmlUtility.ToDoubleValue(value)
        End Set
    End Property

    Public Overridable Property Hidden As Boolean?
        Get
            If Me.ColumnCache Is Nothing OrElse Me.ColumnCache.Hidden.HasValue Then Return Nothing
            Return Me.ColumnCache.Hidden.Value
        End Get
        Set(ByVal value As Boolean?)

            Me.Column.Hidden = XmlUtility.ToBooleanValue(value)
        End Set
    End Property

End Class
