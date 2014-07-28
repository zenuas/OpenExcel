Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLColumns
    Implements IEnumerable(Of XLColumn), IEnumerator(Of XLColumn)


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

#Region "IEnumerator implements"

    Public Overridable Function GetEnumerator() As IEnumerator(Of XLColumn) Implements IEnumerable(Of XLColumn).GetEnumerator

        Me.Reset()
        Return Me
    End Function

    Public Overridable Function GetEnumerator_() As IEnumerator Implements IEnumerable.GetEnumerator

        Return Me.GetEnumerator
    End Function

#End Region

#Region "IEnumerator implements"

    Private current_from_ As UInteger = 0UI
    Private current_to_ As UInteger = 0UI

    Public Overridable ReadOnly Property Current As XLColumn Implements IEnumerator(Of XLColumn).Current
        Get
            Return New XLColumn(Me.Worksheet, Me.current_from_, Me.current_to_)
        End Get
    End Property

    Public Overridable ReadOnly Property Current_ As Object Implements IEnumerator.Current
        Get
            Return Me.Current
        End Get
    End Property

    Public Overridable Function MoveNext() As Boolean Implements IEnumerator.MoveNext

        If Me.To < Me.current_from_ Then Return False
        Me.current_to_ = Me.current_to_ + 1UI
        Me.current_from_ = Me.current_to_
        If Me.To < Me.current_from_ Then Return False

        '               From             To
        '                123456789012345678
        ' case 1  <----> |                |          対象外
        ' case 2      <--+-->             |          分割対象 1～4
        ' case 3         |                |          対象作成 5～6
        ' case 4         |     <---->     |          対象     7～12
        ' case 5         |                |          対象作成 13～14
        ' case 6         |             <--+-->       分割対象 15～18
        ' case 7         |                | <---->   対象外
        ' case 8      <--+----------------+-->       分割対象 1～18

        Dim col = Me.Worksheet.ColumnsData.Elements(Of Column).Where(Function(c) c.Min.Value <= Me.To AndAlso c.Max.Value >= Me.From AndAlso c.Max.Value >= Me.current_from_).FirstOrDefault
        If col Is Nothing Then

            ' case 3、case 5
            Me.current_to_ = Me.To

        ElseIf Me.current_from_ < col.Min.Value Then

            ' case 4
            Me.current_to_ = col.Min.Value - 1UI
        Else

            ' case 2、case 6、case 8
            Me.current_to_ = Math.Min(col.Max.Value, Me.To)
        End If

        Return True
    End Function

    Public Overridable Sub Reset() Implements IEnumerator.Reset

        Me.current_from_ = Me.From - 1UI
        Me.current_to_ = Me.From - 1UI
    End Sub

#End Region

#Region "IDisposable implements"

    Public Sub Dispose() Implements IDisposable.Dispose

    End Sub

#End Region

    Public Overridable WriteOnly Property Width As Double?
        Set(ByVal value As Double?)

            For Each col In Me

                col.Width = value
            Next
        End Set
    End Property

    Public Overridable WriteOnly Property Hidden As Boolean?
        Set(ByVal value As Boolean?)

            For Each col In Me

                col.Hidden = value
            Next
        End Set
    End Property

End Class
