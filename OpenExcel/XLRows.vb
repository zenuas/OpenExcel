Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml.Spreadsheet


Public Class XLRows
    Implements IEnumerable(Of XLRow), IEnumerator(Of XLRow)


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

    Public Overridable Function GetEnumerator() As IEnumerator(Of XLRow) Implements IEnumerable(Of XLRow).GetEnumerator

        Me.Reset()
        Return Me
    End Function

    Public Overridable Function GetEnumerator_() As IEnumerator Implements IEnumerable.GetEnumerator

        Return Me.GetEnumerator
    End Function

#End Region

#Region "IEnumerator implements"

    Private current_index_ As UInteger = 0UI

    Public Overridable ReadOnly Property Current As XLRow Implements IEnumerator(Of XLRow).Current
        Get
            Return New XLRow(Me.Worksheet, Me.current_index_)
        End Get
    End Property

    Public Overridable ReadOnly Property Current_ As Object Implements IEnumerator.Current
        Get
            Return Me.Current
        End Get
    End Property

    Public Overridable Function MoveNext() As Boolean Implements IEnumerator.MoveNext

        If Me.To < Me.current_index_ Then Return False
        Me.current_index_ += 1UI
        If Me.To < Me.current_index_ Then Return False

        Return True
    End Function

    Public Overridable Sub Reset() Implements IEnumerator.Reset

        Me.current_index_ = Me.From - 1UI
    End Sub

#End Region

#Region "IDisposable implements"

    Public Sub Dispose() Implements IDisposable.Dispose

    End Sub

#End Region

    Public Overridable WriteOnly Property Height As Double?
        Set(ByVal value As Double?)

            For Each row In Me

                row.Row.Height = XmlUtility.ToDoubleValue(value)
            Next
        End Set
    End Property

    ''' <summary>
    ''' 行の表示設定
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Overridable WriteOnly Property Hidden As Boolean?
        Set(value As Boolean?)

            For Each x In Me

                x.Row.Hidden = XmlUtility.ToBooleanValue(value)
            Next
        End Set
    End Property

End Class
