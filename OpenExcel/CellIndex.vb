﻿Imports System


Public Class CellIndex

    Private col_ As UInteger = 0UI
    Private row_ As UInteger = 0UI

    Public Sub New(ByVal col As UInteger, ByVal row As UInteger)

        Me.col_ = col
        Me.row_ = row
    End Sub

    Public Overridable ReadOnly Property Column As UInteger
        Get
            Return Me.col_
        End Get
    End Property

    Public Overridable ReadOnly Property Row As UInteger
        Get
            Return Me.row_
        End Get
    End Property

    Public Shared Widening Operator CType(ByVal name As String) As CellIndex

        Return ConvertCellIndex(name)
    End Operator

    Public Shared Function ConvertColumnName(ByVal col As Integer) As String

        Return ConvertColumnName(CUInt(col))
    End Function

    Public Shared Function ConvertColumnName(ByVal col As UInteger) As String

        If col <= 0 Then Throw New ArgumentException("col")

        Dim name As New System.Text.StringBuilder
        Do
            col -= 1UI
            name.Insert(0, Convert.ToChar(col Mod 26UI + Convert.ToUInt32("A"c)))
            col = col \ 26UI

        Loop While col > 0

        Return name.ToString
    End Function

    Public Shared Function ConvertColumnIndex(ByVal name As String) As UInteger

        If String.IsNullOrEmpty(name) Then Throw New ArgumentException("name")

        name = name.ToUpper
        Dim col = 0UI

        For i As Integer = 0 To name.Length - 1

            If name(i) < "A"c OrElse name(i) > "Z"c Then Throw New ArgumentException("cell")
            col = col * 26UI + (Convert.ToUInt32(name(i)) - Convert.ToUInt32("A"c) + 1UI)
        Next

        Return col
    End Function

    Public Shared Function ConvertCellIndex(ByVal name As String) As CellIndex

        If String.IsNullOrEmpty(name) Then Throw New ArgumentException("name")

        name = name.ToUpper
        Dim col As UInteger = 0UI
        Dim row As UInteger = 0UI

        Dim i = 0
        Do While i < name.Length


            If name(i) < "A"c OrElse name(i) > "Z"c Then Exit Do
            col = col * 26UI + (Convert.ToUInt32(name(i)) - Convert.ToUInt32("A"c) + 1UI)
            i += 1
        Loop

        Do While i < name.Length

            If name(i) < "0"c OrElse name(i) > "9"c Then Exit Do
            row = row * 10UI + (Convert.ToUInt32(name(i)) - Convert.ToUInt32("0"c))
            i += 1
        Loop

        Return New CellIndex(col, row)
    End Function

    Public Function ToAddress() As String

        Return ToAddress(Me.Column, Me.Row)
    End Function

    Public Shared Function ToAddress(ByVal col As Integer, ByVal row As Integer) As String

        Return ToAddress(CUInt(col), CUInt(row))
    End Function

    Public Shared Function ToAddress(ByVal col As String, ByVal row As Integer) As String

        Return ToAddress(col, CUInt(row))
    End Function

    Public Shared Function ToAddress(ByVal col As UInteger, ByVal row As UInteger) As String

        Return ToAddress(ConvertColumnName(col), row)
    End Function

    Public Shared Function ToAddress(ByVal col As String, ByVal row As UInteger) As String

        Return String.Format("{0}{1}", col, row)
    End Function

End Class
