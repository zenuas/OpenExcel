Imports System


Public Class CellRange

    Private from_ As CellIndex
    Private to_ As CellIndex

    Public Sub New(ByVal from_col As UInteger, ByVal from_row As UInteger, ByVal to_col As UInteger, ByVal to_row As UInteger)

        Me.from_ = New CellIndex(from_col, from_row)
        Me.to_ = New CellIndex(to_col, to_row)
    End Sub

    Public Sub New(ByVal from As CellIndex, ByVal to_ As CellIndex)

        Me.from_ = from
        Me.to_ = to_
    End Sub

    Public ReadOnly Property From As CellIndex
        Get
            Return Me.from_
        End Get
    End Property

    Public ReadOnly Property [To] As CellIndex
        Get
            Return Me.to_
        End Get
    End Property

    Public Shared Widening Operator CType(ByVal name As String) As CellRange

        Return ConvertCellRange(name)
    End Operator

    Public Shared Function ConvertCellRange(ByVal name As String) As CellRange

        If String.IsNullOrEmpty(name) Then Throw New ArgumentException("name")

        Dim i = 0
        Dim ByIndex = Function()

                          Dim c = 0UI
                          While i < name.Length

                              If name(i) < "0"c OrElse name(i) > "9"c Then Exit While
                              c = c * 10UI + (Convert.ToUInt32(name(i)) - Convert.ToUInt32("0"c))
                              i += 1
                          End While

                          Return c
                      End Function

        Dim ByName = Function()

                         Dim c = 0UI
                         While i < name.Length

                             If name(i) < "A"c OrElse name(i) > "Z"c Then Exit While
                             c = c * 26UI + (Convert.ToUInt32(name(i)) - Convert.ToUInt32("A"c) + 1UI)
                             i += 1
                         End While

                         Return c
                     End Function

        name = name.ToUpper
        Dim from_col As UInteger = ByName()
        Dim from_row As UInteger = ByIndex()
        If name(i) <> ":"c Then Throw New ArgumentException("name")
        i += 1
        Dim to_col As UInteger = ByName()
        Dim to_row As UInteger = ByIndex()

        Return New CellRange(from_col, from_row, to_col, to_row)
    End Function

End Class
