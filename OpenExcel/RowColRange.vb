Imports System


Public Class RowColRange

    Private from_ As UInteger
    Private to_ As UInteger

    Public Sub New(ByVal from As UInteger, ByVal to_ As UInteger)

        Me.from_ = from
        Me.to_ = to_
    End Sub

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

    Public Shared Widening Operator CType(ByVal name As String) As RowColRange

        Return ConvertRange(name)
    End Operator

    Public Shared Function ConvertRange(ByVal name As String) As RowColRange

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

        Dim start = 0UI
        Dim end_ = 0UI
        name = name.ToUpper
        If name(0) < "A"c OrElse name(0) > "Z"c Then

            start = ByIndex()
            If i < name.Length AndAlso name(i) = ":"c Then

                i += 1
                end_ = ByIndex()
            Else

                end_ = start
            End If
        Else

            start = ByName()
            If i < name.Length AndAlso name(i) = ":"c Then

                i += 1
                end_ = ByName()
            Else

                end_ = start
            End If
        End If

        Return New RowColRange(start, end_)
    End Function

End Class
