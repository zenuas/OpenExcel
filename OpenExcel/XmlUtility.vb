Imports DocumentFormat.OpenXml


Public Class XmlUtility

    Public Shared Function ToBooleanValue(ByVal value As Boolean?) As BooleanValue

        If value.HasValue Then Return New BooleanValue(value.Value)
        Return New BooleanValue
    End Function

    Public Shared Function ToDoubleValue(ByVal value As Double?) As DoubleValue

        If value.HasValue Then Return New DoubleValue(value.Value)
        Return New DoubleValue
    End Function

End Class
