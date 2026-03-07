Imports Newtonsoft.Json
Public Class SchResult
    Public Property tf As Integer
    Public Property weight As Double
    Public Property key As String
End Class

Public Class Resp
    Public Property numFound As Integer
    Public Property start As Integer
    Public Property docs As List(Of SchResult)
End Class
Public Class ResParams
    Public Property q As String
    Public Property fl As String
    Public Property sort As String
    Public Property rows As String
End Class

Public Class RespHeader
    Public Property status As Integer
    Public Property QTime As Integer
    ' ... 其他 params 的屬性
    Public Property params As ResParams
End Class

Public Class SearchResponse
    Public Property responseHeader As RespHeader
    Public Property response As Resp
End Class