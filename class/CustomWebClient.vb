Imports System.Net

Public Class CustomWebClient
    Inherits WebClient

    ' 自訂超時時間（以毫秒為單位）
    Public Property Timeout As Integer

    Public Sub New()
        ' 設定預設超時時間為 5 分鐘 (300,000 毫秒)
        Me.Timeout = 300000
    End Sub

    Protected Overrides Function GetWebRequest(ByVal address As Uri) As WebRequest
        ' 呼叫基底類別的方法來建立 WebRequest 物件
        Dim request As WebRequest = MyBase.GetWebRequest(address)

        ' 將 WebRequest 物件轉換為 HttpWebRequest 以存取 Timeout 屬性
        Dim httpWebRequest As HttpWebRequest = TryCast(request, HttpWebRequest)

        If httpWebRequest IsNot Nothing Then
            ' 設定超時時間
            httpWebRequest.Timeout = Me.Timeout
        End If

        Return request
    End Function
End Class
