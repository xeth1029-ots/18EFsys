Partial Class AppError
    Inherits UI.Page  ' 不能繼承 BasePage

    Private logger As ILog = LogManager.GetLogger(GetType(AppError))

    Public Ex As Exception = Nothing
    Public strStackTrace As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim objconn As SqlConnection = Nothing
        'Call TIMS.sUtl_404NOTFOUND(Me, objconn)
        'Return
        If Not IsPostBack Then cCreate1()
    End Sub

    Sub cCreate1()
        Try
            If Session("LastException") IsNot Nothing Then
                Ex = CType(Session("LastException"), Exception)
                If Request.IsLocal Then
                    strStackTrace = Ex.StackTrace.Replace(vbLf, "<br/>")
                    While Ex.InnerException IsNot Nothing
                        Ex = Ex.InnerException
                        If Ex.StackTrace Is Nothing Then Exit While
                        strStackTrace &= "<br/>"
                        strStackTrace &= Ex.StackTrace.Replace(vbLf, "<br/>")
                    End While
                End If
            End If
        Catch ex As Exception
            logger.Error(ex.Message, ex)
            Call TIMS.WriteTraceLog(ex.Message, ex) 'Throw ex
        End Try

        If Ex IsNot Nothing Then labExMessage.Text = Ex.Message
        labStackTrace.Text = strStackTrace
    End Sub

    Private Sub AppError_Error(sender As Object, e As EventArgs) Handles Me.[Error]
        ' Get last error from the server
        Dim exc As Exception = Server.GetLastError

        If Not IsNothing(exc) Then logger.Error(exc.Message, exc)
        If IsNothing(exc) Then logger.Error("AppError_Error: 沒有 Exception 訊息")

        ' 清除 Page Error 狀態, 確保 AppError 頁面不會發生遞迴導向的情況
        Server.ClearError()
    End Sub


End Class
