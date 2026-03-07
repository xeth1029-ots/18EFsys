Imports System.Web.Security

Public Class AuthUtil

    Private Shared logger As ILog = LogManager.GetLogger(GetType(AuthUtil))

    ''' <summary>
    ''' 登入 記錄
    ''' </summary>
    ''' <param name="userId">帳號</param>
    ''' <param name="loginResult">設定登入狀態</param>
    Public Shared Sub LoginLog(ByVal userId As String, ByVal loginResult As Boolean)
        Dim sm As SessionModel = SessionModel.Instance()
        Dim message As String = sm.LastErrorMessage

        ' 因為 SessionModel.LastErrorMessage 一旦被讀取, 就會自動清除
        ' 所以要設回去
        sm.LastErrorMessage = message

        ' 登入記錄寫入DB ' 注意: 
        ' 要更新 auditInfo 的 Property, 須先將整個 AuditInfo 取出
        ' 更新完所有 Property 之後再整個設回去, 才會正確
        Dim auditInfo As AuditLogInfo = sm.AuditInfo

        auditInfo.LastFuncPath = "LOGIN"
        auditInfo.LastAccessTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
        auditInfo.LoginUserId = userId
        auditInfo.LoginMethod = "1"
        auditInfo.LoginResultMessage = message

        auditInfo.LoginResult = "0" '登入結果: 1.成功, 0.失敗
        'true:登入/false:登出
        If loginResult Then
            FormsAuthentication.SetAuthCookie(userId, False)
            auditInfo.LoginResult = "1" '登入結果: 1.成功, 0.失敗
        End If

        sm.AuditInfo = auditInfo
        sm.AuditInfo.WriteLoginLog()

        If Not loginResult Then
            '清除登入狀態
            sm.ClearSession()
        End If
    End Sub

    ''' <summary>
    ''' 登出 記錄
    ''' </summary>
    Public Shared Sub LogoutLog()
        Dim sm As SessionModel = SessionModel.Instance()

        If IsNothing(sm.UserInfo) Then
            logger.Info("LogoutLog: UserInfo NOT exists, Skip.")
            Return 'Exit Sub
        End If

        ' 登入記錄寫入DB ' 注意: 
        ' 要更新 auditInfo 的 Property, 須先將整個 AuditInfo 取出
        ' 更新完所有 Property 之後再整個設回去, 才會正確
        Dim auditInfo As AuditLogInfo = sm.AuditInfo

        auditInfo.LastFuncPath = "LOGOUT"
        auditInfo.LastAccessTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
        auditInfo.LoginUserId = sm.UserInfo.UserID
        auditInfo.LoginMethod = "1"

        auditInfo.LoginResult = ""
        auditInfo.LoginResultMessage = ""

        sm.AuditInfo = auditInfo
        sm.AuditInfo.WriteLoginLog()
    End Sub

End Class
