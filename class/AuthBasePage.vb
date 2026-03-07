''' <summary>
''' 所有需要登入控管的頁面, 一律繼承這個基底類
''' </summary>
Public Class AuthBasePage
    Inherits BasePage

    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    ''' <summary>
    ''' 系統登入頁
    ''' </summary>
    Protected LOGIN_PAGE As String = ResolveUrl("~/MOICA_login")

    ''' <summary>
    ''' 統一在這裡檢查使用者登入 Session 狀態
    ''' </summary>
    ''' <param name="e"></param>
    Protected NotOverridable Overrides Sub OnPreInit(e As EventArgs)
        MyBase.OnPreInit(e)

        If Not sm.IsLogin Then Call TIMS.Chk_TEST_Login2()

        '檢查使用者登入狀態資訊
        If Not sm.IsLogin Then
            ' 沒有登入, 導向登入頁面
            logger.Debug($"{Request.Path} Login Required, Redirect to LOGIN_PAGE")
            sm.LastErrorMessage = "您尚未登入或登入資訊已經遺失，請重新登入!"

            Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
            If flag_test Then LOGIN_PAGE = ResolveUrl("~/login")

            Response.Redirect(LOGIN_PAGE)
            Response.End()
            Return
        End If
    End Sub

    Protected NotOverridable Overrides Sub OnInit(e As EventArgs)
        MyBase.OnInit(e)
    End Sub

End Class