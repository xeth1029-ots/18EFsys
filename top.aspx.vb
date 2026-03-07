Public Class top
    Inherits AuthBasePage

    Const cst_UserID_Login As String = "UserID_Login"
    Const cst_Secret_Login As String = "Secret_Login"
    Const cst_Secret_Login_URL As String = "Secret_Login_URL"
    Const cst_USERID_LOGINaspx As String = "USERID_LOGIN.aspx"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '不啟用https Not TIMS.Get_httpsProtocol 
        If Not TIMS.Get_httpsProtocol Then
            '不啟用https
            a_http.HRef = TIMS.cst_indexB1
        Else
            '啟用https
            a_http.HRef = TIMS.cst_indexA1
        End If

        'HyperLink1 切換計畫鈕
        HyperLink1.NavigateUrl = "Login.aspx" '切換計畫鈕
        HyperLink2.NavigateUrl = "sch.aspx" '功能搜尋
        If Not IsPostBack Then
            Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)

            If Not Session(cst_UserID_Login) Is Nothing Then
                If Session(cst_UserID_Login) = True Then
                    HyperLink1.NavigateUrl = cst_USERID_LOGINaspx
                    Exit Sub
                End If
            End If

            If Not Session(cst_Secret_Login) Is Nothing Then
                If Session(cst_Secret_Login) = True Then
                    Dim Secret_Login_URL As String = ""
                    Secret_Login_URL = ConfigurationSettings.AppSettings(cst_Secret_Login) 'Secret_Login_URL
                    If Session(cst_Secret_Login_URL) <> "" Then Secret_Login_URL = Session(cst_Secret_Login_URL)
                    If Secret_Login_URL <> "" Then Secret_Login_URL = Trim(Secret_Login_URL)
                    If Secret_Login_URL <> "" Then
                        HyperLink1.NavigateUrl = Secret_Login_URL ' Me.ViewState("SecretLogin")
                        Exit Sub
                    End If
                End If
            End If
        End If
    End Sub

End Class