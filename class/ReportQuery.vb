Public Class ReportQuery

    ''' <summary>'取得流水號</summary>
    ''' <param name="page"></param>
    ''' <returns></returns>
    Public Shared Function GetGuid(ByVal page As Web.UI.Page) As String
        Dim cGuid As String = TIMS.GetGUID
        Return cGuid
    End Function

    ''' <summary>取得環境變數(Query) Url。</summary>
    ''' <param name="page"></param>
    ''' <returns></returns>
    Public Shared Function GetUrl(ByVal page As Web.UI.Page) As String
        Dim sUrl As String = TIMS.Utl_GetConfigSet("Query")
        Const cst_r1 As String = "report.do?"
        Const cst_r2 As String = "report?"
        Dim flag_IsIPV6 As Boolean = TIMS.IsIPV6()
        Return If(flag_IsIPV6 AndAlso sUrl.Contains(cst_r1), sUrl.Replace(cst_r1, cst_r2), sUrl)
    End Function

    Public Shared Function GetUrl2() As String
        Return TIMS.Utl_GetConfigSet("Query2")
    End Function

#Region "print iReport"

    ''' <summary>取1個亂數碼</summary>
    ''' <returns></returns>
    Public Shared Function Get_Rnd1() As Integer
        Dim iRst As Integer = 100000 * TIMS.Rnd1X() + 1 '0-100000 +1:1-100001
        Return iRst
    End Function

    Public Shared Function strWOScript(ByVal strWinOpen As String) As String
        Dim iPrintNum As Integer = Get_Rnd1()
        Dim sWindowAtt As String = "toolbar=0,location=0,status=0,menubar=0,resizable=1"
        Dim strScript As String = String.Concat("window.open('", strWinOpen, "','print", iPrintNum, "','", sWindowAtt, "');")
        Return strScript
    End Function

    Public Shared Function strWOScriptC(vWOS As String) As String
        Dim iPrintNum As Integer = Get_Rnd1()
        'titlebar=yes,toolbar=yes,location=yes,status=no,menubar=yes,scrollbars=yes,resizable=yes,width=700,Height=300,left=0,top=0
        'titlebar=0,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=7,Height=3,left=0,top=0
        Dim sWindowAtt As String = "titlebar=0,toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=7,Height=7,left=0,top=0"
        Dim strScript As String = String.Concat("window.open('", vWOS, "','prt", iPrintNum, "','", sWindowAtt, "');")
        Return strScript
    End Function

    Public Shared Function strWOScript1(ByVal strWinOpen As String) As String
        Dim iPrintNum As Integer = Get_Rnd1()
        Dim sWindowAtt As String = "toolbar=0,location=0,status=0,menubar=0,resizable=1"
        Dim strScript As String = "window.open('" & strWinOpen & "','print" & iPrintNum & "','" & sWindowAtt & "');" '不可換行
        Return strScript
    End Function

    ''' <summary> 列印報表 window.open </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="Filename"></param>
    ''' <param name="MyValue"></param>
    Public Shared Sub PrintReport(ByVal MyPage As Page, ByVal Filename As String, ByVal MyValue As String)
        Dim sm As SessionModel = SessionModel.Instance()
        Dim Url As String = ReportQuery.GetUrl(MyPage)
        'Dim strScript As String = ""

        '組合新字串，並確認轉換
        MyValue = GET_ENCODE1(MyPage, MyValue)

        Dim NewStr As String = MyValue

        If NewStr Is Nothing Then NewStr = ""
        '第1個位元加入 & 符號(應該不會執行此項)
        If NewStr.IndexOf("&") <> 0 AndAlso NewStr <> "" Then NewStr = String.Concat("&", NewStr)

        Dim strWinOpen As String = "" '組合連線字串。
        'strWinOpen = Url & "GUID=" & cGuid & "&RptID=" & Filename & NewStr
        strWinOpen = String.Concat(Url, "RptID=", Filename, NewStr)
        If strWinOpen.IndexOf("&UserID=") = -1 Then strWinOpen &= "&UserID=" & sm.UserInfo.UserID

        Try
            'SYS_HISPRINT
            Call TIMS.INSERT_HISPRINT(MyPage, strWinOpen)
        Catch ex As Exception
            Dim sErrMsg1 As String = ""
            sErrMsg1 &= "strWinOpen:" & strWinOpen
            Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
        End Try

        '有單引號 但 沒有斜線單引號 js error
        If strWinOpen.IndexOf("'") > -1 AndAlso strWinOpen.IndexOf("\'") = -1 Then strWinOpen = Replace(strWinOpen, "'", "\'")

        If TIMS.sUtl_ChkTest() Then Common.RespWrite(MyPage, String.Concat(strWinOpen, "<br />")) '測試用

        'Dim sFilenameX1 As String = TIMS.GetMyValue(strWinOpen, "RptID")
        If Filename = "" Then
            Dim strErrmsg As String = ""
            strErrmsg += "/* Public Shared Sub PrintReport(ByVal MyPage As Page, ByVal Filename As String, ByVal MyValue As String) */" & vbCrLf
            strErrmsg += " ERROR: NO Filename!!! " & vbCrLf
            strErrmsg += " strWinOpen: " & strWinOpen & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub 'Return ""
        End If

        Dim strScript As String = String.Concat("<script language=""javascript"">", strWOScript(strWinOpen), "</script>")
        MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Sub

    ''' <summary>組合無外框的報表資訊，可直接取得pdf</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="Filename"></param>
    ''' <param name="MyValue"></param>
    ''' <returns></returns>
    Public Shared Function GetReportUrl2(ByVal MyPage As Page, ByVal Filename As String, ByVal MyValue As String) As String
        Dim sm As SessionModel = SessionModel.Instance()
        Dim Url2 As String = GetUrl2()
        'Dim strScript As String = ""

        '組合新字串，並確認轉換
        MyValue = GET_ENCODE1(MyPage, MyValue)

        Dim NewStr As String = MyValue

        If NewStr Is Nothing Then NewStr = ""
        '第1個位元加入 & 符號(應該不會執行此項)
        If NewStr.IndexOf("&") <> 0 AndAlso NewStr <> "" Then NewStr = String.Concat("&", NewStr)

        Dim strWinOpen As String = "" '組合連線字串。

        strWinOpen = String.Concat(Url2, "RptID=", Filename, NewStr, If(strWinOpen.IndexOf("&UserID=") = -1, String.Concat("&UserID=", sm.UserInfo.UserID), ""))

        Try
            'SYS_HISPRINT
            Call TIMS.INSERT_HISPRINT(MyPage, strWinOpen)
        Catch ex As Exception
            Dim sErrMsg1 As String = ""
            sErrMsg1 &= "strWinOpen:" & strWinOpen
            Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
        End Try

        '有單引號 但 沒有斜線單引號 js error
        If strWinOpen.IndexOf("'") > -1 AndAlso strWinOpen.IndexOf("\'") = -1 Then strWinOpen = Replace(strWinOpen, "'", "\'")

        If TIMS.sUtl_ChkTest() Then Common.RespWrite(MyPage, String.Concat(strWinOpen, "<br />")) '測試用

        'Dim sFilenameX1 As String = TIMS.GetMyValue(strWinOpen, "RptID")
        If Filename = "" Then
            Dim strErrmsg As String = ""
            strErrmsg &= "/* Public Shared Sub PrintReport2(ByVal MyPage As Page, ByVal Filename As String, ByVal MyValue As String) */" & vbCrLf
            strErrmsg &= String.Concat(" ERROR: NO Filename!!! ", " strWinOpen: ", strWinOpen, vbCrLf)
            strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Return ""
        End If

        Return strWinOpen
        'strScript = String.Concat("<script language=""javascript"">", strWOScript(strWinOpen), "</script>")
        'MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Function

    ''' <summary> 組合新字串-UrlEncode </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="s_MyValue"></param>
    ''' <returns></returns>
    Public Shared Function GetNewStr(ByRef MyPage As Page, ByVal s_MyValue As String) As String
        Dim NewStr As String = ""
        If s_MyValue Is Nothing Then Return NewStr
        If s_MyValue.Length = 0 Then Return NewStr
        For i As Integer = 0 To s_MyValue.Length - 1
            NewStr &= If(AscW(s_MyValue.Chars(i)) > 127, MyPage.Server.UrlEncode(s_MyValue.Chars(i)), s_MyValue.Chars(i))
        Next
        Return NewStr
    End Function

    ''' <summary> 組合新字串，並整合參數! </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="s_MyValue"></param>
    ''' <returns></returns>
    Public Shared Function GET_ENCODE1(ByRef MyPage As Page, ByVal s_MyValue As String) As String
        If s_MyValue.Length <= 1 Then Return s_MyValue
        '組合新字串，並確認轉換
        Dim A_MyValue As String() = s_MyValue.Split("&")
        If A_MyValue.Length = 0 Then Return s_MyValue

        Dim rst As String = ""
        For Each sv1 As String In A_MyValue
            Dim s_tmp1 As String = ""
            If sv1.Split("=").Length > 1 Then
                s_tmp1 = GetNewStr(MyPage, sv1.Split("=")(0)) '參數名稱
                s_tmp1 &= "=" & GetNewStr(MyPage, sv1.Split("=")(1)) '參數值
            End If
            If s_tmp1 <> "" Then
                If rst <> "" Then rst &= "&"
                rst &= s_tmp1
            End If
        Next

        '第1個位元加入 & 符號
        If rst <> "" AndAlso rst.Chars(0) <> "&" Then rst = ("&" & rst)

        Return rst
    End Function

    ''' <summary>列印報表 window.open</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="Sys"></param>
    ''' <param name="Filename"></param>
    ''' <param name="MyValue"></param>
    Public Shared Sub PrintReport(ByRef MyPage As Page, ByVal Sys As String, ByVal Filename As String, ByVal MyValue As String)
        PrintReport(MyPage, Filename, MyValue)
    End Sub

    ''' <summary> 專用網頁轉址。 </summary>
    ''' <param name="MyPage"></param>
    Public Shared Sub Redirect(ByVal MyPage As Page)
        Dim sm As SessionModel = SessionModel.Instance()
        Dim sUrl As String = GetUrl(MyPage)
        'Dim NewStr As String = ""
        Dim MyUrl As String = MyPage.Request.Url.ToString
        Dim MyValue As String = Right(MyUrl, MyUrl.Length - MyUrl.IndexOf("?") - 1)
        Dim iPrintNum As Integer = Get_Rnd1()

        '組合新字串，並確認轉換
        MyValue = GET_ENCODE1(MyPage, MyValue)

        Const cst_filename As String = "filename="
        Const cst_RptID As String = "RptID=" '匯出Excel參數 Export=xls
        If MyValue.IndexOf(cst_filename) > -1 Then MyValue = Replace(MyValue, cst_filename, cst_RptID)

        Dim sRed As String = String.Concat(sUrl, "GUID=", iPrintNum, MyValue)

        If sRed.IndexOf("&UserID=") = -1 Then sRed &= "&UserID=" & sm.UserInfo.UserID

        Dim sFilenameX1 As String = TIMS.GetMyValue(MyValue, "RptID")
        If sFilenameX1 = "" Then
            Dim strErrmsg As String = ""
            strErrmsg += "/* Public Shared Sub Redirect(ByVal MyPage As Page) */" & vbCrLf
            strErrmsg += "/*  MyValue: */" & vbCrLf & MyValue & vbCrLf
            strErrmsg += " sRed: " & sRed & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End If

        Try
            Call TIMS.INSERT_HISPRINT(MyPage, sRed)
        Catch ex As Exception
            Dim sErrMsg1 As String = String.Concat("sRed:", sRed)
            Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
        End Try

        MyPage.Response.Redirect(TIMS.sUtl_AntiXss(sRed)) '正式
    End Sub

    ''' <summary> 組合列印報表 SCRIPT </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="Filename"></param>
    ''' <param name="MyValue"></param>
    ''' <param name="EndFlag"></param>
    ''' <returns></returns>
    Public Shared Function ReportScript(ByVal MyPage As Page,
                                        ByVal Filename As String,
                                        ByVal MyValue As String,
                                        ByVal EndFlag As Boolean) As String
        'Dim Str As String = ""
        'Dim NewStr As String = ""
        Dim sm As SessionModel = SessionModel.Instance()
        'Dim cGuid As String = SmartQuery.GetGuid(MyPage)
        Dim Url As String = ReportQuery.GetUrl(MyPage)

        '組合新字串，並確認轉換
        MyValue = GET_ENCODE1(MyPage, MyValue)

        Dim NewStr As String = MyValue
        If NewStr Is Nothing Then NewStr = ""
        '第1個位元加入 & 符號(應該不會執行此項)
        If NewStr.IndexOf("&") <> 0 AndAlso NewStr <> "" Then NewStr = "&" & NewStr

        Dim strWinOpen As String = "" '組合連線字串。
        'strWinOpen = Url & "GUID=" & cGuid & "&RptID=" & Filename & NewStr
        strWinOpen = Url & "RptID=" & Filename & NewStr
        If strWinOpen.IndexOf("&UserID=") = -1 Then strWinOpen &= "&UserID=" & sm.UserInfo.UserID

        'Dim sFilenameX1 As String = TIMS.GetMyValue(strWinOpen, "RptID")
        If Filename = "" Then
            Dim strErrmsg As String = ""
            strErrmsg += "/* public Shared Function ReportScript(ByVal MyPage As Page */" & vbCrLf
            strErrmsg += " strWinOpen: " & strWinOpen & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Return ""
        End If

        'Str &= "window.open('" & strWinOpen & "','print" & iPrintNum & "','" & sWindowAtt & "');"
        Return String.Concat(strWOScript1(strWinOpen), If(EndFlag, "return false;", ""))
    End Function

#End Region

    '組合列印報表 SCRIPT
    Public Shared Function ReportScript(ByVal MyPage As Page,
                                        ByVal Sys As String,
                                        ByVal Filename As String,
                                        ByVal MyValue As String,
                                        ByVal EndFlag As Boolean) As String
        Return ReportScript(MyPage, Filename, MyValue, EndFlag)
    End Function

    '組合列印報表 SCRIPT
    Public Shared Function ReportScript(ByVal MyPage As Page,
                                        ByVal Sys As String,
                                        ByVal Filename As String,
                                        ByVal MyValue As String) As String
        Dim EndFlag As Boolean = True
        Return ReportScript(MyPage, Filename, MyValue, EndFlag)
    End Function

    '組合列印報表 SCRIPT
    Public Shared Function ReportScript(ByVal MyPage As Page,
                                        ByVal Filename As String,
                                        ByVal MyValue As String) As String
        Dim EndFlag As Boolean = True
        Return ReportScript(MyPage, Filename, MyValue, EndFlag)
    End Function

    '(.VB 程式撰寫)報表轉址功能
    Public Shared Sub Redirect(ByVal MyPage As Page, ByVal sUrl As String, ByVal MyValue As String)
        Dim iPrintNum As Integer = Get_Rnd1()
        Dim strWinOpen As String = String.Concat(sUrl, "?", MyValue)

        Dim winopenMsg1 As String = $"window.open('{strWinOpen}','print{iPrintNum}','toolbar=0,location=0,status=0,menubar=0,resizable=1,scrollbars=1')"
        Dim strScript As String = $"<script type=""text/javascript"">{winopenMsg1}</script>{vbCrLf}"
        MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Sub

    'window.open(' 新視窗的網址 ', '新視窗的名稱', config='height=高度,width=寬度');

    Public Shared Sub Redirect(ByVal MyPage As Page, ByVal sUrl As String, ByVal MyValue As String, ByVal width As String, ByVal height As String)
        Dim iPrintNum As Integer = Get_Rnd1()
        Dim strWinOpen As String = String.Concat(sUrl, "?", MyValue)
        Dim s_winopen As String = String.Format(String.Concat("window.open('", strWinOpen, "','print", iPrintNum, "',config='width={0},height={1}');"), width, height)
        Dim strScript As String = $"<script type=""text/javascript"">{s_winopen}</script>{vbCrLf}"
        MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Sub

    '查無資料，關閉視窗window.close();
    Public Shared Sub CloseWin(ByVal MyPage As Page)
        Dim s_winclose As String = "alert('查無資料，關閉視窗');window.opener=null;window.open('','_self');window.close();"
        Dim strScript As String = $"<script type=""text/javascript"">{s_winclose}</script>{vbCrLf}"
        MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Sub

    '查無資料，關閉視窗window.close();
    Public Shared Sub CloseWin2(ByVal MyPage As Page)
        Dim strScript As String = String.Concat("<script type=""text/javascript"">", "window.opener=null;window.open('','_self');window.close();", "</script>")
        MyPage.RegisterStartupScript("PrintRepost", strScript)
    End Sub

    ''' <summary>取得pdf 匯出所要使用的 BaseUrl</summary>
    ''' <param name="MyPage"></param>
    ''' <returns></returns>
    Public Shared Function GetBaseUrl(MyPage As Page) As String
        Dim flag_test_ENVC As Boolean = TIMS.CHK_IS_TEST_ENVC() '檢測為測試環境:true 正式環境為:false
        Dim vScheme As String = MyPage.Request.Url.Scheme
        'Dim vDnsSafeHost As String = MyPage.Request.Url.DnsSafeHost
        Dim vPort As String = String.Concat("", MyPage.Request.Url.Port)
        If flag_test_ENVC Then
            Dim vLOCALADDR As String = TIMS.Get_LOCALADDR(MyPage, 1)
            Return String.Concat(vScheme, "://", vLOCALADDR, If(vPort <> "" AndAlso vPort <> "443", String.Concat(":", vPort), ""), "/")
        Else
            Dim vSERVER_NAME As String = TIMS.GET_SERVER_NAME(MyPage)
            Return String.Concat(vScheme, "://", vSERVER_NAME, If(vPort <> "" AndAlso vPort <> "443", String.Concat(":", vPort), ""), "/")
        End If
    End Function

    '' <summary>get HTTP_REFERER</summary>
    '' <param name="MyPage"></param>
    '' <returns></returns>
    'Friend Shared Function GET_HTTP_REFERER(MyPage As Page) As String
    '    Return MyPage.Request.ServerVariables("HTTP_REFERER")
    'End Function
End Class