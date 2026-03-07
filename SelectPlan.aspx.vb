Imports System.Web.Mvc

Public Class SelectPlan
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        DbAccess.Open(objconn)

        'bt_back1.Visible = False
        '是否為超級使用者
        'flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)

        If Request("OP") = "Ajax" AndAlso Request("YR") <> "" Then
            ' Ajax 載入計畫清單
            Call ResponsePlanIDs(Request("YR"))
            Return
        End If

        Call InitYRList()
    End Sub

    ''' <summary>
    ''' 起始年度選單
    ''' </summary>
    Private Sub InitYRList()
        'Dim sql As String
        'Dim dt As DataTable

        YR.Items.Clear()

        Dim parms As New Hashtable From {{"account", sm.UserInfo.UserID}}

        Dim sql As String = ""
        If sm.UserInfo.LID = 2 Then
            sql &= " SELECT Distinct Years FROM VIEW_LOGINPLAN" & vbCrLf
            sql &= " WHERE PlanID IN (SELECT PlanID FROM Auth_AccRWPlan WHERE Account=@account )" & vbCrLf
            sql &= " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE 1=1 and DATEADD(month, 10, EDate) - getdate()>=0)" & vbCrLf
            sql &= " ORDER BY YEARS DESC"
            '計畫結束日加10個月帳號無法登入
        Else
            sql &= " SELECT Distinct Years FROM VIEW_LOGINPLAN" & vbCrLf
            sql &= " WHERE PlanID IN (SELECT PlanID FROM Auth_AccRWPlan WHERE Account=@account )" & vbCrLf
            sql &= " ORDER BY YEARS DESC"
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then
            sm.LastErrorMessage = "此帳號無計畫可登入!!"
            sm.RedirectUrlAfterBlock = ResolveUrl("~/login.aspx")
            sm.UserInfo = Nothing
            Exit Sub
        End If

        With YR
            .DataSource = dt
            .DataTextField = "Years"
            .DataValueField = "Years"
            .DataBind()
        End With
    End Sub

    ''' <summary> 供 Ajax 呼叫, 以年度查詢計劃別, 並回傳 Select Options (partial html) </summary>
    ''' <param name="year"></param>
    Private Sub ResponsePlanIDs(ByVal year As String)
        DbAccess.Open(objconn)

        Dim selTag As New TagBuilder("select")
        Dim sql As String = ""
        If sm.UserInfo.LID = 2 Then
            '非署(局)、非分署(中心)使用者鎖定結束時間、年度顯示計畫(結束時間10個月後，不可在登入該年度計畫。)
            sql &= " SELECT * FROM VIEW_LOGINPLAN" & vbCrLf
            sql &= " WHERE PlanID IN ( SELECT PlanID FROM ID_Plan WHERE Years=@year" & vbCrLf
            sql &= "  and DATEADD(month, 10, EDate) - getdate()>=0 )" & vbCrLf
            sql &= " and PlanID IN (SELECT PlanID FROM Auth_AccRWPlan WHERE Account=@account )" & vbCrLf
            sql &= " and (Clsyear is null or Clsyear > @year)" & vbCrLf
            sql &= " order by TPlanID,DistID,Seq,PlanID" & vbCrLf
        Else
            '署(局)、分署(中心)不管結束時間、年度顯示計畫
            sql &= " SELECT * FROM VIEW_LOGINPLAN" & vbCrLf
            sql &= " WHERE PlanID IN (SELECT PlanID FROM Auth_AccRWPlan WHERE Account=@account )" & vbCrLf
            sql &= " and (Clsyear is null or Clsyear > @year)" & vbCrLf
            sql &= " and Years=@year" & vbCrLf
            sql &= " order by TPlanID,DistID,Seq,PlanID" & vbCrLf
        End If

        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("year", SqlDbType.VarChar).Value = year '""
            .Parameters.Add("account", SqlDbType.VarChar).Value = sm.UserInfo.UserID '""
            dt.Load(.ExecuteReader())
        End With
        Dim hasData As Boolean = False
        If Not IsNothing(dt) Then
            'Dim dr As DataRow
            For Each dr As DataRow In dt.Rows
                hasData = True
                Dim optTag As TagBuilder = New TagBuilder("option")
                optTag.Attributes.Add("value", dr("PlanID"))
                optTag.InnerHtml = dr("PlanName")
                If sm.UserInfo.TPlanID = Convert.ToString(dr("TPLANID")) Then optTag.Attributes.Add("selected", "selected")
                selTag.InnerHtml += optTag.ToString()
            Next
        End If

        If Not hasData Then
            Dim optTag As TagBuilder = New TagBuilder("option")
            optTag.Attributes.Add("value", "")
            optTag.InnerHtml = "沒有適用的計畫"
            selTag.InnerHtml += optTag.ToString()
        End If

        Response.Clear()
        Response.Write(selTag.InnerHtml)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    ''' <summary>
    ''' 送出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
        ' 因為PLANID 選項, 在前端是透過 Ajax 動能載入, 所以 ViewState 中沒有值
        ' 無法透過 PLANID.SelectedValue 取得選取值
        Dim strPlanID As String = Request.Form(PLANID.UniqueID)

        If strPlanID = "" Then
            sm.LastErrorMessage = "請選擇計畫"
            Exit Sub
        End If

        Dim flag_SetPlanOK As Boolean = SetPlan(objconn, Me.YR.SelectedValue, strPlanID, True)
        If flag_SetPlanOK Then
            ' 選擇計畫成功, 導向登入後首頁
            DbAccess.CloseDbConn(objconn)
            Response.Redirect(ResolveUrl("~/Index"))
        Else
            '失敗, 停留在本頁
            'bt_back1.Visible = True
            sm.LastErrorMessage = "帳號權限資料有誤，請連絡系統管者修正資料!"
            'sm.RedirectUrlAfterBlock = ResolveUrl("~/login.aspx")
            'sm.UserInfo = Nothing
            Exit Sub
        End If

    End Sub

    ''' <summary> 設定當前登入的年度及計畫, 並載入對應的功能權限清單
    ''' <para>回傳: true.設定成功, false.設定失敗, 並將訊息寫入 sm.LastErrorMessage</para>
    ''' </summary>
    ''' <param name="YR"></param>
    ''' <param name="PlanID"></param>
    ''' <param name="flag_UpdateDefault">是否一併更新個人的預設登入年度及計畫</param>
    Public Shared Function SetPlan(ByVal objconn As SqlConnection, ByVal YR As String, ByVal PlanID As String, ByVal flag_UpdateDefault As Boolean) As Boolean
        'Optional ByVal flag_UpdateDefault As Boolean = False
        Dim sm As SessionModel = SessionModel.Instance()
        'Dim objconn As SqlConnection = DbAccess.GetConnection()

        Call TIMS.OpenDbConn(objconn)
        '檢查對應的業務權限
        Dim parms As New Hashtable From {{"Account", sm.UserInfo.UserID}, {"PlanID", PlanID}}
        Dim sql As String = ""
        sql &= " select b.PlanID,e.TPlanID" & vbCrLf
        sql &= " ,e.Years+f.Name+g.PlanName+e.Seq PlanName" & vbCrLf
        sql &= " ,e.Years" & vbCrLf
        sql &= " ,c.RID,c.RELSHIP, c.ORGLEVEL, c.DISTID" & vbCrLf
        sql &= " from Auth_Account a " & vbCrLf
        sql &= " JOIN Auth_AccRWPlan b on a.account=b.account" & vbCrLf
        sql &= " JOIN Auth_Relship c on b.RID=c.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
        sql &= " JOIN ID_Plan e on b.PlanID=e.PlanID" & vbCrLf
        sql &= " JOIN ID_District f on e.DistID=f.DistID" & vbCrLf
        sql &= " JOIN Key_Plan g on e.TPlanID=g.TPlanID" & vbCrLf
        sql &= " where a.IsUsed = 'Y' and a.Account=@Account and b.PlanID=@PlanID" & vbCrLf
        Dim Testdt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Dim result As Boolean = False

        If Testdt.Rows.Count = 1 Then
            Dim drAP As DataRow = Testdt.Rows(0)

            If Convert.ToString(drAP("TPlanID")) <> "" Then
                '計算登入次數' Plan_Count
                TIMS.GetLoginCount(Convert.ToString(drAP("RID")), CInt(PlanID), objconn)
                Dim goNext As Boolean = True '步驟繼續 True
                If flag_UpdateDefault Then
                    '將選擇的 年度/計畫 保存到DB中
                    Dim sql_u As String = "" & vbCrLf
                    sql_u &= " UPDATE AUTH_ACCOUNT" & vbCrLf
                    sql_u &= " SET DEFAULT_YEAR=@YEAR,DEFAULT_PLANID=@PLANID,LAST_LOGINDATE=GETDATE()" & vbCrLf
                    sql_u &= " WHERE ACCOUNT=@ACCOUNT AND ISUSED='Y'" & vbCrLf
                    'YR
                    Dim parms_u As New Hashtable From {{"ACCOUNT", sm.UserInfo.UserID}, {"YEAR", Convert.ToString(drAP("Years"))}, {"PLANID", CInt(PlanID)}}
                    Dim rtn As Integer = DbAccess.ExecuteNonQuery(sql_u, objconn, parms_u)
                    If rtn <> 1 Then
                        goNext = False '(異常步驟停止)
                        sm.LastErrorMessage = "保存選擇的 年度/計畫 失敗"
                        Return result
                    End If
                End If

                '未發生問題(步驟繼續)
                If goNext Then
                    Dim s_ACCOUNT As String = sm.UserInfo.UserID
                    Dim drAA As DataRow = TIMS.sUtl_GetAccount(s_ACCOUNT, objconn, False)
                    If drAA Is Nothing Then
                        sm.LastErrorMessage = "帳號權限資料設定錯誤，請連絡系統管者修正資料!!!"
                        Return result
                    End If

                    ' 帳密登入驗證成功
                    Dim userInfo As New LoginUserInfo With {.LoginSuccess = True}

                    ' 將選擇的 年度/計畫 更新到在 SessionModel 
                    TIMS.SET_SESSIONMODEL2(sm, userInfo, drAA, drAP, objconn)

                    TIMS.LOG.Info("SelectPlan - User Logined:" & vbCrLf & userInfo.ToString())
                    'AuthUtil.LoginLog(s_ACCOUNT, True)

                    result = True
                End If
            Else
                sm.LastErrorMessage = "帳號權限資料設定錯誤，請連絡系統管者修正資料!!"
            End If
        ElseIf Testdt.Rows.Count = 0 Then
            sm.LastErrorMessage = "帳號權限資料設定錯誤，請連絡系統管者修正資料!"
        Else
            sm.LastErrorMessage = "該使用者登入此計畫有多種業務權限，請連絡系統管者修正資料!"
        End If

        Return result
    End Function

    'Protected Sub bt_back1_Click(sender As Object, e As EventArgs) Handles bt_back1.Click
    '    Dim redirectUrl As String = ResolveUrl("~/login")
    '    Response.Redirect(redirectUrl)
    'End Sub
End Class