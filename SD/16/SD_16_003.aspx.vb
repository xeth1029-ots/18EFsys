Partial Class SD_16_003
    Inherits AuthBasePage

    'CLASS_RETRO /AUTH_RENDCLASS
    'select * from CLASS_RETRO where rownum <=10 and modifydate>=DATEADD(DAY, getdate(), -100)
    Dim DGobj1 As DataGrid = Nothing
    Const cst_DGobj1N As String = "DG_ClassInfo"
    Const cst_審核通過 As String = "Y"
    Const cst_審核失敗 As String = "N"

    Const cst_errmsg1 As String = "資料儲存時發生錯誤，請重新查詢!!"

    Dim aDate As String = ""
    Dim dtDist As DataTable = Nothing

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        DGobj1 = Me.FindControl(cst_DGobj1N)
        PageControler1.PageDataGrid = DGobj1

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Exit Sub
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Exit Sub
        '    End If
        'End If

        aDate = TIMS.GetSysDate(objconn)
        Dim sql As String = ""
        sql = "SELECT DISTID,NAME FROM ID_DISTRICT WHERE 1=1 AND DISTID NOT IN ('000','002') ORDER BY DISTID"
        dtDist = DbAccess.GetDataTable(sql, objconn)

        If Not IsPostBack Then
            Call Create1()
            btnSaveY1.Attributes("onclick") = "javascript:return CheckSave1();"
        End If

    End Sub

    '第1次載入
    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'DG
        labmsg.Text = ""
        tbSearch1.Visible = False
        DGobj1.Visible = False

        '搜尋區
        sddlYears = TIMS.Get_Years(sddlYears, objconn)
        Common.SetListItem(sddlYears, sm.UserInfo.Years)
        sddlDistID = TIMS.Get_DistID(sddlDistID, dtDist)
        sddlDistID.Enabled = True
        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
                Call Makeplanlist(sddlPlan, sddlYears.SelectedValue, sddlDistID.SelectedValue, sddlOrgName, Nothing, Nothing)
            Case "1"
                If sm.UserInfo.DistID <> "000" Then
                    '分署限定
                    sddlDistID.Enabled = False
                    Common.SetListItem(sddlDistID, sm.UserInfo.DistID)
                End If
                Call Makeplanlist(sddlPlan, sddlYears.SelectedValue, sddlDistID.SelectedValue, sddlOrgName, Nothing, Nothing)
            Case "2"
                sddlYears.Enabled = False
                Common.SetListItem(sddlYears, sm.UserInfo.Years)
                sddlDistID.Enabled = False
                Common.SetListItem(sddlDistID, sm.UserInfo.DistID)
                Call Makeplanlist(sddlPlan, sddlYears.SelectedValue, sddlDistID.SelectedValue, sddlOrgName, Nothing, Nothing)
                sddlPlan.Enabled = False
                Common.SetListItem(sddlPlan, sm.UserInfo.PlanID)
                trYD1.Disabled = True  '搜尋區
                trPlan1.Disabled = True  '搜尋區
                trYD2.Disabled = True  '編輯區
                trPlan2.Disabled = True  '編輯區
                '補助地方政府限定
                If TIMS.ChkRelship23(Me, objconn) Then
                    Call MakeddlOrgName(sddlOrgName, sddlPlan.SelectedValue, sm.UserInfo.OrgID, sddlDistID.SelectedValue, Nothing, Nothing)
                End If
        End Select

        '編輯區
        ddlYears = TIMS.Get_Years(ddlYears, objconn)
        Common.SetListItem(ddlYears, sm.UserInfo.Years)
        ddlDistID = TIMS.Get_DistID(ddlDistID, dtDist)
        ddlDistID.Enabled = True
        If sm.UserInfo.DistID <> "000" Then
            ddlDistID.Enabled = False
            Common.SetListItem(ddlDistID, sm.UserInfo.DistID)
        End If
        Call Makeplanlist(ddlPlan, ddlYears.SelectedValue, ddlDistID.SelectedValue, ddlOrgName, ddlClassCName, ddlAccount)

        ddlReasonID = TIMS.Get_ReasonID(ddlReasonID, objconn)
        cb_SelFunID = TIMS.Get_FunIDReUse(cb_SelFunID, objconn, "")
    End Sub

    '計畫(依年度轄區 清除機構與班級)
    Sub Makeplanlist(ByRef ddlobj As DropDownList, ByVal Years As String, ByVal DistID As String, _
                     ByRef ddlOrgObj As DropDownList, ByRef ddlClassObj As DropDownList, ByRef ddlAcctObj As DropDownList)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select distinct a.PlanID" & vbCrLf
        sql &= " ,a.Years+b.Name+c.PlanName+a.seq PlanName" & vbCrLf
        sql &= " ,a.DistID" & vbCrLf
        sql &= " from ID_Plan a" & vbCrLf
        sql &= " JOIN ID_District b on a.DistID=b.DistID" & vbCrLf
        sql &= " JOIN Key_Plan c on a.TPlanID=c.TPlanID" & vbCrLf
        sql &= " JOIN Auth_AccRWPlan d on a.PlanID=d.PlanID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.years = '" & Years & "' " & vbCrLf
        sql &= " and a.DistID = '" & DistID & "'" & vbCrLf
        sql &= " order by 2 " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn) 'PlanID/PlanName

        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        If Not ddlOrgObj Is Nothing Then ddlOrgObj.Items.Clear()
        If Not ddlClassObj Is Nothing Then ddlClassObj.Items.Clear()
        If Not ddlAcctObj Is Nothing Then ddlAcctObj.Items.Clear()
    End Sub

    '選擇計畫時 顯示機構(依PlanID,登入者OrgID 清除班級)
    Sub MakeddlOrgName(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal OrgID As String, _
                       ByVal DistID As String, ByRef ddlClassObj As DropDownList, ByRef ddlAcctObj As DropDownList)

        Dim sLID As String = Convert.ToString(sm.UserInfo.LID)
        Dim sRID As String = Convert.ToString(sm.UserInfo.RID)

        'Sql &= " SELECT DISTINCT oo.OrgID" & vbCrLf
        ''(含)補助地方政府
        'sql &= " SELECT DISTINCT case when r23.RID3 IS NOT NULL then oo.OrgName" & vbCrLf
        'sql &= "  else r23.RID3+'-'+r23.OrgName3 end OrgName" & vbCrLf
        'sql &= " ,case when r23.OrgName3 IS NOT NULL then r23.OrgName2+'-'+r23.OrgName3" & vbCrLf
        'sql &= "  else oo.OrgName end OrgName" & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        '(含)補助地方政府
        sql &= " SELECT DISTINCT case when r3.RID3 is not null then r3.RID3+'" & TIMS.cst_spFlag1 & "'+CONVERT(varchar, oo.OrgID)" & vbCrLf
        sql &= " else CONVERT(varchar, oo.OrgID) end OrgID" & vbCrLf
        sql &= " ,case when r3.orgname2 is not null then r3.orgname2+'-'+r3.orgname3" & vbCrLf
        sql &= " else oo.OrgName end OrgName" & vbCrLf
        sql &= " ,c.OrgLevel" & vbCrLf
        sql &= " ,c.RID " & vbCrLf
        sql &= " From Auth_Account a " & vbCrLf
        sql &= " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        sql &= " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        sql &= " JOIN Org_Orginfo oo on oo.OrgID =c.OrgID" & vbCrLf
        sql &= " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID " & vbCrLf
        '補助地方政府
        sql &= " LEFT JOIN VIEW_RELSHIP23 r3 ON r3.RID3=c.RID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and a.IsUsed='Y' " & vbCrLf
        Select Case sLID 'Convert.ToString(sm.UserInfo.LID)
            Case "0" '署
            Case "1" '分署
                sql &= " and a.LID>='" & sLID & "'" & vbCrLf
            Case "2" '補助地方政府
                sql &= " and a.LID>='" & sLID & "'" & vbCrLf
                sql &= " and c.OrgLEVEL='3'" & vbCrLf
                sql &= " and r3.RID2='" & sRID & "'" & vbCrLf
            Case Else '異常
                sql &= " and 1<>1 " & vbCrLf
        End Select
        'sql += " and c.OrgID ='" & OrgID & "'" & vbCrLf
        sql &= " and b.PlanID = '" & PlanID & "' " & vbCrLf
        sql &= " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        'sql += " order by c.OrgLevel,oo.OrgName,c.RID " & vbCrLf
        sql &= " order by c.OrgLevel, OrgName,c.RID " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)

        If ddlobj.Items.Count = 0 Then
            'Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " SELECT DISTINCT case when r3.RID3 is not null then r3.RID3+'-'+CONVERT(varchar, oo.OrgID)" & vbCrLf
            sql &= " else CONVERT(varchar, oo.OrgID) end OrgID" & vbCrLf
            sql &= " ,case when r3.orgname2 is not null then r3.orgname2+'-'+r3.orgname3" & vbCrLf
            sql &= " else oo.OrgName end OrgName" & vbCrLf
            sql &= " FROM AUTH_RELSHIP r1" & vbCrLf
            sql &= " JOIN ORG_ORGINFO oo on oo.OrgID=r1.OrgID" & vbCrLf
            sql &= " LEFT JOIN VIEW_RELSHIP23 r3 ON r3.RID3=r1.RID" & vbCrLf
            sql &= " where 1=1" & vbCrLf
            sql &= " and r1.RID='" & sRID & "'" & vbCrLf
            If OrgID <> "" Then
                sql &= " and oo.OrgID='" & OrgID & "'" & vbCrLf
            End If
            DbAccess.MakeListItem(ddlobj, sql, objconn)
        End If

        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'ddlobj.Items.Insert(0, New ListItem("全部", ""))
        If Not ddlClassObj Is Nothing Then ddlClassObj.Items.Clear()
        If Not ddlAcctObj Is Nothing Then ddlAcctObj.Items.Clear()
    End Sub

    '帳號
    Sub MakeAccount(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String, _
                    ByVal OrgID3 As String)
        '含有RID資訊 
        Dim RID3 As String = ""
        Call TIMS.sUtl_AnalysisOrgID(OrgID3, RID3)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT a.Account" & vbCrLf
        sql &= " ,a.Name+'('+d.Name+')' sName" & vbCrLf
        sql &= " ,a.RoleID" & vbCrLf
        sql &= " ,a.LID" & vbCrLf
        sql &= " From Auth_Account a" & vbCrLf
        sql &= " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        sql &= " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        sql &= " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID" & vbCrLf
        sql &= " LEFT JOIN ID_Role d ON a.RoleID=d.RoleID" & vbCrLf
        sql &= " Where 1=1" & vbCrLf
        sql &= " and a.IsUsed='Y' " & vbCrLf
        If LID <> "" Then
            sql &= " and a.LID>='" & LID & "' " & vbCrLf
        End If
        If RoleID <> "" Then
            sql &= " and a.RoleID>='" & RoleID & "' " & vbCrLf
        End If
        sql &= " and b.PlanID = '" & PlanID & "' " & vbCrLf
        sql &= " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        If OrgID3 <> "" Then
            sql &= " and c.OrgID = '" & OrgID3 & "' " & vbCrLf
        Else
            sql &= " and 1<>1 " & vbCrLf
        End If
        sql &= " order by a.RoleID, sName" & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '查詢班級
    Sub MakeClassName(ByRef ddlobj As DropDownList, _
                      ByVal PlanID As String, ByVal OrgID3 As String, ByVal DistID As String)
        If OrgID3 = "" Then OrgID3 = "0"
        If PlanID = "" Then PlanID = sm.UserInfo.PlanID

        '含有RID資訊 
        Dim RID3 As String = ""
        Call TIMS.sUtl_AnalysisOrgID(OrgID3, RID3)

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT a.OCID" & vbCrLf
        sqlstr &= " ,a.ClassCName+'(第'+a.CyclType+'期)' ClassCName" & vbCrLf
        'sqlstr &= " ,e.OrgName" & vbCrLf
        sqlstr &= " ,case when r3.orgname2 is not null then r3.orgname2+'-'+r3.orgname3" & vbCrLf
        sqlstr &= " else e.OrgName end OrgName" & vbCrLf

        sqlstr &= " ,a.Years" & vbCrLf
        sqlstr &= " ,a.CyclType" & vbCrLf
        sqlstr &= " ,a.ClassNum" & vbCrLf
        sqlstr &= " ,b.ClassID" & vbCrLf
        sqlstr &= " ,a.PlanID" & vbCrLf
        sqlstr &= " ,g.TrainName" & vbCrLf
        sqlstr &= " ,CONVERT(varchar, a.STDate, 111) STDate" & vbCrLf
        sqlstr &= " ,CONVERT(varchar, a.FTDate, 111) FTDate" & vbCrLf
        sqlstr &= " ,a.RID" & vbCrLf
        sqlstr &= " ,a.Years + '0' + b.ClassID + a.CyclType ClassID2" & vbCrLf
        sqlstr &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sqlstr &= " JOIN ID_PLAN IP ON IP.PLANID =a.PLANID" & vbCrLf
        sqlstr &= " JOIN ID_CLASS b ON b.CLSID = a.CLSID" & vbCrLf
        sqlstr &= " JOIN ID_DISTRICT C ON C.DISTID=IP.DISTID" & vbCrLf
        sqlstr &= " JOIN Auth_Relship d on d.RID=a.RID" & vbCrLf
        sqlstr &= " JOIN Org_OrgInfo e on e.OrgID=d.OrgID" & vbCrLf
        sqlstr &= " JOIN Key_TrainType g on g.TMID=a.TMID" & vbCrLf
        sqlstr &= " LEFT JOIN MVIEW_RELSHIP23 r3 on r3.RID3=a.RID and r3.PlanID=a.PlanID" & vbCrLf

        sqlstr &= " Where 1=1 " & vbCrLf
        sqlstr &= " and a.IsSuccess='Y'" & vbCrLf '是否轉入成功
        sqlstr &= " and a.NotOpen='N' " & vbCrLf  '不開班
        sqlstr &= " and ip.PlanID = '" & PlanID & "' " & vbCrLf
        sqlstr &= " and d.OrgID = '" & OrgID3 & "' " & vbCrLf
        sqlstr &= " and ip.DistID = '" & DistID & "'" & vbCrLf
        If RID3 <> "" Then
            sqlstr &= " and a.RID='" & RID3 & "'" & vbCrLf 'RID
        End If
        sqlstr &= " ORDER BY 2,a.OCID" & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sqlstr, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'ddlobj.Items.Insert(0, New ListItem("全部", ""))
        'If Not ddlClassObj Is Nothing Then ddlClassObj.Items.Clear()
    End Sub

    '清除輸入值。
    Sub ClearValue1()
        Hid_RETID.Value = ""
        labAPPLYDATE.Text = "" '
        'ddlYears.SelectedIndex = -1
        'ddlDistID.SelectedIndex = -1
        'ddlPlan.SelectedIndex = -1
        'ddlOrgName.SelectedIndex = -1
        'ddlClassCName.SelectedIndex = -1

        'Account.SelectedIndex = -1
        Common.SetListItem(ddlYears, Convert.ToString(sm.UserInfo.Years))
        Common.SetListItem(ddlDistID, Convert.ToString(sm.UserInfo.DistID))
        Call Makeplanlist(ddlPlan, ddlYears.SelectedValue, ddlDistID.SelectedValue, ddlOrgName, ddlClassCName, ddlAccount)
        'Common.SetListItem(ddlPlan, Convert.ToString(sm.UserInfo.PlanID))

        ddlReasonID.SelectedIndex = -1
        txtReason.Text = ""
        'For i As Integer = 0 To cb_SelFunID.Items.Count - 1
        '    cb_SelFunID.Items(i).Selected = False
        'Next
        Call TIMS.SetCblValue(cb_SelFunID, "")
        EndDate.Text = ""
        '是否歸責單位
        'Common.SetListItem(rblBlameUnit, "Y")
        rblBlameUnit.SelectedIndex = -1
    End Sub

    '關閉所有顯示
    Sub CloseList()
        Panelsch1.Visible = False
        Paneledit1.Visible = False
        btnSaveY1.Visible = False
        btnSaveN1.Visible = False
        'Panelview1.Visible = False
    End Sub

    'SQL 查詢sub
    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DGobj1) '顯示列數不正確

        Dim OrgID3 As String = sddlOrgName.SelectedValue
        Dim RID3 As String = ""
        Call TIMS.sUtl_AnalysisOrgID(OrgID3, RID3)

        sClassName.Text = TIMS.ClearSQM(sClassName.Text)
        sCyclType.Text = TIMS.ClearSQM(sCyclType.Text)
        sAPPLYDATE1.Text = TIMS.ClearSQM(sAPPLYDATE1.Text)
        sAPPLYDATE2.Text = TIMS.ClearSQM(sAPPLYDATE2.Text)
        sENDDATE1.Text = TIMS.ClearSQM(sENDDATE1.Text)
        sENDDATE2.Text = TIMS.ClearSQM(sENDDATE2.Text)
        sAPPLYDATE1.Text = TIMS.Cdate3(sAPPLYDATE1.Text)
        sAPPLYDATE2.Text = TIMS.Cdate3(sAPPLYDATE2.Text)
        sENDDATE1.Text = TIMS.Cdate3(sENDDATE1.Text)
        sENDDATE2.Text = TIMS.Cdate3(sENDDATE2.Text)
        If sCyclType.Text <> "" Then
            sCyclType.Text = Val(sCyclType.Text)
            sCyclType.Text = TIMS.AddZero(sCyclType.Text, 2)
        End If

        ViewState("sddlYears") = sddlYears.SelectedValue
        ViewState("sddlDistID") = sddlDistID.SelectedValue
        ViewState("sddlPlan") = sddlPlan.SelectedValue
        ViewState("sddlOrgName") = OrgID3 'sddlOrgName.SelectedValue
        ViewState("sClassName") = sClassName.Text
        ViewState("sCyclType") = sCyclType.Text
        ViewState("sAPPLYDATE1") = sAPPLYDATE1.Text
        ViewState("sAPPLYDATE2") = sAPPLYDATE2.Text
        ViewState("sENDDATE1") = sENDDATE1.Text
        ViewState("sENDDATE2") = sENDDATE2.Text

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.RETID" & vbCrLf '/*PK*/ 
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,a.PLANID" & vbCrLf
        'sql &= " ,a.ORGID" & vbCrLf
        'sql &= " ,oo.OrgName" & vbCrLf
        sql &= " ,case when r3.RID3 is not null then r3.RID3+'" & TIMS.cst_spFlag1 & "'+CONVERT(varchar, a.OrgID)" & vbCrLf
        sql &= " else CONVERT(varchar, a.OrgID) end ORGID" & vbCrLf
        sql &= " ,case when r3.orgname2 is not null then r3.orgname2+'-'+r3.orgname3" & vbCrLf
        sql &= " else oo.OrgName end OrgName" & vbCrLf

        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,a.APPLYACCT" & vbCrLf
        sql &= " ,CONVERT(varchar, a.APPLYDATE, 111) APPLYDATE" & vbCrLf
        sql &= " ,a.FUNID" & vbCrLf
        'sql &= " ,' ' FUNIDN" & vbCrLf

        sql &= " ,CONVERT(varchar, a.ENDDATE, 111) ENDDATE" & vbCrLf
        sql &= " ,a.REASONID" & vbCrLf
        sql &= " ,a.REASON" & vbCrLf
        'sql += " ,a.USEABLE" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        'sql += " ,a.CREATEACCT" & vbCrLf
        'sql += " ,a.CREATEDATE" & vbCrLf
        'sql += " ,a.MODIFYACCT" & vbCrLf
        'sql += " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.APPLIEDRESULT" & vbCrLf
        sql &= " ,a.RESULTACCT" & vbCrLf
        sql &= " ,a.RESULTDATE" & vbCrLf
        sql &= " ,a.RIGHTID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.CYCLTYPE" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf

        sql &= " ,cc.Years+'0'+b.ClassID+cc.CyclType CLASSID2" & vbCrLf
        sql &= " ,aa.NAME ACCTNAME" & vbCrLf
        sql &= " ,dbo.DECODE6(a.APPLIEDRESULT,'Y','通過','N','不通過','待審') APPLIEDRESULT2" & vbCrLf
        sql &= " ,dbo.DECODE6(a.APPLIEDRESULT,'Y','通過','N','失敗','') APPLIEDRESULT3" & vbCrLf
        sql &= " ,dbo.DECODE6(a.BlameUnit,'Y','是','N','否','') BlameUnitN" & vbCrLf

        sql &= " FROM CLASS_RETRO a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=a.OCID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.OrgID=a.OrgID" & vbCrLf
        sql &= " JOIN ID_CLASS b ON b.CLSID=cc.CLSID" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT aa on aa.ACCOUNT=a.APPLYACCT" & vbCrLf
        sql &= " LEFT JOIN MVIEW_RELSHIP23 r3 on r3.RID3=cc.RID and r3.PlanID =cc.PlanID" & vbCrLf

        sql &= " WHERE 1=1" & vbCrLf
        If ViewState("sddlYears") <> "" Then
            sql &= " AND a.YEARS=@YEARS" & vbCrLf
        End If
        If ViewState("sddlDistID") <> "" Then
            sql &= " AND a.DistID=@DistID" & vbCrLf
        End If
        If ViewState("sddlPlan") <> "" Then
            sql &= " AND a.PlanID=@PlanID" & vbCrLf
        End If
        If ViewState("sddlOrgName") <> "" Then
            sql &= " AND a.ORGID=@ORGID" & vbCrLf
        End If
        If RID3 <> "" Then
            sql &= " AND cc.RID=@RID3" & vbCrLf
        End If

        If ViewState("sClassName") <> "" Then
            sql &= " AND cc.CLASSCNAME like '%'+@CLASSCNAME+'%'" & vbCrLf
        End If
        If ViewState("sCyclType") <> "" Then
            sql &= " AND cc.CYCLTYPE=@CYCLTYPE" & vbCrLf
        End If
        If ViewState("sAPPLYDATE1") <> "" Then
            sql &= " AND a.APPLYDATE >=@APPLYDATE1" & vbCrLf
        End If
        If ViewState("sAPPLYDATE2") <> "" Then
            sql &= " AND a.APPLYDATE <=@APPLYDATE2" & vbCrLf
        End If
        If ViewState("sENDDATE1") <> "" Then
            sql &= " AND a.ENDDATE >=@ENDDATE1" & vbCrLf
        End If
        If ViewState("sENDDATE2") <> "" Then
            sql &= " AND a.ENDDATE <=@ENDDATE2" & vbCrLf
        End If

        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            If ViewState("sddlYears") <> "" Then
                'sql += " AND a.YEARS=@YEARS" & vbCrLf
                .Parameters.Add("YEARS", SqlDbType.VarChar).Value = ViewState("sddlYears")
            End If
            If ViewState("sddlDistID") <> "" Then
                'sql += " AND a.DistID=@DistID" & vbCrLf
                .Parameters.Add("DistID", SqlDbType.VarChar).Value = ViewState("sddlDistID")
            End If
            If ViewState("sddlPlan") <> "" Then
                'sql += " AND a.PlanID=@PlanID" & vbCrLf
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = ViewState("sddlPlan")
            End If
            If ViewState("sddlOrgName") <> "" Then
                'sql += " AND a.ORGID=@ORGID" & vbCrLf
                .Parameters.Add("ORGID", SqlDbType.VarChar).Value = ViewState("sddlOrgName")
            End If
            If RID3 <> "" Then
                .Parameters.Add("RID3", SqlDbType.VarChar).Value = RID3
            End If

            If ViewState("sClassName") <> "" Then
                'sql += " AND cc.CLASSCNAME=@CLASSCNAME" & vbCrLf
                .Parameters.Add("CLASSCNAME", SqlDbType.VarChar).Value = ViewState("sClassName")
            End If
            If ViewState("sCyclType") <> "" Then
                'sql += " AND cc.CYCLTYPE=@CYCLTYPE" & vbCrLf
                .Parameters.Add("CYCLTYPE", SqlDbType.VarChar).Value = ViewState("sCyclType")
            End If
            If ViewState("sAPPLYDATE1") <> "" Then
                'sql += " AND a.APPLYDATE >=@APPLYDATE1" & vbCrLf
                .Parameters.Add("APPLYDATE1", SqlDbType.DateTime).Value = CDate(ViewState("sAPPLYDATE1"))
            End If
            If ViewState("sAPPLYDATE2") <> "" Then
                'sql += " AND a.APPLYDATE <=@APPLYDATE2" & vbCrLf
                .Parameters.Add("APPLYDATE2", SqlDbType.DateTime).Value = CDate(ViewState("sAPPLYDATE2"))
            End If
            If ViewState("sENDDATE1") <> "" Then
                'sql += " AND a.ENDDATE >=@ENDDATE1" & vbCrLf
                .Parameters.Add("ENDDATE1", SqlDbType.DateTime).Value = CDate(ViewState("sENDDATE1"))
            End If
            If ViewState("sENDDATE2") <> "" Then
                'sql += " AND a.ENDDATE <=@ENDDATE2" & vbCrLf
                .Parameters.Add("ENDDATE2", SqlDbType.DateTime).Value = CDate(ViewState("sENDDATE2"))
            End If
            dt.Load(.ExecuteReader())
        End With

        labmsg.Text = "查無資料!!"
        tbSearch1.Visible = False
        DGobj1.Visible = False
        If dt.Rows.Count > 0 Then
            labmsg.Text = ""
            tbSearch1.Visible = True
            DGobj1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    'SQL 儲存
    Function SaveData1(ByRef oTrans As SqlTransaction, ByVal sAPPLIEDRESULT As String) As Integer
        'sAPPLIEDRESULT : Y:通過/N:失敗
        Dim iRst As Integer = 0 '資料異動筆數
        Dim iRst2 As Integer = 0 '資料異動筆數2
        Dim sql As String = ""

        If Hid_RETID.Value = "" Then Return iRst
        Select Case sAPPLIEDRESULT
            Case "Y", "N"
            Case Else
                Return iRst
        End Select

        'sql += " SET YEARS=@YEARS" & vbCrLf
        'sql += " ,DISTID=@DISTID" & vbCrLf
        'sql += " ,TPLANID=@TPLANID" & vbCrLf
        'sql += " ,PLANID=@PLANID" & vbCrLf
        'sql += " ,ORGID=@ORGID" & vbCrLf
        'sql += " ,OCID=@OCID" & vbCrLf
        'sql += " ,APPLYACCT=@APPLYACCT" & vbCrLf
        ''sql += " ,APPLYDATE=dbo.TRUNC_DATETIME(getdate())" & vbCrLf '@APPLYDATE" & vbCrLf
        'sql += " ,FUNID=@FUNID" & vbCrLf
        'sql += " ,ENDDATE=@ENDDATE" & vbCrLf
        'sql += " ,REASONID=@REASONID" & vbCrLf
        'sql += " ,REASON=@REASON" & vbCrLf
        'sql += " ,RID=@RID" & vbCrLf
        'sql += " ,CREATEACCT=@CREATEACCT" & vbCrLf
        'sql += " ,CREATEDATE=@CREATEDATE" & vbCrLf

        '異動 CLASS_RETRO.USEABLE
        sql = "" & vbCrLf
        sql &= " UPDATE CLASS_RETRO" & vbCrLf
        sql &= " SET MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=getdate()" & vbCrLf '@MODIFYDATE" & vbCrLf
        sql &= " ,USEABLE=@USEABLE" & vbCrLf
        sql &= " ,APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        sql &= " ,RESULTACCT=@RESULTACCT" & vbCrLf
        sql &= " ,RESULTDATE=getdate()" & vbCrLf
        'sql += " ,RIGHTID=@RIGHTID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND RETID=@RETID" & vbCrLf
        Dim uCmd As New SqlCommand(sql, oTrans.Connection, oTrans)

        '異動 CLASS_RETRO.RIGHTID
        sql = "" & vbCrLf
        sql &= " UPDATE CLASS_RETRO" & vbCrLf
        sql &= " SET RIGHTID=@RIGHTID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND RETID=@RETID" & vbCrLf
        Dim uCmd2 As New SqlCommand(sql, oTrans.Connection, oTrans)

        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM CLASS_RETRO" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND RETID=@RETID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, oTrans.Connection, oTrans)

        sql = ""
        sql &= " INSERT INTO AUTH_RENDCLASS ("
        sql &= " RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate"
        sql &= " ,BlameUnit" '是否歸責單位
        sql &= " ,RETID"
        sql &= " ) VALUES ("
        sql &= " @RightID,@Years,@PlanID,@DistID,@OCID,@Account,getdate(),@UseAble,@ModifyAcct, getdate(),@Reason,@ReasonID,@FunID,@EndDate"
        sql &= " ,@BlameUnit" '是否歸責單位
        sql &= " ,@RETID"
        sql &= " ) "
        Dim iCmd2 As New SqlCommand(sql, oTrans.Connection, oTrans)

        'Call TIMS.OpenDbConn(objconn)
        'Dim sTPLANID As String = TIMS.GetTPlanID(ddlPlan.SelectedValue, objconn)
        If Hid_RETID.Value <> "" Then
            '修改
            Dim dt1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("RETID", SqlDbType.Int).Value = Val(Hid_RETID.Value)
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count = 1 Then
                With uCmd
                    .Parameters.Clear()
                    '.Parameters.Add("YEARS", SqlDbType.VarChar).Value = ddlYears.SelectedValue
                    '.Parameters.Add("DISTID", SqlDbType.VarChar).Value = ddlDistID.SelectedValue
                    '.Parameters.Add("TPLANID", SqlDbType.VarChar).Value = TIMS.GetValue1(sTPLANID)
                    '.Parameters.Add("PLANID", SqlDbType.VarChar).Value = ddlPlan.SelectedValue
                    '.Parameters.Add("ORGID", SqlDbType.VarChar).Value = ddlOrgName.SelectedValue
                    '.Parameters.Add("OCID", SqlDbType.VarChar).Value = ddlClassCName.SelectedValue
                    '.Parameters.Add("APPLYACCT", SqlDbType.VarChar).Value = ddlAccount.SelectedValue 'sm.UserInfo.UserID
                    '.Parameters.Add("FUNID", SqlDbType.VarChar).Value = TIMS.chkSelFunID(cb_SelFunID)
                    '.Parameters.Add("ENDDATE", SqlDbType.DateTime).Value = TIMS.cdate2(Me.EndDate.Text)
                    '.Parameters.Add("REASONID", SqlDbType.VarChar).Value = ddlReasonID.SelectedValue
                    '.Parameters.Add("REASON", SqlDbType.NVarChar).Value = Me.txtReason.Text
                    '.Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                    '.Parameters.Add("Account", SqlDbType.VarChar).Value = Account.SelectedValue
                    '.Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    Select Case sAPPLIEDRESULT
                        Case cst_審核通過
                            .Parameters.Add("USEABLE", SqlDbType.VarChar).Value = sAPPLIEDRESULT
                            .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = sAPPLIEDRESULT
                            .Parameters.Add("RESULTACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        Case cst_審核失敗
                            .Parameters.Add("USEABLE", SqlDbType.VarChar).Value = sAPPLIEDRESULT
                            .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = sAPPLIEDRESULT
                            .Parameters.Add("RESULTACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        Case Else
                            .Parameters.Add("USEABLE", SqlDbType.VarChar).Value = Convert.DBNull
                            .Parameters.Add("APPLIEDRESULT", SqlDbType.VarChar).Value = Convert.DBNull
                            .Parameters.Add("RESULTACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    End Select
                    .Parameters.Add("RETID", SqlDbType.Int).Value = Val(Hid_RETID.Value)
                    iRst = .ExecuteNonQuery()
                End With
            End If
        End If

        If iRst > 0 AndAlso sAPPLIEDRESULT = cst_審核通過 Then
            'select * from Auth_REndClass where rownum <=10
            '新增 AUTH_RENDCLASS            
            Dim iRightID As Integer = DbAccess.GetNewId(oTrans, "AUTH_RENDCLASS_RIGHTID_SEQ,AUTH_RENDCLASS,RIGHTID")
            With iCmd2
                .Parameters.Clear()
                .Parameters.Add("RightID", SqlDbType.Int).Value = iRightID 'Auth_REndClass_MaxNo()
                .Parameters.Add("Years", SqlDbType.VarChar).Value = ddlYears.SelectedValue
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = ddlPlan.SelectedValue
                .Parameters.Add("DistID", SqlDbType.VarChar).Value = ddlDistID.SelectedValue
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = ddlClassCName.SelectedValue
                .Parameters.Add("Account", SqlDbType.VarChar).Value = ddlAccount.SelectedValue

                .Parameters.Add("UseAble", SqlDbType.VarChar).Value = cst_審核通過 '"Y"
                .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                .Parameters.Add("Reason", SqlDbType.NVarChar).Value = txtReason.Text
                .Parameters.Add("ReasonID", SqlDbType.VarChar).Value = ddlReasonID.SelectedValue
                .Parameters.Add("FunID", SqlDbType.VarChar).Value = TIMS.GetSelFunID(cb_SelFunID)
                .Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.Cdate2(Me.EndDate.Text)
                '是否歸責單位
                .Parameters.Add("BlameUnit", SqlDbType.VarChar).Value = TIMS.GetValue1(rblBlameUnit.SelectedValue)
                .Parameters.Add("RETID", SqlDbType.Int).Value = Val(Hid_RETID.Value) 'Auth_REndClass_MaxNo()
                .ExecuteNonQuery()
            End With
            With uCmd2
                .Parameters.Clear()
                .Parameters.Add("RIGHTID", SqlDbType.Int).Value = iRightID 'Auth_REndClass_MaxNo()
                .Parameters.Add("RETID", SqlDbType.Int).Value = Val(Hid_RETID.Value) 'Auth_REndClass_MaxNo()
                .ExecuteNonQuery()
            End With
        End If

        Return iRst
    End Function

    'SQL 顯示資料
    Sub loaddata1()
        If Hid_RETID.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.RETID" & vbCrLf '/*PK*/ 
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,a.PLANID" & vbCrLf
        'sql &= " ,a.ORGID" & vbCrLf
        sql &= " ,case when r3.RID3 is not null then r3.RID3+'" & TIMS.cst_spFlag1 & "'+CONVERT(varchar, a.OrgID)" & vbCrLf
        sql &= " else CONVERT(varchar, a.OrgID) end OrgID" & vbCrLf

        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,a.APPLYACCT" & vbCrLf
        sql &= " ,CONVERT(varchar, a.APPLYDATE, 111) APPLYDATE" & vbCrLf
        sql &= " ,a.FUNID" & vbCrLf
        sql &= " ,a.ENDDATE" & vbCrLf
        sql &= " ,a.REASONID" & vbCrLf
        sql &= " ,a.REASON" & vbCrLf
        sql &= " ,a.USEABLE" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.APPLIEDRESULT" & vbCrLf
        sql &= " ,a.RESULTACCT" & vbCrLf
        sql &= " ,a.RESULTDATE" & vbCrLf
        sql &= " ,a.RIGHTID" & vbCrLf
        '是否歸責單位
        sql &= " ,a.BlameUnit" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,cc.CYCLTYPE" & vbCrLf

        sql &= " ,dbo.DECODE6(a.APPLIEDRESULT,'Y','通過','N','不通過','待審') APPLIEDRESULT2" & vbCrLf
        sql &= " FROM CLASS_RETRO a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=a.OCID" & vbCrLf
        '補助地方政府
        sql &= " LEFT JOIN VIEW_RELSHIP23 r3 ON r3.RID3=cc.RID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.RETID=@RETID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RETID", SqlDbType.VarChar).Value = Hid_RETID.Value
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            '異常值
            If Hid_RETID.Value <> Convert.ToString(dr("RETID")) Then
                Common.MessageBox(Me, "資料有誤，請重新查詢!!")
                Exit Sub
            End If

            Dim OrgID3 As String = Convert.ToString(dr("OrgID"))
            Dim RID3 As String = ""
            Call TIMS.sUtl_AnalysisOrgID(OrgID3, RID3)

            labAPPLYDATE.Text = TIMS.Cdate3(Convert.ToString(dr("APPLYDATE")))
            labAPPLIEDRESULT2.Text = Convert.ToString(dr("APPLIEDRESULT2"))
            Common.SetListItem(ddlYears, Convert.ToString(dr("Years")))
            Common.SetListItem(ddlDistID, Convert.ToString(dr("DistID")))
            Call Makeplanlist(ddlPlan, ddlYears.SelectedValue, ddlDistID.SelectedValue, ddlOrgName, ddlClassCName, ddlAccount)
            Common.SetListItem(ddlPlan, Convert.ToString(dr("PlanID")))

            Call MakeddlOrgName(ddlOrgName, ddlPlan.SelectedValue, OrgID3, ddlDistID.SelectedValue, ddlClassCName, ddlAccount)
            Common.SetListItem(ddlOrgName, Convert.ToString(dr("OrgID")))

            '查詢班級 依輸入機構
            Call MakeClassName(ddlClassCName, ddlPlan.SelectedValue, Convert.ToString(dr("OrgID")), ddlDistID.SelectedValue)
            '帳號 依輸入機構
            Call MakeAccount(Me.ddlAccount, ddlPlan.SelectedValue, "", "", ddlDistID.SelectedValue, Convert.ToString(dr("OrgID")))

            'Common.SetListItem(ddlOrgName, Convert.ToString(dr("OrgID")))
            Common.SetListItem(ddlClassCName, Convert.ToString(dr("OCID")))
            Common.SetListItem(ddlAccount, Convert.ToString(dr("APPLYACCT")))

            Common.SetListItem(ddlReasonID, Convert.ToString(dr("ReasonID")))
            txtReason.Text = Convert.ToString(dr("Reason"))
            'cb_SelFunID.Text = Convert.ToString(dr("cb_SelFunID"))
            If Convert.ToString(dr("FUNID")) <> "" Then
                Call TIMS.SetCblValue(cb_SelFunID, Convert.ToString(dr("FUNID")))
                'For i As Int16 = 0 To cb_SelFunID.Items.Count - 1
                '    cb_SelFunID.Items(i).Selected = False
                '    If Convert.ToString(dr("FUNID")).IndexOf(cb_SelFunID.Items(i).Value) > -1 Then
                '        cb_SelFunID.Items(i).Selected = True
                '    End If
                'Next
            End If
            EndDate.Text = TIMS.Cdate3(Convert.ToString(dr("EndDate")))
            '是否歸責單位
            Common.SetListItem(rblBlameUnit, Convert.ToString(dr("BlameUnit")))
        End If
    End Sub

    'SERVER端 檢查
    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If sddlYears.SelectedValue = "" Then
            Errmsg += "請選擇 年度 不可為空。" & vbCrLf
        End If
        If sddlDistID.SelectedValue = "" Then
            Errmsg += "請選擇 轄區 不可為空。" & vbCrLf
        End If
        If sddlPlan.SelectedValue = "" Then
            Errmsg += "請選擇 訓練計畫 不可為空。" & vbCrLf
        End If
        'If sddlOrgName.SelectedValue = "" Then
        '    Errmsg += "請選擇 訓練機構 不可為空。" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢鈕
    Protected Sub btnSch_Click(sender As Object, e As EventArgs) Handles btnSch.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sERRMSG As String = ""
        Call CheckData2(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Page, sERRMSG)
            Exit Sub
        End If

        Call sSearch1()
    End Sub

    Protected Sub sddlYears_SelectedIndexChanged(sender As Object, e As EventArgs) Handles sddlYears.SelectedIndexChanged
        Call Makeplanlist(sddlPlan, sddlYears.SelectedValue, sddlDistID.SelectedValue, sddlOrgName, Nothing, Nothing)
    End Sub

    Protected Sub sddlDistID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles sddlDistID.SelectedIndexChanged
        Call Makeplanlist(sddlPlan, sddlYears.SelectedValue, sddlDistID.SelectedValue, sddlOrgName, Nothing, Nothing)
    End Sub

    '選擇計畫列出訓練機構
    Protected Sub sddlPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles sddlPlan.SelectedIndexChanged
        Call MakeddlOrgName(sddlOrgName, sddlPlan.SelectedValue, sm.UserInfo.OrgID, sddlDistID.SelectedValue, Nothing, Nothing)
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If ddlYears.SelectedValue = "" Then
            Errmsg += "請選擇 年度 不可為空。" & vbCrLf
        End If
        If ddlDistID.SelectedValue = "" Then
            Errmsg += "請選擇 轄區 不可為空。" & vbCrLf
        End If
        If ddlPlan.SelectedValue = "" Then
            Errmsg += "請選擇 訓練計畫 不可為空。" & vbCrLf
        End If
        If ddlOrgName.SelectedValue = "" Then
            Errmsg += "請選擇 訓練機構 不可為空。" & vbCrLf
        End If
        If ddlClassCName.SelectedValue = "" Then
            Errmsg += "請選擇 班級名稱 不可為空。" & vbCrLf
        End If
        If ddlAccount.SelectedValue = "" Then
            Errmsg += "請選擇 承辦人員 不可為空。" & vbCrLf
        End If
        If ddlReasonID.SelectedValue = "" Then
            Errmsg += "請選擇 補登資料原因 不可為空。" & vbCrLf
        End If
        If txtReason.Text = "" Then
            Errmsg += "請輸入 補登資料補登資料 不可為空。" & vbCrLf
        End If
        If TIMS.GetSelFunID(cb_SelFunID) = "" Then
            Errmsg += "請選擇 開放功能 不可為空。" & vbCrLf
        End If
        EndDate.Text = TIMS.Cdate3(EndDate.Text)
        If EndDate.Text = "" Then
            Errmsg += "請選擇 結束日期 不可為空。" & vbCrLf
        End If

        If Errmsg = "" Then
            If DateDiff(DateInterval.Day, CDate(aDate), CDate(EndDate.Text)) <= 0 Then
                Errmsg += "結束日期 不可為當日或早於申請日期!" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then
            Rst = False
            Return Rst
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '(審核通過) 儲存鈕
    Protected Sub btnSaveY1_Click(sender As Object, e As EventArgs) Handles btnSaveY1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sERRMSG As String = ""
        Call CheckData1(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Page, sERRMSG)
            Exit Sub
        End If

        ''Call TIMS.OpenDbConn(objconn)
        ''儲存前 前置作業
        'Dim Hid_TPLANID As String = TIMS.GetTPlanID(ddlPlan.SelectedValue, objconn)
        'If Hid_TPLANID = "" Then
        '    Common.MessageBox(Me, cst_errmsg1)
        '    Exit Sub
        'End If

        Dim iRst As Integer = 0
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            iRst = SaveData1(oTrans, cst_審核通過) '儲存鈕
            DbAccess.CommitTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End Try

        If iRst > 0 Then
            Common.MessageBox(Me, "儲存完成!!")
            Call ClearValue1()
            Call CloseList()
            Panelsch1.Visible = True
            Call sSearch1()
            Exit Sub
        Else
            Common.MessageBox(Me, "儲存異常，無資料異動!!")
            Exit Sub
        End If
    End Sub

    '(審核失敗) 儲存鈕
    Protected Sub btnSaveN1_Click(sender As Object, e As EventArgs) Handles btnSaveN1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'Dim sERRMSG As String = ""
        'Call CheckData1(sERRMSG)
        'If sERRMSG <> "" Then
        '    Common.MessageBox(Page, sERRMSG)
        '    Exit Sub
        'End If

        ''儲存前 前置作業
        'Dim Hid_TPLANID As String = TIMS.GetTPlanID(ddlPlan.SelectedValue, objconn)
        'If Hid_TPLANID = "" Then
        '    Common.MessageBox(Me, cst_errmsg1)
        '    Exit Sub
        'End If

        Dim iRst As Integer = 0
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(tConn)
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            iRst = SaveData1(oTrans, cst_審核失敗) '儲存鈕
            DbAccess.CommitTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End Try

        If iRst > 0 Then
            Common.MessageBox(Me, "儲存完成!!")
            Call ClearValue1()
            Call CloseList()
            Panelsch1.Visible = True
            Call sSearch1()
            Exit Sub
        Else
            Common.MessageBox(Me, "儲存異常，無資料異動!!")
            Exit Sub
        End If
    End Sub

    '編輯離開。/檢視離開
    Protected Sub btnQuit1_Click(sender As Object, e As EventArgs) Handles btnQuit1.Click
        Call ClearValue1()
        Call CloseList()
        Panelsch1.Visible = True
    End Sub

    Private Sub DG_ClassInfo_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        Select Case e.CommandName
            Case "UPD1"
                Call ClearValue1()
                Call CloseList()
                btnSaveY1.Visible = True
                btnSaveN1.Visible = True
                Paneledit1.Visible = True
                Dim sCmdArg As String = e.CommandArgument
                Hid_RETID.Value = TIMS.GetMyValue(sCmdArg, "RETID")
                Call loaddata1()

            Case "VIE1"
                Call ClearValue1()
                Call CloseList()
                Paneledit1.Visible = True
                Dim sCmdArg As String = e.CommandArgument
                Hid_RETID.Value = TIMS.GetMyValue(sCmdArg, "RETID")
                Call loaddata1()
        End Select
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        'Dim DG_ClassInfo As DataGrid = e.Row.FindControl("DG_ClassInfo")
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim labseqno As Label = e.Item.FindControl("labseqno")
                Dim labFUNIDN As Label = e.Item.FindControl("labFUNIDN")
                Dim hRETID As HiddenField = e.Item.FindControl("hRETID")
                Dim hOCID As HiddenField = e.Item.FindControl("hOCID")
                Dim lbupdate1 As LinkButton = e.Item.FindControl("lbupdate1")
                Dim lbview1 As LinkButton = e.Item.FindControl("lbview1")
                'Dim lbprint1 As LinkButton = e.Item.FindControl("lbprint1")

                labseqno.Text = TIMS.Get_DGSeqNo(sender, e) '序號
                TIMS.Tooltip(labseqno, Convert.ToString(drv("RETID")))
                'labseqno.Text = Convert.ToString(drv("seqno"))
                hRETID.Value = Convert.ToString(drv("RETID"))
                hOCID.Value = Convert.ToString(drv("OCID"))

                Call TIMS.SetCblValue(cb_SelFunID, Convert.ToString(drv("FUNID")))
                labFUNIDN.Text = TIMS.GetSelFunName(cb_SelFunID)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "RETID", Convert.ToString(drv("RETID")))
                lbupdate1.CommandArgument = sCmdArg
                lbview1.CommandArgument = sCmdArg
                'lbprint1.CommandArgument = sCmdArg

                If Convert.ToString(drv("APPLIEDRESULT")) <> "" Then
                    lbupdate1.CommandArgument = ""
                    lbupdate1.Enabled = False
                    TIMS.Tooltip(lbupdate1, "資料已審核")
                End If

        End Select

    End Sub

    '年度選擇
    Protected Sub ddlYears_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlYears.SelectedIndexChanged
        Call Makeplanlist(ddlPlan, ddlYears.SelectedValue, ddlDistID.SelectedValue, ddlOrgName, ddlClassCName, ddlAccount)
    End Sub

    '轄區選擇
    Protected Sub ddlDistID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlDistID.SelectedIndexChanged
        Call Makeplanlist(ddlPlan, ddlYears.SelectedValue, ddlDistID.SelectedValue, ddlOrgName, ddlClassCName, ddlAccount)
    End Sub

    '選擇計畫
    Protected Sub ddlPlan_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPlan.SelectedIndexChanged
        Call MakeddlOrgName(ddlOrgName, ddlPlan.SelectedValue, sm.UserInfo.OrgID, ddlDistID.SelectedValue, ddlClassCName, ddlAccount)
        Common.SetListItem(ddlOrgName, sm.UserInfo.OrgID)
        '查詢班級 依輸入機構
        Call MakeClassName(ddlClassCName, ddlPlan.SelectedValue, ddlOrgName.SelectedValue, ddlDistID.SelectedValue)
        '帳號 依輸入機構
        Call MakeAccount(Me.ddlAccount, ddlPlan.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, ddlDistID.SelectedValue, ddlOrgName.SelectedValue)
    End Sub

    '選擇機構
    Protected Sub ddlOrgName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOrgName.SelectedIndexChanged
        '查詢班級 依輸入機構
        Call MakeClassName(ddlClassCName, ddlPlan.SelectedValue, ddlOrgName.SelectedValue, ddlDistID.SelectedValue)
        '帳號 依輸入機構
        Call MakeAccount(Me.ddlAccount, ddlPlan.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, ddlDistID.SelectedValue, ddlOrgName.SelectedValue)
    End Sub

    Protected Sub DG_ClassInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DG_ClassInfo.SelectedIndexChanged

    End Sub
End Class
