Partial Class CO_01_002_OP
    Inherits AuthBasePage

    'ORG_PARTY
    Const cst_btnEdit As String = "btnEdit"
    Const cst_btnDel As String = "btnDel"

    '    alter TABLE [dbo].[ORG_PARTY] add PLANSUB INT ,PTYDATE datetime
    'go
    'update [ORG_PARTY] Set PLANSUB =3 where 1=1 
    'go
    '--alter TABLE [dbo].[ORG_PARTY] alter column PLANSUB [varchar](1) COLLATE Chinese_Taiwan_Stroke_CS_AS Not null
    'go
    'alter TABLE [dbo].[ORG_PARTY] ALTER COLUMN PLANSUB INT Not NULL
    'GO

    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    Dim flag_ROC As Boolean = True
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            create1()
            Call sSearch1()
            'Session("MyWrongTable") = Nothing
        End If

        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(sDistID, sm.UserInfo.DistID)
            sDistID.Enabled = False
        End If

    End Sub

    Sub USE_SEARCH_VAL1()
        If Session("CO_01_002_OP_SCH") Is Nothing Then Return
        Dim SchVal1 As String = Session("CO_01_002_OP_SCH")
        Session("CO_01_002_OP_SCH") = Nothing
        If TIMS.GetMyValue(SchVal1, "PRG") <> "CO_01_002_OP" Then Return

        Dim rqDISTID As String = TIMS.GetMyValue(SchVal1, "DISTID")
        If rqDISTID <> "" Then TIMS.SetCblValue(sDistID, rqDISTID)

        Dim rqYEARS As String = TIMS.GetMyValue(SchVal1, "YEARS")
        If rqYEARS <> "" Then Common.SetListItem(sYEARlist, rqYEARS)

        Dim rqHALFYEAR As String = TIMS.GetMyValue(SchVal1, "HALFYEAR")
        If rqHALFYEAR <> "" Then Common.SetListItem(sHALFYEAR, rqHALFYEAR)
    End Sub

    Sub create1()
        divSch1.Visible = True
        divEdt1.Visible = False

        '選擇全部轄區
        sDistID.Attributes("onclick") = "SelectAll('sDistID','sDistHidden');"
        sDistID = TIMS.Get_DistID(sDistID)
        sDistID.Items.Insert(0, New ListItem("全部", ""))
        sDistID.AppendDataBoundItems = True
        'Common.SetListItem(DistID, sm.UserInfo.DistID)

        sYEARlist = TIMS.GetSyear(sYEARlist)
        sYEARlist.AppendDataBoundItems = True
        Common.SetListItem(sYEARlist, sm.UserInfo.Years)
        'url1 &= "&YEARS=" & vYEARS
        'url1 &= "&HALFYEAR=" & vHALFYEAR
        'Dim rqYEARS As String = TIMS.ClearSQM(Request("YEARS"))
        'Dim rqHALFYEAR As String = TIMS.ClearSQM(Request("HALFYEAR"))
        'If rqYEARS <> "" Then Common.SetListItem(sYEARlist, rqYEARS)
        'If rqHALFYEAR <> "" Then Common.SetListItem(sHALFYEAR, rqHALFYEAR)
        Call USE_SEARCH_VAL1()
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        'Dim MRQ_ID As String = TIMS.Get_MRqID(Me)
        'Dim url1 As String = "CO_01_002.aspx?ID=" & MRQ_ID
        'TIMS.Utl_Redirect(Me, objconn, url1)
        divSch1.Visible = True
        divEdt1.Visible = False
    End Sub

    Sub sSearch1()
        divSch1.Visible = True
        divEdt1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = TIMS.cst_NODATAMsg1

        schPTYDATE1.Text = TIMS.ClearSQM(schPTYDATE1.Text)
        schPTYDATE2.Text = TIMS.ClearSQM(schPTYDATE2.Text)
        schPTYDATE1.Text = If(flag_ROC, TIMS.Cdate7(schPTYDATE1.Text), TIMS.Cdate3(schPTYDATE1.Text))
        schPTYDATE2.Text = If(flag_ROC, TIMS.Cdate7(schPTYDATE2.Text), TIMS.Cdate3(schPTYDATE2.Text))
        sPTYNAME.Text = TIMS.ClearSQM(sPTYNAME.Text)

        Dim v_sDistID As String = TIMS.GetCblValue(sDistID)
        v_sDistID = TIMS.CombiSQM2IN(v_sDistID)
        Dim v_sYEARlist As String = TIMS.GetListValue(sYEARlist)
        Dim v_sHALFYEAR As String = TIMS.GetListValue(sHALFYEAR)

        Dim sql As String = ""
        sql &= " SELECT a.PTYID" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.Years) YEARS_ROC" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        sql &= " ,d.NAME DISTNAME" & vbCrLf
        sql &= " ,a.HALFYEAR" & vbCrLf
        sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,a.PTYNAME" & vbCrLf
        sql &= " ,a.PLANSUB" & vbCrLf
        'PLANSUB_N'「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 varchar(1)
        sql &= " ,case a.PLANSUB when 1 then '產業人才投資計畫' when 2 then '提升勞工自主學習計畫' when 3 then '不區分' end PLANSUB_N" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,dbo.RT_DataFormat(a.PTYDATE) PTYDATE_RC" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,CASE WHEN o.PTYID IS NULL THEN '1' END CanDelete" & vbCrLf '未使用可刪除
        sql &= " FROM ORG_PARTY a" & vbCrLf
        sql &= " LEFT JOIN ID_DISTRICT d ON d.DISTID=a.DISTID" & vbCrLf
        sql &= " LEFT JOIN (SELECT PTYID,COUNT(1) ORGCNT1 FROM ORG_PARTYORG WITH(NOLOCK) GROUP BY PTYID) o ON o.PTYID=a.PTYID" & vbCrLf

        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        If v_sDistID <> "" Then sql &= " AND a.DISTID IN (" & v_sDistID & ")" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        If v_sHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf

        If (sPTYNAME.Text <> "") Then sql &= " AND a.PTYNAME like '%'+@PTYNAME+'%' " & vbCrLf
        If (schPTYDATE1.Text <> "") Then sql &= " AND a.PTYDATE >= convert(date,@PTYDATE1)" & vbCrLf
        If (schPTYDATE2.Text <> "") Then sql &= " AND a.PTYDATE <= convert(date,@PTYDATE2)" & vbCrLf

        sql &= " ORDER BY a.PTYDATE DESC" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("YEARS", v_sYEARlist)
        'parms.Add("DISTID", sm.UserInfo.DistID)
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        If v_sHALFYEAR <> "" Then parms.Add("HALFYEAR", v_sHALFYEAR)

        If (sPTYNAME.Text <> "") Then parms.Add("PTYNAME", sPTYNAME.Text)
        If (schPTYDATE1.Text <> "") Then parms.Add("PTYDATE1", If(flag_ROC, TIMS.Cdate18(schPTYDATE1.Text), TIMS.Cdate2(schPTYDATE1.Text)))
        If (schPTYDATE2.Text <> "") Then parms.Add("PTYDATE2", If(flag_ROC, TIMS.Cdate18(schPTYDATE2.Text), TIMS.Cdate2(schPTYDATE2.Text)))

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Sub sSaveData1()
        If Hid_PTYID.Value = "" Then Exit Sub
        Dim iPTYID As Integer = Val(TIMS.ClearSQM(Hid_PTYID.Value))
        tPTYDATE.Text = If(flag_ROC, TIMS.Cdate7(tPTYDATE.Text), TIMS.Cdate3(tPTYDATE.Text))

        Dim uSql As String = ""
        uSql &= " UPDATE ORG_PARTY" & vbCrLf
        uSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " ,PTYNAME=@PTYNAME" & vbCrLf
        uSql &= " ,PTYDATE=convert(date,@PTYDATE)" & vbCrLf
        'uSql &= " WHERE ORGID=@ORGID and PTYID=@PTYID" & vbCrLf
        uSql &= " WHERE PTYID=@PTYID" & vbCrLf

        'Dim dSql As String = ""
        'dSql = "" & vbCrLf
        'dSql &= " DELETE ORG_PARTYORG" & vbCrLf
        'dSql &= " WHERE ORGID=@ORGID and PTYID=@PTYID" & vbCrLf
        '--insert

        'Dim iPTYOID As Integer = Val(dt.Rows(0)("PTYOID"))
        Dim u_parms As New Hashtable
        u_parms.Clear()
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("PTYNAME", tPTYNAME.Text)
        u_parms.Add("PTYDATE", If(flag_ROC, TIMS.Cdate18(tPTYDATE.Text), TIMS.Cdate2(tPTYDATE.Text)))
        u_parms.Add("PTYID", iPTYID)
        DbAccess.ExecuteNonQuery(uSql, objconn, u_parms)
    End Sub

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        sSearch1()
    End Sub

    Protected Sub BtnBack2_Click(sender As Object, e As EventArgs) Handles BtnBack2.Click
        divSch1.Visible = True
        divEdt1.Visible = False

        Dim MRQ_ID As String = TIMS.Get_MRqID(Me)
        Dim url1 As String = "CO_01_002.aspx?ID=" & MRQ_ID
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Function CheckSaveData1(ByRef sReason As String) As Boolean
        Dim rst As Boolean = True
        Const cst_必須填寫 As String = "必須填寫"
        '辦理活動場次日期
        Dim aPTYDATE As String = TIMS.ClearSQM(tPTYDATE.Text)
        If aPTYDATE <> "" Then
            If flag_ROC Then
                If Not TIMS.IsDate7(aPTYDATE) Then
                    sReason &= "辦理活動場次日期 必須為正確的日期格式 民國年月日(TW_YEAR/MM/dd)" & vbCrLf
                Else
                    aPTYDATE = TIMS.Cdate7(aPTYDATE)
                End If
            Else
                If Not TIMS.IsDate1(aPTYDATE) Then
                    sReason &= "辦理活動場次日期 必須為正確的日期格式(yyyy/MM/dd)" & vbCrLf
                Else
                    aPTYDATE = TIMS.Cdate3(aPTYDATE)
                End If
            End If
        Else
            sReason &= cst_必須填寫 & "辦理活動場次日期" & vbCrLf
        End If

        '活動場次名稱
        Dim aPTYNAME As String = TIMS.ClearSQM(tPTYNAME.Text)
        If aPTYNAME <> "" Then
            If aPTYNAME.Length > 100 Then
                sReason += "活動場次名稱字串長度不可超過100" & vbCrLf
            End If
        Else
            sReason += cst_必須填寫 & "活動場次名稱" & vbCrLf
        End If

        If sReason <> "" Then rst = False
        Return rst
    End Function
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Dim sErrMsg1 As String = ""
        If Not CheckSaveData1(sErrMsg1) Then
            Common.MessageBox(Me, sErrMsg1)
            Exit Sub
        End If
        'divSch1.Visible = True
        'divEdt1.Visible = False
        sSaveData1()

        divSch1.Visible = True
        divEdt1.Visible = False
        sSearch1()
        sm.LastResultMessage = "儲存完畢"

    End Sub

    ''' <summary> 單一資料 </summary>
    ''' <param name="sCmdArg"></param>
    Sub sLoadData1(ByRef sCmdArg As String)
        'Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Hid_PTYID.Value = TIMS.GetMyValue(sCmdArg, "PTYID")
        Hid_YEARS.Value = TIMS.GetMyValue(sCmdArg, "YEARS")
        Hid_DISTID.Value = TIMS.GetMyValue(sCmdArg, "DISTID")
        Hid_TPLANID.Value = TIMS.GetMyValue(sCmdArg, "TPLANID")
        Hid_HALFYEAR.Value = TIMS.GetMyValue(sCmdArg, "HALFYEAR") '1:上年度 /2:下年度

        Dim sql As String = ""
        sql &= " SELECT a.PTYID" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.Years) YEARS_ROC" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        sql &= " ,d.NAME DISTNAME" & vbCrLf
        sql &= " ,a.HALFYEAR" & vbCrLf
        sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,a.PTYNAME" & vbCrLf
        sql &= " ,a.PLANSUB" & vbCrLf
        'PLANSUB_N'「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 varchar(1)
        sql &= " ,case a.PLANSUB when 1 then '產業人才投資計畫' when 2 then '提升勞工自主學習計畫' when 3 then '不區分' end PLANSUB_N" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,dbo.RT_DataFormat(a.PTYDATE) PTYDATE_RC" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        'sql &= " ,CASE WHEN o.PTYID IS NULL THEN '1' END CanDelete" & vbCrLf '未使用可刪除
        sql &= " FROM ORG_PARTY a" & vbCrLf
        sql &= " LEFT JOIN ID_DISTRICT d ON d.DISTID=a.DISTID" & vbCrLf
        'sql &= " LEFT JOIN (SELECT PTYID,COUNT(1) ORGCNT1 FROM ORG_PARTYORG WITH(NOLOCK) GROUP BY PTYID) o ON o.PTYID=a.PTYID" & vbCrLf

        sql &= " WHERE a.PTYID=@PTYID" & vbCrLf
        sql &= " AND a.YEARS=@YEARS" & vbCrLf
        sql &= " AND a.DISTID=@DISTID" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf

        'If (sPTYNAME.Text <> "") Then sql &= " AND a.PTYNAME like '%'+@PTYNAME+'%' " & vbCrLf
        'If (schPTYDATE1.Text <> "") Then sql &= " AND a.PTYDATE >= convert(date,@PTYDATE1)" & vbCrLf
        'If (schPTYDATE2.Text <> "") Then sql &= " AND a.PTYDATE <= convert(date,@PTYDATE2)" & vbCrLf
        'sql &= " ORDER BY a.PTYDATE DESC" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PTYID", Hid_PTYID.Value)
        parms.Add("YEARS", Hid_YEARS.Value)
        parms.Add("DISTID", Hid_DISTID.Value)
        parms.Add("TPLANID", Hid_TPLANID.Value)
        parms.Add("HALFYEAR", Hid_HALFYEAR.Value)
        'If v_sHALFYEAR <> "" Then parms.Add("HALFYEAR", v_sHALFYEAR)
        'If (sPTYNAME.Text <> "") Then parms.Add("PTYNAME", sPTYNAME.Text)
        'If (schPTYDATE1.Text <> "") Then parms.Add("PTYDATE1", schPTYDATE1.Text)
        'If (schPTYDATE2.Text <> "") Then parms.Add("PTYDATE2", schPTYDATE2.Text)
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        'PageControler1.Visible = False
        'DataGridTable.Visible = False
        'msg1.Text = "查無資料"
        'If dt.Rows.Count = 0 Then Exit Sub
        divSch1.Visible = True
        divEdt1.Visible = False
        If dt.Rows.Count = 0 Then
            sm.LastErrorMessage = "查無資料"
            Exit Sub
        End If

        divSch1.Visible = False
        divEdt1.Visible = True
        Dim dr1 As DataRow = dt.Rows(0)
        Call sClearlist1()
        Call sShowData1(dr1)
    End Sub

    Sub sShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub

        Const cst_上半年度 As String = "上半年度"
        Const cst_下半年度 As String = "下半年度"
        Dim s_HALFYEAR As String = ""
        Select Case Convert.ToString(dr("HALFYEAR"))
            Case "1"
                s_HALFYEAR = cst_上半年度
            Case "2"
                s_HALFYEAR = cst_下半年度
        End Select
        LabPartyYears.Text = Convert.ToString(dr("YEARS_ROC"))
        LabDISTNAME.Text = Convert.ToString(dr("DISTNAME"))
        LabhalfYear.Text = s_HALFYEAR 'Convert.ToString(dr("HALFYEAR"))
        labPLANSUB.Text = Convert.ToString(dr("PLANSUB_N"))
        tPTYDATE.Text = Convert.ToString(dr("PTYDATE_RC"))
        tPTYNAME.Text = Convert.ToString(dr("PTYNAME"))
    End Sub

    Sub sClearlist1()
        LabPartyYears.Text = "" ' Convert.ToString(dr("YEARS"))
        LabDISTNAME.Text = "" 'Convert.ToString(dr("DISTNAME"))
        LabhalfYear.Text = "" 's_HALFYEAR 'Convert.ToString(dr("HALFYEAR"))
        labPLANSUB.Text = "" 'Convert.ToString(dr("PLANSUB_N"))
        tPTYDATE.Text = "" 'Convert.ToString(dr("PTYDATE_RC"))
        tPTYNAME.Text = "" 'Convert.ToString(dr("PTYNAME"))
    End Sub

    Sub sDeleteData1(ByVal sCmdArg As String)
        If sCmdArg = "" Then Exit Sub
        Hid_PTYID.Value = TIMS.GetMyValue(sCmdArg, "PTYID")
        Hid_YEARS.Value = TIMS.GetMyValue(sCmdArg, "YEARS")
        Hid_DISTID.Value = TIMS.GetMyValue(sCmdArg, "DISTID")
        Hid_TPLANID.Value = TIMS.GetMyValue(sCmdArg, "TPLANID")
        Hid_HALFYEAR.Value = TIMS.GetMyValue(sCmdArg, "HALFYEAR") '1:上年度 /2:下年度

        Dim sql As String = ""
        sql &= " SELECT a.PTYID" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.Years) YEARS_ROC" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        'sql &= " ,d.NAME DISTNAME" & vbCrLf
        sql &= " ,a.HALFYEAR" & vbCrLf
        sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,a.PTYNAME" & vbCrLf
        sql &= " ,a.PLANSUB" & vbCrLf
        'PLANSUB_N'「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 varchar(1)
        sql &= " ,case a.PLANSUB when 1 then '產業人才投資計畫' when 2 then '提升勞工自主學習計畫' when 3 then '不區分' end PLANSUB_N" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,dbo.RT_DataFormat(a.PTYDATE) PTYDATE_RC" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        'sql &= " ,CASE WHEN o.PTYID IS NULL THEN '1' END CanDelete" & vbCrLf '未使用可刪除
        sql &= " FROM ORG_PARTY a" & vbCrLf
        'sql &= " LEFT JOIN (SELECT PTYID,COUNT(1) ORGCNT1 FROM ORG_PARTYORG WITH(NOLOCK) GROUP BY PTYID) o ON o.PTYID=a.PTYID" & vbCrLf

        sql &= " WHERE a.PTYID=@PTYID" & vbCrLf
        sql &= " AND a.YEARS=@YEARS" & vbCrLf
        sql &= " AND a.DISTID=@DISTID" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf

        'If (sPTYNAME.Text <> "") Then sql &= " AND a.PTYNAME like '%'+@PTYNAME+'%' " & vbCrLf
        'If (schPTYDATE1.Text <> "") Then sql &= " AND a.PTYDATE >= convert(date,@PTYDATE1)" & vbCrLf
        'If (schPTYDATE2.Text <> "") Then sql &= " AND a.PTYDATE <= convert(date,@PTYDATE2)" & vbCrLf
        'sql &= " ORDER BY a.PTYDATE DESC" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PTYID", Hid_PTYID.Value)
        parms.Add("YEARS", Hid_YEARS.Value)
        parms.Add("DISTID", Hid_DISTID.Value)
        parms.Add("TPLANID", Hid_TPLANID.Value)
        parms.Add("HALFYEAR", Hid_HALFYEAR.Value)
        'If v_sHALFYEAR <> "" Then parms.Add("HALFYEAR", v_sHALFYEAR)
        'If (sPTYNAME.Text <> "") Then parms.Add("PTYNAME", sPTYNAME.Text)
        'If (schPTYDATE1.Text <> "") Then parms.Add("PTYDATE1", schPTYDATE1.Text)
        'If (schPTYDATE2.Text <> "") Then parms.Add("PTYDATE2", schPTYDATE2.Text)
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        'PageControler1.Visible = False
        'DataGridTable.Visible = False
        'msg1.Text = "查無資料"
        'If dt.Rows.Count = 0 Then Exit Sub
        divSch1.Visible = True
        divEdt1.Visible = False
        If dt.Rows.Count = 1 Then
            Dim dSql As String = " DELETE ORG_PARTY WHERE PTYID=@PTYID"
            Dim d_parms As New Hashtable
            d_parms.Clear()
            d_parms.Add("PTYID", Val(Hid_PTYID.Value))
            DbAccess.ExecuteNonQuery(dSql, objconn, d_parms)

            sm.LastResultMessage = "資料已刪除"
            Exit Sub
        End If

        'divSch1.Visible = False
        'divEdt1.Visible = True
        'Dim dr1 As DataRow = dt.Rows(0)
        'Call sClearlist1()
        'Call sShowData1(dr1)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Call sClearlist1()

        'Dim sCmdArg As String = ""
        Select Case e.CommandName
            Case cst_btnEdit
                sCmdArg = e.CommandArgument
                sLoadData1(sCmdArg)
            Case cst_btnDel
                sCmdArg = e.CommandArgument
                sDeleteData1(sCmdArg)
                'divSch1.Visible = True
                'divEdt1.Visible = False
                sSearch1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDel1 As LinkButton = e.Item.FindControl("lbtDel1")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PTYID", Convert.ToString(drv("PTYID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "TPLANID", Convert.ToString(drv("TPLANID")))
                TIMS.SetMyValue(sCmdArg, "DISTID", Convert.ToString(drv("DISTID")))
                TIMS.SetMyValue(sCmdArg, "HALFYEAR", Convert.ToString(drv("HALFYEAR")))

                lbtEdit.CommandArgument = sCmdArg
                lbtDel1.Visible = False
                If Convert.ToString(drv("CanDelete")) = "1" Then
                    lbtDel1.Visible = True
                    lbtDel1.CommandArgument = sCmdArg
                    lbtDel1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                End If
        End Select
    End Sub

End Class
