Public Class CO_01_002
    Inherits AuthBasePage 'System.Web.UI.Page

    'ORG_PARTY ORG_PARTYORG
    Dim gdtParty As DataTable = Nothing
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    'Const Cst_ComIDNO As Integer = 0
    Const Cst_Years As Integer = 0 '年度
    Const Cst_HALFYEAR As Integer = 1 '上下半年'1:上年度 /2:下年度
    Const Cst_PLANSUB As Integer = 2 '「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 INT
    Const Cst_PTYDATE As Integer = 3 '辦理活動場次日期
    Const Cst_PTYNAME As Integer = 4 '活動場次名稱
    Const cst_filedColumnNum As Integer = 5

    Dim aYears As String = "" 'colArray(Cst_Years).ToString '年度
    Dim aHALFYEAR As String = "" 'colArray(Cst_HALFYEAR).ToString '上下半年
    Dim aPLANSUB As String = "" '「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分
    Dim aPTYDATE As String = "" '辦理活動場次日期
    Dim aPTYNAME As String = "" 'colArray(Cst_PTYNAME).ToString '活動場次名稱

    'alter TABLE [dbo].[ORG_PARTY] add PLANSUB [varchar](1) COLLATE Chinese_Taiwan_Stroke_CS_AS ,PTYDATE datetime
    'go
    'update [ORG_PARTY] Set PLANSUB ='3' where 1!=1 
    'go
    'alter TABLE [dbo].[ORG_PARTY] alter column PLANSUB [varchar](1) COLLATE Chinese_Taiwan_Stroke_CS_AS not null
    'go
    'alter TABLE [dbo].[ORG_PARTY] ALTER COLUMN PLANSUB INT NOT NULL
    '匯入檔的欄位配置如下：
    Const Cst_COMIDNO As Integer = 0
    Const Cst_PTYID As Integer = 1
    ''' <summary>
    ''' 匯入檔的欄位總數
    ''' </summary>
    Const cst_filedColumnNum2 As Integer = 2
    'Dim aCOMIDNO As String = ""

    '匯入時共用的 存取值
    Dim aORGID As String = ""
    Dim aPTYID As String = ""

    'Const cst_btnEdit As String = "btnEdit"
    'Const cst_btnAddt As String = "btnAddt"
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        'Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        HyperLink1.NavigateUrl = "../../Doc/OrgPARTYv21.zip"
        HyperLink2.NavigateUrl = "../../Doc/OrgPARTYORGv21.zip"

        If Not IsPostBack Then
            Call SCreate1()
        End If

        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(sDistID, sm.UserInfo.DistID)
            sDistID.Enabled = False
        End If

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    'center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
        '    'HistoryRID.Attributes("onclick") = "ShowFrame();"
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If

    End Sub

    Sub SCreate1()
        divSch1.Visible = True
        divEdt1.Visible = False
        msg1.Text = ""
        PageControler1.Visible = False

        '選擇全部轄區
        sDistID.Attributes("onclick") = "SelectAll('sDistID','sDistHidden');"
        sDistID = TIMS.Get_DistID(sDistID)
        sDistID.Items.Insert(0, New ListItem("全部", ""))
        sDistID.AppendDataBoundItems = True
        'Common.SetListItem(DistID, sm.UserInfo.DistID)

        SYEARlist = TIMS.GetSyear(SYEARlist)
        Common.SetListItem(SYEARlist, sm.UserInfo.Years)

        '(加強操作便利性)
        'RIDValue.Value = sm.UserInfo.RID
        'center.Text = sm.UserInfo.OrgName
        'Orgidvalue.Value = sm.UserInfo.OrgID
    End Sub

    ''' <summary>單筆資料查詢顯示1</summary>
    ''' <param name="sCmdArg"></param>
    Sub SLoadData1(ByRef sCmdArg As String)
        'Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        Dim vDISTID As String = TIMS.GetMyValue(sCmdArg, "DISTID") '單1
        Dim vTPLANID As String = TIMS.GetMyValue(sCmdArg, "TPLANID")
        Dim vHALFYEAR As String = TIMS.GetMyValue(sCmdArg, "HALFYEAR") '1:上年度 /2:下年度
        Dim vOrgID As String = TIMS.GetMyValue(sCmdArg, "OrgID")

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR" & vbCrLf
        '「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 varchar(1)
        'sql &= " ,COUNT(1) CNT1" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN PLANSUB IN (1,3) THEN 1 END) CNTG1" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN PLANSUB IN (2,3) THEN 1 END) CNTW1" & vbCrLf
        sql &= " FROM dbo.ORG_PARTY a" & vbCrLf
        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf
        sql &= " AND a.DISTID=@DISTID" & vbCrLf '單1
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        sql &= " GROUP BY a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR )" & vbCrLf

        sql &= " SELECT a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.Years) Years_ROC" & vbCrLf
        '「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分 varchar(1)
        'sql &= " ,CASE WHEN rr.ORGKIND2='G' THEN a.CNTG1 WHEN rr.ORGKIND2='W' THEN a.CNTW1 ELSE a.CNT1 END CNT1 " & vbCrLf '應出席場次
        sql &= " ,rr.ORGKIND2" & vbCrLf
        'sql &= " ,a.CNT1 ,a.CNTG1 ,a.CNTW1" & vbCrLf
        sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf '1:上年度 /2:下年度
        sql &= " ,rr.RID,rr.OrgID,rr.OrgName" & vbCrLf
        sql &= " ,dbo.FN_GET_PARTYCNT1(a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR,rr.ORGID,rr.ORGKIND2) CNT1" & vbCrLf '應出席場次
        sql &= " ,dbo.FN_GET_PARTYCNT2(a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR,rr.ORGID) CNT2" & vbCrLf '實際出席場次
        sql &= " FROM dbo.VIEW_RIDNAME rr" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.PlanID=rr.PlanID" & vbCrLf
        sql &= " JOIN WC1 a on ip.Years=a.Years and ip.DistID=a.DistID and ip.TPlanID=ip.TPlanID" & vbCrLf
        sql &= " WHERE ip.YEARS=@YEARS" & vbCrLf
        sql &= " AND ip.DISTID=@DISTID" & vbCrLf '單1
        sql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND rr.OrgID=@OrgID" & vbCrLf

        Dim parms As New Hashtable From {
            {"YEARS", vYEARS},
            {"DISTID", vDISTID}, 'sm.UserInfo.DistID)
            {"TPLANID", vTPLANID}, ' sm.UserInfo.TPlanID)
            {"OrgID", vOrgID}
        }
        If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

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
        Call SClearlist1()
        Call SShowData1(dr1)
    End Sub

    ''' <summary>查詢鈕多筆資料LIST顯示</summary>
    Sub SSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        'ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        Dim vsDistID As String = TIMS.GetCblValue(sDistID)
        vsDistID = TIMS.CombiSQM2IN(vsDistID) 'SQL使用

        Dim vYEARS As String = TIMS.GetListValue(SYEARlist) '.SelectedValue)
        Dim vHALFYEAR As String = TIMS.GetListValue(halfYear) '.SelectedValue) '1:上年度 /2:下年度
        Dim v_ORGNAMElk As String = TIMS.ClearSQM(txtORGNAME.Text)
        'Dim v_ORGID As String = ""
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'If Len(RIDValue.Value) <> 1 Then
        '    v_ORGID = TIMS.Get_OrgID(RIDValue.Value, objconn)
        '    v_ORGID = TIMS.VAL1(v_ORGID)
        '    If v_ORGID <> "" AndAlso Val(v_ORGID) <= 0 Then v_ORGID = "" '異常值，清空
        'End If

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR" & vbCrLf
        sql &= " FROM dbo.ORG_PARTY a" & vbCrLf '場次資訊
        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        If vsDistID <> "" Then sql &= " AND a.DISTID IN (" & vsDistID & ")" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf
        sql &= " GROUP BY a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR )" & vbCrLf

        sql &= " ,WC2 AS (SELECT a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR" & vbCrLf
        'sql &= " ,dbo.FN_CYEAR2(a.Years) Years_ROC" & vbCrLf
        sql &= " ,rr.ORGKIND2" & vbCrLf
        'sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf
        sql &= " ,rr.RID,rr.ORGID,rr.ORGNAME,ip.DISTNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_PARTYCNT1(a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR,rr.ORGID,rr.ORGKIND2) CNT1" & vbCrLf '應出席場次
        sql &= " ,dbo.FN_GET_PARTYCNT2(a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR,rr.ORGID) CNT2" & vbCrLf '實際出席場次
        sql &= " FROM dbo.VIEW_RIDNAME rr" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.PlanID=rr.PlanID" & vbCrLf
        sql &= " JOIN WC1 a on ip.Years=a.Years and ip.DistID=a.DistID and ip.TPlanID=ip.TPlanID" & vbCrLf
        sql &= " WHERE ip.YEARS=@YEARS" & vbCrLf
        'sql &= " AND ip.DISTID=@DISTID" & vbCrLf
        If vsDistID <> "" Then sql &= " AND ip.DISTID IN (" & vsDistID & ")" & vbCrLf
        'If v_ORGID <> "" Then sql &= " AND rr.ORGID=@ORGID" & vbCrLf
        If v_ORGNAMElk <> "" Then sql &= " AND  rr.ORGNAME like '%'+@ORGNAMElk+'%'" & vbCrLf
        sql &= " AND ip.TPLANID=@TPLANID )" & vbCrLf

        sql &= " SELECT a.YEARS,a.DISTID,a.TPLANID,a.HALFYEAR" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) Years_ROC" & vbCrLf
        sql &= " ,a.ORGKIND2" & vbCrLf
        sql &= " ,case a.HALFYEAR when 1 then '上年度' when 2 then '下年度' end halfYearN" & vbCrLf
        sql &= " ,a.RID,a.ORGID,a.ORGNAME,a.DISTNAME" & vbCrLf
        sql &= " ,a.CNT1,a.CNT2" & vbCrLf
        sql &= " FROM WC2 a" & vbCrLf
        sql &= " WHERE (a.CNT1>0 OR a.CNT2>0)" & vbCrLf

        'parms.Add("DISTID", sm.UserInfo.DistID)
        Dim parms As New Hashtable From {{"YEARS", vYEARS}, {"TPLANID", sm.UserInfo.TPlanID}}
        If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR)
        'If v_ORGID <> "" Then parms.Add("ORGID", v_ORGID)
        If v_ORGNAMElk <> "" Then parms.Add("ORGNAMElk", v_ORGNAMElk)

        'If TIMS.sUtl_ChkTest() Then
        '    TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms), vbCrLf, "--##CO_01_002.aspx, sql:", vbCrLf, sql))
        'End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        'PageControler1.Visible = False
        'DataGridTable.Visible = False
        'msg1.Text = "查無資料"
        If dt.Rows.Count = 0 Then Exit Sub

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Sub SSaveData1()
        If Hid_ORGID.Value = "" Then Exit Sub
        Dim iORGID As Integer = Val(TIMS.ClearSQM(Hid_ORGID.Value))

        Dim parms As New Hashtable
        Dim sSql As String = ""
        sSql = " SELECT PTYOID FROM ORG_PARTYORG WHERE ORGID=@ORGID and PTYID=@PTYID" & vbCrLf

        Dim iSql As String = ""
        iSql &= " INSERT INTO ORG_PARTYORG (PTYOID,ORGID,PTYID,MODIFYACCT,MODIFYDATE)" & vbCrLf
        iSql &= " VALUES(@PTYOID,@ORGID,@PTYID,@MODIFYACCT,GETDATE())" & vbCrLf

        Dim uSql As String = ""
        uSql &= " UPDATE ORG_PARTYORG" & vbCrLf
        uSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        'uSql &= " WHERE ORGID=@ORGID and PTYID=@PTYID" & vbCrLf
        uSql &= " WHERE PTYOID=@PTYOID" & vbCrLf

        Dim dSql As String = ""
        dSql &= " DELETE ORG_PARTYORG" & vbCrLf
        dSql &= " WHERE ORGID=@ORGID and PTYID=@PTYID" & vbCrLf
        '--insert

        'ORG_PARTYORG
        For Each eItem As ListItem In chkbPTYID.Items
            Dim iPTYID As Integer = Val(eItem.Value)
            parms.Clear()
            parms.Add("ORGID", iORGID)
            parms.Add("PTYID", iPTYID)
            Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, parms)
            If eItem.Selected Then
                'insert/update
                If dt.Rows.Count = 0 Then
                    Dim iPTYOID As Integer = DbAccess.GetNewId(objconn, "ORG_PARTYORG_PTYOID_SEQ,ORG_PARTYORG,PTYOID")
                    parms.Clear()
                    parms.Add("PTYOID", iPTYOID)
                    parms.Add("ORGID", iORGID)
                    parms.Add("PTYID", iPTYID)
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(iSql, objconn, parms)
                Else
                    Dim iPTYOID As Integer = Val(dt.Rows(0)("PTYOID"))
                    parms.Clear()
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    'parms.Add("ORGID", iORGID)
                    'parms.Add("PTYID", iPTYID)
                    parms.Add("PTYOID", iPTYOID)
                    DbAccess.ExecuteNonQuery(uSql, objconn, parms)
                End If
            Else
                If dt.Rows.Count > 0 Then
                    'delete
                    parms.Clear()
                    parms.Add("ORGID", iORGID)
                    parms.Add("PTYID", iPTYID)
                    DbAccess.ExecuteNonQuery(dSql, objconn, parms)
                End If
            End If
        Next

    End Sub

    Sub SClearlist1()
        Hid_YEARS.Value = ""
        Hid_DISTID.Value = ""
        Hid_TPLANID.Value = ""
        Hid_HALFYEAR.Value = ""
        Hid_ORGID.Value = ""

        LabOrgName.Text = "" 'Convert.ToString(dr("OrgName"))
        LabPartyYears.Text = "" ' Convert.ToString(dr("YEARS"))
        LabhalfYear.Text = "" 'vHALFYEAR 'Convert.ToString(dr("HALFYEAR"))
        LabShouldTimes.Text = "" 'Convert.ToString(dr("CNT1"))
        LabActTimes.Text = "" 'Convert.ToString(dr("CNT2"))
        LabPrjDegree.Text = "" '"0%" 'Convert.ToString(dr("PrjDegree"))

    End Sub

    Function Get_dtPARTY(ByRef parms As Hashtable) As DataTable
        Dim vHALFYEAR As String = TIMS.GetMyValue2(parms, "HALFYEAR")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")

        Dim sql As String = ""
        sql &= " SELECT PTYID ,PTYNAME"
        sql &= " FROM ORG_PARTY WITH(NOLOCK)"
        sql &= " WHERE YEARS=@YEARS"
        sql &= " AND DISTID=@DISTID"
        sql &= " AND TPLANID=@TPLANID"
        If vHALFYEAR <> "" Then sql &= " AND HALFYEAR=@HALFYEAR"
        Select Case vORGKIND2
            Case "G"
                sql &= " AND PLANSUB IN (1,3)"
            Case "W"
                sql &= " AND PLANSUB IN (2,3)"
        End Select
        sql &= " ORDER BY PTYID,PTYNAME"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    Function Get_ddlPARTY(ByVal obj As CheckBoxList, ByRef parms As Hashtable) As CheckBoxList
        Dim dt As DataTable = Get_dtPARTY(parms)
        With obj
            .Items.Clear()
            .DataSource = dt
            .DataTextField = "PTYNAME"
            .DataValueField = "PTYID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", 0))
        End With
        Return obj
    End Function

    Function Get_PrjDegree(ByVal dbCNT1 As Double, ByVal dbCNT2 As Double) As String
        Dim rst As String = "0%"
        If dbCNT1 <> 0 AndAlso dbCNT2 <> 0 Then
            rst = TIMS.ROUND(Convert.ToString((dbCNT2 / dbCNT1) * 100), 2) & "%"
        End If
        Return rst
    End Function

    Sub SShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub

        Const cst_上半年度 As String = "上半年度"
        Const cst_下半年度 As String = "下半年度"
        Dim txtHALFYEAR As String = ""
        Select Case Convert.ToString(dr("HALFYEAR"))
            Case "1"
                txtHALFYEAR = cst_上半年度
            Case "2"
                txtHALFYEAR = cst_下半年度
        End Select

        Hid_YEARS.Value = Convert.ToString(dr("YEARS")) '""
        Hid_DISTID.Value = Convert.ToString(dr("DISTID")) ' ""
        Hid_TPLANID.Value = Convert.ToString(dr("TPLANID")) '""
        Hid_HALFYEAR.Value = Convert.ToString(dr("HALFYEAR")) '""
        Hid_ORGID.Value = Convert.ToString(dr("ORGID")) '""

        LabOrgName.Text = Convert.ToString(dr("OrgName"))
        LabPartyYears.Text = Convert.ToString(dr("YEARS_ROC"))
        LabhalfYear.Text = txtHALFYEAR 'Convert.ToString(dr("HALFYEAR"))
        LabShouldTimes.Text = Convert.ToString(dr("CNT1"))
        LabActTimes.Text = Convert.ToString(dr("CNT2"))

        Dim dbCNT1 As Double = TIMS.VAL1(dr("CNT1"))
        Dim dbCNT2 As Double = TIMS.VAL1(dr("CNT2"))
        LabPrjDegree.Text = "0%" 'Convert.ToString(dr("PrjDegree"))
        If dbCNT1 <> 0 AndAlso dbCNT2 <> 0 Then
            Dim vPrjDegree As String = TIMS.ROUND(Convert.ToString((dbCNT2 / dbCNT1) * 100), 2) & "%"
            LabPrjDegree.Text = vPrjDegree '"0%" 'Convert.ToString(dr("PrjDegree"))
        End If

        'btnRecal.Text = Convert.ToString(dr("HALFYEAR"))
        'chkbPTYID.Text = Convert.ToString(dr("HALFYEAR"))

        Dim vYEARS As String = Convert.ToString(dr("YEARS"))
        Dim vDISTID As String = Convert.ToString(dr("DISTID"))
        Dim vTPLANID As String = Convert.ToString(dr("TPLANID"))
        Dim vHALFYEAR As String = Convert.ToString(dr("HALFYEAR"))
        Dim vORGKIND2 As String = Convert.ToString(dr("ORGKIND2"))
        'Dim vOrgID As String = Convert.ToString(dr("OrgID"))

        'parms.Clear()
        Dim parms As New Hashtable From {
            {"YEARS", vYEARS},
            {"DISTID", vDISTID}, 'sm.UserInfo.DistID)
            {"TPLANID", vTPLANID}, ' sm.UserInfo.TPlanID)
            {"HALFYEAR", vHALFYEAR},
            {"ORGKIND2", vORGKIND2}
        }
        'parms.Add("OrgID", vOrgID)
        chkbPTYID = Get_ddlPARTY(chkbPTYID, parms)
        chkbPTYID_hid.Value = "0"
        chkbPTYID.Attributes("onclick") = "SelectAll('chkbPTYID','chkbPTYID_hid');"

        Dim iORGID As Integer = Val(Hid_ORGID.Value)
        For Each eItem As ListItem In chkbPTYID.Items
            Dim iPTYID As Integer = Val(eItem.Value)
            Dim parms2 As New Hashtable From {{"ORGID", iORGID}, {"PTYID", iPTYID}}
            Dim sSql2 As String = " SELECT PTYOID FROM ORG_PARTYORG WHERE ORGID=@ORGID AND PTYID=@PTYID" & vbCrLf
            Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, parms2)
            eItem.Selected = TIMS.dtHaveDATA(dt2)
            'If dt.Rows.Count > 0 Then eItem.Selected = True
        Next

    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Call SClearlist1()

        'Dim CSCID As String = TIMS.GetMyValue(sCmdArg, "CSCID")
        'Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        'Dim vOrgID As String = TIMS.GetMyValue(sCmdArg, "OrgID")
        'Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        'Dim PlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        Dim vDISTID As String = TIMS.GetMyValue(sCmdArg, "DISTID")
        Dim vTPLANID As String = TIMS.GetMyValue(sCmdArg, "TPLANID")
        Dim vHALFYEAR As String = TIMS.GetMyValue(sCmdArg, "HALFYEAR") '1:上年度 /2:下年度
        Dim vOrgID As String = TIMS.GetMyValue(sCmdArg, "OrgID")

        'Dim ACT As String = TIMS.GetMyValue(sCmdArg, "ACT")
        'Select Case e.CommandName
        '    Case cst_btnAddt
        '        sLoadData1(sCmdArg)
        '    Case cst_btnEdit
        '        sLoadData1(sCmdArg)
        'End Select
        SLoadData1(sCmdArg)

    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "DISTID", Convert.ToString(drv("DISTID")))
                TIMS.SetMyValue(sCmdArg, "TPLANID", Convert.ToString(drv("TPLANID")))
                TIMS.SetMyValue(sCmdArg, "HALFYEAR", Convert.ToString(drv("HALFYEAR")))
                TIMS.SetMyValue(sCmdArg, "OrgID", Convert.ToString(drv("OrgID")))

                Dim dbCNT1 As Double = TIMS.VAL1(drv("CNT1"))
                Dim dbCNT2 As Double = TIMS.VAL1(drv("CNT2"))
                Dim iLabPrjDegree As Label = e.Item.FindControl("iLabPrjDegree")
                iLabPrjDegree.Text = Get_PrjDegree(dbCNT1, dbCNT2) '"0%" 'Convert.ToString(dr("PrjDegree"))
                'Dim sAct As String = cst_btnEdit
                'lbtEdit.Text = "編輯"
                'lbtEdit.CommandName = cst_btnEdit
                'TIMS.SetMyValue(sCmdArg, "ACT", sAct)
                lbtEdit.CommandArgument = sCmdArg
        End Select

    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        divSch1.Visible = True
        divEdt1.Visible = False
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click

        SSaveData1()

        divSch1.Visible = True
        divEdt1.Visible = False
        SSearch1()
        sm.LastResultMessage = "儲存完畢"

    End Sub

    Function ChangeImportDate(ByVal colArray As Array) As Array
        'Const cst_000 As String = "00000000"
        '廠商統一編號
        'colArray(Cst_ComIDNO) = Right(cst_000 & colArray(Cst_ComIDNO).ToString, 8)
        Return colArray
    End Function

    Function CheckImportData1(ByRef colArray As Array) As String
        Const cst_必須填寫 As String = "必須填寫"
        Dim Reason As String = ""
        If colArray.Length < cst_filedColumnNum Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)<BR>"
            Reason &= "欄位對應有誤<BR>"
            Reason &= "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        aYears = TIMS.ClearSQM(colArray(Cst_Years)) '年度
        aHALFYEAR = TIMS.ClearSQM(colArray(Cst_HALFYEAR)) '上下半年
        aPLANSUB = TIMS.ClearSQM(colArray(Cst_PLANSUB)) '「計畫別」：1.產業人才投資計畫、2.提升勞工自主學習計畫、3.不區分
        aPTYDATE = TIMS.ClearSQM(colArray(Cst_PTYDATE)) '辦理活動場次日期
        aPTYNAME = TIMS.ClearSQM(colArray(Cst_PTYNAME)) '活動場次名稱
        If aYears <> "" Then
            If Not TIMS.IsNumeric2(aYears) Then
                Reason += "年度必須為正確的數字格式<BR>"
            End If
        Else
            Reason += cst_必須填寫 & "年度<Br>"
        End If

        If aHALFYEAR <> "" Then
            If Not TIMS.IsNumeric2(aHALFYEAR) Then
                Reason += "上下半年必須為正確的數字格式<BR>"
            End If
            Select Case aHALFYEAR
                Case "1", "2"
                Case Else
                    Reason += "上下半年只能為1:上半年/2:下半年<BR>"
            End Select
        Else
            Reason += cst_必須填寫 & "上下半年<Br>"
        End If

        If aPLANSUB <> "" Then
            If Not TIMS.IsNumeric2(aPLANSUB) Then
                Reason += "計畫別 必須為正確的數字格式<BR>"
            End If
            Select Case aPLANSUB
                Case "1", "2", "3"
                Case Else
                    Reason += "計畫別 只能為 1:產業人才投資計畫／2:提升勞工自主學習計畫／3:不區分<BR>"
            End Select
        Else
            Reason += cst_必須填寫 & "計畫別<Br>"
        End If

        '辦理活動場次日期
        If aPTYDATE <> "" Then
            If Not TIMS.IsDate1(aPTYDATE) Then
                Reason += "辦理活動場次日期 必須為正確的日期格式(yyyy/MM/dd)<BR>"
            Else
                aPTYDATE = TIMS.Cdate3(aPTYDATE)
            End If
        Else
            Reason += cst_必須填寫 & "辦理活動場次日期<Br>"
        End If

        '活動場次名稱
        If aPTYNAME <> "" Then
            If aPTYNAME.Length > 100 Then
                Reason += "活動場次名稱字串長度不可超過100<BR>"
            End If
        Else
            Reason += cst_必須填寫 & "活動場次名稱<Br>"
        End If

        If Reason <> "" Then Return Reason

        '2018	28	000	2	活動2_19
        'ORG_PARTY 'YEARS,TPLANID,DISTID,HALFYEAR,PTYNAME
        Dim sql As String = ""
        sql &= " SELECT 'X' FROM ORG_PARTY"
        sql &= " WHERE YEARS=@YEARS AND TPLANID=@TPLANID AND DISTID=@DISTID"
        sql &= " AND HALFYEAR=@HALFYEAR AND PTYNAME=@PTYNAME"
        Dim parms As New Hashtable From {
            {"YEARS", aYears},
            {"TPLANID", sm.UserInfo.TPlanID},
            {"DISTID", sm.UserInfo.DistID},
            {"HALFYEAR", aHALFYEAR},
            {"PTYNAME", aPTYNAME}
        }
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then
            '已匯入
            Reason += "活動場次名稱已重複<Br>"
        End If
        Return Reason
    End Function

    Sub SImport1(ByRef FullFileName1 As String)
        '上傳檔案
        File1.PostedFile.SaveAs(FullFileName1)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            Const cst_FirstCol1 As String = "年度"
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName1, "", Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then
            dt_xls = TIMS.GetDataTable_ODSFile(FullFileName1)
        End If
        '刪除檔案
        'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName1)
        Reason = TIMS.Chk_DTXLS1(dt_xls, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If
        'If dt_xls.Rows.Count = 0 Then
        '    Common.MessageBox(Me, "資料有誤，故無法匯入，請修正匯入檔案，謝謝")
        '    Exit Sub
        'End If

        'xls 方式 讀取寫入資料庫
        '將檔案讀出放入記憶體

        '取出資料庫的所有欄位--------   Start
        'Dim da As SqlDataAdapter = Nothing
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        'Dim Reason As String = "" '儲存錯誤的原因
        Dim iRowIndex As Integer = 1
        'Dim colArray As Array
        For Each dr1 As DataRow In dt_xls.Rows
            Reason = ""
            'colArray = dt.Rows(i).ItemArray
            Dim colArray As Array = dr1.ItemArray
            '轉換正確欄位值
            'colArray = ChangeImportDate(colArray)
            '檢查正確欄位值
            Reason += CheckImportData1(colArray)

            '通過檢查，開始輸入資料---------------------Start
            If Reason = "" Then
                Dim iPTYID As Integer = DbAccess.GetNewId(objconn, "ORG_PARTY_PTYID_SEQ,ORG_PARTY,PTYID")
                Dim sql As String = ""
                sql &= " INSERT INTO ORG_PARTY(PTYID,YEARS,TPLANID,DISTID,HALFYEAR,PLANSUB,PTYDATE,PTYNAME,MODIFYACCT,MODIFYDATE)" & vbCrLf
                sql &= " VALUES (@PTYID,@YEARS,@TPLANID,@DISTID,@HALFYEAR,@PLANSUB,@PTYDATE,@PTYNAME,@MODIFYACCT,GETDATE())" & vbCrLf
                Dim parms As New Hashtable From {
                    {"PTYID", iPTYID},
                    {"YEARS", aYears},
                    {"TPLANID", sm.UserInfo.TPlanID},
                    {"DISTID", sm.UserInfo.DistID},
                    {"HALFYEAR", aHALFYEAR},
                    {"PLANSUB", aPLANSUB},
                    {"PTYDATE", TIMS.Cdate2(aPTYDATE)},
                    {"PTYNAME", aPTYNAME},
                    {"MODIFYACCT", sm.UserInfo.UserID}
                }
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = iRowIndex
                drWrong("Reason") = Reason
            End If
            iRowIndex += 1
        Next

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        Dim explain2 As String = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        '開始判別欄位存入------------   End
        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, explain)
            Exit Sub
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視失敗原因?')){window.open('CO_01_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")

    End Sub

    ''' <summary>匯入</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Btn_XlsImport1_Click(sender As Object, e As EventArgs) Handles Btn_XlsImport1.Click
        Dim sMyFileName As String = ""
        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SImport1(FullFileName1)
    End Sub

    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Call SClearlist1()
        Call SSearch1()
    End Sub

    Function CheckImportData2(ByRef colArray As Array) As String
        'Const cst_filedNum = 8
        Const cst_必須填寫 As String = "必須填寫"
        Dim Reason As String = ""
        'Dim sql As String = ""
        'Dim dr As DataRow = Nothing
        If colArray.Length < cst_filedColumnNum2 Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        Dim v_COMIDNO As String = ""
        v_COMIDNO = colArray(Cst_COMIDNO).ToString
        v_COMIDNO = TIMS.ClearSQM(v_COMIDNO)
        If v_COMIDNO <> "" Then
            '使用匯入統編取得機構代碼ID--ORGID
            Dim t_OrgID As String = TIMS.Get_OrgIDforComIDNO(objconn, v_COMIDNO)
            If t_OrgID <> "" Then aORGID = t_OrgID
        End If

        If aORGID = "" Then
            Reason += String.Concat("(統編資料)輸入有誤,", v_COMIDNO, "<BR>")
            Return Reason
        End If
        'If gdtParty.Rows.Count = 0 Then
        '    Reason += "輸入資料有誤<BR>"
        '    Return Reason
        'End If
        'Dim aYears As String = colArray(Cst_Years).ToString '年度
        'Dim aHALFYEAR As String = colArray(Cst_HALFYEAR).ToString '上下半年
        'Dim aPTYNAME As String = colArray(Cst_PTYNAME).ToString '活動場次名稱
        aPTYID = colArray(Cst_PTYID).ToString
        aPTYID = TIMS.ClearSQM(aPTYID)
        If aPTYID <> "" Then
            If Not TIMS.IsNumeric2(aPTYID) Then
                Reason += String.Concat("(場次代碼)必須為正確的數字格式,", aPTYID, "<BR>")
            End If
        Else
            Reason += cst_必須填寫 & "場次代碼<Br>"
        End If
        If Reason <> "" Then Return Reason

        '2018	28	000	2	活動2_19
        'ORG_PARTY 'YEARS,TPLANID,DISTID,HALFYEAR,PTYNAME
        'parms.Clear()
        Dim parms As New Hashtable From {{"ORGID", aORGID}, {"PTYID", aPTYID}}
        Dim sql As String = " SELECT 'X' FROM ORG_PARTYORG WHERE ORGID=@ORGID AND PTYID=@PTYID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then
            '已匯入
            Reason += "該機構活動場次代碼已重複(已匯入)<Br>"
        End If
        Return Reason
    End Function

    Sub SImport2(ByRef FullFileName1 As String)
        '上傳檔案
        File2.PostedFile.SaveAs(FullFileName1)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            Const cst_FirstCol1 As String = "場次代碼"
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName1, "", Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then
            dt_xls = TIMS.GetDataTable_ODSFile(FullFileName1)
        End If
        '刪除檔案
        'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName1)
        Reason = TIMS.Chk_DTXLS1(dt_xls, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        'xls 方式 讀取寫入資料庫
        '將檔案讀出放入記憶體

        '取出資料庫的所有欄位--------   Start
        'Dim da As SqlDataAdapter = Nothing
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        Dim iRowIndex As Integer = 1
        'Dim colArray As Array
        For Each dr1 As DataRow In dt_xls.Rows
            Reason = ""
            'colArray = dt.Rows(i).ItemArray
            Dim colArray As Array = dr1.ItemArray
            '轉換正確欄位值
            'colArray = ChangeImportDate(colArray)
            '檢查正確欄位值
            Reason += CheckImportData2(colArray)

            '通過檢查，開始輸入資料---------------------Start
            If Reason = "" Then
                'insert
                Dim iPTYOID As Integer = DbAccess.GetNewId(objconn, "ORG_PARTYORG_PTYOID_SEQ,ORG_PARTYORG,PTYOID")
                'Dim iORGID As Integer = Val(Orgidvalue.Value)
                'Exit Sub
                'parms.Clear()
                Dim parms As New Hashtable From {{"PTYOID", iPTYOID}, {"ORGID", aORGID}, {"PTYID", aPTYID}, {"MODIFYACCT", sm.UserInfo.UserID}}
                Dim sql As String = ""
                sql &= " INSERT INTO ORG_PARTYORG(PTYOID,ORGID,PTYID,MODIFYACCT,MODIFYDATE)" & vbCrLf
                sql &= " VALUES (@PTYOID,@ORGID,@PTYID,@MODIFYACCT,GETDATE())" & vbCrLf
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = iRowIndex
                drWrong("Reason") = Reason
            End If

            iRowIndex += 1
        Next

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        '開始判別欄位存入------------   End
        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, explain)
            Exit Sub
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視失敗原因?')){window.open('CO_01_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")

    End Sub

    Function CheckImpData2(ByRef sErrMsg2 As String) As Boolean
        Dim rst As Boolean = True
        sErrMsg2 = ""

        'If RIDValue.Value.Length = 1 Then
        '    sErrMsg2 &= "請先選擇有效的訓練機構!" & vbCrLf
        '    Return False
        'End If
        Select Case sm.UserInfo.LID
            Case 0
                sErrMsg2 &= "登入機構與業務權限有誤!(署，暫不提供匯入)" & vbCrLf
            Case 1
                Dim s_RIDValue As String = sm.UserInfo.RID
                '確認RID 業務權限 與PLANID 計畫 為登入權限 sm.UserInfo.PlanID 
                Dim flagOK As Boolean = TIMS.CheckRIDsPLAN(Me, s_RIDValue, objconn)
                If Not flagOK Then sErrMsg2 += "登入機構與業務權限有誤!(分署查無計畫資訊!)" & vbCrLf
            Case 2
                sErrMsg2 &= "登入機構與業務權限有誤!(委訓單位，暫不提供匯入)" & vbCrLf
            Case Else
                sErrMsg2 &= "登入機構與業務!(權限有誤，暫不提供匯入)" & vbCrLf
        End Select
        If sErrMsg2 <> "" Then Return False

        'aORGID = TIMS.Get_OrgID(RIDValue.Value, objconn)
        'Dim iORGID As Integer = TIMS.VAL1(aORGID)
        'If aORGID = "" OrElse iORGID = 0 OrElse iORGID = -1 Then sErrMsg2 += "登入權限有誤，請選擇機構" & vbCrLf
        'If sErrMsg2 <> "" Then Return False
        'If aORGID = -1 Then sErrMsg2 += "請優先選擇機構" & vbCrLf
        'If sErrMsg2 <> "" Then Return False

        'Dim drR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        Dim drR As DataRow = TIMS.GetPlanID1(sm.UserInfo.PlanID, objconn)
        If drR Is Nothing Then
            sErrMsg2 &= "登入計畫與業務權限有誤!(查無計畫資訊)" & vbCrLf
            Return False
        End If

        Dim parms As New Hashtable From {{"YEARS", drR("YEARS")}, {"DISTID", drR("DISTID")}, {"TPLANID", drR("TPLANID")}}
        gdtParty = Get_dtPARTY(parms)
        If gdtParty.Rows.Count = 0 Then
            sErrMsg2 &= "(登入計畫與業務權限有誤)尚未匯入總場次!" & vbCrLf
            Return False
        End If

        Return rst
    End Function

    ''' <summary>匯入</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Btn_XlsImport2_Click(sender As Object, e As EventArgs) Handles Btn_XlsImport2.Click
        Dim sErrMsg2 As String = ""
        aORGID = ""
        CheckImpData2(sErrMsg2)
        If sErrMsg2 <> "" Then
            aORGID = ""
            Common.MessageBox(Me, sErrMsg2)
            Exit Sub
        End If

        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File2, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "xls", 1) Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "ods", 1) Then Return
        End If

        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File2.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        '(匯入)
        Call SImport2(FullFileName1)
    End Sub

    Function Get_dtPARTY2(ByRef parms As Hashtable) As DataTable
        'parms.Add("YEARS", vYEARS)
        'parms.Add("DISTID", vsDistID)
        'parms.Add("TPLANID", sm.UserInfo.TPlanID)
        'parms.Add("HALFYEAR", vHALFYEAR)

        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        vDISTID = TIMS.CombiSQM2IN(vDISTID)
        Dim vHALFYEAR As String = TIMS.GetMyValue2(parms, "HALFYEAR")
        Dim sql As String = ""
        sql &= " SELECT PTYID 場次代碼"
        sql &= " ,PTYNAME 活動場次名稱"
        sql &= " ,YEARS 年度"
        sql &= " ,HALFYEAR 上下半年"
        sql &= " ,''''+DISTID 分署代碼"
        sql &= " ,''''+TPLANID 計畫代碼"
        sql &= " FROM ORG_PARTY"
        sql &= " WHERE YEARS=@YEARS"
        If vDISTID <> "" Then sql &= " And DISTID in (" & vDISTID & ")"
        sql &= " And TPLANID=@TPLANID"
        If vHALFYEAR <> "" Then sql &= " And HALFYEAR=@HALFYEAR"
        sql &= " ORDER BY PTYID,PTYNAME"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    ''' <summary>匯出場次代碼-SUB</summary>
    Sub SExprot2()
        '匯出excel 
        Dim sFileName1 As String = "匯出總場次"

        Dim dtXls As DataTable = Nothing
        Dim vYEARS As String = TIMS.GetListValue(SYEARlist) '.SelectedValue)
        Dim vHALFYEAR As String = TIMS.GetListValue(halfYear) '.SelectedValue) '1:上年度 /2:下年度

        Dim vsDistID As String = TIMS.GetCblValue(sDistID)
        vsDistID = TIMS.CombiSQM2IN(vsDistID) 'SQL使用
        'Select Case sm.UserInfo.LID
        '    Case 0
        '    Case Else
        'End Select

        Dim parms As New Hashtable From {{"YEARS", vYEARS}, {"DISTID", vsDistID}, {"TPLANID", sm.UserInfo.TPlanID}, {"HALFYEAR", vHALFYEAR}}
        dtXls = Get_dtPARTY2(parms)
        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '匯出excel 
        'Dim filename As String = "ExpFile" & TIMS.GetRnd6Eng
        'Dim strfileext As String = ".xls"
        'HttpContext.Current.Response.ContentType = "application/vnd.ms-excel"
        'HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5")
        'HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment; filename=" & filename & strfileext)
        'HttpContext.Current.Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        '先把分頁關掉
        Dim GridView1 As New GridView With {
            .AllowPaging = False,
            .DataSource = dtXls
        }
        GridView1.DataBind()

        'Get the HTML for the control.
        Dim tw As IO.StringWriter = New IO.StringWriter()
        Dim hw As HtmlTextWriter = New HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Dim strHTML As String = ""
        strHTML &= (tw.ToString())

        'Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
        'parmsExp.Add("strSTYLE", strSTYLE)
        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
            {"FileName", sFileName1},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'TIMS.CloseDbConn(objconn)
        'HttpContext.Current.Response.End()
        'GridView1.AllowPaging = True
        'bindgv()
    End Sub

    ''' <summary>匯出場次代碼</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles btnExp1.Click
        Call SExprot2()
    End Sub

    Sub KEEP_SEARCH_VAL1()
        Dim vsDistID As String = TIMS.GetCblValue(sDistID)
        'vsDistID = TIMS.CombiSQM2IN(vsDistID) 'SQL使用
        Dim vYEARS As String = TIMS.GetListValue(SYEARlist) '.SelectedValue)
        Dim vHALFYEAR As String = TIMS.GetListValue(halfYear) '.SelectedValue) '1:上年度 /2:下年度

        Dim SchVal1 As String = ""
        SchVal1 &= "&PRG=CO_01_002_OP"
        SchVal1 &= "&DISTID=" & vsDistID
        SchVal1 &= "&YEARS=" & vYEARS
        SchVal1 &= "&HALFYEAR=" & vHALFYEAR
        Session("CO_01_002_OP_SCH") = SchVal1
    End Sub
    ''' <summary>查詢總場次</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSchOP1_Click(sender As Object, e As EventArgs) Handles BtnSchOP1.Click
        'Dim v_SYEARlist As String = TIMS.GetListValue(SYEARlist)
        'Dim v_halfYear As String = TIMS.GetListValue(halfYear)
        Call KEEP_SEARCH_VAL1()

        Dim MRQ_ID As String = TIMS.Get_MRqID(Me)
        Dim url1 As String = "CO_01_002_OP.aspx?ID=" & MRQ_ID

        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

End Class
