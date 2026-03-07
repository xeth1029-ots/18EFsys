Public Class CO_01_001
    Inherits AuthBasePage 'System.Web.UI.Page

    '審查計分表排程 '排程[Co_OrgScoring] Co_OrgScoring.exe.config 
    '\SVN\WDAIIP\SRC\Batch\Co_OrgScoring
    'CLASS_SCORE

    Const cst_btnEdit As String = "btnEdit"
    Const cst_btnAddt As String = "btnAddt"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        '產投/非產投判斷
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Me.LabTMID.Text = "訓練業別"

        If Not IsPostBack Then
            CCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            'HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub CCreate1()
        divSch1.Visible = True
        divEdt1.Visible = False
        msg1.Text = ""
        PageControler1.Visible = False
        DataGridTable.Visible = False
        '(加強操作便利性)
        RIDValue.Value = sm.UserInfo.RID
        center.Text = sm.UserInfo.OrgName
        ddlISPASS1 = TIMS.GET_ddlISPASSCNT_N(ddlISPASS1)
        ddlISPASS2 = TIMS.GET_ddlISPASSCNT_N(ddlISPASS2)
        ddlISPASS3 = TIMS.GET_ddlISPASSCNT_N(ddlISPASS3)

        '申請階段
        'Dim v_APPSTAGE As String = If(Now.Month < 7, "1", "2")
        Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_01)
        sch_ddlAPPSTAGE = TIMS.Get_APPSTAGE2(sch_ddlAPPSTAGE)
        Common.SetListItem(sch_ddlAPPSTAGE, v_APPSTAGE)
    End Sub

    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Call sClearlist1()
        Call sSearch1()
    End Sub

    ''' <summary>查詢單筆資料</summary>
    ''' <param name="sCmdArg"></param>
    Sub SLoadData1(ByRef sCmdArg As String)
        'Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim CSCID As String = TIMS.GetMyValue(sCmdArg, "CSCID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        'Dim OrgID As String = TIMS.GetMyValue(sCmdArg, "OrgID")
        'Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        'Dim PlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        'Dim DistID As String = TIMS.GetMyValue(sCmdArg, "DistID")
        'Dim ACT As String = TIMS.GetMyValue(sCmdArg, "ACT")

        Dim parms As New Hashtable From {{"TPlanID", sm.UserInfo.TPlanID}, {"Years", sm.UserInfo.Years}}
        Dim sql As String = ""
        sql &= " SELECT a.CSCID" & vbCrLf
        'sql &= " ,a.DATAKIND" & vbCrLf
        sql &= " ,a.SENDACCT1,a.SENDDATE1,a.STATUS1,a.ISPASS1,a.MODIFYACCT1,a.MODIFYDATE1" & vbCrLf
        sql &= " ,a.SENDACCT2,a.SENDDATE2,a.STATUS2,a.ISPASS2,a.MODIFYACCT2,a.MODIFYDATE2" & vbCrLf
        sql &= " ,a.SENDACCT3,a.SENDDATE3,a.STATUS3,a.ISPASS3,a.MODIFYACCT3,a.MODIFYDATE3" & vbCrLf
        sql &= " ,a.OVERWEEK1 ,a.OVERWEEK2 ,a.OVERWEEK3" & vbCrLf

        sql &= " ,CONVERT(varchar,cc.STDATE,111) STDATE" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.FTDATE,111) FTDATE" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.STDATE+14,111) STDATE14" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.FTDATE+21,111) FTDATE21" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.STDATE,111)+'~'+CONVERT(varchar,cc.FTDATE,111) SFTDATE" & vbCrLf

        sql &= " ,cc.OCID,cc.OrgID,cc.RID,cc.PlanID,cc.DistID" & vbCrLf
        sql &= " ,cc.OrgName,cc.ClassID,cc.CLASSCNAME,cc.CLASSCNAME2,cc.CyclType,cc.CJOBNAME" & vbCrLf
        sql &= " ,ISNULL(cc.TRAINNAME,cc.JOBNAME) JOBNAME" & vbCrLf
        'APPSTAGE_N
        sql &= " ,cc.APPSTAGE,CASE cc.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        sql &= " FROM dbo.VIEW2 cc " & vbCrLf
        sql &= " LEFT JOIN CLASS_SCORE a on a.OCID=cc.OCID" & vbCrLf
        sql &= " WHERE cc.TPlanID=@TPlanID AND cc.Years=@Years" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sql &= " AND cc.PlanID=@PlanID" & vbCrLf
                sql &= " AND cc.DistID=@DistID" & vbCrLf
                parms.Add("PlanID", sm.UserInfo.PlanID)
                parms.Add("DistID", sm.UserInfo.DistID)
        End Select
        If OCID <> "" Then
            parms.Add("OCID", Val(OCID))
            sql &= " AND cc.OCID=@OCID" & vbCrLf
        End If
        If Hid_CSCID.Value <> "" AndAlso CSCID <> "" Then
            parms.Add("CSCID", Val(CSCID))
            sql &= " AND a.CSCID=@CSCID" & vbCrLf
        End If
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
        Call sClearlist1()
        Call sShowData1(dr1)
    End Sub

    ''' <summary>查詢 List</summary>
    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        'cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        STDATE1.Text = TIMS.Cdate3(TIMS.ClearSQM(STDATE1.Text))
        STDATE2.Text = TIMS.Cdate3(TIMS.ClearSQM(STDATE2.Text))
        Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)

        Dim parms As New Hashtable From {{"TPlanID", sm.UserInfo.TPlanID}, {"Years", sm.UserInfo.Years}}
        Dim sql As String = ""
        sql &= " SELECT a.CSCID" & vbCrLf
        'sql &= " ,a.DATAKIND" & vbCrLf
        sql &= " ,a.SENDACCT1,a.SENDDATE1,a.STATUS1,a.ISPASS1,a.MODIFYACCT1,a.MODIFYDATE1" & vbCrLf
        sql &= " ,a.SENDACCT2,a.SENDDATE2,a.STATUS2,a.ISPASS2,a.MODIFYACCT2,a.MODIFYDATE2" & vbCrLf
        sql &= " ,a.SENDACCT3,a.SENDDATE3,a.STATUS3,a.ISPASS3,a.MODIFYACCT3,a.MODIFYDATE3" & vbCrLf
        sql &= " ,a.OVERWEEK1,a.OVERWEEK2,a.OVERWEEK3" & vbCrLf

        sql &= " ,CONVERT(varchar,cc.STDATE,111) STDATE" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.FTDATE,111) FTDATE" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.STDATE+14,111) STDATE14" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.FTDATE+21,111) FTDATE21" & vbCrLf
        sql &= " ,CONVERT(varchar,cc.STDATE,111)+'~'+CONVERT(varchar,cc.FTDATE,111) SFTDATE" & vbCrLf

        sql &= " ,cc.OCID,cc.OrgID,cc.RID,cc.PlanID,cc.DistID" & vbCrLf
        sql &= " ,cc.OrgName,cc.CLASSCNAME,cc.CLASSCNAME2,cc.CJOBNAME,cc.JOBNAME" & vbCrLf
        'APPSTAGE_N
        sql &= " ,cc.APPSTAGE,CASE cc.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        sql &= " FROM dbo.VIEW2 cc " & vbCrLf
        sql &= " LEFT JOIN CLASS_SCORE a on a.OCID=cc.OCID" & vbCrLf
        sql &= " WHERE cc.TPlanID=@TPlanID AND cc.Years=@Years" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sql &= " AND cc.PlanID=@PlanID" & vbCrLf
                sql &= " AND cc.DistID=@DistID" & vbCrLf
                parms.Add("PlanID", sm.UserInfo.PlanID)
                parms.Add("DistID", sm.UserInfo.DistID)
        End Select

        If RIDValue.Value <> "" Then
            sql &= " AND cc.RID='" & RIDValue.Value & "'" & vbCrLf
        End If
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If jobValue.Value <> "" Then
        '        sql &= " AND cc.TMID='" & jobValue.Value & "'" & vbCrLf
        '    End If
        'Else
        '    If trainValue.Value <> "" Then
        '        sql &= " AND cc.TMID='" & trainValue.Value & "'" & vbCrLf
        '    End If
        'End If
        'If Me.txtCJOB_NAME.Text <> "" Then
        '    sql &= " and cc.CJOB_UNKEY='" & Me.cjobValue.Value & "'" & vbCrLf
        'End If
        If ClassName.Text <> "" Then
            sql &= " and cc.CLASSCNAME like '%'+'" & ClassName.Text & "'+'%'" & vbCrLf
        End If
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            sql &= " and cc.CyclType='" & CyclType.Text & "'" & vbCrLf
        End If
        If (STDATE1.Text <> "") Then
            sql &= " and cc.STDATE >=@STDATE1" & vbCrLf
            parms.Add("STDATE1", TIMS.Cdate2(STDATE1.Text))
        End If
        If (STDATE2.Text <> "") Then
            sql &= " and cc.STDATE <=@STDATE2" & vbCrLf
            parms.Add("STDATE2", TIMS.Cdate2(STDATE2.Text))
        End If
        If v_sch_ddlAPPSTAG <> "" Then
            sql &= " AND cc.APPSTAGE=@APPSTAGE" & vbCrLf
            parms.Add("APPSTAGE", v_sch_ddlAPPSTAG)
        End If

        sql &= " ORDER BY cc.STDATE" & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)

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

    Sub sSaveData1()
        If Hid_OCID.Value = "" Then Return

        Dim sSql As String = "SELECT CSCID FROM CLASS_SCORE WHERE OCID=@OCID" & vbCrLf
        Dim i_sql As String = ""
        i_sql &= " INSERT INTO CLASS_SCORE(CSCID,OCID) VALUES (@CSCID,@OCID)" & vbCrLf

        Dim usSql1 As String = ""
        usSql1 &= " UPDATE CLASS_SCORE" & vbCrLf
        usSql1 &= " SET SENDACCT1=@SENDACCT" & vbCrLf
        usSql1 &= " ,SENDDATE1=@SENDDATE" & vbCrLf
        usSql1 &= " ,STATUS1=@STATUS" & vbCrLf
        usSql1 &= " ,ISPASS1=@ISPASS" & vbCrLf
        usSql1 &= " ,OVERWEEK1=@OVERWEEK" & vbCrLf
        usSql1 &= " ,MODIFYACCT1=@MODIFYACCT" & vbCrLf
        usSql1 &= " ,MODIFYDATE1=GETDATE()" & vbCrLf
        usSql1 &= " WHERE CSCID=@CSCID AND OCID=@OCID" & vbCrLf
        Dim usSql2 As String = ""
        usSql2 &= " UPDATE CLASS_SCORE" & vbCrLf
        usSql2 &= " SET SENDACCT2=@SENDACCT" & vbCrLf
        usSql2 &= " ,SENDDATE2=@SENDDATE" & vbCrLf
        usSql2 &= " ,STATUS2=@STATUS" & vbCrLf
        usSql2 &= " ,ISPASS2=@ISPASS" & vbCrLf
        usSql2 &= " ,OVERWEEK2=@OVERWEEK" & vbCrLf
        usSql2 &= " ,MODIFYACCT2=@MODIFYACCT" & vbCrLf
        usSql2 &= " ,MODIFYDATE2=GETDATE()" & vbCrLf
        usSql2 &= " WHERE CSCID=@CSCID AND OCID=@OCID" & vbCrLf
        Dim usSql3 As String = ""
        usSql3 &= " UPDATE CLASS_SCORE" & vbCrLf
        usSql3 &= " SET SENDACCT3=@SENDACCT" & vbCrLf
        usSql3 &= " ,SENDDATE3=@SENDDATE" & vbCrLf
        usSql3 &= " ,STATUS3=@STATUS" & vbCrLf
        usSql3 &= " ,ISPASS3=@ISPASS" & vbCrLf
        usSql3 &= " ,OVERWEEK3=@OVERWEEK" & vbCrLf
        usSql3 &= " ,MODIFYACCT3=@MODIFYACCT" & vbCrLf
        usSql3 &= " ,MODIFYDATE3=GETDATE()" & vbCrLf
        usSql3 &= " WHERE CSCID=@CSCID AND OCID=@OCID" & vbCrLf

        If Hid_CSCID.Value = "" Then
            Dim parms As New Hashtable 'parms.Clear()
            parms.Add("OCID", CInt(Hid_OCID.Value))
            Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, parms)
            Dim dr1 As DataRow = Nothing
            If dt1.Rows.Count > 0 Then
                dr1 = dt1.Rows(0)
                Hid_CSCID.Value = Val(dr1("CSCID"))
            End If
        End If

        '---insert
        If Hid_CSCID.Value = "" Then
            Dim iCSCID As Integer = DbAccess.GetNewId(objconn, "CLASS_SCORE_CSCID_SEQ,CLASS_SCORE,CSCID")
            Hid_CSCID.Value = iCSCID
            Dim i_parms As New Hashtable
            i_parms.Add("CSCID", iCSCID)
            i_parms.Add("OCID", CInt(Hid_OCID.Value))
            DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        End If

        Dim vddlISPASS1 As String = TIMS.GetListValue(ddlISPASS1)
        Dim vddlISPASS2 As String = TIMS.GetListValue(ddlISPASS2)
        Dim vddlISPASS3 As String = TIMS.GetListValue(ddlISPASS3)

        '---updata
        Dim uParms1 As New Hashtable
        uParms1.Add("SENDACCT", sm.UserInfo.UserID)
        uParms1.Add("SENDDATE", TIMS.Cdate2(SENDDATE1.Text))
        uParms1.Add("STATUS", STATUS1.SelectedValue)
        uParms1.Add("ISPASS", If(vddlISPASS1 <> "", vddlISPASS1, Convert.DBNull))
        uParms1.Add("OVERWEEK", OVERWEEK1.SelectedValue)
        uParms1.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms1.Add("CSCID", CInt(Hid_CSCID.Value))
        uParms1.Add("OCID", CInt(Hid_OCID.Value))
        If ChkboxSave_1.Checked Then DbAccess.ExecuteNonQuery(usSql1, objconn, uParms1)

        Dim uParms2 As New Hashtable
        uParms2.Add("SENDACCT", sm.UserInfo.UserID)
        uParms2.Add("SENDDATE", TIMS.Cdate2(SENDDATE2.Text))
        uParms2.Add("STATUS", STATUS2.SelectedValue)
        uParms2.Add("ISPASS", If(vddlISPASS2 <> "", vddlISPASS2, Convert.DBNull))
        uParms2.Add("OVERWEEK", OVERWEEK2.SelectedValue)
        uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms2.Add("CSCID", CInt(Hid_CSCID.Value))
        uParms2.Add("OCID", CInt(Hid_OCID.Value))
        If ChkboxSave_2.Checked Then DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)

        Dim uParms3 As New Hashtable
        uParms3.Add("SENDACCT", sm.UserInfo.UserID)
        uParms3.Add("SENDDATE", TIMS.Cdate2(SENDDATE3.Text))
        uParms3.Add("STATUS", STATUS3.SelectedValue)
        uParms3.Add("ISPASS", If(vddlISPASS3 <> "", vddlISPASS3, Convert.DBNull))
        uParms3.Add("OVERWEEK", OVERWEEK3.SelectedValue)
        uParms3.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms3.Add("CSCID", CInt(Hid_CSCID.Value))
        uParms3.Add("OCID", CInt(Hid_OCID.Value))
        If ChkboxSave_3.Checked Then DbAccess.ExecuteNonQuery(usSql3, objconn, uParms3)
    End Sub

    Sub sClearlist1()
        Hid_CSCID.Value = "" 'Convert.ToString(dr("CSCID")) 'CSCID
        Hid_OCID.Value = "" 'Convert.ToString(dr("OCID")) 'OCID
        'Hid_OrgID.Value = "" 'Convert.ToString(dr("OrgID")) 'OrgID
        'Hid_ACT.Value = Convert.ToString(dr("ACT")) 'ACT
        Hid_stdate.Value = "" 'TIMS.cdate3(dr("stdate"))
        Hid_stdate14.Value = "" 'TIMS.cdate3(dr("stdate14"))
        Hid_ftdate.Value = "" 'TIMS.cdate3(dr("stdate"))
        Hid_ftdate21.Value = "" 'TIMS.cdate3(dr("stdate14"))

        LabOrgName.Text = "" ' Convert.ToString(dr("OrgName"))
        LabClassID.Text = "" 'Convert.ToString(dr("ClassID"))
        LabCLASSCNAME.Text = "" ' Convert.ToString(dr("CLASSCNAME"))
        LabCyclType.Text = "" 'Convert.ToString(dr("CyclType"))
        LabCJOBNAME.Text = "" 'Convert.ToString(dr("CJOBNAME"))
        LabJOBNAME.Text = "" 'Convert.ToString(dr("JOBNAME"))
        LabSFTDATE.Text = ""

        'If Hid_CSCID.Value = "" Then Exit Sub
        ChkboxSave_1.Checked = False
        ChkboxSave_2.Checked = False
        ChkboxSave_3.Checked = False

        SENDDATE1.Text = "" 'TIMS.cdate3(dr("SENDDATE1"))
        STATUS1.SelectedIndex = -1
        'ISPASS1.SelectedIndex = -1
        Common.SetListItem(STATUS1, "")
        Common.SetListItem(ddlISPASS1, "")
        'Common.SetListItem(ISPASS1, "")
        SENDDATE2.Text = "" 'TIMS.cdate3(dr("SENDDATE2"))
        STATUS2.SelectedIndex = -1
        'ISPASS2.SelectedIndex = -1
        Common.SetListItem(STATUS2, "")
        Common.SetListItem(ddlISPASS2, "")
        'Common.SetListItem(ISPASS2, "")
        SENDDATE3.Text = "" ' TIMS.cdate3(dr("SENDDATE3"))
        STATUS3.SelectedIndex = -1
        'ISPASS3.SelectedIndex = -1
        Common.SetListItem(STATUS3, "")
        Common.SetListItem(ddlISPASS3, "")
        'Common.SetListItem(ISPASS3, "")

        OVERWEEK1.SelectedIndex = -1
        OVERWEEK2.SelectedIndex = -1
        OVERWEEK3.SelectedIndex = -1
        Common.SetListItem(OVERWEEK1, "")
        Common.SetListItem(OVERWEEK2, "")
        Common.SetListItem(OVERWEEK3, "")
    End Sub

    Sub sShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub
        Hid_CSCID.Value = Convert.ToString(dr("CSCID")) 'CSCID
        Hid_OCID.Value = Convert.ToString(dr("OCID")) 'OCID
        'Hid_OrgID.Value = Convert.ToString(dr("OrgID")) 'OrgID
        'Hid_ACT.Value = Convert.ToString(dr("ACT")) 'ACT
        Hid_stdate.Value = TIMS.Cdate3(dr("stdate"))
        Hid_stdate14.Value = TIMS.Cdate3(dr("stdate14"))
        Hid_ftdate.Value = TIMS.Cdate3(dr("ftdate"))
        Hid_ftdate21.Value = TIMS.Cdate3(dr("ftdate21"))

        LabOrgName.Text = Convert.ToString(dr("OrgName"))
        LabClassID.Text = Convert.ToString(dr("ClassID"))
        LabCLASSCNAME.Text = Convert.ToString(dr("CLASSCNAME"))
        LabCyclType.Text = Convert.ToString(dr("CyclType"))
        LabCJOBNAME.Text = Convert.ToString(dr("CJOBNAME"))
        LabJOBNAME.Text = Convert.ToString(dr("JOBNAME"))
        LabSFTDATE.Text = Convert.ToString(dr("SFTDATE"))

        If Hid_CSCID.Value = "" Then Exit Sub

        If Convert.ToString(dr("SENDACCT1")) <> "" Then
            SENDDATE1.Text = TIMS.Cdate3(dr("SENDDATE1"))
            Common.SetListItem(STATUS1, Convert.ToString(dr("STATUS1")))
            Common.SetListItem(ddlISPASS1, Convert.ToString(dr("ISPASS1")))
            Common.SetListItem(OVERWEEK1, Convert.ToString(dr("OVERWEEK1")))
        End If
        If Convert.ToString(dr("SENDACCT2")) <> "" Then
            SENDDATE2.Text = TIMS.Cdate3(dr("SENDDATE2"))
            Common.SetListItem(STATUS2, Convert.ToString(dr("STATUS2")))
            Common.SetListItem(ddlISPASS2, Convert.ToString(dr("ISPASS2")))
            Common.SetListItem(OVERWEEK2, Convert.ToString(dr("OVERWEEK2")))
        End If
        If Convert.ToString(dr("SENDACCT3")) <> "" Then
            SENDDATE3.Text = TIMS.Cdate3(dr("SENDDATE3"))
            Common.SetListItem(STATUS3, Convert.ToString(dr("STATUS3")))
            Common.SetListItem(ddlISPASS3, Convert.ToString(dr("ISPASS3")))
            Common.SetListItem(OVERWEEK3, Convert.ToString(dr("OVERWEEK3")))
        End If

    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Call sClearlist1()

        Dim CSCID As String = TIMS.GetMyValue(sCmdArg, "CSCID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim OrgID As String = TIMS.GetMyValue(sCmdArg, "OrgID")
        Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim PlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim DistID As String = TIMS.GetMyValue(sCmdArg, "DistID")
        'Dim ACT As String = TIMS.GetMyValue(sCmdArg, "ACT")

        Select Case e.CommandName
            Case cst_btnAddt
                SLoadData1(sCmdArg)
            Case cst_btnEdit
                SLoadData1(sCmdArg)
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CSCID", Convert.ToString(drv("CSCID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))

                TIMS.SetMyValue(sCmdArg, "OrgID", Convert.ToString(drv("OrgID")))
                TIMS.SetMyValue(sCmdArg, "RID", Convert.ToString(drv("RID")))
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "DistID", Convert.ToString(drv("DistID")))

                Dim sAct As String = cst_btnEdit
                lbtEdit.Text = "編輯"
                lbtEdit.CommandName = cst_btnEdit
                If Convert.ToString(drv("CSCID")) = "" Then
                    sAct = cst_btnAddt
                    lbtEdit.Text = "新增"
                    lbtEdit.CommandName = cst_btnAddt
                End If
                TIMS.SetMyValue(sCmdArg, "ACT", sAct)
                lbtEdit.CommandArgument = sCmdArg
        End Select

    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        divSch1.Visible = True
        divEdt1.Visible = False

    End Sub

    '儲存
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        SENDDATE1.Text = TIMS.Cdate3(SENDDATE1.Text)
        SENDDATE2.Text = TIMS.Cdate3(SENDDATE2.Text)
        SENDDATE3.Text = TIMS.Cdate3(SENDDATE3.Text)

        Call sSaveData1()

        divSch1.Visible = True
        divEdt1.Visible = False
        sm.LastResultMessage = "儲存完畢"

        Call sClearlist1()
        Call sSearch1()
    End Sub

End Class