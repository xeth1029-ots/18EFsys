Partial Class SD_03_013
    Inherits AuthBasePage

    Const cst_vsSearchStr1 As String = "vsSearchStr1"

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn) '開啟連線
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call Create1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    '第1次載入
    Sub Create1()
        DataGridTable1.Visible = False

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Msg1.Text = ""

        'form1'FrameTable'Table1'TitleLab1'TitleLab2'table_sch'center
        'Button8'Button4'Button3'RIDValue'HistoryList2'HistoryRID
        'TMID1'OCID1'OCIDValue1'TMIDValue1'HistoryList'HistoryTable
        SCH_ddlELFORMID = TIMS.Get_ddlELFORMID(SCH_ddlELFORMID, objConn)
        Common.SetListItem(SCH_ddlELFORMID, "1")

        SCH_IDNO.Text = ""
        SCH_CNAME.Text = ""
        Common.SetListItem(SCH_rblSIGN_YN, "")
        ' labPageSize' TxtPageSize' BtnSearch1' 查詢' Msg1' DataGridTable1' DataGrid1' 序號>'CNAME'姓名>'IDNO_MK'身分證號碼>'ENAME'表單文件>
        'SIGN_TXT'是否簽署>'SIGN_TW_TIME'簽署時間>'BtnUCLS2'取消簽名'BtnUCLS2'BtnSHOW1'顯示簽名'BtnSHOW1'BtnDOWNL1'下載簽名文件'BtnDOWNL1'PageControler1

        If sm.UserInfo.LID <> "2" Then
            '若只有管理一個班級，自動協助帶出班級
            Call TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objConn)
        Else
            'Button12_Click(sender, e)
            center.Enabled = False
            Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objConn)
        End If

        USEKEEPSEARCH()

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

    End Sub

    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objConn)
        '判斷機構是否只有一個班級 '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = Convert.ToString(dr("trainname"))
        OCID1.Text = Convert.ToString(dr("classname"))
        TMIDValue1.Value = Convert.ToString(dr("trainid"))
        OCIDValue1.Value = Convert.ToString(dr("ocid"))
    End Sub

    '查詢 - SQL
    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim v_SCH_ELNO As String = TIMS.GetListValue(SCH_ddlELFORMID)
        SCH_IDNO.Text = TIMS.ClearSQM(SCH_IDNO.Text)
        SCH_CNAME.Text = TIMS.ClearSQM(SCH_CNAME.Text)
        Dim v_SIGN_YN As String = TIMS.GetListValue(SCH_rblSIGN_YN)

        Dim ERRMSG As String = ""
        If RIDValue.Value = "" Then ERRMSG &= "請選擇訓練機構!" & vbCrLf
        If OCIDValue1.Value = "" Then ERRMSG &= "請選擇職類/班別!" & vbCrLf
        If v_SCH_ELNO = "" Then ERRMSG &= "請選擇表單文件不可為空!" & vbCrLf
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Return
        End If

        Dim parms As New Hashtable() From {{"YEARS", sm.UserInfo.Years}, {"OCID", OCIDValue1.Value}, {"RID", RIDValue.Value}, {"ELNO", v_SCH_ELNO}}

        Dim sSql As String = ""
        sSql &= " WITH WE1 AS (SELECT x.ELNO, x.ENAME FROM KEY_ELFORM x WHERE x.ELNO=@ELNO)" & vbCrLf
        sSql &= " SELECT cs.OCID,cs.TPlanID,cs.RID,cs.SOCID,cs.SID,cs.ORGKINDGW" & vbCrLf
        sSql &= " ,cs.YEARS,dbo.FN_CYEAR2(cs.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,w1.ELNO,w1.ENAME" & vbCrLf
        sSql &= " ,cs.CLASSCNAME2" & vbCrLf
        sSql &= " ,cs.NAME CNAME" & vbCrLf
        sSql &= " ,cs.IDNO,dbo.FN_GET_MASK1(cs.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,e1.CSELNO,e1.P1_LINK,e1.FILEPATH1 ,e1.SIGNDACCT,e1.SIGNDATE" & vbCrLf
        sSql &= " ,case when e1.ELNO IS NOT NULL then '是' else '否' end SIGN_TXT" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1C(e1.SIGNDATE) SIGNDATE_TWTIME" & vbCrLf
        sSql &= " FROM V_STUDENTINFO cs" & vbCrLf
        sSql &= " CROSS JOIN WE1 w1" & vbCrLf
        sSql &= " LEFT JOIN STUD_ELFORM e1 on e1.SOCID=cs.SOCID and e1.OCID=cs.OCID and e1.ELNO=w1.ELNO" & vbCrLf
        'sSql &= " /* 'sSql &= "" WHERE cs.OCID=120874"" & vbCrLf */" & vbCrLf
        sSql &= " WHERE cs.YEARS=@YEARS AND cs.OCID=@OCID AND cs.RID=@RID" & vbCrLf

        If SCH_IDNO.Text <> "" Then
            parms.Add("IDNO", SCH_IDNO.Text)
            sSql &= " AND cs.IDNO=@IDNO" & vbCrLf
        End If
        If SCH_CNAME.Text <> "" Then
            parms.Add("NAME", SCH_CNAME.Text)
            sSql &= " AND cs.NAME like '%'+@NAME+'%'" & vbCrLf
        End If
        Select Case v_SIGN_YN
            Case "Y"
                sSql &= " AND e1.ELNO IS NOT NULL" & vbCrLf
            Case "N"
                sSql &= " AND e1.ELNO IS NULL" & vbCrLf
        End Select

        DataGridTable1.Visible = False
        Msg1.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, parms)

        If TIMS.dtNODATA(dt) Then Return 'Exit Sub

        DataGridTable1.Visible = True
        Msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        Call Search1()
    End Sub
    Sub USEKEEPSEARCH()
        If Session(cst_vsSearchStr1) Is Nothing Then Return

        Dim strSearchStr1 As String = Session(cst_vsSearchStr1)
        Session(cst_vsSearchStr1) = Nothing

        Dim MyValue As String = TIMS.GetMyValue(strSearchStr1, "prg")
        If MyValue = "SD_03_013" Then
            center.Text = TIMS.GetMyValue(strSearchStr1, "center")
            RIDValue.Value = TIMS.GetMyValue(strSearchStr1, "RIDValue")
            TMID1.Text = TIMS.GetMyValue(strSearchStr1, "TMID1")
            OCID1.Text = TIMS.GetMyValue(strSearchStr1, "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "OCIDValue1")

            MyValue = TIMS.GetMyValue(strSearchStr1, "SCH_ddlELFORMID")
            Common.SetListItem(SCH_ddlELFORMID, MyValue)
            SCH_IDNO.Text = TIMS.GetMyValue(strSearchStr1, "SCH_IDNO")
            SCH_CNAME.Text = TIMS.GetMyValue(strSearchStr1, "SCH_CNAME")
            MyValue = TIMS.GetMyValue(strSearchStr1, "SCH_rblSIGN_YN")
            Common.SetListItem(SCH_rblSIGN_YN, MyValue)

            MyValue = TIMS.GetMyValue(strSearchStr1, "submit")
            If MyValue = "1" Then Call Search1()
        End If
    End Sub
    Sub KEEPSEARCH()
        Session(cst_vsSearchStr1) = Nothing

        Dim xSearchStr As String = "prg=SD_03_013"
        xSearchStr &= "&center=" & TIMS.ClearSQM(center.Text)
        xSearchStr &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        xSearchStr &= "&TMID1=" & TIMS.ClearSQM(TMID1.Text)
        xSearchStr &= "&OCID1=" & TIMS.ClearSQM(OCID1.Text)
        xSearchStr &= "&OCIDValue1=" & TIMS.ClearSQM(OCIDValue1.Value)
        xSearchStr &= "&TMIDValue1=" & TIMS.ClearSQM(TMIDValue1.Value)
        'xSearchStr += "&IDNO=" & TIMS.ChangeIDNO(IDNO.Text))
        xSearchStr &= "&SCH_ddlELFORMID=" & TIMS.GetListValue(SCH_ddlELFORMID)
        xSearchStr &= "&SCH_IDNO=" & TIMS.ClearSQM(SCH_IDNO.Text)
        xSearchStr &= "&SCH_CNAME=" & TIMS.ClearSQM(SCH_CNAME.Text)
        xSearchStr &= "&SCH_rblSIGN_YN=" & TIMS.GetListValue(SCH_rblSIGN_YN)
        xSearchStr &= If(DataGridTable1.Visible, "&submit=1", "&submit=0")

        Session(cst_vsSearchStr1) = xSearchStr
    End Sub

    '取消簽名
    Private Sub UCLS2_SIGN(sCmdArg As String)
        Dim ELNO As String = TIMS.GetMyValue(sCmdArg, "ELNO")
        Dim CSELNO As String = TIMS.GetMyValue(sCmdArg, "CSELNO")
        'Dim ORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        'Dim TPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        'Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        'Dim SID As String = TIMS.GetMyValue(sCmdArg, "SID")
        'Dim YEARS_ROC As String = TIMS.GetMyValue(sCmdArg, "YEARS_ROC", Convert.ToString(drv("YEARS_ROC")))

        'CHECK STUD_ELFORM
        Dim pms1 As New Hashtable From {{"CSELNO", Val(CSELNO)}, {"ELNO", Val(ELNO)}, {"SOCID", Val(SOCID)}, {"OCID", Val(OCID)}}
        Dim sSql As String = " SELECT 1 FROM STUD_ELFORM WHERE CSELNO=@CSELNO AND ELNO=@ELNO AND SOCID=@SOCID AND OCID=@OCID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objConn, pms1)
        If TIMS.dtNODATA(dt1) Then Return

        'BAK STUD_ELFORMDEL
        Dim pms_i As New Hashtable From {{"CSELNO", Val(CSELNO)}, {"ELNO", Val(ELNO)}, {"SOCID", Val(SOCID)}, {"OCID", Val(OCID)}}
        pms_i.Add("MODIFYACCT", sm.UserInfo.UserID)
        Dim sSql_i As String = ""
        sSql_i &= " INSERT INTO STUD_ELFORMDEL(CSELNO, ELNO, SOCID, OCID, IDNO, P1_LINK, CREATEACCT, CREATEDATE, SIGNDACCT, SIGNDATE, SENDACCT, SENDDATE, MODIFYACCT, MODIFYDATE, FILEPATH1)" & vbCrLf
        sSql_i &= " SELECT CSELNO, ELNO, SOCID, OCID, IDNO, P1_LINK, CREATEACCT, CREATEDATE, SIGNDACCT, SIGNDATE, SENDACCT, SENDDATE,@MODIFYACCT MODIFYACCT,GETDATE() MODIFYDATE, FILEPATH1" & vbCrLf
        sSql_i &= " FROM STUD_ELFORM WHERE CSELNO=@CSELNO AND ELNO=@ELNO AND SOCID=@SOCID AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(sSql_i, objConn, pms_i)

        'DEL STUD_ELFORM
        Dim pms_d As New Hashtable From {{"CSELNO", Val(CSELNO)}, {"ELNO", Val(ELNO)}, {"SOCID", Val(SOCID)}, {"OCID", Val(OCID)}}
        Dim sSql_d As String = ""
        sSql_d &= " DELETE STUD_ELFORM FROM STUD_ELFORM" & vbCrLf
        sSql_d &= " WHERE CSELNO=@CSELNO AND ELNO=@ELNO AND SOCID=@SOCID AND OCID=@OCID" & vbCrLf
        Dim iRst As Integer = DbAccess.ExecuteNonQuery(sSql_d, objConn, pms_d)
        If iRst > 0 Then Common.MessageBox(Me, "已取消簽名，確認狀態!")
    End Sub

    '顯示簽名
    Private Sub SHOW1_IMG(sCmdArg As String)
        'Dim ELNO As String = TIMS.GetMyValue(sCmdArg, "ELNO")
        Dim CSELNO As String = TIMS.GetMyValue(sCmdArg, "CSELNO")
        ''Dim ORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        ''Dim TPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        ''Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        'Dim SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        'Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        'Dim SID As String = TIMS.GetMyValue(sCmdArg, "SID")
        Dim sUrl_1 As String = "../../Common/imgShow1.aspx?"
        'Dim xBlockN As String = TIMS.xBlockName()
        Dim xBlockN As String = String.Concat("OjtImgS1x", CSELNO)
        Dim url1 As String = String.Concat(sUrl_1, sCmdArg)
        Dim s_Specs1 As String = "width=1100,height=350,top=200,left=200,location=0,status=0,menubar=0,scrollbars=1,resizable=0,scrollbars=0"
        Call TIMS.OpenWin1(Me, url1, xBlockN, s_Specs1)
    End Sub

    '下載簽名文件
    Private Sub DOWNL1_RPT1(sCmdArg As String)
        Dim v_ELNO As String = TIMS.GetMyValue(sCmdArg, "ELNO")
        'Dim v_CSELNO As String = TIMS.GetMyValue(sCmdArg, "CSELNO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        'Dim TPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        Dim v_RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim v_SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim v_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim v_SID As String = TIMS.GetMyValue(sCmdArg, "SID")
        Dim YEARS_ROC As String = TIMS.GetMyValue(sCmdArg, "YEARS_ROC")
        '--parms = "RptID=" + RPTN1 + "&TPlanID=" + data1.TPLANID + "&RID=" + data1.RID + "&SOCID=" + data1.SOCID + "&OCID=" + data1.OCID + "&SID=" + data1.SID + "&Years=" + data1.YEARS_ROC;
        Dim RPTN1 As String = Get_PRINTFILENAME1(v_ELNO, vORGKINDGW)
        If RPTN1 = "" Then Exit Sub

        Dim s_YTEST As String = TIMS.Utl_GetConfigSet("YTEST")
        Dim myValue1 As String = ""
        TIMS.SetMyValue(myValue1, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(myValue1, "RID", v_RID)
        TIMS.SetMyValue(myValue1, "SOCID", v_SOCID)
        TIMS.SetMyValue(myValue1, "OCID", v_OCID)
        TIMS.SetMyValue(myValue1, "SID", v_SID)
        TIMS.SetMyValue(myValue1, "Years", YEARS_ROC)
        If (s_YTEST <> "") Then TIMS.SetMyValue(myValue1, "YTEST", s_YTEST)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, RPTN1, myValue1)
    End Sub


    Private Function Get_PRINTFILENAME1(vELNO As String, vORGKINDGW As String) As String
        Dim rst As String = ""
        Const cst_RTE As String = "D"
        Select Case vELNO
            Case "1"
                If cst_RTE = "T" Then
                    rst = If(vORGKINDGW.Equals("G"), "OJTSD1421G3", "OJTSD1421W3")
                ElseIf cst_RTE = "D" Then
                    rst = If(vORGKINDGW.Equals("G"), "OJTSD1421G3S", "OJTSD1421W3S")
                End If
            Case "2"
                If cst_RTE = "T" Then
                    rst = "OJTSD1405B1"
                ElseIf cst_RTE = "D" Then
                    rst = "OJTSD1405B1S"
                End If
            Case "3"
                If cst_RTE = "T" Then
                    rst = If(vORGKINDGW.Equals("G"), "OJTSD140138G", "OJTSD140138W")
                ElseIf cst_RTE = "D" Then
                    rst = If(vORGKINDGW.Equals("G"), "OJTSD140138GS", "OJTSD140138WS")
                End If
        End Select
        Return rst
    End Function

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument

        Call KEEPSEARCH()

        Dim vELNO As String = TIMS.GetMyValue(sCmdArg, "ELNO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim TPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim SID As String = TIMS.GetMyValue(sCmdArg, "SID")
        Dim YEARS_ROC As String = TIMS.GetMyValue(sCmdArg, "YEARS_ROC")

        Select Case e.CommandName
            Case "BtnUCLS2" '取消簽名
                UCLS2_SIGN(sCmdArg)
                Call Search1()
            Case "BtnSHOW1" '顯示簽名
                SHOW1_IMG(sCmdArg)
            Case "BtnDOWNL1" '下載簽名文件
                DOWNL1_RPT1(sCmdArg)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim BtnUCLS2 As LinkButton = e.Item.FindControl("BtnUCLS2") '取消簽名
                Dim BtnSHOW1 As LinkButton = e.Item.FindControl("BtnSHOW1") '顯示簽名
                Dim BtnDOWNL1 As LinkButton = e.Item.FindControl("BtnDOWNL1") '下載簽名文件
                Dim iDGSeqNo As Integer = TIMS.Get_DGSeqNo(sender, e)
                e.Item.Cells(0).Text = iDGSeqNo
                '--parms = "RptID=" + RPTN1 + "&TPlanID=" + data1.TPLANID + "&RID=" + data1.RID + "&SOCID=" + data1.SOCID + "&OCID=" + data1.OCID + "&SID=" + data1.SID + "&Years=" + data1.YEARS_ROC;

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "ELNO", Convert.ToString(drv("ELNO")))
                TIMS.SetMyValue(sCmdArg, "CSELNO", Convert.ToString(drv("CSELNO")))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "TPlanID", Convert.ToString(drv("TPlanID")))
                TIMS.SetMyValue(sCmdArg, "RID", Convert.ToString(drv("RID")))
                TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "SID", Convert.ToString(drv("SID")))
                TIMS.SetMyValue(sCmdArg, "YEARS_ROC", Convert.ToString(drv("YEARS_ROC")))

                Dim SIGNDATE_TWTIME As String = Convert.ToString(drv("SIGNDATE_TWTIME"))
                Dim fg_Enabled As Boolean = (SIGNDATE_TWTIME <> "")
                Const cst_Enabledmsg1 As String = "尚未簽署"
                BtnUCLS2.Enabled = fg_Enabled
                BtnUCLS2.Attributes("onclick") = String.Concat("return confirm('您確定要刪除第", iDGSeqNo, "筆資料嗎?');")
                BtnSHOW1.Enabled = fg_Enabled
                BtnDOWNL1.Enabled = fg_Enabled
                If Not BtnUCLS2.Enabled Then TIMS.Tooltip(BtnUCLS2, cst_Enabledmsg1, True)
                If Not BtnSHOW1.Enabled Then TIMS.Tooltip(BtnSHOW1, cst_Enabledmsg1, True)
                If Not BtnDOWNL1.Enabled Then TIMS.Tooltip(BtnDOWNL1, cst_Enabledmsg1, True)

                If BtnUCLS2.Enabled Then BtnUCLS2.CommandArgument = sCmdArg
                If BtnSHOW1.Enabled Then BtnSHOW1.CommandArgument = sCmdArg
                If BtnDOWNL1.Enabled Then BtnDOWNL1.CommandArgument = sCmdArg
        End Select
    End Sub

    Protected Sub SCH_ddlELFORMID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SCH_ddlELFORMID.SelectedIndexChanged
        DataGridTable1.Visible = False
        Msg1.Text = ""
    End Sub
End Class
