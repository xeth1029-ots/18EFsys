Partial Class SD_16_004
    Inherits AuthBasePage

    'CLASS_MAJOR
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

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

        If Not IsPostBack Then
            Call sCreate1()
        End If

    End Sub

    '第1次載入
    Sub sCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        TIMS.GET_KEYMAJORKIND(CBL_NAPPROV, "NAPPROV", objconn)
        TIMS.GET_KEYMAJORKIND(CBL_CEXCEP, "CEXCEP", objconn)
        TIMS.GET_KEYMAJORKIND(CBL_OTHNAPPROV, "OTHNAPPROV", objconn)
        TIMS.GET_KEYMAJORKIND(CBL_OTHMAJOR, "OTHMAJOR", objconn)

        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        Dim s_javascript_btn6 As String = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button6.Attributes("onclick") = s_javascript_btn6
        btn_Save1.Attributes("onclick") = "javascript:return checkSave()"

        tb_DataGrid1.Visible = False
    End Sub

    '顯示狀況
    Sub sUtl_PanelList(ByVal iType As Integer)
        'iType:1 搜尋 2:'新增/修改 3:'檢視
        Panel1.Visible = False '搜尋
        Panel2.Visible = False '新增/修改
        Select Case iType
            Case 1
                Panel1.Visible = True '搜尋
            Case 2
                Panel2.Visible = True '新增/修改
        End Select
    End Sub

    '查詢鈕
    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        Call sSearch1()
        'Sch_Mark.Value = "1"
    End Sub

    '查詢sub
    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        'iType :0:一般查詢/1:匯出(查詢結果，將身分證字號中間6碼調整為隱碼顯示)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Hid_CMID.Value = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        sch_VERIFYDATE1.Text = TIMS.ClearSQM(sch_VERIFYDATE1.Text)
        sch_VERIFYDATE2.Text = TIMS.ClearSQM(sch_VERIFYDATE2.Text)
        sch_VERIFYDATE1.Text = TIMS.Cdate3(sch_VERIFYDATE1.Text)
        sch_VERIFYDATE2.Text = TIMS.Cdate3(sch_VERIFYDATE2.Text)

        Dim pms_s As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}}
        Dim sSql As String = ""
        sSql &= " SELECT a.CMID,a.OCID,a.SEQNO" & vbCrLf
        sSql &= " ,a.CREATEACCT,a.CREATEDATE,dbo.FN_CDATE1B(a.CREATEDATE) CREATEDATE_ROC" & vbCrLf
        sSql &= " ,a.VERIFYDATE,dbo.FN_CDATE1B(a.VERIFYDATE) VERIFYDATE_ROC ,a.RESULT" & vbCrLf
        sSql &= " ,a.NAPPROV,a.CEXCEP,a.OTHNAPPROV,a.OTHMAJOR" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " ,cc.ORGNAME,cc.CLASSCNAME2,cc.RID" & vbCrLf
        sSql &= " ,cc.STDATE,cc.FTDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(cc.STDATE) STDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(cc.FTDATE) FTDATE_ROC" & vbCrLf
        sSql &= " FROM CLASS_MAJOR a" & vbCrLf
        sSql &= " JOIN VIEW2 cc on cc.OCID=a.OCID" & vbCrLf
        sSql &= " WHERE cc.TPLANID=@TPLANID AND cc.YEARS=@YEARS" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
                Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                If RIDValue.Value.Length > 1 Then
                    pms_s.Add("DISTID", s_DISTID)
                    pms_s.Add("RID", RIDValue.Value)
                    sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                    sSql &= " AND cc.RID=@RID" & vbCrLf
                ElseIf RIDValue.Value = "A" Then
                    sSql &= " AND 1=1" & vbCrLf
                Else
                    pms_s.Add("DISTID", s_DISTID)
                    sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                End If
            Case 1
                If RIDValue.Value.Length > 1 Then
                    pms_s.Add("DISTID", sm.UserInfo.DistID)
                    pms_s.Add("RID", RIDValue.Value)
                    sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                    sSql &= " AND cc.RID=@RID" & vbCrLf
                Else
                    pms_s.Add("DISTID", sm.UserInfo.DistID)
                    sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                End If
            Case Else
                pms_s.Add("DISTID", sm.UserInfo.DistID)
                pms_s.Add("RID", sm.UserInfo.RID)
                sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
                sSql &= " AND cc.RID=@RID" & vbCrLf
        End Select
        If OCIDValue1.Value <> "" Then
            pms_s.Add("OCID", OCIDValue1.Value)
            sSql &= " AND a.OCID=@OCID" & vbCrLf
        End If
        If sch_VERIFYDATE1.Text <> "" Then
            pms_s.Add("VERIFYDATE1", sch_VERIFYDATE1.Text)
            sSql &= " AND a.VERIFYDATE>=@VERIFYDATE1" & vbCrLf
        End If
        If sch_VERIFYDATE2.Text <> "" Then
            pms_s.Add("VERIFYDATE2", sch_VERIFYDATE2.Text)
            sSql &= " AND a.VERIFYDATE<=@VERIFYDATE2" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_s)

        Call sUtl_PanelList(1)
        tb_DataGrid1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count = 0 Then Return
        'If dt.Rows.Count > 0 Then End If

        tb_DataGrid1.Visible = True
        msg.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim rCMID As String = TIMS.GetMyValue(sCmdArg, "CMID")
        Dim rOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim rSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
        If rCMID = "" OrElse rOCID = "" Then Exit Sub

        Dim sql As String = ""
        Select Case e.CommandName
            Case "edit" '修改
                Call sUtl_PanelList(2) '修改
                Call Show_Detail1(rCMID, rOCID, rSEQNO) 'I:新增/V:檢視/E:編輯
            Case "delt" '刪除
                Call DEL_CLASS_MAJOR(rCMID, rOCID, rSEQNO)
                Call sSearch1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header, ListItemType.Footer
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim H_SEQNO As HiddenField = e.Item.FindControl("H_SEQNO")
                Dim H_OCID As HiddenField = e.Item.FindControl("H_OCID")
                Dim H_CMID As HiddenField = e.Item.FindControl("H_CMID")
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDelt As LinkButton = e.Item.FindControl("lbtDelt")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                H_SEQNO.Value = Convert.ToString(drv("SEQNO"))
                H_OCID.Value = Convert.ToString(drv("OCID"))
                H_CMID.Value = Convert.ToString(drv("CMID"))

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CMID", Convert.ToString(drv("CMID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "SEQNO", Convert.ToString(drv("SEQNO")))
                lbtEdit.Visible = True
                lbtEdit.CommandArgument = sCmdArg ' drv("SBSN")

                lbtDelt.Visible = False
                Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                If sm.UserInfo.LID = 0 Then
                    lbtDelt.Visible = True
                    TIMS.Tooltip(lbtDelt, "署有權限可刪除", True)
                    lbtDelt.CommandArgument = sCmdArg ' drv("SBSN")
                    lbtDelt.Attributes("onclick") = "javascript:return confirm('此動作會刪除此筆資料，是否確定刪除?');"
                ElseIf flagS1 Then
                    lbtDelt.Visible = If(flagS1, True, False)
                    lbtDelt.CommandArgument = sCmdArg ' drv("SBSN")
                    lbtDelt.Style.Item("display") = "none"
                    lbtDelt.Attributes("onclick") = "javascript:return confirm('此動作會刪除此筆資料，是否確定刪除?');"
                End If

        End Select
    End Sub

    Private Sub DEL_CLASS_MAJOR(rCMID As String, rOCID As String, rSEQNO As String)
        Dim pms_1 As New Hashtable From {{"CMID", rCMID}, {"OCID", rOCID}, {"SEQNO", rSEQNO}}
        Dim sSql As String = ""
        sSql &= " SELECT a.CMID,a.OCID,a.SEQNO" & vbCrLf
        sSql &= " FROM CLASS_MAJOR a" & vbCrLf
        sSql &= " WHERE a.CMID=@CMID AND a.OCID=@OCID AND a.SEQNO=@SEQNO" & vbCrLf
        Dim drDATA As DataRow = DbAccess.GetOneRow(sSql, objconn, pms_1)
        If drDATA Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim pms_I1 As New Hashtable From {{"CMID", rCMID}, {"OCID", rOCID}, {"SEQNO", rSEQNO}}
        Dim sM1A As String = "CMID ,OCID,SEQNO,CREATEACCT,CREATEDATE,VERIFYDATE,RESULT,NAPPROV,CEXCEP,OTHNAPPROV,OTHMAJOR,MODIFYACCT,MODIFYDATE"
        Dim sM1B As String = String.Concat("CMID ,OCID,SEQNO,CREATEACCT,CREATEDATE,VERIFYDATE,RESULT,NAPPROV,CEXCEP,OTHNAPPROV,OTHMAJOR,'", sm.UserInfo.UserID, "' MODIFYACCT,GETDATE() MODIFYDATE")
        Dim sSqlI1 As String = String.Concat(" INSERT INTO CLASS_MAJORDEL(", sM1A, ") SELECT ", sM1B, " FROM CLASS_MAJOR WHERE CMID=@CMID AND OCID=@OCID AND SEQNO=@SEQNO")
        DbAccess.ExecuteNonQuery(sSqlI1, objconn, pms_I1)

        Dim pms_I2 As New Hashtable From {{"CMID", rCMID}}
        Dim sM2A As String = "CMID ,MKDSEQ,MODIFYACCT,MODIFYDATE"
        Dim sM2B As String = String.Concat("CMID ,MKDSEQ,'", sm.UserInfo.UserID, "' MODIFYACCT,GETDATE() MODIFYDATE")
        Dim sSqlI2 As String = String.Concat(" INSERT INTO CLASS_MAJORDETAILDEL(", sM2A, ") SELECT ", sM2B, " FROM CLASS_MAJORDETAIL WHERE CMID=@CMID")
        DbAccess.ExecuteNonQuery(sSqlI2, objconn, pms_I2)

        Dim pms_D2 As New Hashtable From {{"CMID", rCMID}}
        Dim sSqlD2 As String = " DELETE CLASS_MAJORDETAIL WHERE CMID=@CMID" & vbCrLf
        DbAccess.ExecuteNonQuery(sSqlD2, objconn, pms_D2)

        Dim pms_D As New Hashtable From {{"CMID", rCMID}, {"OCID", rOCID}, {"SEQNO", rSEQNO}}
        Dim sSqlD As String = " DELETE CLASS_MAJOR WHERE CMID=@CMID AND OCID=@OCID AND SEQNO=@SEQNO" & vbCrLf
        DbAccess.ExecuteNonQuery(sSqlD, objconn, pms_D)
    End Sub

    '單1資料顯示。
    Sub Show_Detail1(s_CMID As String, s_OCID As String, s_SEQNO As String)
        s_CMID = TIMS.ClearSQM(s_CMID)
        s_OCID = TIMS.ClearSQM(s_OCID)
        s_SEQNO = TIMS.ClearSQM(s_SEQNO)
        If s_CMID = "" OrElse s_OCID = "" OrElse s_SEQNO = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim pms_1 As New Hashtable From {{"CMID", s_CMID}, {"OCID", s_OCID}, {"SEQNO", s_SEQNO}}

        Dim sSql As String = ""
        sSql &= " SELECT a.CMID ,a.OCID,a.SEQNO" & vbCrLf
        sSql &= " ,a.CREATEACCT,a.CREATEDATE,dbo.FN_CDATE1B(a.CREATEDATE) CREATEDATE_ROC" & vbCrLf
        sSql &= " ,a.VERIFYDATE,dbo.FN_CDATE1B(a.VERIFYDATE) VERIFYDATE_ROC,a.RESULT" & vbCrLf
        sSql &= " ,a.NAPPROV,a.CEXCEP,a.OTHNAPPROV,a.OTHMAJOR" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " ,cc.STDATE,cc.FTDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(cc.STDATE) STDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(cc.FTDATE) FTDATE_ROC" & vbCrLf
        sSql &= " FROM CLASS_MAJOR a" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=a.OCID" & vbCrLf
        sSql &= " WHERE a.CMID=@CMID AND a.OCID=@OCID AND a.SEQNO=@SEQNO" & vbCrLf
        Dim drDATA As DataRow = DbAccess.GetOneRow(sSql, objconn, pms_1)

        If drDATA Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If
        'If drDATA Is Nothing Then Exit Sub

        Dim drCC As DataRow = TIMS.GetOCIDDate(Convert.ToString(drDATA("OCID")), objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Call Clear_value1()

        Call Show_Detail2(drCC, 2, drDATA)
    End Sub

    ''' <summary>顯示查詢資料</summary>
    ''' <param name="drCC">一定要有班級資訊</param>
    ''' <param name="iTYPE">1 :新增使用 / 2 :查詢使用</param>
    ''' <param name="drMJ">CLASS_MAJOR 資料</param>
    Public Sub Show_Detail2(ByRef drCC As DataRow, iTYPE As Integer, ByRef drMJ As DataRow)
        If drCC Is Nothing Then Return
        'iTYPE = 1 :新增使用 / 2 :查詢使用
        VERIFYDATE.Enabled = True
        span3.Visible = True

        labORGNAME.Text = drCC("ORGNAME")
        labCLASSCNAME2.Text = drCC("CLASSCNAME2")
        'Dim s_SFTDATE_ROC As String = String.Concat(TIMS.cdate3(drCC("STDATE")), "~", TIMS.cdate3(drCC("FTDATE")))
        labSFTDATE.Text = String.Concat(TIMS.Cdate3(drCC("STDATE")), "~", TIMS.Cdate3(drCC("FTDATE")))

        If iTYPE = 1 Then
            Hid_CMID.Value = ""
            Hid_OCID.Value = Convert.ToString(drCC("OCID"))
            Hid_SEQNO.Value = "1" 'GET_NEW_SEQNO(Hid_OCID.Value) ' "1"

            labCREATEDATE.Text = TIMS.Cdate3(drCC("TODAY1"))
        Else
            If drMJ Is Nothing Then Return 'CLASS_MAJOR (為空,資料異常)
            Hid_CMID.Value = Convert.ToString(drMJ("CMID"))
            Hid_OCID.Value = Convert.ToString(drMJ("OCID"))
            Hid_SEQNO.Value = Convert.ToString(drMJ("SEQNO"))
            '登錄日期 
            labCREATEDATE.Text = TIMS.Cdate3(drMJ("CREATEDATE"))
            '經查核確認日期 
            VERIFYDATE.Text = TIMS.Cdate3(drMJ("VERIFYDATE"))
            span3.Visible = False
            VERIFYDATE.Enabled = False
            '查核結果說明
            RESULT.Text = Convert.ToString(drMJ("RESULT"))
            '重要工作事項未依核定課程施訓: 
            NAPPROV.Text = Convert.ToString(drMJ("NAPPROV"))
            '課程異常狀況: 
            CEXCEP.Text = Convert.ToString(drMJ("CEXCEP"))
            '其他未依核定課程施訓: 
            OTHNAPPROV.Text = Convert.ToString(drMJ("OTHNAPPROV"))
            '其他重大異常狀況:-->
            OTHMAJOR.Text = Convert.ToString(drMJ("OTHMAJOR"))
            'labCREATEDATE_ROC,VERIFYDATE2,span3,RESULT,CBL_NAPPROV,NAPPROV,CBL_CEXCEP,CEXCEP,CBL_OTHNAPPROV,OTHNAPPROV,CBL_OTHMAJOR,OTHMAJOR,

            Dim iCMID As Integer = Convert.ToString(drMJ("CMID"))

            For Each lstMKDSEQ As ListItem In CBL_NAPPROV.Items
                lstMKDSEQ.Selected = CHK_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value))
            Next
            For Each lstMKDSEQ As ListItem In CBL_CEXCEP.Items
                lstMKDSEQ.Selected = CHK_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value))
            Next
            For Each lstMKDSEQ As ListItem In CBL_OTHNAPPROV.Items
                lstMKDSEQ.Selected = CHK_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value))
            Next
            For Each lstMKDSEQ As ListItem In CBL_OTHMAJOR.Items
                lstMKDSEQ.Selected = CHK_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value))
            Next
        End If
    End Sub

    Private Function GET_NEW_SEQNO(rOCID As String) As Integer
        Dim pms_1 As New Hashtable From {{"OCID", rOCID}}
        Dim sSql As String = " SELECT COUNT(1)+1 SEQNO FROM CLASS_MAJOR WHERE OCID=@OCID"
        Dim drDATA As DataRow = DbAccess.GetOneRow(sSql, objconn, pms_1)
        Dim rst As String = If(drDATA Is Nothing, "1", Convert.ToString(drDATA("SEQNO")))
        If rst = "" Then Return "1"
        Return Val(rst)
    End Function

    '新增鈕
    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim s_ERROR1 As String = ""
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            s_ERROR1 = String.Concat(TIMS.cst_NODATAMsg1, "(新增請先選擇班級)!")
            Common.MessageBox(Me, s_ERROR1)
            Return
        End If

        Call sUtl_PanelList(2) '新增

        Call Clear_value1()

        Call Show_Detail2(drCC, 1, Nothing)
    End Sub

    ''' <summary>儲存, 回傳要儲存 Hid_CMID.Value</summary>
    Sub SAVE_CLASS_MAJOR()
        Hid_CMID.Value = TIMS.ClearSQM(Hid_CMID.Value)
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Hid_SEQNO.Value = TIMS.ClearSQM(Hid_SEQNO.Value)

        VERIFYDATE.Text = TIMS.ClearSQM(VERIFYDATE.Text)
        RESULT.Text = TIMS.ClearSQM(RESULT.Text)
        NAPPROV.Text = TIMS.ClearSQM(NAPPROV.Text)
        CEXCEP.Text = TIMS.ClearSQM(CEXCEP.Text)
        OTHNAPPROV.Text = TIMS.ClearSQM(OTHNAPPROV.Text)
        OTHMAJOR.Text = TIMS.ClearSQM(OTHMAJOR.Text)
        VERIFYDATE.Text = TIMS.Cdate3(VERIFYDATE.Text)
        Dim fg_NEW As Boolean = (Hid_CMID.Value = "")
        Dim fg_UPDATE As Boolean = (Hid_CMID.Value <> "")
        Dim iSEQNO As Integer = If(Hid_SEQNO.Value <> "", Val(Hid_SEQNO.Value), 1)

        Dim sql As String = ""
        If fg_NEW Then
            iSEQNO = GET_NEW_SEQNO(Hid_OCID.Value) ' "1"
            Dim iCMID As Integer = DbAccess.GetNewId(objconn, "CLASS_MAJOR_CMID_SEQ,CLASS_MAJOR,CMID")
            Hid_CMID.Value = iCMID.ToString()
            'iParms.Add("CREATEDATE", CREATEDATE)
            Dim iParms As New Hashtable From {
                {"CMID", iCMID},
                {"OCID", Val(Hid_OCID.Value)},
                {"SEQNO", iSEQNO},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"VERIFYDATE", If(VERIFYDATE.Text <> "", VERIFYDATE.Text, Convert.DBNull)},
                {"RESULT", If(RESULT.Text <> "", RESULT.Text, Convert.DBNull)},
                {"NAPPROV", If(NAPPROV.Text <> "", NAPPROV.Text, Convert.DBNull)},
                {"CEXCEP", If(CEXCEP.Text <> "", CEXCEP.Text, Convert.DBNull)},
                {"OTHNAPPROV", If(OTHNAPPROV.Text <> "", OTHNAPPROV.Text, Convert.DBNull)},
                {"OTHMAJOR", If(OTHMAJOR.Text <> "", OTHMAJOR.Text, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            Dim isSql As String = ""
            isSql &= " INSERT INTO CLASS_MAJOR(CMID, OCID, SEQNO, CREATEACCT, CREATEDATE, VERIFYDATE, RESULT, NAPPROV, CEXCEP, OTHNAPPROV, OTHMAJOR, MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSql &= " VALUES(@CMID,@OCID,@SEQNO,@CREATEACCT,GETDATE(),@VERIFYDATE,@RESULT,@NAPPROV,@CEXCEP,@OTHNAPPROV,@OTHMAJOR,@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)

        ElseIf fg_UPDATE Then
            Dim iCMID As Integer = Val(Hid_CMID.Value)
            Dim uParms As New Hashtable From {
                {"CMID", iCMID},
                {"OCID", Val(Hid_OCID.Value)},
                {"SEQNO", iSEQNO},
                {"RESULT", If(RESULT.Text <> "", RESULT.Text, Convert.DBNull)},
                {"NAPPROV", If(NAPPROV.Text <> "", NAPPROV.Text, Convert.DBNull)},
                {"CEXCEP", If(CEXCEP.Text <> "", CEXCEP.Text, Convert.DBNull)},
                {"OTHNAPPROV", If(OTHNAPPROV.Text <> "", OTHNAPPROV.Text, Convert.DBNull)},
                {"OTHMAJOR", If(OTHMAJOR.Text <> "", OTHMAJOR.Text, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim usSql As String = ""
            usSql &= " UPDATE CLASS_MAJOR" & vbCrLf
            usSql &= " SET RESULT=@RESULT,NAPPROV=@NAPPROV" & vbCrLf
            usSql &= " ,CEXCEP=@CEXCEP,OTHNAPPROV=@OTHNAPPROV,OTHMAJOR=@OTHMAJOR" & vbCrLf
            usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE CMID=@CMID AND OCID=@OCID AND SEQNO=@SEQNO" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        End If

    End Sub

    '儲存鈕
    Private Sub btn_Save1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save1.Click
        Hid_CMID.Value = TIMS.ClearSQM(Hid_CMID.Value)
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Hid_SEQNO.Value = TIMS.ClearSQM(Hid_SEQNO.Value)
        VERIFYDATE.Text = TIMS.ClearSQM(VERIFYDATE.Text)
        RESULT.Text = TIMS.ClearSQM(RESULT.Text)
        NAPPROV.Text = TIMS.ClearSQM(NAPPROV.Text)
        CEXCEP.Text = TIMS.ClearSQM(CEXCEP.Text)
        OTHNAPPROV.Text = TIMS.ClearSQM(OTHNAPPROV.Text)
        OTHMAJOR.Text = TIMS.ClearSQM(OTHMAJOR.Text)
        VERIFYDATE.Text = TIMS.Cdate3(VERIFYDATE.Text)
        Dim fg_NEW As Boolean = (Hid_CMID.Value = "")
        Dim fg_UPDATE As Boolean = (Hid_CMID.Value <> "")

        Dim sErrmsg As String = ""
        If VERIFYDATE.Text = "" Then sErrmsg &= "請輸入 經查核確認日期"
        If fg_NEW AndAlso Hid_OCID.Value <> "" AndAlso VERIFYDATE.Text <> "" Then
            Dim pms_s1 As New Hashtable From {{"OCID", Hid_OCID.Value}, {"VERIFYDATE", VERIFYDATE.Text}}
            Dim fg_EXISTS1 As Boolean = CHK_CLASS_MAJOR(pms_s1)
            If (fg_EXISTS1) Then sErrmsg &= String.Format("班級、經查核確認日期已有資料，請使用修改!({0}),({1})", Hid_OCID.Value, VERIFYDATE.Text)
        End If
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Call SAVE_CLASS_MAJOR()

        Dim iCMID As Integer = Val(Hid_CMID.Value)
        'Dim vCBL_NAPPROV As String = TIMS.GetCblValue(CBL_NAPPROV)
        'Dim vCBL_CEXCEP As String = TIMS.GetCblValue(CBL_CEXCEP)
        'Dim vCBL_OTHNAPPROV As String = TIMS.GetCblValue(CBL_OTHNAPPROV)
        'Dim vCBL_OTHMAJOR As String = TIMS.GetCblValue(CBL_OTHMAJOR)
        For Each lstMKDSEQ As ListItem In CBL_NAPPROV.Items
            SAVE_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value), lstMKDSEQ.Selected)
        Next
        For Each lstMKDSEQ As ListItem In CBL_CEXCEP.Items
            SAVE_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value), lstMKDSEQ.Selected)
        Next
        For Each lstMKDSEQ As ListItem In CBL_OTHNAPPROV.Items
            SAVE_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value), lstMKDSEQ.Selected)
        Next
        For Each lstMKDSEQ As ListItem In CBL_OTHMAJOR.Items
            SAVE_CLASS_MAJORDETAIL(iCMID, Val(lstMKDSEQ.Value), lstMKDSEQ.Selected)
        Next

        Common.MessageBox(Me, "儲存成功")
        Call sSearch1()
    End Sub

    Private Function CHK_CLASS_MAJOR(pms_r As Hashtable) As Boolean
        Dim rst As Boolean = False
        Dim v_OCID As String = TIMS.GetMyValue2(pms_r, "OCID")
        Dim v_VERIFYDATE As String = TIMS.Cdate3(TIMS.GetMyValue2(pms_r, "VERIFYDATE"))
        If v_OCID = "" OrElse v_OCID = "" Then Return rst

        Dim pms_s1 As New Hashtable From {{"OCID", Hid_OCID.Value}, {"VERIFYDATE", VERIFYDATE.Text}}
        Dim ssSql As String = ""
        ssSql &= " SELECT 1 FROM CLASS_MAJOR WHERE OCID=@OCID AND VERIFYDATE=@VERIFYDATE" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(ssSql, objconn, pms_s1)
        Return (dr1 IsNot Nothing)
    End Function

    Function CHK_CLASS_MAJORDETAIL(iCMID As Integer, iMKDSEQ As Integer) As Boolean
        Dim pms_s As New Hashtable From {{"CMID", iCMID}, {"MKDSEQ", iMKDSEQ}}
        Dim ssSql As String = ""
        ssSql &= " SELECT 1 FROM CLASS_MAJORDETAIL WHERE CMID=@CMID AND MKDSEQ=@MKDSEQ" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(ssSql, objconn, pms_s)
        Return (dr1 IsNot Nothing)
    End Function

    Sub SAVE_CLASS_MAJORDETAIL(iCMID As Integer, iMKDSEQ As Integer, fg_CHK As Boolean)
        Dim pms_s As New Hashtable From {{"CMID", iCMID}, {"MKDSEQ", iMKDSEQ}}
        Dim ssSql As String = ""
        ssSql &= " SELECT 1 FROM CLASS_MAJORDETAIL WHERE CMID=@CMID AND MKDSEQ=@MKDSEQ" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(ssSql, objconn, pms_s)

        Dim pms_i As New Hashtable From {{"CMID", iCMID}, {"MKDSEQ", iMKDSEQ}, {"MODIFYACCT", sm.UserInfo.UserID}}
        Dim isSql As String = ""
        isSql &= " INSERT INTO CLASS_MAJORDETAIL(CMID, MKDSEQ, MODIFYACCT, MODIFYDATE)" & vbCrLf
        isSql &= " VALUES(@CMID,@MKDSEQ,@MODIFYACCT,GETDATE())" & vbCrLf
        Dim pms_d As New Hashtable From {{"CMID", iCMID}, {"MKDSEQ", iMKDSEQ}}
        Dim dsSql As String = ""
        dsSql &= " DELETE CLASS_MAJORDETAIL WHERE CMID=@CMID AND MKDSEQ=@MKDSEQ" & vbCrLf
        If fg_CHK Then
            If dr1 Is Nothing Then DbAccess.ExecuteNonQuery(isSql, objconn, pms_i)
        Else
            If dr1 IsNot Nothing Then DbAccess.ExecuteNonQuery(dsSql, objconn, pms_d)
        End If
    End Sub

    '清除編輯值。
    Sub Clear_value1()
        '訓練期間 
        labSFTDATE.Text = "" 'String.Concat(dr("STDATE_ROC"), "~", dr("FTDATE_ROC"))
        '登錄日期 
        labCREATEDATE.Text = "" 'Convert.ToString(dr("CREATEDATE_ROC"))
        '經查核確認日期 
        VERIFYDATE.Text = "" 'Convert.ToString(dr("VERIFYDATE"))
        '查核結果說明
        RESULT.Text = "" 'Convert.ToString(dr("RESULT"))
        '--重要工作事項未依核定課程施訓: 
        NAPPROV.Text = "" 'Convert.ToString(dr("NAPPROV"))
        '課程異常狀況: 
        CEXCEP.Text = "" 'Convert.ToString(dr("CEXCEP"))
        '其他未依核定課程施訓: 
        OTHNAPPROV.Text = "" 'Convert.ToString(dr("OTHNAPPROV"))
        '其他重大異常狀況:-->
        OTHMAJOR.Text = "" 'Convert.ToString(dr("OTHMAJOR"))

        TIMS.SetCblValue(CBL_NAPPROV, "")
        TIMS.SetCblValue(CBL_CEXCEP, "")
        TIMS.SetCblValue(CBL_OTHNAPPROV, "")
        TIMS.SetCblValue(CBL_OTHMAJOR, "")

        Hid_CMID.Value = "" ' Convert.ToString(dr("CMID"))
        Hid_OCID.Value = "" 'Convert.ToString(dr("OCID"))
        Hid_SEQNO.Value = "" ' Convert.ToString(dr("SEQNO"))
    End Sub

    '儲存離開鈕
    Private Sub btn_Leave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Leave1.Click
        Call sUtl_PanelList(1) '搜尋
        Call Clear_value1()
    End Sub

    Private Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        'BtnGETvalue2
        '判斷機構是否只有一個班級 '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = Convert.ToString(dr("TRAINNAME"))
        OCID1.Text = Convert.ToString(dr("CLASSNAME"))
        TMIDValue1.Value = Convert.ToString(dr("TRAINID"))
        OCIDValue1.Value = Convert.ToString(dr("OCID"))
    End Sub

End Class
