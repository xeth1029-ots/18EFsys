Partial Class SD_14_019
    Inherits AuthBasePage

    'SyncLock
    'Private Shared ReadOnly CreatTable3_lock As New Object

    Const Cst_006_ComIDNO As String = "76250772"        '南區職業訓練中心  '付款人統一編號**
    Const Cst_006_BankCode2 As String = "0040440"       '南區職業訓練中心  '付款人銀行總行代號**'付款人銀行分行代號**
    'Const Cst_006_BankCode3 As String = "044036070048" '南區職業訓練中心  '付款人帳號**
    Const Cst_006_BankCode3 As String = "044036070023"  '南區職業訓練中心  '付款人帳號**

    Dim dtOrg As DataTable = Nothing
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        PageControler1.PageDataGrid = DataGrid1
        'dtOrg = TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'", objconn)

        If Not IsPostBack Then
            cCreate1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'Years.Value = sm.UserInfo.Years - 1911
        'PlanID.Value = sm.UserInfo.PlanID
    End Sub

    Sub cCreate1()
        msg.Text = ""
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "0")
        txtNowDate1.Text = Common.FormatDate(Now)

        Dim s_javascript_btn2 As String = ""
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button2.Attributes("onclick") = s_javascript_btn2

        'print_type.Attributes("onclick") = "printkind();" '列印時排序方式
        Button1.Attributes("onclick") = "return CheckSearch();"
        'Button5.Attributes("onclick") = "CheckPrint('" & ReportQuery.GetSmartQueryPath & "');"
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        dtOrg = TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'", objconn)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim dt As DataTable

        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim pms_s1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"RelShip", RelShip}}
        Dim sql As String = ""
        sql &= " SELECT cc.OCID,cc.COMIDNO,cc.PLANID,ip.DISTID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,cc.STDate ,cc.FTDate,cc.AppliedResultM" & vbCrLf
        sql &= " ,vr.ORGNAME" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PlanID = cc.PlanID" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME vr ON vr.RID = cc.RID" & vbCrLf
        sql &= " WHERE cc.NOTOPEN='N'" & vbCrLf
        sql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND vr.RelShip LIKE concat(@RelShip,'%')" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.YEARS=@YEARS" & vbCrLf
            pms_s1.Add("YEARS", sm.UserInfo.Years)
        Else
            sql &= " AND cc.PLANID=@PLANID" & vbCrLf
            pms_s1.Add("PLANID", sm.UserInfo.PlanID)
        End If

        '28:產業人才投資方案
        'KindValue.Value = TIMS.GetTPlanName(sm.UserInfo.TPlanID)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sql &= " AND vr.OrgKind <> '10'" & vbCrLf
                Case "2"
                    '提升勞工自主學習計畫
                    sql &= " AND vr.OrgKind = '10'" & vbCrLf
            End Select
        End If

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value <> "" Then sql &= " AND cc.OCID = '" & OCIDValue1.Value & "'" & vbCrLf
        If STDate1.Text <> "" Then sql &= " AND cc.STDate >= " & TIMS.To_date(STDate1.Text) '& "','yyyy/mm/dd')" & vbCrLf
        If STDate2.Text <> "" Then sql &= " AND cc.STDate <= " & TIMS.To_date(STDate2.Text) 'convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
        If FTDate1.Text <> "" Then sql &= " AND cc.FTDate >= " & TIMS.To_date(FTDate1.Text) 'convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
        If FTDate2.Text <> "" Then sql &= " AND cc.FTDate <= " & TIMS.To_date(FTDate2.Text) 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
        sql &= " ORDER BY cc.STDate" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn, pms_s1)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count = 0 Then Exit Sub
        msg.Text = ""
        DataGridTable.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '匯出按鈕
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub '(狀況異常)
        Dim sCmdArg As String = e.CommandArgument
        Dim mvOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        If mvOCID = "" Then Exit Sub '(狀況異常)

        dtOrg = TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'", objconn)
        Select Case e.CommandName
            Case "Export"
                If Not IsDate(txtNowDate1.Text) Then
                    Common.MessageBox(Me, "付款日期有誤!!!")
                    Exit Sub
                End If
                Select Case Me.ExportType.SelectedValue
                    Case "1"
                        '依【中國信託文字檔】格式
                        CreatTable1(mvOCID)
                    Case "2"
                        '依【臺灣銀行XML檔】格式(台灣銀行FXML)
                        CreatTable2(mvOCID)
                    Case "3"
                        Dim strErrmsg As String = ""
                        '依【ach】格式
                        CreatTable3(mvOCID, strErrmsg)
                        'SyncLock CreatTable3_lock End SyncLock
                        If strErrmsg <> "" Then
                            Common.MessageBox(Me, strErrmsg)
                        Else
                            'If strErrmsg = "" Then TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
                            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
                        End If
                End Select
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"

            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputHidden = e.Item.FindControl("OCID")
                Dim btnExport As Button = e.Item.FindControl("btnExport") 'Export(匯出)

                Dim s_AppliedResultM As String = Convert.ToString(drv("AppliedResultM"))
                OCID.Value = Convert.ToString(drv("OCID"))
                If OCID.Value = "" Then Exit Sub '(狀況異常)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                btnExport.CommandArgument = sCmdArg

                Dim i_CHK_Cnt As Integer = ChkSubCount(drv("OCID")) ' checkCount(drv("OCID"))

                Dim fg_btnExpEnabled As Boolean = True '(匯出功能卡控)
                Dim s_tit1 As String = String.Concat(i_CHK_Cnt, " 筆資料。")
                If i_CHK_Cnt = 0 Then
                    fg_btnExpEnabled = False
                    s_tit1 = "查無資料!!"
                ElseIf (s_AppliedResultM <> "Y") Then
                    fg_btnExpEnabled = False
                    s_tit1 = "尚未完成補助審核，無法匯出銀行匯款資料"
                    'ElseIf i_CHK_Cnt > 0 AndAlso s_AppliedResultM = "Y" Then
                End If
                btnExport.Enabled = fg_btnExpEnabled
                TIMS.Tooltip(btnExport, s_tit1)

        End Select
    End Sub

    '檢查是否有資料
    Function ChkSubCount(ByVal OCIDval As String) As Integer
        Dim pms_s As New Hashtable From {{"OCID", OCIDval}}
        Dim sql As String = ""
        sql &= " SELECT COUNT(1) CNT" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS c WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_SUBSIDYCOST t WITH(NOLOCK) ON t.SOCID = c.SOCID" & vbCrLf
        sql &= " WHERE c.OCID=@OCID " & vbCrLf
        Return DbAccess.ExecuteScalar(sql, objconn, pms_s)
    End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    '依【中國信託文字檔】格式
    Sub CreatTable1(ByVal OCIDVal As String)
        'SelectValue
        'Dim strfield As String = "0,11,17,"
        'Dim dt1 As DataTable
        'Dim i As Integer = 0
        'AcctMode	0:郵政1:金融
        'PostNo	    郵政-局號 
        'AcctHeadNo	金融-總代號 
        'AcctExNo	金融-分支代號 
        'AcctNo     帳號
        Dim sParms As New Hashtable From {{"OCID", OCIDVal}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID ,d.IDNO ,c.MIdentityID ,c.StudentID" & vbCrLf
        'BankCode2
        sql &= " ,CASE WHEN h.AcctMode=1 THEN REPLACE(ISNULL(CONVERT(VARCHAR, h.AcctHeadNo),'000') + RTRIM(LTRIM(ISNULL(CONVERT(VARCHAR, h.AcctExNo),'0000'))),'-','')" & vbCrLf
        sql &= "  WHEN h.AcctMode=0 THEN '7000021'" & vbCrLf '0:郵政 '7000021(郵政總行分行代號)
        sql &= "  ELSE ' ' END BankCode2" & vbCrLf
        'BankCode3
        sql &= "  ,REPLACE(CONVERT(VARCHAR, ISNULL(h.PostNo,'')) + CONVERT(VARCHAR, ISNULL(h.AcctNo,'')),'-','') BankCode3" & vbCrLf
        sql &= "  ,d.Name" & vbCrLf
        sql &= "  ,t.SumOfMoney" & vbCrLf '產投補助。
        sql &= " FROM Class_Classinfo a" & vbCrLf
        sql &= " JOIN Plan_PlanInfo b ON b.ComIDNO = a.ComIDNO AND b.PlanID = a.PlanID AND b.SeqNO = a.SeqNO" & vbCrLf
        sql &= " JOIN ID_Class e ON e.CLSID = a.CLSID" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME v1 ON v1.RID = b.RID" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass c ON c.OCID = a.OCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo d ON d.SID = c.SID" & vbCrLf
        sql &= " JOIN Stud_ServicePlace h ON h.SOCID = c.SOCID" & vbCrLf
        sql &= " JOIN Key_Identity g ON g.IdentityID = c.MIdentityID" & vbCrLf
        sql &= " JOIN Stud_SubSidyCost t ON t.SOCID = c.SOCID" & vbCrLf
        sql &= " WHERE a.AppliedResultR='Y'" & vbCrLf
        sql &= " AND a.OCID=@OCID" & vbCrLf
        sql &= " ORDER BY c.StudentID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, sParms)
        If dt1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        Me.ViewState("ComIDNO") = If(dtOrg.Rows.Count > 0, dtOrg.Rows(0)("ComIDNO"), "")
        'Me.ViewState("ComIDNO") = TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'", objconn).Rows(0)("ComIDNO")
        Me.ViewState("OrgName") = sm.UserInfo.OrgName 'TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'").Rows(0)("OrgName")
        'Me.ViewState("NowDate1") = Common.FormatDate2Roc(Now, "")
        Me.ViewState("NowDate1") = Common.FormatDate2Roc(txtNowDate1.Text, "")

        Dim MyFileName As String = String.Concat("Result", ".txt")
        Dim sMemo As String = ""
        sMemo &= "&ACT=" & MyFileName & vbCrLf
        'sMemo &= "&TPLANID=" & sm.UserInfo.TPlanID & vbCrLf
        sMemo &= "&OCIDValue1=" & OCIDVal & vbCrLf
        sMemo &= "&COUNT=" & dt1.Rows.Count & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDVal, sMemo, objconn)

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("Result", System.Text.Encoding.UTF8) & ".txt")
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.ASCII))
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        '建立輸出文字
        Dim ExportStr As String = ""
        'ExportStr = "准考證號碼" & vbTab & "姓名" & vbTab & "身分證號碼" & vbTab & "報名日期" & vbTab & "筆試成績" & vbTab & "口試成績" & vbTab & "總成績" & vbTab
        'ExportStr &= vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim i As Integer = 1
        For Each dr As DataRow In dt1.Rows
            '整理特殊狀況
            dr("BankCode2") = Convert.ToString(dr("BankCode2")).Replace(" ", "")
            dr("BankCode3") = Convert.ToString(dr("BankCode3")).Replace(" ", "")
            dr("BankCode2") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode2"))))
            dr("BankCode3") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode3"))))
            ExportStr = ""
            ExportStr &= Convert.ToString(dr("IDNO"))
            ExportStr &= String.Concat(" ", Format(i, "0#####"))
            '01:自行 / '11:跨行
            ExportStr &= If(Left(dr("BankCode2").ToString, 3) = "822", "01", "11")
            ExportStr &= dr("BankCode2").ToString
            ExportStr &= Format(dr("SumOfMoney"), "0##########")
            ExportStr &= "002"
            ExportStr &= dr("BankCode3").ToString()
            ExportStr &= TIMS.Str_Repeat(" ", 54 - TIMS.LENB(ExportStr))
            ExportStr &= ReplaceSpecialName(dr("Name").ToString())
            ExportStr &= TIMS.Str_Repeat(" ", 134 - TIMS.LENB(ExportStr))
            ExportStr &= Me.ViewState("ComIDNO").ToString()
            ExportStr &= TIMS.Str_Repeat(" ", 155 - TIMS.LENB(ExportStr))
            ExportStr &= Me.ViewState("OrgName").ToString()
            ExportStr &= TIMS.Str_Repeat(" ", 315 - TIMS.LENB(ExportStr))
            ExportStr &= Me.ViewState("NowDate1").ToString()
            ExportStr &= vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            i += 1
        Next
        Call TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

    '依【臺灣銀行XML檔】格式
    Sub CreatTable2(ByVal OCIDVal As String)
        'SelectValue
        'Dim strfield As String = "0,11,17,"
        'Dim sql As String = ""
        'Dim dt1 As DataTable
        'Dim i As Integer = 0
        'AcctMode	0:郵政1:金融
        'PostNo	    郵政-局號 
        'AcctHeadNo	金融-總代號 
        'AcctExNo	金融-分支代號 
        'AcctNo     帳號

        Dim sParms As New Hashtable From {{"OCID", OCIDVal}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID ,d.IDNO ,c.MIdentityID ,c.StudentID" & vbCrLf
        'BankCode2
        sql &= " ,CASE WHEN h.AcctMode=1 THEN REPLACE(ISNULL(CONVERT(VARCHAR, h.AcctHeadNo),'000') + RTRIM(LTRIM(ISNULL(CONVERT(VARCHAR, h.AcctExNo),'0000'))),'-','')" & vbCrLf
        sql &= "  WHEN h.AcctMode=0 THEN '7000021'" & vbCrLf '0:郵政 '7000021(郵政總行分行代號)
        sql &= "  ELSE ' ' END BankCode2" & vbCrLf
        'BankCode3
        sql &= " ,REPLACE(CONVERT(VARCHAR, ISNULL(h.PostNo,'')) + CONVERT(VARCHAR, ISNULL(h.AcctNo,'')),'-','') BankCode3" & vbCrLf
        '手續費(非台銀10/台銀0)
        sql &= " ,CASE WHEN CONVERT(VARCHAR, h.AcctHeadNo) LIKE '004%' THEN 0 ELSE 10 END Money2" & vbCrLf
        sql &= " ,d.Name" & vbCrLf
        sql &= " ,t.SumOfMoney" & vbCrLf
        sql &= " ,ip.DistID" & vbCrLf
        sql &= " FROM Class_Classinfo a" & vbCrLf
        sql &= " JOIN Plan_PlanInfo b ON b.ComIDNO = a.ComIDNO AND b.PlanID = a.PlanID AND b.SeqNO = a.SeqNO" & vbCrLf
        sql &= " JOIN ID_Class e ON e.CLSID = a.CLSID" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID = b.PlanID" & vbCrLf
        sql &= " JOIN View_RIDName v1 ON v1.RID = b.RID" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass c ON c.OCID = a.OCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo d ON d.SID = c.SID" & vbCrLf
        sql &= " JOIN Stud_ServicePlace h ON h.SOCID = c.SOCID" & vbCrLf
        sql &= " JOIN Key_Identity g ON g.IdentityID = c.MIdentityID" & vbCrLf
        sql &= " JOIN Stud_SubSidyCost t ON t.SOCID = c.SOCID" & vbCrLf
        sql &= " WHERE a.AppliedResultR='Y'" & vbCrLf
        sql &= " AND a.OCID=@OCID" & vbCrLf
        sql &= " ORDER BY c.StudentID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, sParms)
        If dt1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        Select Case sm.UserInfo.DistID
            Case "006" '南區職業訓練中心
                Me.ViewState("ComIDNO") = Cst_006_ComIDNO '南區職業訓練中心'付款人統一編號**
            Case Else
                Me.ViewState("ComIDNO") = ""
                If dtOrg.Rows.Count > 0 Then
                    Me.ViewState("ComIDNO") = dtOrg.Rows(0)("ComIDNO")
                End If
                'Me.ViewState("ComIDNO") = TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'").Rows(0)("ComIDNO")
        End Select

        Me.ViewState("OrgName") = sm.UserInfo.OrgName 'TIMS.Get_KeyTable("Org_OrgInfo", "OrgID='" & sm.UserInfo.OrgID & "'").Rows(0)("OrgName")
        'Me.ViewState("NowDate1") = Common.FormatDate2Roc(Now, "")
        'Me.ViewState("NowDate2") = Common.FormatDate(Now, "")
        Me.ViewState("NowDate1") = Common.FormatDate2Roc(txtNowDate1.Text, "")
        Me.ViewState("NowDate2") = Common.FormatDate(txtNowDate1.Text, "")

        Dim MyFileName As String = String.Concat("FM", Me.ViewState("NowDate1"), ".TXT")
        Dim sMemo As String = ""
        sMemo &= "&ACT=" & MyFileName & vbCrLf
        'sMemo &= "&TPLANID=" & sm.UserInfo.TPlanID & vbCrLf
        sMemo &= "&OCIDValue1=" & OCIDVal & vbCrLf
        sMemo &= "&COUNT=" & dt1.Rows.Count & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDVal, sMemo, objconn)

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("Result", System.Text.Encoding.UTF8) & ".txt")
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.ASCII))
        Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("big5")

        Dim ExportStr As String = "" '建立輸出文字
        'ExportStr = "准考證號碼" & vbTab & "姓名" & vbTab & "身分證號碼" & vbTab & "報名日期" & vbTab & "筆試成績" & vbTab & "口試成績" & vbTab & "總成績" & vbTab
        'ExportStr &= vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim i As Integer = 1
        For Each dr As DataRow In dt1.Rows
            '整理特殊狀況
            dr("BankCode2") = Convert.ToString(dr("BankCode2")).Replace(" ", "") '空白清除
            dr("BankCode3") = Convert.ToString(dr("BankCode3")).Replace(" ", "") '空白清除
            dr("BankCode2") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode2"))))
            dr("BankCode3") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode3"))))

            ExportStr = ""
            ExportStr &= dr("StudentID").ToString '交易序號**
            ExportStr &= TIMS.Str_Repeat(" ", 12 - TIMS.LENB(ExportStr))
            ExportStr &= Me.ViewState("NowDate2") '付款日**

            Select Case dr("DistID").ToString
                Case "006" '南區職業訓練中心
                    ExportStr &= Cst_006_BankCode2 '付款人銀行總行代號**'付款人銀行分行代號**
                    ExportStr &= TIMS.Str_Repeat(" ", 27 - TIMS.LENB(ExportStr))
                    ExportStr &= Cst_006_BankCode3 '付款人帳號**
                    ExportStr &= TIMS.Str_Repeat(" ", 43 - TIMS.LENB(ExportStr))
                Case Else
                    ExportStr &= dr("BankCode2").ToString '付款人銀行總行代號**'付款人銀行分行代號**
                    ExportStr &= TIMS.Str_Repeat(" ", 27 - TIMS.LENB(ExportStr))
                    ExportStr &= dr("BankCode3").ToString '付款人帳號**
                    ExportStr &= TIMS.Str_Repeat(" ", 43 - TIMS.LENB(ExportStr))
            End Select

            ExportStr &= Me.ViewState("OrgName").ToString '付款人戶名**
            ExportStr &= TIMS.Str_Repeat(" ", 123 - TIMS.LENB(ExportStr))
            ExportStr &= Me.ViewState("ComIDNO").ToString '付款人統一編號**
            ExportStr &= TIMS.Str_Repeat(" ", 133 - TIMS.LENB(ExportStr))
            ExportStr &= TIMS.Str_Repeat(" ", 169 - TIMS.LENB(ExportStr))
            ExportStr &= dr("BankCode2").ToString '收款人銀行總行代號**'收款人銀行分行代號**
            ExportStr &= TIMS.Str_Repeat(" ", 176 - TIMS.LENB(ExportStr))
            ExportStr &= dr("BankCode3").ToString '收款人帳號**
            ExportStr &= TIMS.Str_Repeat(" ", 192 - TIMS.LENB(ExportStr))
            ExportStr &= ConvertBig5(ReplaceSpecialName(dr("Name").ToString)) '收款人戶名**
            ExportStr &= TIMS.Str_Repeat(" ", 272 - TIMS.LENB(ExportStr))
            ExportStr &= dr("IDNO").ToString '收款人統一編號**
            ExportStr &= TIMS.Str_Repeat(" ", 282 - TIMS.LENB(ExportStr))
            ExportStr &= Format(dr("SumOfMoney"), "0###########") '付款金額
            ExportStr &= "0" '手續費分攤方式 0:付款人付擔 1:收款人付擔
            ExportStr &= TIMS.Str_Repeat(" ", 527 - TIMS.LENB(ExportStr))
            ExportStr &= Format(dr("Money2"), "0###") '手續費4
            ExportStr &= vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            i += 1
        Next
        TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

    ''' <summary>依【ach】格式 '匯出鈕</summary>
    ''' <param name="OCIDVal"></param>
    ''' <param name="strErrmsg"></param>
    Sub CreatTable3(ByVal OCIDVal As String, ByRef strErrmsg As String)
        Dim sParms As New Hashtable From {{"OCID", OCIDVal}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID ,d.IDNO ,c.MIdentityID ,c.StudentID" & vbCrLf
        sql &= " ,CASE WHEN h.AcctMode=1 THEN REPLACE(CONVERT(VARCHAR, h.AcctHeadNo) + CONVERT(VARCHAR, h.AcctExNo),'-','')" & vbCrLf
        sql &= "  WHEN h.AcctMode=0 THEN '7000021' ELSE ' ' END BankCode2" & vbCrLf
        sql &= " ,REPLACE(CONVERT(VARCHAR, ISNULL(h.PostNo,'')) + CONVERT(VARCHAR, ISNULL(h.AcctNo,'')) ,'-','') BankCode3" & vbCrLf
        sql &= " ,d.Name ,t.SumOfMoney" & vbCrLf
        sql &= " ,CASE WHEN c.BudgetID = '03' THEN '1' ELSE '' END AS Budget03" & vbCrLf
        sql &= " ,CASE WHEN c.BudgetID = '02' THEN '1' ELSE '' END AS Budget02" & vbCrLf
        sql &= " FROM Class_Classinfo a" & vbCrLf
        sql &= " JOIN Plan_PlanInfo b ON b.ComIDNO = a.ComIDNO AND b.PlanID = a.PlanID AND b.SeqNO = a.SeqNO" & vbCrLf
        sql &= " JOIN ID_Class e ON e.CLSID = a.CLSID" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME v1 ON v1.RID = b.RID" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass c ON c.OCID = a.OCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo d ON d.SID = c.SID" & vbCrLf
        sql &= " JOIN Stud_ServicePlace h ON h.SOCID = c.SOCID" & vbCrLf
        sql &= " JOIN Key_Identity g ON g.IdentityID = c.MIdentityID" & vbCrLf
        sql &= " JOIN Stud_SubSidyCost t ON t.SOCID = c.SOCID" & vbCrLf
        'sql &= " LEFT JOIN (SELECT SOCID ,SUM(Hours) AS TtlHours FROM Stud_Turnout GROUP BY SOCID) st ON st.SOCID = c.SOCID" & vbCrLf
        sql &= " WHERE a.AppliedResultR='Y' AND a.OCID =@OCID" & vbCrLf
        sql &= " ORDER BY c.StudentID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, sParms)
        If dt1.Rows.Count = 0 Then
            strErrmsg = "查無資料!!"
            Return ' Common.MessageBox(Me, "查無資料!!") Exit Sub
        End If

        '收受行代號	收受者帳號	收受者統編	金額	用戶號碼  	公司股市代號	發動者專用區	查詢專用區	收付款結果	就保基金	就安基金
        Dim dtACH As New DataTable
        dtACH.Columns.Add(New DataColumn("收受行代號"))
        dtACH.Columns.Add(New DataColumn("收受者帳號"))
        dtACH.Columns.Add(New DataColumn("收受者統編"))
        dtACH.Columns.Add(New DataColumn("金額"))
        dtACH.Columns.Add(New DataColumn("用戶號碼"))
        dtACH.Columns.Add(New DataColumn("公司股市代號"))
        dtACH.Columns.Add(New DataColumn("發動者專用區"))
        dtACH.Columns.Add(New DataColumn("查詢專用區"))
        dtACH.Columns.Add(New DataColumn("收付款結果"))
        dtACH.Columns.Add(New DataColumn("就保基金"))
        dtACH.Columns.Add(New DataColumn("就安基金"))

        For Each dr As DataRow In dt1.Rows
            dr("BankCode2") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode2"))))
            dr("BankCode3") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode3"))))

            Dim sBankCode2 As String = TIMS.ClearSQM(dr("BankCode2").ToString)
            Dim sBankCode3 As String = TIMS.ClearSQM(dr("BankCode3").ToString)
            Dim sIDNO As String = TIMS.ClearSQM(dr("IDNO").ToString)
            Dim sSumOfMoney As String = TIMS.ClearSQM(dr("SumOfMoney").ToString)
            Dim sName As String = TIMS.ClearSQM(dr("Name").ToString)
            Dim sBudget03 As String = TIMS.ClearSQM(dr("Budget03").ToString)
            Dim sBudget02 As String = TIMS.ClearSQM(dr("Budget02").ToString)

            '收受行代號	收受者帳號	收受者統編	金額	用戶號碼  	公司股市代號	發動者專用區	查詢專用區	收付款結果	就保基金	就安基金
            Dim drACH As DataRow = dtACH.NewRow
            dtACH.Rows.Add(drACH)
            'drWrong("Index") = RowIndex
            drACH("收受行代號") = sBankCode2
            drACH("收受者帳號") = sBankCode3
            drACH("收受者統編") = sIDNO
            drACH("金額") = sSumOfMoney
            drACH("用戶號碼") = ""
            drACH("公司股市代號") = ""
            drACH("發動者專用區") = ""
            drACH("查詢專用區") = sName
            drACH("收付款結果") = ""
            drACH("就保基金") = sBudget03
            drACH("就安基金") = sBudget02
        Next

        Dim dsACH As New DataSet
        dsACH.Tables.Add(dtACH)

        Dim s_fileName1 As String = String.Concat("ACH轉出檔", Replace(Replace(Replace(OCIDVal, ")", ""), "(", ""), "/", ""), ".xlsx")
        ExpClass1.Utl_Export2_XLSX_Direct(Me, dsACH, s_fileName1)

        'Dim rPMS As New Hashtable From {{"OCIDVal", OCIDVal}, {"Sql", sql}}
        'Export_OleDbCMD1(Me, objconn, dt1, rPMS)
    End Sub


    Public Shared Sub Export_OleDbCMD1(MyPage As Page, oConn As SqlConnection, dt1 As DataTable, rPMS As Hashtable)
        Dim vOCIDVal As String = TIMS.GetMyValue2(rPMS, "OCIDVal")
        Dim vSql As String = TIMS.GetMyValue2(rPMS, "Sql")

        Dim sFileName As String = String.Concat("~\SD\07\", TIMS.GetDateNo(), "_", vOCIDVal, ".xls")
        Dim MyPath As String = MyPage.Server.MapPath(sFileName)
        Call TIMS.MyFileDelete(MyPath)
        Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        IO.File.Copy(MyPage.Server.MapPath("~\SD\07\sample3.xls"), MyPath, True)
        Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)

        Dim MyFileName As String = String.Concat("ACH轉出檔", Replace(Replace(Replace(vOCIDVal, ")", ""), "(", ""), "/", ""), ".xls")
        Dim sMemo As String = ""
        sMemo &= "&ACT=" & MyFileName & vbCrLf
        'sMemo &= "&TPLANID=" & sm.UserInfo.TPlanID & vbCrLf
        sMemo &= "&OCIDValue1=" & vOCIDVal & vbCrLf
        sMemo &= "&COUNT=" & dt1.Rows.Count & vbCrLf
        TIMS.SubInsAccountLog1(MyPage, MyPage.Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip2, vOCIDVal, sMemo, oConn)

        '建立資料面
        Dim fg_DATA_OK As Boolean = True

        Dim strErrmsg As String = ""
        Dim OleDbConnStr As String = TIMS.Get_OleDbStr(MyPath)
        Using MyConn As New OleDb.OleDbConnection(OleDbConnStr)
            'MyConn.ConnectionString = TIMS.Get_OleDbStr(MyPath)
            Try
                MyConn.Open()
            Catch ex As Exception
                fg_DATA_OK = False
                TIMS.LOG.Error(ex.Message, ex)
                'Dim strErrmsg As String = ""
                strErrmsg &= String.Concat("/* ex.Message: */", vbCrLf, ex.Message, vbCrLf, "sql:", vbCrLf, vSql, vbCrLf)
                If MyConn IsNot Nothing Then strErrmsg &= "conn.ConnectionString:" & MyConn.ConnectionString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(MyPage, ex) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                'Common.MessageBox(Me, "Excel資料庫無法連線(*.xls)")
                strErrmsg = "Excel資料庫無法連線(*.xls)"
                'Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                If MyConn IsNot Nothing AndAlso MyConn.State = ConnectionState.Open Then MyConn.Close()
                If MyConn IsNot Nothing Then MyConn.Dispose() ': MyConn = Nothing
                Return ' Exit Sub
            End Try

            strErrmsg = ""
            For Each dr As DataRow In dt1.Rows
                dr("BankCode2") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode2"))))
                dr("BankCode3") = TIMS.ChangABC123(TIMS.ChangeIDNO(Convert.ToString(dr("BankCode3"))))

                Dim sBankCode2 As String = TIMS.ClearSQM(dr("BankCode2").ToString)
                Dim sBankCode3 As String = TIMS.ClearSQM(dr("BankCode3").ToString)
                Dim sIDNO As String = TIMS.ClearSQM(dr("IDNO").ToString)
                Dim sSumOfMoney As String = TIMS.ClearSQM(dr("SumOfMoney").ToString)
                Dim sName As String = TIMS.ClearSQM(dr("Name").ToString)
                Dim sBudget03 As String = TIMS.ClearSQM(dr("Budget03").ToString)
                Dim sBudget02 As String = TIMS.ClearSQM(dr("Budget02").ToString)

                '收受行代號	收受者帳號	收受者統編	金額	用戶號碼  	公司股市代號	發動者專用區	查詢專用區	收付款結果	就保基金	就安基金

                Dim sql_odb As String = ""
                sql_odb &= " Insert Into [Sheet1$] (收受行代號,收受者帳號,收受者統編,金額,查詢專用區,就保基金,就安基金)"
                sql_odb &= " Values ('" & sBankCode2 & "','" & sBankCode3 & "','" & sIDNO & "','" & sSumOfMoney & "','" & sName & "','" & sBudget03 & "','" & sBudget02 & "')"
                Using OleDbCmd As New OleDb.OleDbCommand(sql_odb, MyConn)
                    Try
                        OleDbCmd.ExecuteNonQuery()
                    Catch ex As Exception
                        TIMS.LOG.Error(ex.Message, ex)
                        strErrmsg &= String.Concat(If(strErrmsg <> "", "、", ""), sName, "-匯出有誤!")
                    End Try
                End Using
            Next
            If MyConn IsNot Nothing AndAlso MyConn.State = ConnectionState.Open Then MyConn.Close()
            If MyConn IsNot Nothing Then MyConn.Dispose() ': MyConn = Nothing
        End Using
        If strErrmsg <> "" Then Return

        '將新建立的excel存入記憶體下載-----   Start
        'Dim strErrmsg As String = ""
        strErrmsg = ""
        Try
            'Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
            Using fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
                'Dim br As New System.IO.BinaryReader(fr)
                Dim buf(fr.Length) As Byte
                fr.Read(buf, 0, fr.Length)
                fr.Close()

                With MyPage.Response
                    .Clear()
                    .ClearHeaders()
                    .Buffer = False ' False 停用緩衝，確保檔案立即傳送到瀏覽器。True: 緩衝響應，提高效能。 
                    .AddHeader("content-disposition", String.Concat("attachment;filename=", HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8)))
                    .ContentType = "Application/vnd.ms-Excel"
                    'Common.RespWrite(Me, br.ReadBytes(fr.Length))
                    .BinaryWrite(buf)
                    .Flush()
                    .Clear()
                End With
            End Using
        Catch ex As Exception
            strErrmsg = String.Concat("無法存取該檔案!!!", vbCrLf, " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) ", vbCrLf, ex.Message & vbCrLf)
            TIMS.LOG.Warn(strErrmsg, ex)
            TIMS.LOG.Error(ex.Message, ex)
            Return
        End Try

        '刪除Temp中的資料
        Call TIMS.MyFileDelete(MyPath)
        '將新建立的excel存入記憶體下載-----   End
    End Sub

    ''' <summary>
    ''' (特)換字
    ''' </summary>
    ''' <param name="Namestr"></param>
    ''' <returns></returns>
    Function ReplaceSpecialName(ByVal Namestr As String) As String
        '將特殊字轉換   '廖眾峯
        Dim str1 As String
        str1 = Namestr.Trim
        str1 = Replace(str1, "黄", "黃")
        str1 = Replace(str1, "啓", "啟")
        Return str1
    End Function

#Region "(No Use)"

    'Function ReplaceText(ByVal MyText As String) As String
    '    MyText = Replace(MyText, "'", "''")
    '    If MyText Is Nothing Then
    '        MyText = ""
    '    End If
    '    Return MyText
    'End Function

#End Region

    Public Function ConvertBig5(ByVal strUtf As String) As String
        Dim utf81 As Encoding = Encoding.GetEncoding("utf-8")
        Dim big51 As Encoding = Encoding.GetEncoding("big5")
        Response.ContentEncoding = big51
        Dim strUtf81 As Byte() = utf81.GetBytes(strUtf.Trim())
        Dim strBig51 As Byte() = Encoding.Convert(utf81, big51, strUtf81)
        Dim big5Chars1 As Char() = New Char(big51.GetCharCount(strBig51, 0, strBig51.Length) - 1) {}
        big51.GetChars(strBig51, 0, strBig51.Length, big5Chars1, 0)
        Dim tempString1 As New String(big5Chars1)
        Return tempString1
    End Function

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class