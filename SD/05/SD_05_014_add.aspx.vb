Partial Class SD_05_014_add
    Inherits AuthBasePage

    'STUD_TURNOUT2 (產投出缺勤使用)
    'SELECT * FROM STUD_TURNOUT2 WHERE ROWNUM <=10
    'SELECT COUNT(1) CNT  FROM STUD_TURNOUT2 WHERE LEAVEID IS NOT NULL
    'SELECT * FROM KEY_LEAVE  
    Const cst_LEAVEID99 As String = "99"  '99(使用者輸入)

    Const cst_ttip01 As String = "學員經費審核結果已經通過，不可修改"
    Const cst_ttip02 As String = "學員經費審核結果已經通過，但系統管理者以上權限可以修改"
    Const cst_ttip03 As String = "學員經費審核結果尚未通過，可修改"
    Const cst_ttip04 As String = "此班級已結訓，且送出補助申請，不可再修改出缺勤作業。"
    Const cst_ttip05 As String = "班級學員經費審核結果已經通過，不可修改"
    Const cst_ttip06 As String = "班級學員經費審核結果已經通過，但系統管理者以上權限可以修改"

    '.當訓練單位送出該班的補助申請作業(不管送出幾個人的補助申請)後
    '該班在學員出缺勤作業的新增或修改作業即予以鎖定，儲存鈕灰階不可按，或是儲存時，顯示告警，不可儲存
    '顯示"此班級已結訓，且送出補助申請，不可再修改出缺勤作業。
    '如訓練單位需修改學員出缺勤資料，則需取消補助申請，或分署整班退回後，才可修改出缺勤作業資料。

    Dim ff3 As String = ""
    Dim dtST2 As DataTable = Nothing
    Dim oTest_flag As Boolean = False  '測試
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'If TIMS.sUtl_ChkTest() Then oTest_flag = True  '測試

        If Not IsPostBack Then
            msg.Text = ""
            Button3.Visible = False
            DataGridTable.Visible = False
            LeaveDate.Text = Now.Date
            If Not Session("_SearchStr") Is Nothing Then
                Me.ViewState("_SearchStr") = Session("_SearchStr")
                'Session("_SearchStr") = Nothing
            End If
            Dim r_STOID As String = TIMS.ClearSQM(Request("STOID"))
            If r_STOID = "" Then
                '新增
                AddTable.Visible = True
                EditTable.Visible = False
                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID

                'Session("_SearchStr") 
                If Session("_SearchStr") IsNot Nothing Then
                    Dim strSearch1 As String = Session("_SearchStr")
                    center.Text = TIMS.GetMyValue(strSearch1, "center")
                    RIDValue.Value = TIMS.GetMyValue(strSearch1, "RIDValue")
                    TMID1.Text = TIMS.GetMyValue(strSearch1, "TMID1")
                    OCID1.Text = TIMS.GetMyValue(strSearch1, "OCID1")
                    TMIDValue1.Value = TIMS.GetMyValue(strSearch1, "TMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(strSearch1, "OCIDValue1")
                    Call Search1()
                End If
            Else
                '修改
                Dim r_SOCID As String = TIMS.ClearSQM(Request("SOCID"))
                Dim r_OCID As String = TIMS.ClearSQM(Request("OCID"))
                STOIDvalue.Value = If(r_STOID <> "", r_STOID, STOIDvalue.Value)
                SOCIDvalue.Value = If(r_STOID <> "", r_SOCID, SOCIDvalue.Value)
                OCIDValue1.Value = If(r_STOID <> "", r_OCID, OCIDValue1.Value)

                AddTable.Visible = False
                EditTable.Visible = True
                Call Create1()
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Style("display") = "none"
        Button3.Attributes("onclick") = "return CheckData();"
    End Sub

    Function Get_TURNOUT2x1(ByVal vOCID As String) As DataTable
        Dim rst As New DataTable
        If vOCID = "" Then Return rst

        Dim sql As String = ""
        sql &= " WITH WCS1 AS (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID)" & vbCrLf
        sql &= " SELECT a.*" & vbCrLf
        sql &= " FROM STUD_TURNOUT2 a" & vbCrLf
        sql &= " JOIN WCS1 c ON c.SOCID = a.SOCID" & vbCrLf
        sql &= " ORDER BY a.LeaveDate" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = vOCID
            rst.Load(.ExecuteReader())
        End With
        Return rst
    End Function

    ''' <summary>(目前累積缺席時數)排除喪假(LEAVEID:05)。 總時數</summary>
    ''' <param name="SOCID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function Get_CountHours1(ByVal SOCID As String) As String
        Dim rst As String = "0"
        Dim sql As String = "" 'sql &= " SELECT SOCID"
        sql &= " SELECT SUM(HOURS) COUNTHOURS "
        sql &= " FROM STUD_TURNOUT2 "
        sql &= " WHERE LEAVEID IS NULL " '排除喪假(LEAVEID:05)。
        sql &= " AND SOCID=@SOCID "
        sql &= " GROUP BY SOCID "
        Call TIMS.OpenDbConn(objconn)
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            rst = Val(dr("COUNTHOURS")) '目前累積缺席時數
        End If
        Return rst
    End Function

    '查詢(依學員) 單1學員資料取得到
    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim ReqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim ReqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        If ReqOCID = "" OrElse ReqSOCID = "" Then Exit Sub

        Dim parms As New Hashtable
        parms.Add("OCID", ReqOCID)
        parms.Add("SOCID", ReqSOCID)
        Dim sql As String = ""
        sql &= " SELECT a.SOCID ,b.STOID" & vbCrLf
        sql &= " ,a.ClassCName" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.StudentID) StudentID" & vbCrLf
        sql &= " ,a.Name" & vbCrLf
        sql &= " ,c.THours" & vbCrLf
        'sql &= "  ,ISNULL(d.COUNTHOURS,0) COUNTHOURS" & vbCrLf
        sql &= " ,b.LeaveDate" & vbCrLf
        sql &= " ,b.Hours" & vbCrLf
        sql &= " ,b.NIHOURS" & vbCrLf
        sql &= " ,b.NIREASONS" & vbCrLf
        sql &= " ,c.AppliedResultM" & vbCrLf
        sql &= " ,b.LEAVEID" & vbCrLf '喪假(LEAVEID:05)。
        sql &= " ,c.OCID" & vbCrLf
        sql &= " ,c.ISCLOSED" & vbCrLf
        sql &= " FROM VIEW_STUDENTBASICDATA a" & vbCrLf
        sql &= " JOIN STUD_TURNOUT2 b ON a.SOCID = b.SOCID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON a.OCID = c.OCID" & vbCrLf
        'sql &= " LEFT JOIN (SELECT SOCID,SUM(HOURS) COUNTHOURS FROM STUD_TURNOUT2 WHERE SOCID = '" & ReqSOCID & "' GROUP BY SOCID) d ON a.SOCID = d.SOCID" & vbCrLf
        sql &= " WHERE a.OCID =@OCID" & vbCrLf
        sql &= " AND a.SOCID =@SOCID" & vbCrLf
        sql &= " ORDER BY b.LeaveDate" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        '系統管理者
        Button5.Enabled = True '儲存
        If dt.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt.Rows(0)
            OCIDValue1.Value = dr1("OCID")

            Dim flag_ISCLOSED As Boolean = If(Convert.ToString(dr1("ISCLOSED")).Equals("Y"), True, False)
            Dim flag_ChkSUBSIDYCOST As Boolean = If(TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, objconn), True, False)
            Dim flag_AppliedResultM As Boolean = If(Convert.ToString(dr1("AppliedResultM")).Equals("Y"), True, False)
            Dim flag_Lock_Edit1 As Boolean = False
            If flag_ISCLOSED OrElse flag_ChkSUBSIDYCOST OrElse flag_AppliedResultM Then flag_Lock_Edit1 = True
            If Not flag_ISCLOSED AndAlso Not flag_AppliedResultM Then flag_Lock_Edit1 = False

            'Dim flag_ISCLOSED As Boolean = If(Convert.ToString(dr("ISCLOSED")).Equals("Y"), True, False)
            'Dim flag_ChkSUBSIDYCOST_1 As Boolean = If(TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, objconn), True, False)
            ''Dim flag_ChkSUBSIDYCOST_2 As Boolean = If(TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, ReqSOCID, objconn), True, False)
            'If flag_ChkSUBSIDYCOST_1 AndAlso flag_ChkSUBSIDYCOST_2 Then Hid_SUBSIDYCOST_FLG.Value = "Y"

            If Not oTest_flag Then
                If flag_Lock_Edit1 Then
                    'Common.MessageBox(Me, cst_ttip04)
                    'Exit Sub
                    Button3.Enabled = False '(不可)儲存-新增
                    TIMS.Tooltip(Button3, cst_ttip04)
                    Button5.Enabled = False '(不可)儲存-修改
                    TIMS.Tooltip(Button5, cst_ttip04)
                    Common.MessageBox(Me, cst_ttip04)
                    'Exit Sub
                End If
            End If
            'Call sUtl_SetSaveButton35(OCIDValue1.Value)
        End If

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Button5.Visible = False
        If dt.Rows.Count > 0 Then
            Button5.Visible = True

            Dim dr As DataRow = dt.Rows(0)
            ClassCName2.Text = dr("ClassCName").ToString
            StudentID2.Text = dr("StudentID").ToString
            Name2.Text = dr("Name").ToString
            THours.Text = dr("THours").ToString
            '排除喪假(LEAVEID:05)。 總時數
            '目前累積缺席時數
            CountHours.Text = Get_CountHours1(Convert.ToString(dr("SOCID"))) 'dr("CountHours").ToString

            DataGrid2.DataSource = dt
            DataGrid2.DataKeyField = "STOID"
            DataGrid2.DataBind()

            If Convert.ToString(dr("AppliedResultM")) = "Y" Then
                If sm.UserInfo.RoleID > 1 Then
                    Button5.Enabled = False
                    TIMS.Tooltip(Button5, cst_ttip05)
                Else
                    '系統管理者
                    TIMS.Tooltip(Button5, cst_ttip06)
                End If
                'Me.Button5.ToolTip = "學員經費審核結果已經通過，不可修改"
            End If
        End If
    End Sub

    '查詢 (依班) 新增
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        If OCIDValue1.Value = "" Then Exit Sub
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "輸入班級有誤!!")
            Exit Sub
        End If

        Dim flag_ISCLOSED As Boolean = If(Convert.ToString(drCC("ISCLOSED")).Equals("Y"), True, False)
        Dim flag_ChkSUBSIDYCOST As Boolean = If(TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, objconn), True, False)
        Dim flag_AppliedResultM As Boolean = If(Convert.ToString(drCC("AppliedResultM")).Equals("Y"), True, False)
        Dim flag_Lock_Edit1 As Boolean = False
        If flag_ISCLOSED OrElse flag_ChkSUBSIDYCOST OrElse flag_AppliedResultM Then flag_Lock_Edit1 = True
        If Not flag_ISCLOSED AndAlso Not flag_AppliedResultM Then flag_Lock_Edit1 = False

        If Not oTest_flag Then
            If flag_Lock_Edit1 Then
                Button3.Enabled = False '(不可)儲存-新增
                TIMS.Tooltip(Button3, cst_ttip04)
                Button5.Enabled = False '(不可)儲存-修改
                TIMS.Tooltip(Button5, cst_ttip04)
                Common.MessageBox(Me, cst_ttip04)
            End If
        End If

        'Call sUtl_SetSaveButton35(OCIDValue1.Value)
        'If TIMS.sUtl_ChkSUBSIDYCOST(OCIDValue1.Value, objconn) Then Hid_SUBSIDYCOST_FLG.Value = "Y"
        'If Not oTest_flag AndAlso Hid_SUBSIDYCOST_FLG.Value = "Y" Then
        '    Button3.Enabled = False '(不可)儲存
        '    TIMS.Tooltip(Button3, cst_ttip04)
        '    Button5.Enabled = False '(不可)儲存
        '    TIMS.Tooltip(Button5, cst_ttip04)
        '    Common.MessageBox(Me, cst_ttip04)
        '    'Exit Sub
        'End If

        'Dim dt As DataTable = Nothing
        Dim hPMS As New Hashtable From {{"OCID", Val(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT a.SOCID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.StudentID) StudentID" & vbCrLf
        sql &= " ,a.Name" & vbCrLf
        sql &= " ,a.AppliedResultM" & vbCrLf
        sql &= " ,ct.AppliedStatusM AppliedStatusM" & vbCrLf
        sql &= " ,ct.APPLIEDSTATUS" & vbCrLf
        sql &= " ,ISNULL(b.CountHours,0) CountHours" & vbCrLf '目前累積缺席時數
        sql &= " FROM VIEW_STUDENTBASICDATA a" & vbCrLf

        '目前累積缺席時數
        sql &= " LEFT JOIN ( SELECT t.SOCID" & vbCrLf
        sql &= "   ,SUM(t.Hours) COUNTHOURS" & vbCrLf '全部時數不區分 ('不含喪假(LEAVEID:05)。)
        sql &= "   FROM STUD_TURNOUT2 t" & vbCrLf
        sql &= "   JOIN CLASS_STUDENTSOFCLASS cs ON cs.SOCID=t.SOCID" & vbCrLf
        sql &= "   WHERE t.LEAVEID IS NULL " '排除喪假(LEAVEID:05)。
        sql &= "   AND cs.OCID=@OCID" & vbCrLf
        sql &= "   GROUP BY t.SOCID" & vbCrLf
        sql &= " ) b ON a.SOCID = b.SOCID" & vbCrLf

        sql &= " LEFT JOIN STUD_SUBSIDYCOST ct on ct.SOCID=a.SOCID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID AND a.STUDSTATUS NOT IN (2,3)" & vbCrLf '排除離退訓學員輸入資料 by AMU 20090916
        sql &= " ORDER BY StudentID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, hPMS)

        Button3.Visible = False
        msg.Text = "查無此班學員資料"
        DataGridTable.Visible = False

        If dt.Rows.Count > 0 Then
            'Dim dtST2 As DataTable
            dtST2 = Get_TURNOUT2x1(OCIDValue1.Value)

            Button3.Visible = True
            msg.Text = ""
            DataGridTable.Visible = True


            scrollDiv.Attributes.Add("class", "DivHeight")
            'msg.Text = ""
            'Button3.Visible = True
            'DataGridTable.Visible = True
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '查詢隱藏學員
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "FixedTitleRow"
                'e.Item.CssClass = "SD_TD1"
                'e.Item.Attributes.Add("class", "FixedTitleRow")  '20100204 andy edit 固定第一列
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim drv As DataRowView = e.Item.DataItem

                Dim Hours As TextBox = e.Item.FindControl("Hours") '缺席時數
                Dim hidLeaveDate3 As HtmlInputHidden = e.Item.FindControl("hidLeaveDate3")
                Dim HidSOCID As HiddenField = e.Item.FindControl("HidSOCID")
                Dim NIHOURS As TextBox = e.Item.FindControl("NIHOURS") '不列入缺席時數
                Dim NIREASONS As TextBox = e.Item.FindControl("NIREASONS") '不列入缺席原因
                Dim vvvSOCID As String = TIMS.ClearSQM(drv("SOCID"))
                HidSOCID.Value = Convert.ToString(vvvSOCID)

                ff3 = "SOCID='" & vvvSOCID & "'"
                hidLeaveDate3.Value = ""
                If dtST2 Is Nothing Then dtST2 = Get_TURNOUT2x1(OCIDValue1.Value)
                If Not dtST2 Is Nothing AndAlso dtST2.Select(ff3, "LeaveDate").Length > 0 Then
                    For Each dr As DataRow In dtST2.Select(ff3, "LeaveDate")
                        Me.ViewState("LeaveDate3") = Common.FormatDate(dr("LeaveDate"))
                        If Me.ViewState("LeaveDate3") <> "" Then
                            If hidLeaveDate3.Value <> "" Then hidLeaveDate3.Value &= ","
                            hidLeaveDate3.Value &= Me.ViewState("LeaveDate3")
                        End If
                    Next
                End If

                e.Item.Cells(0).Attributes("onclick") = "wopen('SD_05_014_His.aspx?SOCID=" & vvvSOCID & "','His',400,300,1);"
                e.Item.Cells(0).Style("CURSOR") = "hand"

                e.Item.Cells(1).Attributes("onclick") = "wopen('SD_05_014_His.aspx?SOCID=" & vvvSOCID & "','His',400,300,1);"
                e.Item.Cells(1).Style("CURSOR") = "hand"

                If Convert.ToString(drv("AppliedStatusM")).Equals("Y") Then
                    Hours.Enabled = False
                    NIHOURS.Enabled = False
                    NIREASONS.Enabled = False
                    TIMS.Tooltip(Hours, cst_ttip01)
                    TIMS.Tooltip(NIHOURS, cst_ttip01)
                    TIMS.Tooltip(NIREASONS, cst_ttip01)
                ElseIf Convert.ToString(drv("APPLIEDSTATUS")).Equals("1") Then
                    Hours.Enabled = False
                    NIHOURS.Enabled = False
                    NIREASONS.Enabled = False
                    TIMS.Tooltip(Hours, cst_ttip04)
                    TIMS.Tooltip(NIHOURS, cst_ttip04)
                    TIMS.Tooltip(NIREASONS, cst_ttip04)
                End If

        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "SD_TD1"
                e.Item.CssClass = "head_navy"  'edit，by:20181121
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"

                Dim drv As DataRowView = e.Item.DataItem
                Dim LeaveDate2 As TextBox = e.Item.FindControl("LeaveDate2")
                Dim Hours2 As TextBox = e.Item.FindControl("Hours2")
                'Dim NIHOURS2 As TextBox = e.Item.FindControl("NIHOURS2")
                Dim NIREASONS As TextBox = e.Item.FindControl("NIREASONS")

                Dim IMG1 As HtmlImage = e.Item.FindControl("IMG1")
                Dim Change As HtmlInputHidden = e.Item.FindControl("Change") '判斷是否有異動情況。
                Dim HidSTOID As HiddenField = e.Item.FindControl("HidSTOID") '序號
                HidSTOID.Value = Convert.ToString(drv("STOID"))
                'Dim chk_LEAVEID05 As CheckBox = e.Item.FindControl("chk_LEAVEID05")

                If drv("STOID") = Request("STOID") Then e.Item.BackColor = Color.FromName("#FFF4FF")
                'LeaveDate2.Text = FormatDateTime(drv("LeaveDate"), 2)
                'chk_LEAVEID05.Checked = False '喪假(LEAVEID:05)。
                'If Convert.ToString(drv("LEAVEID")) = "05" Then chk_LEAVEID05.Checked = True
                LeaveDate2.Text = TIMS.Cdate3(drv("LeaveDate"))
                If Convert.ToString(drv("Hours")) <> "" Then Hours2.Text = Val(drv("Hours"))
                If Convert.ToString(drv("NIHOURS")) <> "" Then
                    If Val(drv("NIHOURS")) > 0 Then Hours2.Text = Val(drv("NIHOURS"))
                End If
                NIREASONS.Text = Convert.ToString(drv("NIREASONS"))

                IMG1.Attributes("onclick") = "show_calendar('" & LeaveDate2.ClientID & "','','','CY/MM/DD');document.getElementById('" & Change.ClientID & "').value='1';"
                LeaveDate2.Attributes("onchange") = "document.getElementById('" & Change.ClientID & "').value='1';"
                Hours2.Attributes("onchange") = "document.getElementById('" & Change.ClientID & "').value='1';"
                'NIHOURS2.Attributes("onchange") = "document.getElementById('" & Change.ClientID & "').value='1';"
                NIREASONS.Attributes("onchange") = "document.getElementById('" & Change.ClientID & "').value='1';"
                'chk_LEAVEID05.Attributes("onclick") = "document.getElementById('" & Change.ClientID & "').value='1';"
        End Select
    End Sub

    '新增儲存前 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True

        LeaveDate.Text = TIMS.ClearSQM(LeaveDate.Text)
        If LeaveDate.Text = "" Then
            Errmsg += "未出席 日期必須填寫" & vbCrLf
            Return False
        End If
        If Not TIMS.IsDate1(LeaveDate.Text) Then
            Errmsg += "未出席 日期格式有誤，應為yyyy/MM/dd" & vbCrLf
            Return False
        End If
        LeaveDate.Text = CDate(LeaveDate.Text).ToString("yyyy/MM/dd")

        Dim vsSTDate As String = ""
        Dim vsFTDate As String = ""
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value)
        If drCC Is Nothing Then
            Errmsg += "班級選擇有誤" & vbCrLf
            Return False
        End If
        vsSTDate = TIMS.Cdate3(drCC("STDate")) 'CDate(dr("STDate")).ToString("yyyy/MM/dd")
        vsFTDate = TIMS.Cdate3(drCC("FTDate")) 'CDate(dr("FTDate")).ToString("yyyy/MM/dd")

        If Errmsg = "" Then
            If LeaveDate.Text <> "" AndAlso vsFTDate <> "" Then
                If CDate(LeaveDate.Text) > CDate(vsFTDate) Then Errmsg += "【未出席日期】不得大於【結訓日期】!!" & vbCrLf
            End If
        End If
        If Errmsg = "" Then
            If LeaveDate.Text <> "" AndAlso vsSTDate <> "" Then
                If CDate(LeaveDate.Text) < CDate(vsSTDate) Then Errmsg += "【未出席日期】不得小於【開訓日期】!!" & vbCrLf
            End If
        End If

        Dim hPms2 As New Hashtable From {{"OCID", Val(OCIDValue1.Value)}}
        Dim sql As String = "SELECT SOCID, OCID FROM CLASS_STUDENTSOFCLASS WHERE OCID =@OCID"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn, hPms2)

        Dim iX As Integer = 0 'i = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            iX += 1
            Dim sSeqMsg As String = "第" & iX & "筆資料："
            Dim Hours As TextBox = eItem.FindControl("Hours")
            Dim HidSOCID As HiddenField = eItem.FindControl("HidSOCID")
            'Dim chk_LEAVEID05 As CheckBox = eItem.FindControl("chk_LEAVEID05")
            Dim NIHOURS As TextBox = eItem.FindControl("NIHOURS")
            Dim NIREASONS As TextBox = eItem.FindControl("NIREASONS")

            'If Hours.Text.Trim() = "" Then
            '    Hours.Text = "0"
            'End If
            'SOCID
            If Not IsNumeric(HidSOCID.Value) Then
                Errmsg += "學員學號有誤!!" & vbCrLf
                Exit For
            End If
            If HidSOCID.Value = "" Then
                Errmsg += "學員學號有誤!!" & vbCrLf
                Exit For
            End If
            ff3 = "SOCID='" & HidSOCID.Value & "'"
            If dt2.Select(ff3).Length = 0 Then
                Errmsg += "班級學員選擇有誤!!" & vbCrLf
                Exit For
            End If

            Const cst_sub1 As String = "缺席時數"
            Const cst_sub2 As String = "不列入缺席時數"
            Const cst_sub3 As String = "不列入缺席原因"
            Hours.Text = TIMS.ClearSQM(Hours.Text)
            NIHOURS.Text = TIMS.ClearSQM(NIHOURS.Text)
            NIREASONS.Text = TIMS.ClearSQM(NIREASONS.Text)
            If Hours.Text = "0" Then Hours.Text = ""
            If NIHOURS.Text = "0" Then NIHOURS.Text = ""
            '有填資料才檢核
            If Hours.Text <> "" OrElse NIHOURS.Text <> "" OrElse NIREASONS.Text <> "" Then
                If (Hours.Text <> "" OrElse NIHOURS.Text = "") AndAlso NIREASONS.Text <> "" Then
                    Errmsg += sSeqMsg & cst_sub1 & "有填或是" & cst_sub2 & "未填，" & cst_sub3 & "有填，邏輯資料有誤!!" & vbCrLf
                    Exit For
                End If
                If (Hours.Text = "" OrElse NIHOURS.Text <> "") AndAlso NIREASONS.Text = "" Then
                    Errmsg += sSeqMsg & cst_sub1 & "未填或是" & cst_sub2 & "有填，" & cst_sub3 & "未填，邏輯資料有誤!!" & vbCrLf
                    Exit For
                End If
                If Hours.Text <> "" AndAlso NIHOURS.Text <> "" Then
                    Errmsg += sSeqMsg & cst_sub1 & "有填，" & cst_sub2 & "有填，邏輯資料有誤!!" & vbCrLf
                    Exit For
                End If
                If Hours.Text = "" AndAlso NIHOURS.Text = "" Then
                    Errmsg += sSeqMsg & cst_sub1 & "未填，" & cst_sub2 & "未填(時數格式有誤，不得為空)" & vbCrLf
                    Exit For
                End If
                '20080528  Andy
                '又改可以輸入小數點了
                If Not TIMS.Chk_GetHoursNum5(cst_sub1, Hours.Text, Errmsg) Then Exit For
                If Not TIMS.Chk_GetHoursNum5(cst_sub2, NIHOURS.Text, Errmsg) Then Exit For
                'If Hours.Text = "" AndAlso chk_LEAVEID05.Checked Then
                '    Errmsg += "勾選喪假未填寫時數，資料有誤" & vbCrLf
                '    Exit For
                'End If
            End If
            'i += 1
        Next

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '修改儲存前檢查
    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True

        Dim vsSTDate As String = ""
        Dim vsFTDate As String = ""
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Errmsg += "班級選擇有誤" & vbCrLf
            Return False
        End If
        vsSTDate = CDate(drCC("STDate")).ToString("yyyy/MM/dd")
        vsFTDate = CDate(drCC("FTDate")).ToString("yyyy/MM/dd")

        SOCIDvalue.Value = TIMS.ClearSQM(SOCIDvalue.Value)

        Dim hPms As New Hashtable From {{"SOCID", Val(SOCIDvalue.Value)}}
        Dim sql1 As String = " SELECT * FROM STUD_TURNOUT2 WHERE SOCID =@SOCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql1, objconn, hPms)
        If dt.Rows.Count = 0 Then
            Errmsg += "學員選擇有誤" & vbCrLf
            Return False
        End If

        Dim hPms2 As New Hashtable From {{"SOCID", Val(SOCIDvalue.Value)}, {"OCID", Val(OCIDValue1.Value)}}
        Dim sql2 As String = ""
        sql2 &= " SELECT 'x' FROM CLASS_STUDENTSOFCLASS "
        sql2 &= " WHERE SOCID =@SOCID AND OCID =@OCID"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, hPms2)
        If dt2.Rows.Count <> 1 Then
            Errmsg += "班級學員選擇有誤" & vbCrLf
            Return False
        End If

        'i = 0
        Dim iX As Integer = 0
        For Each eItem As DataGridItem In DataGrid2.Items
            iX += 1
            Dim LeaveDate2 As TextBox = eItem.FindControl("LeaveDate2")
            Dim Change As HtmlInputHidden = eItem.FindControl("Change")
            Dim Hours2 As TextBox = eItem.FindControl("Hours2")
            'Dim NIHOURS2 As TextBox = eItem.FindControl("NIHOURS2")
            Dim NIREASONS As TextBox = eItem.FindControl("NIREASONS")
            Dim HidSTOID As HiddenField = eItem.FindControl("HidSTOID")
            'Dim chk_LEAVEID05 As CheckBox = item.FindControl("chk_LEAVEID05")

            If Change.Value = "1" Then
                If HidSTOID.Value = "" Then
                    Errmsg += "學員選擇代碼有誤" & vbCrLf
                    Exit For
                End If
                If dt.Select("STOID='" & HidSTOID.Value & "'").Length = 0 Then
                    Errmsg += "學員選擇代碼有誤" & vbCrLf
                    Exit For
                End If

                Dim sSeqMsg As String = "第" & iX & "筆資料："
                Const cst_sub1 As String = "缺席時數"
                'Const cst_sub2 As String = "不列入缺席時數"
                'Const cst_sub3 As String = "不列入缺席原因"
                Hours2.Text = TIMS.ClearSQM(Hours2.Text)
                'NIHOURS2.Text = TIMS.ClearSQM(NIHOURS2.Text)
                NIREASONS.Text = TIMS.ClearSQM(NIREASONS.Text)
                If Hours2.Text = "0" Then Hours2.Text = ""
                'If NIHOURS2.Text = "0" Then NIHOURS2.Text = ""
                '有填資料才檢核

                '20080528  Andy
                '又改可以輸入小數點了
                If Not TIMS.Chk_GetHoursNum5(cst_sub1, Hours2.Text, Errmsg) Then Exit For

                LeaveDate2.Text = TIMS.ClearSQM(LeaveDate2.Text)
                'LeaveDate2.Text = Trim(LeaveDate2.Text)
                If LeaveDate2.Text <> "" Then
                    If Not TIMS.IsDate1(LeaveDate2.Text) Then
                        Errmsg += sSeqMsg & "未出席日期格式有誤，應為yyyy/MM/dd" & vbCrLf
                        Exit For
                    Else
                        LeaveDate2.Text = CDate(LeaveDate2.Text).ToString("yyyy/MM/dd")
                    End If
                Else
                    'LeaveDate2.Text = ""
                    Errmsg += sSeqMsg & "未出席日期必須填寫" & vbCrLf
                    Exit For
                End If

                If Errmsg = "" Then
                    If LeaveDate2.Text <> "" AndAlso vsFTDate <> "" Then
                        If CDate(LeaveDate2.Text) > CDate(vsFTDate) Then
                            Errmsg += sSeqMsg & "【未出席日期】不得大於【結訓日期】!!" & vbCrLf
                            Exit For
                        End If
                    End If
                End If
                If Errmsg = "" Then
                    If LeaveDate2.Text <> "" AndAlso vsSTDate <> "" Then
                        If CDate(LeaveDate2.Text) < CDate(vsSTDate) Then
                            Errmsg += sSeqMsg & "【未出席日期】不得小於【開訓日期】!!" & vbCrLf
                            Exit For
                        End If
                    End If
                End If
            End If
            'i += 1
        Next

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存1(新增儲存)
    Sub SaveData1()
        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, "儲存失敗有誤!" & ex.ToString)
        'End Try
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow
        Dim sql As String
        sql = " SELECT * FROM STUD_TURNOUT2 WHERE 1<>1 "
        dt = DbAccess.GetDataTable(sql, da, objconn)
        '2006/03/28 add conn by matt
        'Dim i As Integer = 0
        'i = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Hours As TextBox = eItem.FindControl("Hours")
            Dim hidLeaveDate3 As HtmlInputHidden = eItem.FindControl("hidLeaveDate3")
            Dim HidSOCID As HiddenField = eItem.FindControl("HidSOCID")
            Dim NIHOURS As TextBox = eItem.FindControl("NIHOURS")
            Dim NIREASONS As TextBox = eItem.FindControl("NIREASONS")

            'Dim Hours As TextBox = item.FindControl("Hours")
            'Dim HidSOCID As HiddenField = item.FindControl("HidSOCID")
            'Dim chk_LEAVEID05 As CheckBox = item.FindControl("chk_LEAVEID05")
            NIREASONS.Text = TIMS.ClearSQM(NIREASONS.Text)
            If NIREASONS.Text.Length > 10 Then NIREASONS.Text = NIREASONS.Text.Substring(0, 10)
            HidSOCID.Value = TIMS.ClearSQM(HidSOCID.Value)

            '必定輸入時數
            If Hours.Text <> "" OrElse NIHOURS.Text <> "" OrElse NIREASONS.Text <> "" Then
                Dim iSTOID As Integer = DbAccess.GetNewId(objconn, "STUD_TURNOUT2_STOID_SEQ,STUD_TURNOUT2,STOID")
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("STOID") = iSTOID
                dr("SOCID") = Val(HidSOCID.Value) 'DataGrid1.DataKeys(i) 'SOCID
                dr("LeaveDate") = TIMS.Cdate2(LeaveDate.Text)
                'dr("Hours") = Val(Hours.Text)
                'dr("LEAVEID") = IIf(chk_LEAVEID05.Checked, "05", Convert.DBNull)
                If NIREASONS.Text <> "" Then
                    dr("Hours") = 0
                    dr("NIHOURS") = Val(NIHOURS.Text)
                    dr("NIREASONS") = NIREASONS.Text
                    dr("LEAVEID") = cst_LEAVEID99 '"99"
                Else
                    dr("Hours") = Val(Hours.Text)
                    dr("NIHOURS") = Convert.DBNull
                    dr("NIREASONS") = Convert.DBNull
                    dr("LEAVEID") = Convert.DBNull
                End If
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
            End If
            'i += 1
        Next
        DbAccess.UpdateDataTable(dt, da)

        Session("_SearchStr") = Me.ViewState("_SearchStr")
        Common.RespWrite(Me, "<script>alert('儲存成功');location.href='SD_05_014.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '儲存2(修改儲存)
    Sub SaveData2()
        Dim ReqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        '學員資料不同(不儲存)
        If SOCIDvalue.Value = "" OrElse SOCIDvalue.Value <> ReqSOCID Then Exit Sub
        If SOCIDvalue.Value <> ReqSOCID Then Exit Sub '學員資料不同(不儲存)

        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = ""
        sql &= " UPDATE STUD_TURNOUT2 "
        sql &= " SET LeaveDate = @LeaveDate "
        sql &= " ,Hours = @Hours "
        sql &= " ,NIHOURS = @NIHOURS "
        sql &= " ,NIREASONS = @NIREASONS "
        sql &= " ,LEAVEID = @LEAVEID "
        sql &= " ,ModifyAcct = @ModifyAcct "
        sql &= " ,ModifyDate = GETDATE() "
        sql &= " WHERE STOID = @STOID "
        Dim uCmd As New SqlCommand(sql, objconn)
        '刪除未開發
        'Dim dCmd As New SqlCommand(sql, objconn)

        For Each eItem As DataGridItem In DataGrid2.Items
            Dim LeaveDate2 As TextBox = eItem.FindControl("LeaveDate2")
            Dim Hours2 As TextBox = eItem.FindControl("Hours2")
            'Dim NIHOURS2 As TextBox = eItem.FindControl("NIHOURS2")
            Dim NIREASONS As TextBox = eItem.FindControl("NIREASONS")
            'Dim IMG1 As HtmlImage = e.Item.FindControl("IMG1")
            Dim Change As HtmlInputHidden = eItem.FindControl("Change") '1:有異動 '判斷是否有異動情況。
            'Dim chk_LEAVEID05 As CheckBox = item.FindControl("chk_LEAVEID05")
            Dim HidSTOID As HiddenField = eItem.FindControl("HidSTOID") '序號

            NIREASONS.Text = TIMS.ClearSQM(NIREASONS.Text)
            If NIREASONS.Text.Length > 10 Then NIREASONS.Text = NIREASONS.Text.Substring(0, 10)

            HidSTOID.Value = TIMS.ClearSQM(HidSTOID.Value)
            If Change.Value = "1" Then
                If Val(Hours2.Text) = -1 Then
                    '刪除未開發
                Else
#Region "(No Use)"

                    '又改可以輸入小數點了 by andy 20080623  
                    'If Hours2.Text <> "" OrElse NIHOURS2.Text <> "" OrElse NIREASONS.Text <> "" Then
                    '    With uCmd
                    '        .Parameters.Clear()
                    '        .Parameters.Add("LeaveDate", SqlDbType.DateTime).Value = CDate(TIMS.cdate2(LeaveDate2.Text))
                    '        '.Parameters.Add("Hours", SqlDbType.VarChar).Value = Val(Hours2.Text)
                    '        '.Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = IIf(chk_LEAVEID05.Checked, "05", Convert.DBNull)
                    '        If NIREASONS.Text <> "" Then
                    '            .Parameters.Add("Hours", SqlDbType.Int).Value = 0
                    '            .Parameters.Add("NIHOURS", SqlDbType.Int).Value = Val(NIHOURS2.Text)
                    '            .Parameters.Add("NIREASONS", SqlDbType.VarChar).Value = NIREASONS.Text
                    '            .Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = cst_LEAVEID99 '"99"0
                    '        Else
                    '            .Parameters.Add("Hours", SqlDbType.Int).Value = Val(Hours2.Text)
                    '            .Parameters.Add("NIHOURS", SqlDbType.Int).Value = Convert.DBNull
                    '            .Parameters.Add("NIREASONS", SqlDbType.VarChar).Value = Convert.DBNull
                    '            .Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = Convert.DBNull
                    '        End If

                    '        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '        .Parameters.Add("STOID", SqlDbType.VarChar).Value = HidSTOID.Value
                    '        .ExecuteNonQuery()
                    '    End With
                    'End If

#End Region

                    If Hours2.Text <> "" OrElse NIREASONS.Text <> "" Then
                        With uCmd
                            .Parameters.Clear()
                            .Parameters.Add("LeaveDate", SqlDbType.DateTime).Value = CDate(TIMS.Cdate2(LeaveDate2.Text))
                            '.Parameters.Add("Hours", SqlDbType.VarChar).Value = Val(Hours2.Text)
                            '.Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = IIf(chk_LEAVEID05.Checked, "05", Convert.DBNull)
                            If NIREASONS.Text <> "" Then
                                .Parameters.Add("Hours", SqlDbType.Float).Value = 0
                                .Parameters.Add("NIHOURS", SqlDbType.Float).Value = Val(Hours2.Text)
                                .Parameters.Add("NIREASONS", SqlDbType.VarChar).Value = NIREASONS.Text
                                .Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = cst_LEAVEID99 '"99"0
                            Else
                                .Parameters.Add("Hours", SqlDbType.Float).Value = Val(Hours2.Text)
                                .Parameters.Add("NIHOURS", SqlDbType.Float).Value = Convert.DBNull
                                .Parameters.Add("NIREASONS", SqlDbType.VarChar).Value = Convert.DBNull
                                .Parameters.Add("LEAVEID", SqlDbType.VarChar).Value = Convert.DBNull
                            End If

                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("STOID", SqlDbType.VarChar).Value = HidSTOID.Value
                            '.ExecuteNonQuery()
                            DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)  'edit，by:20181031
                        End With
                    End If
                End If
            End If
            'i += 1
        Next
        'DbAccess.UpdateDataTable(dt, da)
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        Common.RespWrite(Me, "<script>alert('修改成功');location.href='SD_05_014.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '班級未出席儲存 (新增儲存 整班)
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
    End Sub

    '回上一頁
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        TIMS.Utl_Redirect1(Me, "SD_05_014.aspx?ID=" & Request("ID") & "")
    End Sub

    '學員未出席儲存(單1學員 修改儲存)
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Errmsg As String = ""
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData2()
    End Sub

    '回上一頁。
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        TIMS.Utl_Redirect1(Me, "SD_05_014.aspx?ID=" & Request("ID") & "")
    End Sub

End Class