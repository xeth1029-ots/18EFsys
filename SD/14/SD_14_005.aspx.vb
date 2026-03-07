Partial Class SD_14_005
    Inherits AuthBasePage

    '學員資料表(學員基本資料表)
    'iReport 'ReportQuery 'SQControl.aspx
    'OLD: 'SD_14_005_2012    '2014: 'SD_14_005_2012_b
    '2016: 'SD_14_005_2016_b
    'Dim iPYNum14 As Integer=1 'TIMS.sUtl_GetPYNum14(Me)
    'Const cst_printFN1 As String="SD_14_005_2016_b"
    Const cst_printFN1 As String = "SD_14_005_2021"
    'Const cst_printFN1O As String="OJTSD1405B1"(報名網空白)
    'Const cst_printFN1S As String="OJTSD1405B1S"(報名網簽名)
    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1
        'iPYNum14=TIMS.sUtl_GetPYNum14(Me)
        Years.Value = sm.UserInfo.Years - 1911

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "0")

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button4_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        Button1.Attributes("onclick") = "return CheckSearch();"
        'Button3.Attributes("onclick")="return CheckPrint('" & ReportQuery.GetSmartQueryPath & "','" & prtFilename & "');"
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效訓練機構/班級!") 'Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        'Dim dt As DataTable
        SOCIDValue.Value = ""

        Dim pParms As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sSql As String = ""
        sSql &= " SELECT a.OCID,a.SOCID,a.StudentID,a.SID" & vbCrLf
        sSql &= " ,a.STUDID2,a.CLASSCNAME2,a.STDate,a.FTDate" & vbCrLf
        sSql &= " ,a.NAME,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,format(a.Birthday,'yyyy/MM/dd') BIRTHDAY" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK2(a.Birthday) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,a.Sex,a.StudStatus,a.AppliedResultM ,a.PlanID ,a.RID" & vbCrLf
        sSql &= " FROM dbo.VIEW_STUDENTBASICDATA a" & vbCrLf
        '排除離退訓學員
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3) AND a.OCID=@OCID" & vbCrLf
        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sSql &= " AND a.OrgKind2='G'" & vbCrLf
                Case "2"
                    '提升勞工自主學習計畫
                    sSql &= " AND a.OrgKind2='W'" & vbCrLf
            End Select
        End If
        sSql &= " ORDER BY a.StudentID " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        '查詢原因
        'Dim v_INQUIRY As String=TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        '查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDID2,NAME,IDNO,BIRTHDAY")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, "", objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        If dt.Rows.Count = 0 Then Return '(沒資料就算了)

        '28:產業人才投資方案
        'KindValue.Value=TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    Dim PNAME As String=PlanPoint.SelectedItem.Text
        '    Select Case PlanPoint.SelectedValue
        '        Case "1", "2"
        '            KindValue.Value=PNAME
        '    End Select
        'End If
        msg.Text = ""
        DataGridTable.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'Call TIMS.set_row_color(DataGrid1)
    End Sub

    Function COMP1(ByRef s_SOCID As String, ByRef s_SOCIDValue As String) As Boolean
        Dim rst As Boolean = False
        If s_SOCID = "" Then Return rst
        If s_SOCIDValue = "" Then Return rst
        Dim SOCIDArray As String() = Split(s_SOCIDValue, ",")
        If SOCIDArray.Length = 0 Then Return rst
        For Each str1 As String In SOCIDArray
            If str1.Equals(s_SOCID) Then
                rst = True
                Return rst
            End If
        Next
        Return rst
    End Function

    Function GET_SEX_N(ByRef s_sex As String) As String
        Dim rst As String = s_sex
        Select Case s_sex
            Case "M"
                rst = "男"
            Case "F"
                rst = "女"
        End Select
        Return rst
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header e.Item.CssClass="head_navy"
        'If e.Item.ItemType=ListItemType.Item Then e.Item.CssClass=""
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim SOCID As HtmlInputCheckBox = e.Item.FindControl("SOCID")
                SOCID.Value = Convert.ToString(drv("SOCID"))
                SOCID.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                SOCID.Checked = COMP1(SOCID.Value, SOCIDValue.Value)

                'e.Item.Cells(1).Text=Right(drv("StudentID"), 2)
                'e.Item.Cells(1).Text=Convert.ToString(drv("STUDID2"))
                e.Item.Cells(4).Text = GET_SEX_N($"{drv("Sex")}")
                e.Item.Cells(6).Text = TIMS.GET_STUDSTATUS_N($"{drv("StudStatus")}")

        End Select
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
    End Sub

    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效訓練機構/班級!")
            Exit Sub
            'Common.MessageBox(Me, TIMS.cst_NODATAMsg1) Exit Sub
        End If

        '列印
        SOCIDValue.Value = TIMS.CombiSQLINM3(SOCIDValue.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Years.Value = TIMS.ClearSQM(Years.Value)
        'Dim tSOCIDValue1 As String = ""
        'If SOCIDValue.Value = "" Then
        '    For Each eItem As DataGridItem In DataGrid1.Items
        '        Dim SOCID As HtmlInputCheckBox = eItem.FindControl("SOCID")
        '        If tSOCIDValue1 <> "" Then tSOCIDValue1 &= ","
        '        tSOCIDValue1 &= SOCID.Value
        '    Next
        'End If
        If SOCIDValue.Value = "" Then
            Common.MessageBox(Me, "請選擇要列印的學員")
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Dim pParms As New Hashtable From {{"OCID", OCIDValue1.Value}, {"RID", RIDValue.Value}}
        Dim sSql As String = ""
        sSql &= " SELECT a.OCID,a.SOCID,a.StudentID,a.SID" & vbCrLf
        sSql &= " ,a.STUDID2,a.CLASSCNAME2,a.STDate,a.FTDate" & vbCrLf
        sSql &= " ,a.NAME,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK2(a.Birthday) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,a.Sex,a.StudStatus,a.AppliedResultM,a.PlanID,a.RID" & vbCrLf
        sSql &= " FROM dbo.VIEW_STUDENTBASICDATA a" & vbCrLf
        '排除離退訓學員
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3) AND a.OCID=@OCID AND a.RID=@RID" & vbCrLf
        Dim V_SOCIDVALUE As String = TIMS.CombiSQLINM3(SOCIDValue.Value)
        If V_SOCIDVALUE <> "" Then sSql &= $" AND a.SOCID IN ({V_SOCIDVALUE})" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDID2,NAME,IDNO")
        'Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(myValue, "SOCID", SOCIDValue.Value)
        TIMS.SetMyValue(myValue, "Years", Years.Value)
        sMemo = myValue
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)

    End Sub
End Class