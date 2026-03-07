Partial Class SD_14_016
    Inherits AuthBasePage

    'ReportQuery 'SD_14_016*.jrxml -- @BussinessTrain 參訓學員補助費簽領清冊
    Const cst_reportFN1 As String = "SD_14_016"

    Const cst_列印說明 As Integer = 4
    Const cst_Msg_a1 As String = "學員經費尚未審核"
    Const cst_Msg_a2 As String = "僅供系統指定計畫列印<br>(其它計畫暫不提供列印)."
    Const cst_Msg_a3 As String = "非產業人才投資計畫(不提供列印)."

    Const cst_Msg_a4 As String = "學員經費審核.不通過"
    Const cst_Msg_a5 As String = "學員經費審核.退件修正"
    Const cst_Msg_a6 As String = "學員經費尚未審核."
    Const cst_Msg_a7 As String = "學員經費審核.其他"

    Const cst_errMsg1 As String = "參數異常，請重新查詢!!"

    Dim str_title1 As String = ""
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
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            print_orderyby.Value = If(print_type.SelectedValue = "2", "c.StudentID", "d.IDNO")

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2

            print_type.Attributes("onclick") = "printkind();" '列印時排序方式
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '使用民國年。
        'Years.Value = sm.UserInfo.Years - 1911
        'PlanID.Value = sm.UserInfo.PlanID
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)

        If STDate1.Text <> "" Then
            'STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If STDate2.Text <> "" Then
            'STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate1.Text <> "" Then
            'FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate2.Text <> "" Then
            'FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If Errmsg = "" Then
            If Me.STDate1.Text <> "" AndAlso Me.STDate2.Text <> "" Then
                If DateDiff(DateInterval.Day, CDate(STDate1.Text), CDate(STDate2.Text)) < 0 Then
                    Errmsg += "開訓期間 日期起迄，迄日需大起日" & vbCrLf
                End If
            End If
        End If

        If Errmsg = "" Then
            If Me.FTDate1.Text <> "" AndAlso Me.FTDate2.Text <> "" Then
                If DateDiff(DateInterval.Day, CDate(FTDate1.Text), CDate(FTDate2.Text)) < 0 Then
                    Errmsg += "結訓期間 日期起迄，迄日需大起日" & vbCrLf
                End If
            End If
        End If

        'If Syear.SelectedValue = "" _
        '    AndAlso STDate1.Text = "" _
        '    AndAlso STDate2.Text = "" _
        '    AndAlso FTDate1.Text = "" _
        '    AndAlso FTDate2.Text = "" _
        '    Then
        '    Errmsg += "[年度]、[開訓區間]、[結訓區間],請擇一輸入查詢" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim pms1 As New Hashtable From {{"RelShip", RelShip}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) ClassCName" & vbCrLf
        sql &= " ,a.STDate ,a.FTDate" & vbCrLf
        sql &= " ,a.AppliedResultM" & vbCrLf
        sql &= " ,b.OrgName ,b.OrgKind" & vbCrLf
        sql &= " FROM Class_ClassInfo a" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID =a.PlanID" & vbCrLf
        sql &= " WHERE b.RelShip like @RelShip+'%'" & vbCrLf

        If sm.UserInfo.RID = "A" Then
            sql &= " and ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " and ip.Years ='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " and ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If

        If STDate1.Text <> "" Then
            sql &= " and a.STDate>= " & TIMS.To_date(STDate1.Text) & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql &= " and a.STDate<= " & TIMS.To_date(STDate2.Text) & vbCrLf 'convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            sql &= " and a.FTDate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf 'convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql &= " and a.FTDate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            str_title1 = TIMS.Get_PName28(Me, "W", objconn)

            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Return
        'If sCmdArg = "" Then
        '    Common.MessageBox(Me, cst_errMsg1)
        '    Return 'Exit Sub
        'End If

        Const cst_printCmd1 As String = "print"
        Select Case e.CommandName
            Case cst_printCmd1 '"print"
                Dim sYears As String = TIMS.GetMyValue(sCmdArg, "Years")
                Dim sOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
                Dim v_print_type As String = TIMS.GetListValue(print_type)
                print_orderyby.Value = If(v_print_type = "2", "c.StudentID", "d.IDNO")
                v_print_type = If(v_print_type = "2", v_print_type, "1") '(過濾值)
                Dim myValue As String = ""
                myValue &= "&Years=" & sYears
                myValue &= "&OCID=" & sOCID
                myValue &= "&Printtype=" & print_orderyby.Value
                Select Case v_print_type
                    Case "2"
                        myValue &= "&Printtype2=" & print_orderyby.Value
                    Case Else '"1"
                        myValue &= "&Printtype1=" & print_orderyby.Value
                End Select
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN1, myValue)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""

                Dim drv As DataRowView = e.Item.DataItem
                Dim BtnPrint As Button = e.Item.FindControl("BtnPrint") '列印'e.Item.Cells(cst_列印說明)
                Dim labMsg1 As Label = e.Item.FindControl("labMsg1")
                labMsg1.Text = "" '(列印說明)一般狀況不顯示資訊。
                BtnPrint.Visible = False '一般狀況不顯示列印按鈕。
                'btn.Visible = False
                If Convert.ToString(drv("AppliedResultM")) = "" Then
                    'labMsg1.Text = cst_Msg_a1 '因為會蓋過按鈕，所以只有在無值的情況下，才顯示此訊息。
                    labMsg1.Text = cst_Msg_a1 '因為會蓋過按鈕，所以只有在無值的情況下，才顯示此訊息。
                End If

                If Convert.ToString(drv("AppliedResultM")) <> "" Then
                    Select Case Convert.ToString(drv("AppliedResultM"))
                        Case "Y"
                            If TIMS.Cst_CanPrintSD_14_016.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                If Convert.ToString(drv("OrgKind")) = "10" Then
                                    '提升在職勞工自主學習計畫
                                    'btn.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_016&path=TIMS&Years=" & (sm.UserInfo.Years - 1911) & "&OCID=" & drv("OCID") & "');"
                                    Dim cmdArg As String = ""
                                    cmdArg = ""
                                    cmdArg &= "&Years=" & CStr(CInt(sm.UserInfo.Years) - 1911)
                                    cmdArg &= "&OCID=" & CStr(drv("OCID"))
                                    'btn.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_14_016", MyValue)
                                    BtnPrint.Visible = True '顯示列印按鈕。
                                    BtnPrint.CommandArgument = cmdArg
                                Else
                                    '提升在職勞工自主學習計畫
                                    'labMsg1.Text = "產業人才投資計畫(不提供列印)."
                                    labMsg1.Text = cst_Msg_a2
                                    '"提升在職勞工自主學習計畫"
                                    TIMS.Tooltip(labMsg1, str_title1)
                                End If
                            Else
                                labMsg1.Text = cst_Msg_a3
                            End If
                        Case "N"
                            labMsg1.Text = cst_Msg_a4 '"學員經費審核.不通過"
                        Case "R"
                            labMsg1.Text = cst_Msg_a5 '"學員經費審核.退件修正"
                        Case " "
                            labMsg1.Text = cst_Msg_a6 '"學員經費尚未審核."
                        Case Else
                            labMsg1.Text = cst_Msg_a7 & Convert.ToString(drv("AppliedResultM")) '"學員經費審核.其他" & Convert.ToString(drv("AppliedResultM"))
                    End Select

                End If

        End Select
    End Sub

End Class
