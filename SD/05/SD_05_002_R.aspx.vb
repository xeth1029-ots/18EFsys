Partial Class SD_05_002_R
    Inherits AuthBasePage

    'SD_05_002_R_Rpt.aspx '出缺勤明細表
    'excuse_list '請假、缺曠課累計時數統計表
    'SD_05_002_R2 '請假、缺曠課累計時數統計表 2017

    'Const cst_printFN1 As String = "excuse_list"
    Const cst_printFN2 As String = "SD_05_002_R2"
    Const cst_prg_printFN1_aspx As String = "SD_05_002_R_Rpt.aspx"

#Region "Function1"
    Sub ChkSysGlobalVar(ByRef MyPage As Page)
        'Dim Rst As Boolean = True
        'Errmsg = ""
        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim item1 As Integer = 0
        Dim item2 As Integer = 0
        Dim item3 As Integer = 0
        Dim item4 As Integer = 0
        sql = "SELECT * FROM Sys_GlobalVar WHERE GVID=4 and DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            item1 = 0
            item2 = 1
            If Not IsDBNull(dr("ItemVar1")) Then
                If (dr("ItemVar1") = "0") Then
                    MyPage.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('請系統管理者至首頁>>系統管理>>系統參數管理>>參數設定,設定此計畫出缺勤警示!!!')</Script>")
                Else
                    item1 = Int(Split(dr("ItemVar1"), "/")(0))
                    If Int(Split(dr("ItemVar1"), "/")(1)) <> 0 Then
                        item2 = Int(Split(dr("ItemVar1"), "/")(1))
                    Else
                        item2 = 1
                    End If
                End If
            End If

            item3 = 0
            item4 = 1
            If Not IsDBNull(dr("ItemVar2")) Then
                If (dr("ItemVar2") = "0") Then
                    MyPage.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('請系統管理者至首頁>>系統管理>>系統參數管理>>參數設定,設定此計畫出缺勤警示!!!')</Script>")
                Else
                    item3 = Int(Split(dr("ItemVar2"), "/")(0))
                    If Int(Split(dr("ItemVar2"), "/")(1)) <> 0 Then
                        item4 = Int(Split(dr("ItemVar2"), "/")(1))
                    Else
                        item4 = 1
                    End If
                End If
            End If
            Me.Hiditem1.Value = CStr(item1)
            Me.Hiditem2.Value = CStr(item2)
            Me.Hiditem3.Value = CStr(item3)
            Me.Hiditem4.Value = CStr(item4)
        Else
            '查無資料。
            MyPage.RegisterStartupScript("Startup ", "<script language=""javascript"">alert('請系統管理者至首頁>>系統管理>>系統參數管理>>參數設定,設定此計畫出缺勤警示!!!')</Script>")
            Me.Hiditem1.Value = CStr(item1)
            Me.Hiditem2.Value = "1"
            Me.Hiditem3.Value = CStr(item3)
            Me.Hiditem4.Value = "1"
        End If
        'If Errmsg <> "" Then Rst = False
        'Return Rst
    End Sub

#End Region


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
        '檢查Session是否存在 End

        'Dim Errmsg As String = ""
        'Errmsg = ""
        Call ChkSysGlobalVar(Me)

        Button1.Attributes("onclick") = "javascript:return print();"

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

        'Button1.Attributes("onclick") =     ReportQuery.ReportScript(Me, "list", "fall_vacant_list", "OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'&start_date='+document.getElementById('start_date').value+'&end_date='+document.getElementById('end_date').value+'&RID='+document.getElementById('RIDValue').value+'&TPlanID=" & sm.UserInfo.TPlanID & "&item1=" & item1 & "&item2=" & item2 & "&item3=" & item3 & "&item4=" & item4)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Common.SetListItem(RadioButtonList1, "1")
            Common.SetListItem(RadioButtonList2, "1")

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        RadioButtonList1.Attributes.Add("onclick", "checkRBL1VAL();")
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        'If Trim(start_date.Text) <> "" Then start_date.Text = Trim(start_date.Text) Else start_date.Text = ""
        'If Trim(end_date.Text) <> "" Then end_date.Text = Trim(end_date.Text) Else end_date.Text = ""

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "時間區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "時間區間 起始日期 為必填" & vbCrLf
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "時間區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "時間區間 迄止日期 為必填" & vbCrLf
        End If

        If start_date.Text = "" AndAlso end_date.Text = "" Then
            Errmsg += "時間區間 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then
                    Errmsg += "【時間區間】的起日 不得大於 迄日!!" & vbCrLf
                End If
            End If
        End If

        '列印格式 1:出缺勤明細表 / 2:請假、缺曠課累計時數統計表
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)
        Select Case v_RadioButtonList1'.SelectedValue
            Case "1", "2"
            Case Else
                '列印格式
                Errmsg += "請選擇【列印格式】!!" & vbCrLf
        End Select

        'If Convert.ToString(OCID.SelectedValue) = "" _
        '    OrElse Not IsNumeric(OCID.SelectedValue) Then
        '    Errmsg += "統計對象 為必選" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("TMID")
        OCIDValue1.Value = dr("ocid")
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '列印格式 1:出缺勤明細表 / 2:請假、缺曠課累計時數統計表
        Dim v_RadioButtonList1 As String = TIMS.GetListValue(RadioButtonList1)

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            If v_RadioButtonList1 = "2" Then RadioButtonList2.Style.Add("display", "none")
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Const cst_def_iPrtPageSize1 As Integer = 42
        Const cst_def_iPrtPageSize2 As Integer = 26

        Dim querystr As String = ""
        Dim RID As String = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hiditem2.Value = TIMS.ClearSQM(Hiditem2.Value)
        Hiditem4.Value = TIMS.ClearSQM(Hiditem4.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        TMIDValue1.Value = TIMS.ClearSQM(TMIDValue1.Value)

        RID = If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value)
        If Hiditem2.Value = "0" OrElse Hiditem2.Value = "" Then Hiditem2.Value = "1"
        If Hiditem4.Value = "0" OrElse Hiditem4.Value = "" Then Hiditem4.Value = "1"

        '列印版型 1:直版 42 2:橫版 26
        Dim v_RadioButtonList2 As String = TIMS.GetListValue(RadioButtonList2)
        If v_RadioButtonList2 = "" Then v_RadioButtonList2 = "2"
        Dim i_prtPageSize As Integer = If(v_RadioButtonList2 = "2", cst_def_iPrtPageSize2, cst_def_iPrtPageSize1)
        'prtPageSize.Text = TIMS.ClearSQM(prtPageSize.Text)
        'If Not TIMS.IsNumeric2(prtPageSize.Text) Then prtPageSize.Text = cst_def_prtPageSize

        querystr = "xxx=zzz"
        If Me.cjobValue.Value <> "" Then
            '補查詢通俗職類
            querystr &= "&CJOB_UNKEY=" & Me.cjobValue.Value
            querystr &= "&PlanID=" & sm.UserInfo.PlanID
        End If

        querystr += "&OCID=" & Me.OCIDValue1.Value
        querystr += "&TMID=" & Me.TMIDValue1.Value
        querystr += "&TPlanID=" & sm.UserInfo.TPlanID
        querystr += "&RID=" & RID
        querystr += "&start_date=" & Me.start_date.Text
        querystr += "&end_date=" & Me.end_date.Text

        querystr += "&item1=" & Hiditem1.Value
        querystr += "&item2=" & Hiditem2.Value
        querystr += "&item3=" & Hiditem3.Value
        querystr += "&item4=" & Hiditem4.Value
        querystr += "&UserID=" & sm.UserInfo.UserID
        querystr += "&prtPageSize=" & i_prtPageSize '.Text

        Select Case v_RadioButtonList1'.SelectedValue
            Case "1"
                '出缺勤明細表
                Dim sUrl As String = cst_prg_printFN1_aspx
                ReportQuery.Redirect(Me, sUrl, querystr)
                'Dim s_width As String = "743"
                'Dim s_height As String = "1033"
                'ReportQuery.Redirect(Me, sUrl, querystr, s_width, s_height)
            Case "2"
                'Dim flagYear2017 As Boolean = False
                'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)
                'Dim sPrintName1 As String = cst_printFN1
                'If flagYear2017 Then sPrintName1 = cst_printFN2
                '請假、缺曠課累計時數統計表
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, querystr)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, querystr)
        End Select
    End Sub

    Protected Sub btnEnterDate_Click(sender As Object, e As EventArgs) Handles btnEnterDate.Click
        If Me.OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請先選擇班級。")
            Exit Sub
        End If
        Dim dr As DataRow = TIMS.GetOCIDDate(Me.OCIDValue1.Value, objconn)
        Me.start_date.Text = Common.FormatDate(dr("STDate"))
        Me.end_date.Text = Common.FormatDate(dr("FTDate"))

    End Sub
End Class
