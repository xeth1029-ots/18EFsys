Partial Class SD_04_012_R
    Inherits AuthBasePage

    'SD_04_012_R /MV_CLASS_SCHEDULE2
    'Dim objconn As SqlConnection
    'Dim sql As String
    'Dim dr As DataRow

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            Printtype.Attributes("onclick") = "ChangeMode();"
            Button1.Attributes("onclick") = "javascript:return ReportPrint();"

            Years = TIMS.GetSyear(Years, sm.UserInfo.Years - 5, sm.UserInfo.Years + 2, True)
            Common.SetListItem(Years, Now.Year)
            For i As Integer = 1 To 12
                Months.Items.Add(New ListItem(i & "月份", i))
            Next
            Months.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Common.SetListItem(Months, Now.Month)

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Allday_TR.Style("display") = "none" '全期的隱藏
            Button3_Click(sender, e)
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

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim RID, Year As String
        'Dim STdate, FTdate As String
        Dim RID As String = ""
        Dim Year As String = ""
        Dim STdate As String = ""
        Dim FTdate As String = ""
        Dim start_month As String = ""
        Dim end_month As String = ""

        If Years.SelectedValue <> "" And Months.SelectedValue <> "" Then
            start_month = Years.SelectedValue & "/" & Months.SelectedValue & "/1"
            end_month = Common.FormatDate(DateAdd(DateInterval.Month, 1, CDate(start_month)))
            Year = (Years.SelectedValue - 1911)
        End If

        start_month = TIMS.Cdate3(start_month)
        end_month = TIMS.Cdate3(end_month)
        If s_date.Text <> "" And e_date.Text <> "" Then
            STdate = TIMS.Cdate3(s_date.Text)
            FTdate = TIMS.Cdate3(e_date.Text)
        End If

        If Me.RIDValue.Value = "" Then
            RID = sm.UserInfo.RID
        Else
            RID = Me.RIDValue.Value
        End If

        'http://163.29.199.77:8080/ReportServer3/report.do?RptID=SD_04_012_R&RID=A&TPlanID=06&SDate=2018/11/1&EDate=2018/12/01&Years=107&Months=11&OCID=&STdate=&FTdate=&CJOB_UNKEY=&UserID=F223624622
        Dim myValue As String = ""
        myValue = "RID=" & RID
        myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        myValue &= "&Years=" & Year
        myValue &= "&Months=" & Months.SelectedValue
        myValue &= "&OCID=" & OCIDValue1.Value
        myValue &= "&SDate=" & start_month
        myValue &= "&EDate=" & end_month

        myValue &= "&STdate=" & STdate
        myValue &= "&FTdate=" & FTdate
        myValue &= "&CJOB_UNKEY=" & cjobValue.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_04_012_R", myValue)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Printtype_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Printtype.SelectedIndexChanged

        Select Case Printtype.SelectedIndex
            Case 0
                'Months_TR.Style("display") = "inline"
                '上面是原寫法
                Months_TR.Style("display") = ""
                Allday_TR.Style("display") = "none"
                s_date.Text = ""
                e_date.Text = ""

            Case 1
                'If OCIDValue1.Value = "" Then
                '    Common.MessageBox(Me.Page, "尚未選擇職類/班別")
                '    Printtype.SelectedValue = "0"
                'Else
                'End If
                Dim SQL As String
                Dim dr As DataRow = Nothing

                If OCIDValue1.Value <> "" Then
                    SQL = "Select STDate,FTDate From class_classinfo where OCID = '" & OCIDValue1.Value & "' "
                    dr = DbAccess.GetOneRow(SQL, objconn)

                    If dr IsNot Nothing Then
                        s_date.Text = dr("STDate")
                        e_date.Text = dr("FTDate")
                    End If
                End If

                Months_TR.Style("display") = "none"
                'Allday_TR.Style("display") = "inline"
                '上面是原寫法
                Allday_TR.Style("display") = ""
                Years.SelectedValue = ""
                Months.SelectedValue = ""
        End Select

    End Sub
End Class

