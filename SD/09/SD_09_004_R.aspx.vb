Partial Class SD_09_004_R
    Inherits AuthBasePage

    'roll_book (週) OCID,start_Date,end_Date
    'roll_book_1 (天)

    Const cst_printFN1 As String = "roll_book"
    Const cst_printFN2 As String = "roll_book_1"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        Button1.Attributes("onclick") = "return search();"
        'Button1.Attributes("onclick") = "if(search()){"
        'Button1.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "roll_book", "OCID='+document.getElementById('OCIDValue1').value+'&start_Date='+document.getElementById('start_date').value+'&end_Date='+document.getElementById('end_date').value+'")
        'Button1.Attributes("onclick") += "}return false;"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        If Not IsPostBack Then Button2_Click(sender, e)
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Select Case RadioButtonList1.SelectedValue
            Case "1" '週
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, "OCID='+document.getElementById('OCIDValue1').value+'&start_Date='+document.getElementById('start_date').value+'&end_Date='+document.getElementById('end_date').value+'")
                ReportQuery.PrintReport(Me, cst_printFN1, "OCID=" + OCIDValue1.Value + "&start_Date=" + start_date.Text.Trim + "&end_Date=" + end_date.Text.Trim + "")
            Case "2" '天
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, "OCID='+document.getElementById('OCIDValue1').value+'&start_Date='+document.getElementById('start_date').value+'&end_Date='+document.getElementById('end_date').value+'")
                ReportQuery.PrintReport(Me, cst_printFN2, "OCID=" + OCIDValue1.Value + "&start_Date=" + start_date.Text.Trim + "&end_Date=" + end_date.Text.Trim + "")
            Case Else
                Common.MessageBox(Me, "請選擇列印方式!")
        End Select
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class