Partial Class SD_02_007_R
    Inherits AuthBasePage

    'Check_in_list
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            SelResult = TIMS.Get_SelResult(SelResult, 0, objconn)
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button13_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg.aspx?name=Stu_Maintain');"
        Const cst_javascript_openOrg_FMT2 As String = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        Button8.Attributes("onclick") = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, cst_javascript_openOrg_FMT1, cst_javascript_openOrg_FMT2)

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
        '檢查班級是以成績或者以報名順序來排序
        Dim msg As String = ""
        If OCIDValue1.Value = "" Then msg += "請選擇職類/班別!" & vbCrLf
        If SelResult.SelectedIndex = -1 Then msg += "請選擇錄取總類!" & vbCrLf
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        If (OCIDValue1.Value <> "") And (SelResult.SelectedIndex <> -1) Then
            Dim sqlstr As String = ""
            sqlstr = "SELECT SUM(SumOfGrad) total FROM Stud_SelResult WHERE OCID='" & OCIDValue1.Value & "' "
            Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
            Dim TotalGrade As Integer = 0

            'Microsoft.VisualBasic.Information.IsNothing()
            TotalGrade = 0
            If Not IsDBNull(dr(0)) Then TotalGrade = dr(0)

            Dim Order As String = ""
            If TotalGrade = 0 Then
                Order = "Order By h.SelResultID,b.NotExam DESC,b.TRNDMode Desc,b.TRNDType,b.RelEnterDate,b.ExamNo"
            Else
                Order = "Order By h.SelResultID,b.NotExam DESC,b.TRNDMode Desc,b.TRNDType,b.TotalResult DESC,b.ExamNo"
            End If

            Dim v_SelResult As String = TIMS.GetListValue(SelResult) '.SelectedValue
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "Check_in_list", "OCID=" & OCIDValue1.Value & "&SelResultID=" & v_SelResult & "&Order=" & Order)
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '判斷機構是否只有一個班級  不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub
End Class