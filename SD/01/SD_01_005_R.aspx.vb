Partial Class SD_01_005_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        'Dim dr As DataRow
        If Not IsPostBack Then
            msg.Text = ""
            Button2.Attributes("onclick") = "javascript:return search()"

            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button13_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OCIDValue1.Value <> "" Then OCIDValue1.Value = Trim(OCIDValue1.Value)
        If IDNO.Text <> "" Then IDNO.Text = Trim(IDNO.Text)
        If OCIDValue1.Value = "" AndAlso IDNO.Text = "" Then
            Common.MessageBox(Me, "請擇1選擇 職類/班別 或 身分證號碼!!")
            Exit Sub
        End If

        Dim MyValue As String = "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&OCID=" & OCIDValue1.Value & "&IDNO=" & IDNO.Text
        ReportQuery.PrintReport(Me, "Member", "SD_01_005_R", MyValue)

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse Convert.ToString(dr("total")) <> "1" Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

End Class
