Partial Class SD_09_010_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_09_010_R"

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

        Button1.Attributes("onclick") = "javascript:return search1();"
        submit.Attributes("onclick") = "return ReportPrint('SQ_AutoLogout=true&path=TIMS&sys=Member&filename=" & cst_printFN1 & "&DistID=" & sm.UserInfo.DistID & "');"

        If Not IsPostBack Then
            msg.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Button3_Click(sender, e)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Table4.Visible = True

        Dim dt As DataTable
        Dim SqlStr As String = ""
        SqlStr = "" & vbCrLf
        SqlStr += " select b.SOCID,b.Rank,b.studentid" & vbCrLf
        SqlStr += " ,c.name,c.EngName,c.idno" & vbCrLf
        SqlStr += " ,b.OCID,b.StudStatus  " & vbCrLf
        SqlStr += " from class_classinfo a " & vbCrLf
        SqlStr += " join class_studentsofclass b on a.ocid=b.ocid " & vbCrLf
        SqlStr += " join stud_studentinfo c on b.sid=c.sid " & vbCrLf
        SqlStr += " where c.PassPortNO=2 " & vbCrLf
        SqlStr += " and b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        dt = DbAccess.GetDataTable(SqlStr, objconn)

        DG_stud.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DG_stud.Visible = True
            msg.Text = ""

            DG_stud.DataSource = dt
            DG_stud.DataBind()
        End If

    End Sub

    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim objCheckbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                objCheckbox1.Value = Convert.ToString(drv("idno"))
                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                e.Item.Cells(4).Text = STUDSTATUS_N '"在訓"
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
