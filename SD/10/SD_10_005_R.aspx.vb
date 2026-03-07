Partial Class SD_10_005_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_10_005_R"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            CCreate1()
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        tr_center.Visible = (sm.UserInfo.LID = 0)
    End Sub

    Sub CCreate1()
        msg.Text = ""
        Button1.Attributes("onclick") = "javascript:return search1();"
        Table4.Visible = False
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        Dim s_javascript_btn2 As String = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button5.Attributes("onclick") = s_javascript_btn2
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Table4.Visible = True

        Dim sqlstr_stud As String = "select b.SOCID,b.Rank,b.studentid,c.name,c.EngName,b.OCID,b.StudStatus  from class_classinfo a join class_studentsofclass b on a.ocid=b.ocid join stud_studentinfo c on b.sid=c.sid where c.PassPortNO=2 and b.OCID='" & OCIDValue1.Value & "'"
        Dim stud_table As DataTable = DbAccess.GetDataTable(sqlstr_stud, objconn)
        If stud_table.Rows.Count = 0 Then
            DG_stud.Visible = False
            msg.Text = "查無資料!!"
        Else
            msg.Visible = False
            Me.submit.Visible = True
            DG_stud.Visible = True
            DG_stud.DataSource = stud_table
            DG_stud.DataBind()
        End If
    End Sub

    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Const Cst_學員狀態 As Integer = 4

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim oCheckbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim Hid_StudentID As HiddenField = e.Item.FindControl("Hid_StudentID")
                oCheckbox1.Value = drv("StudentID").ToString
                Hid_StudentID.Value = Convert.ToString(drv("StudentID"))
                e.Item.Cells(Cst_學員狀態).Text = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
        End Select
    End Sub

    Protected Sub Submit_Click(sender As Object, e As EventArgs) Handles submit.Click
        'submit.Attributes("onclick") = "return ReportPrint('SQ_AutoLogout=true&path=TIMS&sys=Member&filename=SD_10_005_R
        '&DistID=" & sm.UserInfo.DistID & "');"
        'window.open('../../SQControl.aspx?' + url + '&StudentID=' + StudentID + '&OCID=' + document.getElementById('OCIDValue1').value, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1')
        'Dim OCIDstr As String = ""
        Dim StudentIDstr As String = ""
        'StudentIDstr = ""
        For Each eItem As DataGridItem In Me.DG_stud.Items
            Dim oCheckbox1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            Dim Hid_StudentID As HiddenField = eItem.FindControl("Hid_StudentID")
            If oCheckbox1.Checked Then
                'OCIDstr = myitem.Cells(6).Text '只有1筆

                If StudentIDstr <> "" Then StudentIDstr &= ","
                StudentIDstr &= "\'" & Hid_StudentID.Value & "\'"
                'StudentIDstr = StudentIDstr & Convert.ToString("\'" & myitem.Cells(7).Text & "\'" & ",")
            End If
        Next

        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇職類班別!!!")
            Exit Sub
        End If
        If StudentIDstr = "" Then
            Common.MessageBox(Me, "請選取要套印的學員!!!")
            Exit Sub
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&DistID=" & sm.UserInfo.DistID
        MyValue &= "&StudentID=" & StudentIDstr
        MyValue &= "&OCID=" & OCIDValue1.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
