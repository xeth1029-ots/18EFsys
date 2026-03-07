Partial Class SD_10_001_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "student_inclass" '套印在訓證明
    'Const cst_printFN2 As String = "stdinclass_16"
    'student_inclass
    '2016 針對特定計畫:職前訓練(02','14','17','20','21','26','34','37','47','53','55','58','59','61','62','64','65')共計17支計畫，修改在訓證明、受訓證明、結訓證明等3張表件
    'stdinclass_16
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        'student_inclass
        If Not IsPostBack Then
            CCreate1()
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", , "search")
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
        search.Attributes("onclick") = "javascript:return search1()"

        '在訓'證明字號
        ProveNum.Text = TIMS.GetGlobalVar(Me, "5", "1", objconn)
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        Dim s_javascript_btn2 As String = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button5.Attributes("onclick") = s_javascript_btn2
    End Sub
    Private Sub Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim parms As New Hashtable From {{"OCIDV1", OCIDValue1.Value}}
        Dim sql As String = ""
        sql &= " SELECT b.StudentID,c.name,c.EngName ,b.OCID ,b.StudStatus" & vbCrLf
        sql &= " FROM class_classinfo a" & vbCrLf
        sql &= " join class_studentsofclass b on a.ocid=b.ocid " & vbCrLf
        sql &= " join stud_studentinfo c on b.sid=c.sid " & vbCrLf
        sql &= " WHERE b.STUDSTATUS NOT IN (2,3)" & vbCrLf
        sql &= " AND b.STUDSTATUS!=5" & vbCrLf
        sql &= " AND b.OCID=@OCIDV1 " & vbCrLf
        sql &= " ORDER BY b.StudentID" & vbCrLf
        Dim stud_table As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Panel1.Visible = True

        If TIMS.dtNODATA(stud_table) Then
            DG_stud.Visible = False
            msg.Text = "查無資料!!"
            submit.Visible = False
            Exit Sub
        End If

        DG_stud.Visible = True
        msg.Text = ""
        submit.Visible = True

        DG_stud.DataSource = stud_table
        DG_stud.DataBind()
    End Sub

    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Const Cst_學員狀態 As Integer = 4

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim oCheckbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                oCheckbox1.Value = drv("StudentID").ToString
                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                e.Item.Cells(Cst_學員狀態).Text = STUDSTATUS_N '"在訓"
        End Select
    End Sub

    Private Sub submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles submit.Click
        'OCIDValue1.Value
        Dim StudentIDstr As String = ""
        Dim objDG As DataGrid = DG_stud
        Const cst_StudentID As Integer = 1
        For Each myitem As DataGridItem In objDG.Items
            'Dim objCheckbox As CheckBox
            Dim objCheckbox As HtmlInputCheckBox = myitem.FindControl("Checkbox1")
            If objCheckbox.Checked Then
                'OCIDstr = myitem.Cells(6).Text '只有1筆
                If StudentIDstr <> "" Then StudentIDstr += ","
                StudentIDstr += "\'" & myitem.Cells(cst_StudentID).Text & "\'"
            End If
        Next
        If StudentIDstr = "" Then
            Common.MessageBox(Me, "請選取要套印的學員!!!")
            Exit Sub
        End If

        'student_inclass
        Dim RptFN1 As String = cst_printFN1
        'Dim w16R1 As String = TIMS.Utl_GetConfigSet("w16R1")
        'If w16R1 = "Y" Then RptFN1 = cst_printFN2
        Dim v_rblYearType1 As String = TIMS.GetListValue(rblYearType1) '列印格式 1:西元年 2:民國年
        Dim v_RTE As String = If(v_rblYearType1 = "1", "E", "C") '列印格式 E:西元年 C:民國年 #{RTE} RTE
        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= String.Concat("&DistID=", sm.UserInfo.DistID)
        MyValue &= String.Concat("&OCID=", OCIDValue1.Value)
        MyValue &= String.Concat("&StudentID=", StudentIDstr)
        MyValue &= String.Concat("&ProveNum=", ProveNum.Text)
        MyValue &= String.Concat("&rblYearType1=", v_rblYearType1) '列印格式 1:西元年 2:民國年
        MyValue &= String.Concat("&RTE=", v_RTE) '列印格式 E:西元年 C:民國年 #{RTE} RTE
        MyValue &= "&Type=2" '$P{Type}=="1"?$F{PN}+" 補發":$F{PN}
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, RptFN1, MyValue)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
