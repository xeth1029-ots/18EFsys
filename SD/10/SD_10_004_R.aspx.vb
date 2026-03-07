Partial Class SD_10_004_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "price"
    'price
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

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
        Table4.Visible = False
        msg.Text = ""
        Button1.Attributes("onclick") = "javascript:return search1();"

        '獎狀字號
        ProveNum.Text = TIMS.GetGlobalVar(Me, "12", "1", objconn)
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        Dim s_javascript_btn2 As String = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
        Button5.Attributes("onclick") = s_javascript_btn2
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Table4.Visible = True
        Dim sql As String = ""
        sql = ""
        sql &= " select b.SOCID,b.Rank,b.studentid,c.name,c.EngName"
        sql &= " ,b.OCID,b.StudStatus"
        sql &= " from class_classinfo a"
        sql &= " join class_studentsofclass b on a.ocid=b.ocid"
        sql &= " join stud_studentinfo c on b.sid=c.sid"
        sql &= " where 1=1"
        sql &= " and b.OCID='" & OCIDValue1.Value & "'"
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        Table4.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            Table4.Visible = True
            msg.Text = ""

            If Me.ViewState("sort") Is Nothing Then
                Me.ViewState("sort") = "StudentID"
            End If
            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr_stud, "SOCID", "StudentID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SOCID"
            PageControler1.Sort = "StudentID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const Cst_學員狀態 As Integer = 4
        'Dim objControl As HtmlInputCheckBox 'HtmlInputCheckBox'CheckBox

        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not Me.ViewState("sort") Is Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer
                    Select Case Me.ViewState("sort")
                        Case "StudentID", "StudentID desc"
                            i = 1
                        Case "Rank", "Rank desc"
                            i = 5
                    End Select

                    If Me.ViewState("sort").ToString.IndexOf("desc") = -1 Then
                        img.ImageUrl = "../../images/SortUp.gif"
                    Else
                        img.ImageUrl = "../../images/SortDown.gif"
                    End If
                    e.Item.Cells(i).Controls.Add(img)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim oCheckbox As HtmlInputCheckBox = e.Item.FindControl("Checkbox_select")
                oCheckbox.Value = Convert.ToString(drv("StudentID"))
                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                e.Item.Cells(Cst_學員狀態).Text = If(STUDSTATUS_N <> "", STUDSTATUS_N, "該學員狀態有誤，請重新查詢！")

        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If e.SortExpression = Me.ViewState("sort") Then
            Me.ViewState("sort") = e.SortExpression & " desc"
        Else
            Me.ViewState("sort") = e.SortExpression
        End If
        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim OCIDstr As String = ""
        Dim StudentIDstr As String = ""
        StudentIDstr = ""
        For Each myitem As DataGridItem In Me.DataGrid1.Items
            'Dim objCheckbox As CheckBox
            Dim objCheckbox As HtmlInputCheckBox 'HtmlInputCheckBox'CheckBox
            objCheckbox = myitem.FindControl("Checkbox_select")
            If objCheckbox.Checked Then
                OCIDstr = myitem.Cells(6).Text '只有1筆

                If StudentIDstr <> "" Then StudentIDstr &= ","
                StudentIDstr &= "\'" & myitem.Cells(7).Text & "\'"
                'StudentIDstr = StudentIDstr & Convert.ToString("\'" & myitem.Cells(7).Text & "\'" & ",")
            End If
        Next

        If StudentIDstr = "" Then
            Common.MessageBox(Me, "請選取要套印的學員!!!")
            Exit Sub
        End If
        '有學員資料
        Dim strScript As String = ""
        strScript = ""
        strScript &= "&OCID=" & OCIDstr
        strScript &= "&StudentID=" & StudentIDstr '多筆
        strScript &= "&distid=" & sm.UserInfo.DistID
        strScript &= "&ProveNum=" & ProveNum.Text
        strScript &= "&PrintDate=" & PrintDate.Text
        strScript &= "&rblYearType1=" & rblYearType1.SelectedValue '列印格式 1:西元年 2:民國年
        Select Case Convert.ToString(Me.ViewState("sort"))
            Case "StudentID"
                '依學號
                strScript &= "&Parameter1=1"
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, strScript)
            Case "Rank"
                '依名次
                strScript &= "&Parameter=1"
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, strScript)
            Case Else
                Common.MessageBox(Me, "資料異常，請重新查詢資料！")
        End Select

    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
