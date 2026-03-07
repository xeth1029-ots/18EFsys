'Imports System.Data.SqlClient
'Imports System.Data
'Imports Turbo
Partial Class SD_11_002
    Inherits AuthBasePage

    Dim sqlAdapter As SqlDataAdapter
    Dim stud_table As DataTable
    Dim ProcessType As String
    'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        msg.Text = ""
        search.Attributes("onclick") = "javascript:return search1()"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Dim re_ocid As String ', ClassName 
        re_ocid = Request("ocid")
        ProcessType = Request("ProcessType")

#Region "(No Use)"

        '分頁設定 Start
        'DataGridPage1.MyDataGrid = DG_stud
        '分頁設定 End

        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt

        '        If Not Request("ID") Is Nothing Then

        '            Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '            If FunDrArray.Length = 0 Then
        '                Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '                Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '            Else
        '                FunDr = FunDrArray(0)
        '                If FunDr("Adds") = "1" Then
        '                    check_add.Value = "1"
        '                Else
        '                    check_add.Value = "0"
        '                End If
        '                If FunDr("Sech") = "1" Then
        '                    search.Enabled = True
        '                    check_search.Value = "1"
        '                Else
        '                    search.Enabled = False
        '                    check_search.Value = "0"
        '                End If
        '                If FunDr("Del") = "1" Then
        '                    check_del.Value = "1"
        '                Else
        '                    check_del.Value = "0"
        '                End If
        '                If FunDr("Mod") = "1" Then
        '                    check_mod.Value = "1"
        '                Else
        '                    check_mod.Value = "0"
        '                End If
        '            End If
        '        End If
        '    End If
        'End If

#End Region

        If Not Page.IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            hidSearchTag.Value = ""
            StudentTable.Style.Item("display") = "none"

            If ProcessType = "Back" Then
                If Session("QuestionarySearchStr2") IsNot Nothing Then
                    Dim s_search1 As String = Session("QuestionarySearchStr2")
                    Dim MyValue As String = ""
                    Session("QuestionarySearchStr2") = Nothing
                    center.Text = TIMS.GetMyValue(s_search1, "center")
                    RIDValue.Value = TIMS.GetMyValue(s_search1, "RIDValue")
                    TMID1.Text = TIMS.GetMyValue(s_search1, "TMID1")
                    OCID1.Text = TIMS.GetMyValue(s_search1, "OCID1")
                    TMIDValue1.Value = TIMS.GetMyValue(s_search1, "TMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(s_search1, "OCIDValue1")
                    'MyValue = TIMS.GetMyValue(s_search1, "StudentTable")
                    MyValue = TIMS.GetMyValue(s_search1, "Button1")
                    If MyValue = "True" Then search_Click(sender, e)
                    'Session("QuestionarySearchStr2") = Nothing
                End If
            End If
        End If

        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        If center.Text = sm.UserInfo.OrgName And hidSearchTag.Value = "" Then
            Button7_Click(sender, e)
        Else
            hidSearchTag.Value = ""
        End If
    End Sub

    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Me.Panel1.Visible = True
        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False

        Dim sqlstr_stud As String = ""
        sqlstr_stud = ""
        sqlstr_stud += " SELECT a.OCID, a.CyclType, a.LevelType, a.ClassCName, a.FTDate, b.total, ISNULL(c.num,0) AS num FROM class_classinfo a "
        sqlstr_stud += " JOIN (SELECT OCID, COUNT(1) AS total FROM class_studentsofclass WHERE StudStatus IN (1,4,5) GROUP BY OCID ) b ON a.ocid = b.ocid "
        sqlstr_stud += " LEFT JOIN (SELECT OCID, COUNT(1) AS num FROM Stud_QuestionEpt GROUP BY OCID) c ON c.ocid = a.ocid "
        If sm.UserInfo.RID = "A" Then
            sqlstr_stud += " WHERE PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID = '" & sm.UserInfo.TPlanID & "' AND Years = '" & sm.UserInfo.Years & "') "
        Else
            sqlstr_stud += " WHERE PlanID = '" & sm.UserInfo.PlanID & "' "
        End If

        If RIDValue.Value <> "" Then sqlstr_stud += " AND a.RID = '" & RIDValue.Value & "' "
        If OCIDValue1.Value <> "" Then
            sqlstr_stud += " AND a.OCID = '" & OCIDValue1.Value & "' "
        End If
        sqlstr_stud += " ORDER BY a.OCID "
        stud_table = DbAccess.GetDataTable(sqlstr_stud, objconn)

        If stud_table.Rows.Count = 0 Then
            DataGrid1.Style.Item("display") = "none"
            msg.Text = "查無資料!!"
        Else
            msg.Visible = False
            DataGrid1.Style.Item("display") = "" '"inline"
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.Visible = True
            DataGrid1.DataSource = stud_table
            DataGrid1.DataBind()
            ''分頁用-   Start
            'DataGridPage1.MyDataTable = stud_table
            'DataGridPage1.FirstTime()
            ''分頁用-   End
        End If
    End Sub

    Sub GetSearchStr2()
        Dim s_search1 As String = ""
        s_search1 += "center=" & center.Text
        s_search1 += "&RIDValue=" & RIDValue.Value
        's_search1= "TMID1=" & TMID1.Text
        s_search1 += "&TMID1=" & TMID1.Text
        s_search1 += "&OCID1=" & OCID1.Text
        s_search1 += "&TMIDValue1=" & TMIDValue1.Value
        s_search1 += "&OCIDValue1=" & OCIDValue1.Value
        s_search1 += "&Button1=" & DG_stud.Visible
        's_search1 += "&StudentTable=" & StudentTable.Style.Item("display")
        Session("QuestionarySearchStr2") = s_search1
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            Dim drv As DataRowView = e.Item.DataItem

            e.Item.Cells(1).Text = TIMS.GET_CLASSNAME(Convert.ToString(drv("ClassCName")), Convert.ToString(drv("CyclType")))

            Dim mybut1 As Button = e.Item.Cells(4).FindControl("Button1")

            mybut1.CommandArgument = DataGrid1.DataKeys(e.Item.ItemIndex)
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "view" Then
            Dim sqlstr_stud As String
            sqlstr_stud = ""
            sqlstr_stud &= " SELECT a.*, ISNULL(b.StudentCount,0) AS StudentCount, ISNULL(c.TrainCount,0) AS TrainCount, ISNULL(d.LeaveCount,0) AS LeaveCount "
            sqlstr_stud += " FROM (SELECT * FROM Class_ClassInfo WHERE OCID = '" & e.CommandArgument & "') a "
            sqlstr_stud += " LEFT JOIN (SELECT OCID, COUNT(1) StudentCount FROM Class_StudentsOfClass WHERE OCID = '" & e.CommandArgument & "' GROUP BY OCID) b ON a.OCID = b.OCID "
            sqlstr_stud += " LEFT JOIN (SELECT OCID, COUNT(1) TrainCount FROM Class_StudentsOfClass WHERE OCID = '" & e.CommandArgument & "' AND StudStatus IN (1,4,5) GROUP BY OCID) c ON a.OCID = c.OCID "
            sqlstr_stud += " LEFT JOIN (SELECT OCID, COUNT(1) LeaveCount FROM Class_StudentsOfClass WHERE OCID = '" & e.CommandArgument & "' AND StudStatus IN (2,3) GROUP BY OCID) d ON a.OCID = d.OCID "
            Dim dr As DataRow = DbAccess.GetOneRow(sqlstr_stud, objconn)

            If dr IsNot Nothing Then
                Label1.Text = "班別：" & TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
                If Not IsDBNull(dr("LevelType")) Then
                    If CInt(dr("LevelType")) <> 0 Then Label1.Text += "第" & TIMS.GetChtNum(CInt(dr("LevelType"))) & "階段"
                End If
                Label1.Text += "(開訓人數:" & dr("StudentCount").ToString & "&nbsp;&nbsp;在結訓人數:" & dr("TrainCount").ToString & "&nbsp;&nbsp;離退訓人數:" & dr("LeaveCount").ToString & ")"
            End If

            sqlstr_stud = " SELECT b.studentid, b.StudStatus, c.name, b.OCID, b.RejectTDate1, b.RejectTDate2 "
            sqlstr_stud += " FROM class_classinfo a "
            sqlstr_stud += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
            sqlstr_stud += " JOIN stud_studentinfo c ON b.sid = c.sid "
            sqlstr_stud += " WHERE a.ocid = '" & e.CommandArgument & "' "
            sqlstr_stud += " ORDER BY b.studentid "
            Dim dt As DataTable = DbAccess.GetDataTable(sqlstr_stud, objconn)

            If dt.Rows.Count = 0 Then
                msg2.Text = "查無此班學生資料!"
                StudentTable.Style.Item("display") = "none"
                Label1.Visible = False
            Else
                Label1.Visible = True
                StudentTable.Style.Item("display") = "" '"inline"
                'Session("DTable_Stuednt2") = dt
                'DG_stud.DataKeyField = "SOCID"
                DG_stud.DataSource = dt
                DG_stud.DataBind()
            End If
        End If
    End Sub

    Private Sub DG_stud_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Dim drv As DataRowView
        drv = e.Item.DataItem
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
            Dim but5 As Button
            Dim but4, but6 As Button
            but4 = e.Item.Cells(3).FindControl("Button4") '新增
            but5 = e.Item.Cells(3).FindControl("Button5") '查看
            but6 = e.Item.Cells(3).FindControl("Button6") '清除重填
            If drv("RejectTDate1").ToString <> "" Then e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate1"), 2) & ")"
            If drv("RejectTDate2").ToString <> "" Then e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate2"), 2) & ")"
            Dim sqlstr_list As String = "select * from Stud_QuestionEpt where OCID='" & e.Item.Cells(4).Text & "' and StudID= '" & e.Item.Cells(5).Text & "'"
            If DbAccess.GetCount(sqlstr_list) > 0 Then '已有資料
                e.Item.Cells(2).Text = "是"
                but4.Enabled = False '不可新增
                but5.Enabled = True '可以查看
                but6.Enabled = True '可清除重填
#Region "(因頁面權限目前被拿掉的關係,導致按鈕無法使用,所以暫時先將Enabled的判斷拿掉，by:20180913)"

                'If check_mod.Value = "0" And check_del.Value = "0" Then '兩者功能皆沒有時,不能使用
                '    but6.Enabled = False
                'Else
                '    but6.Enabled = True
                'End If
                'If check_search.Value = "1" Then
                '    but5.Enabled = True
                'Else
                '    but5.Enabled = False
                'End If

#End Region
            Else
                e.Item.Cells(2).Text = "否"
                but4.Enabled = True '可新增
                but5.Enabled = False '不可查看
                but6.Enabled = False '不可清除重填
#Region "(因頁面權限目前被拿掉的關係,導致按鈕無法使用,所以暫時先將Enabled的判斷拿掉，by:20180913)"

                'If check_add.Value = "1" Then
                '    but4.Enabled = True
                'Else
                '    but4.Enabled = False
                'End If

#End Region
            End If
        End If
    End Sub

    Private Sub DG_stud_ItemCommand1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        If e.CommandName = "insert" Then
            GetSearchStr2()
            TIMS.Utl_Redirect1(Me, "SD_11_002_add.aspx?ProcessType=Insert&Stuedntid=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID"))
        ElseIf e.CommandName = "clear" Then
            GetSearchStr2()
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "if (window.confirm('此動作會刪除訓練成效追蹤調查表資料，是否確定刪除?')){" + vbCrLf
            strScript += "location.href ='SD_11_002_add.aspx?ProcessType=del&Stuedntid=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID") & "';}" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
        Else
            GetSearchStr2()
            TIMS.Utl_Redirect1(Me, "SD_11_002_add.aspx?ProcessType=check&Stuedntid=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID"))
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        DataGrid1.Style.Item("display") = "none"
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        GetSearchStr2()
        'TIMS.Utl_Redirect1(Me, "SD_11_002_add.aspx?ProcessType=Print&ID=" & Request("ID"))

        '========== (因頁面改為框架式的關係,無法使用轉頁方式,所以調整成開新視窗的方式處理，by:20180913)
        Dim transferPage As String = "SD_11_002_add.aspx?ProcessType=Print&ID=" & Request("ID")
        Dim strScript As String = "<script language=""javascript"">window.open('" + transferPage + "', '_blank', 'height=900,width=1000,toolbar=1');</script>"
        Page.RegisterStartupScript("", strScript)
        '================================================= End
    End Sub
End Class