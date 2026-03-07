Partial Class CP_04_003_01
    Inherits AuthBasePage

    '"window.open('CP_04_003_01.aspx?ID=" & Request("ID") & "&Student_Data=" & drv("OCID") & "','OCID','width=500,height=500'); return false;"
    'Button1.Attributes("onclick") = "window.open('CP_04_003_01_History.aspx?ID=" & Request("ID") & "&OCID=" & Request("Student_Data") & "','history','width=750,height=600,scrollbars=1')"
    '班級學員查詢
    Dim strBlockName As String = ""

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
        'PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            btnClose2.Attributes("onclick") = "window.close();"

            Dim vOCID As String = Convert.ToString(Request("Student_Data"))
            vOCID = TIMS.ClearSQM(vOCID)
            Button1.Attributes("onclick") = "window.open('CP_04_003_01_History.aspx?ID=" & Request("ID") & "&OCID=" & vOCID & "','history','width=750,height=600,scrollbars=1'); return false;"

            '查學生資料
            Call sUtl_Create1(vOCID)

        End If
    End Sub

    '查學生資料
    Sub sUtl_Create1(ByVal OCID As String)
        OCID = TIMS.ClearSQM(OCID)
        msg.Text = "查無學生資料!"
        Me.Table2.Visible = False
        If OCID = "" Then Exit Sub

        Dim parms As New Hashtable
        parms.Add("OCID", OCID)
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        Dim sql As String = ""
        sql &= " SELECT c.Years+'0'+ISNULL(f.ClassID2,f.ClassID)+ISNULL(c.CyclType,'01') FWStudentID" & vbCrLf
        sql &= " ,c.OCID ,a.SOCID" & vbCrLf
        sql &= " ,a.StudentID StdID" & vbCrLf
        sql &= " ,REPLACE(a.StudentID, c.Years + '0' + ISNULL(f.ClassID2,f.ClassID) + ISNULL(c.CyclType,'01'), '') StudentID " & vbCrLf
        sql &= " ,b.SID ,b.Name " & vbCrLf
        sql &= " ,(CASE WHEN b.Sex = 'M' THEN '男' WHEN b.Sex = 'F' THEN '女' ELSE ' ' END) Sex " & vbCrLf
        sql &= " ,format(b.Birthday ,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " ,b.IDNO ,d.Years " & vbCrLf
        sql &= " ,c.ClassCName ,e.PlanName " & vbCrLf
        sql &= " ,a.MODIFYDATE " & vbCrLf
        sql &= " FROM Class_StudentsOfClass a " & vbCrLf
        sql &= " JOIN Stud_StudentInfo b ON a.SID = b.SID " & vbCrLf
        sql &= " JOIN Class_ClassInfo c ON a.OCID = c.OCID " & vbCrLf
        sql &= " JOIN ID_Plan d ON c.PlanID = d.PlanID " & vbCrLf
        sql &= " JOIN Key_Plan e ON d.TPlanID = e.TPlanID " & vbCrLf
        sql &= " JOIN ID_CLass f ON c.CLSID = f.CLSID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.MAKESOCID IS NULL " & vbCrLf
        sql &= " AND a.OCID =@OCID" & vbCrLf
        sql &= " AND d.TPLANID=@TPLANID" & vbCrLf
        sql &= " ORDER BY a.StudentID " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無學生資料!"
        Table2.Visible = False
        If dt.Rows.Count = 0 Then Exit Sub

        'If dt.Rows.Count > 0 Then
        msg.Text = ""
        Table2.Visible = True

        'Dim strBlockName As String = "MESSAGE_" + New Random().GetHashCode().ToString("x")
        strBlockName = "MESSAGE_" + New Random().GetHashCode().ToString("x")

        dt.DefaultView.Sort = "StudentID"
        DataGrid1.AllowPaging = False
        DataGrid1.DataSource = dt.DefaultView
        DataGrid1.DataBind()

        'PageControler1.PageDataTable = dt
        'PageControler1.Sort = "StudentID"
        'PageControler1.ControlerLoad()
        Dim dr As DataRow = dt.Rows(0)
        Me.YearLabel.Text = Convert.ToString(dr("Years"))
        Me.TrainPlanLabel.Text = Convert.ToString(dr("PlanName"))
        Me.ClassNameLabel.Text = Convert.ToString(dr("ClassCName"))
    End Sub

    'Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
    '    Select Case e.CommandName
    '        Case "list"
    '            Dim sCmdArg As String = e.CommandArgument '
    '            Dim MyValue As String = ""
    '            MyValue = TIMS.GetMyValue(sCmdArg, "Student_History")
    '            MyValue = TIMS.ClearSQM(MyValue)
    '            MyValue = TIMS.ChangeIDNO(MyValue)
    '            Session("Student_History") = MyValue 'IDNO
    '            Dim strBlockName As String = "MESSAGE_" + New Random().GetHashCode().ToString("x")
    '            Dim strScript As String = ""
    '            strScript = "<script language=""javascript"">"
    '            strScript += "window.open('CP_04_003_01_01.aspx?ID=" & Request("ID") & sCmdArg & "');"
    '            strScript += "</script>"
    '            TIMS.RegisterStartupScript(Me, strBlockName, strScript)
    '    End Select
    'End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '學員參訓歷史
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim Button3 As Button = e.Item.FindControl("Button3")
                Dim sCmdArg As String = ""
                sCmdArg &= "&Student_History=" & Convert.ToString(drv("IDNO")) 'IDNO
                Button3.CommandArgument = sCmdArg

                'Dim strBlockName As String = "MESSAGE_" + New Random().GetHashCode().ToString("x")
                If Convert.ToString(drv("IDNO")) <> "" Then
                    Button3.Attributes("onclick") = "window.open('CP_04_003_01_01.aspx?ID=" & Request("ID") & sCmdArg & "','" & strBlockName & "','width=750,height=600,scrollbars=1'); return false;"
                End If

        End Select
    End Sub
End Class