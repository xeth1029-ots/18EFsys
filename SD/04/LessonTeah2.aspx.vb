Partial Class LessonTeah2
    Inherits AuthBasePage

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Me.modifytype.Value = Me.Request("type")
        Me.fieldname.Value = Me.Request("fieldname")
        Me.hiddenname.Value = Me.Request("hiddenname")

        If Not Page.IsPostBack Then
            Call Search_Query("")
        End If
        txtSearch1.Attributes("onkeypress") = "Search_click();return false;"

        Me.btnSearch.Attributes("onclick") = "Search_click2();return false;"
        Me.btnSearch.Attributes("onkeypress") = "Search_click2();return false;"
    End Sub

    Sub Search_Query(ByVal TeachCNameVal As String)
        Dim rqRID As String = TIMS.ClearSQM(Request("RID"))
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select tt.TechID,tt.TeachCName,tt.KindEngage,te.TechID TechIDEmp " & vbCrLf
        sql += " from Teach_TeacherInfo tt   " & vbCrLf
        sql += " left join (" & vbCrLf
        sql += " 	select DISTINCT TechID from Teacher_Employs   " & vbCrLf
        sql += " 	where TEDate2 >= getdate()) te on tt.TechID = te.TechID " & vbCrLf
        sql += " where tt.WorkStatus = '1'" & vbCrLf
        'If sqlTechType <> "" Then '產投任課教師 或 '產投助教
        '    Sql += sqlTechType & vbCrLf
        'End If
        If rqRID = "" Then 'RID (業務權限)
            sql += "and tt.RID = '" & sm.UserInfo.RID & "'" & vbCrLf
        Else
            sql += "and tt.RID = '" & rqRID & "'" & vbCrLf
        End If
        ''排除1名
        'If Me.ExistTech.Value <> "" Then
        '    Sql += " and  tt.TechID <> " & Me.ExistTech.Value & vbCrLf
        'End If
        '搜尋名稱
        TeachCNameVal = TIMS.ClearSQM(TeachCNameVal)
        If TeachCNameVal <> "" Then sql &= String.Format(" and tt.TeachCName like N'%{0}%'", TeachCNameVal) & vbCrLf
        'sql += " ORDER BY tt.TeachCName" & vbCrLf

        Dim objtable As DataTable = DbAccess.GetDataTable(sql, objconn)
        '有在外聘師資管理功能建資料且聘約迄日尚未到期的老師
        'objtable2 = DbAccess.GetDataTable(objstr2, objconn)

        Dim mydv1 As New DataView(objtable)
        Dim mydv2 As New DataView(objtable)
        Dim mydv3 As New DataView(objtable)
        'KindEngage:(1.內;2.外)
        mydv1.RowFilter = "KindEngage = '1'" '1.內
        '有在外聘師資管理功能建資料且聘約迄日尚未到期的老師
        mydv2.RowFilter = "KindEngage = '2' and TechIDEmp is NOT NULL" '1.內 (有簽約)'有在外聘師資管理功能建資料且聘約迄日尚未到期的老師
        mydv3.RowFilter = "KindEngage = '2' and TechIDEmp is NULL" '1.內(沒簽約)

        mydv1.Sort = "TeachCName,TechID"
        mydv2.Sort = "TeachCName,TechID"
        mydv3.Sort = "TeachCName,TechID"

        Me.LessonTeahList1.DataSource = mydv1
        Me.LessonTeahList1.DataTextField = "TeachCName"
        Me.LessonTeahList1.DataValueField = "TechID"
        Me.LessonTeahList1.DataBind()

        Me.LessonTeahList2.DataSource = mydv2
        Me.LessonTeahList2.DataTextField = "TeachCName"
        Me.LessonTeahList2.DataValueField = "TechID"
        Me.LessonTeahList2.DataBind()

        Me.LessonTeahList3.DataSource = mydv3
        Me.LessonTeahList3.DataTextField = "TeachCName"
        Me.LessonTeahList3.DataValueField = "TechID"
        Me.LessonTeahList3.DataBind()
    End Sub

    Private Sub LessonTeahList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList1.SelectedIndexChanged
        Dim strScript As String

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('" & Me.LessonTeahList1.SelectedValue & "','" & Me.LessonTeahList1.Items(Me.LessonTeahList1.SelectedIndex).Text & "');" + vbCrLf
        strScript += "</script>"

        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub LessonTeahList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList2.SelectedIndexChanged
        Dim strScript As String

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('" & Me.LessonTeahList2.SelectedValue & "','" & Me.LessonTeahList2.Items(Me.LessonTeahList2.SelectedIndex).Text & "');" + vbCrLf
        strScript += "</script>"

        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub LessonTeahList3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LessonTeahList3.SelectedIndexChanged
        Dim strScript As String

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('" & Me.LessonTeahList3.SelectedValue & "','" & Me.LessonTeahList3.Items(Me.LessonTeahList3.SelectedIndex).Text & "');" + vbCrLf
        strScript += "</script>"

        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        Dim strScript As String

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('','');" + vbCrLf
        strScript += "</script>"

        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub hbtnSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hbtnSearch.ServerClick
        Dim InputVal As String = ""
        'Me.txtSearch1.Text = Trim(Me.txtSearch1.Text)
        Me.txtSearch1.Text = TIMS.ClearSQM(Me.txtSearch1.Text)
        InputVal = Me.txtSearch1.Text
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        End If

        InputVal = TIMS.ClearSQM(InputVal)
        If InputVal <> "" Then
            Search_Query(InputVal)
        Else
            Search_Query("")
        End If
    End Sub

End Class
