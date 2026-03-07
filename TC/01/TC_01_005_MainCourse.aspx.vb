Partial Class TC_01_005_MainCourse
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
        '檢查Session是否存在 End

        If Me.ViewState("MySort") Is Nothing Then
            ViewState("MySort") = "CourseID"
        Else
            If ViewState("MySort") = "CourseID" Then
                ViewState("MySort") = "CourseID DESC"
            Else
                ViewState("MySort") = "CourseID"
            End If
        End If

        'objconn = DbAccess.GetConnection()
        If Not IsPostBack Then
            Call create()
        End If

    End Sub

    Sub create()
        Dim aryrow() As String = {"課程代碼", "課程名稱"}
        Dim cell As New HtmlTableCell
        Dim row As New HtmlTableRow
        Dim i As Integer = 0
        Dim rqRID As String
        Dim rqOrgid As String
        Dim rqClass_id As String
        'RID = Request("rid")
        rqRID = Request("RID")
        rqOrgid = Request("orgid")
        rqClass_id = Request("classid")
        rqRID = TIMS.ClearSQM(rqRID)
        rqOrgid = TIMS.ClearSQM(rqOrgid)
        rqClass_id = TIMS.ClearSQM(rqClass_id)

        For i = 0 To aryrow.Length - 1
            cell = New HtmlTableCell
            If i = 0 Then
                Dim Linkbutton1 As New LinkButton
                Dim Image1 As New System.Web.UI.WebControls.Image
                Linkbutton1.ID = "Linkbutton1"
                Linkbutton1.Text = aryrow(i)
                cell.Controls.Add(Linkbutton1)

                If ViewState("MySort") = "CourseID DESC" Then
                    Image1.ImageUrl = "../../images/SortDown.gif"
                ElseIf ViewState("MySort") = "CourseID" Then
                    Image1.ImageUrl = "../../images/SortUp.gif"
                End If
                cell.Controls.Add(Image1)
            Else
                cell.InnerHtml = aryrow(i)
            End If
            cell.Attributes("Class") = "head_navy"
            row.Cells.Add(cell)
            row.Style("Color") = "#ffffff"
        Next
        row.Align = "center"
        row.BgColor = "#999900"
        row.Attributes("Class") = "head_navy"
        Me.search_tbl.Rows.Add(row)

        Dim planid As String = ""
        'Dim sqlAdapter As SqlDataAdapter
        Dim course_info As DataTable

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " select CourseID,CourseName,CourID  " & vbCrLf
        sqlstr += " from Course_CourseInfo" & vbCrLf
        sqlstr += " where MainCourID is NULL " & vbCrLf '空的才能為父層
        sqlstr += " and Valid='Y' " & vbCrLf '目前為有效值

        Me.trainValue.Value = TIMS.ClearSQM(Me.trainValue.Value)
        If Me.trainValue.Value <> "" Then
            sqlstr += " and TMID = '" & Me.trainValue.Value & "' " & vbCrLf
        End If

        ClassID.Text = TIMS.ClearSQM(ClassID.Text)
        'Me.ClassID.Text = Trim(Me.ClassID.Text)
        Me.ViewState("ClassID") = Replace(Me.ClassID.Text, "'", "''")
        If Me.ViewState("ClassID") <> "" Then
            sqlstr += " and CourseID like  '%" & Me.ViewState("ClassID") & "%' " & vbCrLf
        End If
        '自已不能為父層
        sqlstr += " and CourseID<>'" & rqClass_id & "'" & vbCrLf

        If rqRID = "" Then
            '暫依舊規則
            If rqOrgid <> "" Then
                sqlstr += " and orgid='" & rqOrgid & "' " & vbCrLf
            Else
                sqlstr += " and orgid='" & sm.UserInfo.OrgID & "'" & vbCrLf
            End If
            'sqlstr += " and (orgid='" & sm.UserInfo.OrgID & "' or orgid='" & orgid & "') " & vbCrLf
        Else
            sqlstr += " and rid='" & rqRID & "' " & vbCrLf '2010依RID決定使用單位、年度、計畫
        End If
        sqlstr += " order by CourseID" & vbCrLf

        Dim da As SqlDataAdapter = Nothing
        course_info = DbAccess.GetDataTable(sqlstr, da, objconn)
        i = 1
        For Each dr As DataRow In course_info.Select(Nothing, ViewState("MySort"))
            'Dim classid As String = ""
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerHtml = String.Format("<a href=""javascript:returnValue('{1}','{2}');""><font color='#000000'>{0}</font></a>", dr("CourseID"), dr("CourID"), dr("CourseName"))
            row.Cells.Add(cell)
            cell = New HtmlTableCell
            cell.InnerText = TIMS.ClearSQM(dr("CourseName"))
            row.Cells.Add(cell)
            row.Align = "center"
            Me.search_tbl.Rows.Add(row)
            If i Mod 2 = 0 Then row.BgColor = "WhiteSmoke"
            i = i + 1
        Next
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TB_career_id.Text = ""
        trainValue.Value = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call create()
    End Sub

    Protected Sub TB_career_id_TextChanged(sender As Object, e As EventArgs)

    End Sub
End Class
