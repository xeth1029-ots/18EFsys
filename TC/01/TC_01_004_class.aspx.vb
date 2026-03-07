Partial Class TC_01_004_class
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

        If Not IsPostBack Then
            Call Search_Query()

            'txtSearch1.Attributes("onkeypress") = "Search_click();return false;"
            'txtSearch2.Attributes("onkeypress") = "Search_click();return false;"
            txtSearch1.Attributes("onkeypress") = "Search_click();"
            txtSearch2.Attributes("onkeypress") = "Search_click();"

            Me.btnSearch.Attributes("onclick") = "Search_click2();return false;"
            Me.btnSearch.Attributes("onkeypress") = "Search_click2();return false;"
        End If

    End Sub

    Sub Search_Query(Optional ByVal ClassNameVal As String = "", Optional ByVal ClassIDVal As String = "") 'As String
        Dim plan As DataRow
        Dim sqlstr_A As String = ""
        Dim TMID1 As String = Request("TMID") 'TrainIDTMID
        TMID1 = TIMS.ClearSQM(TMID1)
        Dim course_info As DataTable
        Dim sqlstr As String

        Dim i As Integer = 0
        Dim aryrow() As String = {"班別代碼", "班別名稱"}
        Dim cell As New HtmlTableCell
        Dim row As New HtmlTableRow

        sqlstr_A = "" & vbCrLf
        sqlstr_A += " select b.PlanName" & vbCrLf
        sqlstr_A += " ,b.TPlanID" & vbCrLf
        sqlstr_A += " from ID_Plan a " & vbCrLf
        sqlstr_A += " join Key_Plan b on a.TPlanID=b.TPlanID " & vbCrLf
        sqlstr_A += " where a.PlanID='" & sm.UserInfo.PlanID & "'"
        plan = DbAccess.GetOneRow(sqlstr_A, objconn)

        Me.ProecessType.Text = plan("PlanName")

        sqlstr = "" & vbCrLf
        sqlstr &= " select a.CLSID" & vbCrLf
        sqlstr &= " ,a.ClassID" & vbCrLf
        sqlstr &= " ,a.ClassName" & vbCrLf
        sqlstr &= " ,a.ClassEName" & vbCrLf
        sqlstr &= " ,c.TMID" & vbCrLf
        sqlstr &= " ,CASE when c.JobID is null then c.TrainID else c.JobID end TrainID" & vbCrLf
        sqlstr &= " ,CASE when c.JobID is null then '['+c.TrainID+']'+c.trainName else '['+c.JobID+']'+c.JobName end TrainName" & vbCrLf
        sqlstr &= " ,CASE when c.JobID is null then c.TrainID else c.JobID end JobID" & vbCrLf
        sqlstr &= " ,CASE when c.JobID is null then '['+c.TrainID+']'+c.trainName else '['+c.JobID+']'+c.JobName end JobName" & vbCrLf
        sqlstr &= " from ID_Class a " & vbCrLf
        sqlstr &= " join Key_Plan b on a.TPlanID=b.TPlanID " & vbCrLf
        sqlstr &= " join Key_TrainType c on a.TMID=c.TMID" & vbCrLf
        sqlstr &= " where 1=1" & vbCrLf
        sqlstr &= " and a.TPlanID='" & plan("TPlanID") & "'" & vbCrLf
        sqlstr &= " and a.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        sqlstr &= " and a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        'penny 加入職類判斷
        If TMID1 <> "" Then
            ' 97年前產業人才投資方案 TMID1使用  TrainID的tmid for (tims/產學訓)
            ' 97年後產業人才投資方案將 TMID1改為 JobID的tmid for (產學訓)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '**by Milor 20080512--PM應客戶問題，要讓產學訓班別代碼選擇時不受職類限制，所以將此條件MARK----start
                'sqlstr += " AND c.TMID IN ('" & TMID1 & "') " & vbCrLf
                '**by Milor 20080512----end
            End If
        End If
        sqlstr &= " ORDER BY a.ClassID" & vbCrLf
        course_info = DbAccess.GetDataTable(sqlstr, objconn)

        For i = 0 To aryrow.Length - 1
            cell = New HtmlTableCell
            cell.InnerText = aryrow(i)
            row.Cells.Add(cell)
            row.Attributes("Class") = "head_navy"
        Next
        row.Align = "center"
        Me.search_tbl.Rows.Add(row)

        i = 1
        For Each dr As DataRow In course_info.Rows
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerHtml = String.Format("<a href=""javascript:returnValue('{0}','{1}','{2}','{3}','{4}');""><font color='#000000'>{1}</font></a>", dr("CLSID"), dr("ClassID"), dr("ClassEName"), dr("TMID"), dr("TrainName"))
            row.Cells.Add(cell)
            cell = New HtmlTableCell
            cell.InnerText = dr("ClassName")
            row.Cells.Add(cell)
            row.Align = "center"
            Me.search_tbl.Rows.Add(row)
            If i Mod 2 = 0 Then row.BgColor = "#F5F5F5"
            i = i + 1
        Next
    End Sub

    Private Sub hbtnSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hbtnSearch.ServerClick
        Dim InputVal As String = ""
        Me.txtSearch1.Text = Trim(Me.txtSearch1.Text)
        Me.txtSearch2.Text = Trim(Me.txtSearch2.Text)

        InputVal = Me.txtSearch1.Text
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        End If

        InputVal = Me.txtSearch2.Text
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        End If

        Search_Query(Me.txtSearch1.Text, Me.txtSearch2.Text)
    End Sub
End Class
