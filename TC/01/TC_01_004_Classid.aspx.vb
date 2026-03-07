Partial Class TC_01_004_Classid
    Inherits AuthBasePage

    '(產投使用)
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        If Not IsPostBack Then
            Call Search_Query()
            'txtSearch1.Attributes("onkeypress") = "Search_click();return false;"
            'txtSearch2.Attributes("onkeypress") = "Search_click();return false;"
            txtSearch1.Attributes("onkeypress") = "Search_click();"
            txtSearch2.Attributes("onkeypress") = "Search_click();"
            btnSearch.Attributes("onclick") = "Search_click2();return false;"
            btnSearch.Attributes("onkeypress") = "Search_click2();return false;"
        End If
    End Sub

    Sub Search_Query(Optional ByVal ClassNameVal As String = "", Optional ByVal ClassIDVal As String = "") 'As String
#Region "(No Use)"

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
        '    Exit Sub
        'End If

#End Region

        Dim course_info As DataTable

        Dim aryrow() As String = {"選擇", "班別代碼", "班別名稱"}
        Dim cell As New HtmlTableCell
        Dim row As New HtmlTableRow
        Dim i As Integer

        Dim TMID1 As String = TIMS.ClearSQM(Request("TMID"))
        Hid_PlanID.Value = TIMS.ClearSQM(Request("PlanID"))
        If Hid_PlanID.Value = "" Then Hid_PlanID.Value = sm.UserInfo.PlanID

        Dim parms_A As New Hashtable
        parms_A.Add("PlanID", Hid_PlanID.Value)
        Dim sqlstr_A As String = ""
        sqlstr_A = "" & vbCrLf
        sqlstr_A &= " SELECT b.PlanName ,b.TPlanID,a.DISTID " & vbCrLf
        sqlstr_A &= " FROM ID_Plan a " & vbCrLf
        sqlstr_A &= " JOIN Key_Plan b ON a.TPlanID = b.TPlanID " & vbCrLf
        sqlstr_A &= " WHERE a.PlanID = @PlanID "
        Dim drPlan As DataRow = DbAccess.GetOneRow(sqlstr_A, objconn, parms_A)
        If drPlan Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        ProecessType.Text = drPlan("PlanName").ToString()
        Hid_DistID.Value = drPlan("DISTID").ToString()
        If Hid_DistID.Value = "" Then Hid_DistID.Value = sm.UserInfo.DistID

        'penny 加入職類判斷
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("TPlanID", drPlan("TPlanID").ToString())
        parms.Add("DistID", Hid_DistID.Value)
        parms.Add("Years", sm.UserInfo.Years.ToString())

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT a.CLSID " & vbCrLf
        sqlstr &= " ,a.ClassID " & vbCrLf
        sqlstr &= " ,a.ClassName " & vbCrLf
        sqlstr &= " ,a.Content " & vbCrLf
        sqlstr &= " ,c.TMID " & vbCrLf
        If iPYNum >= 3 Then
            sqlstr &= " ,c.TrainID " & vbCrLf
            sqlstr &= " ,c.TrainName " & vbCrLf
            sqlstr &= " ,c.JobID " & vbCrLf
            sqlstr &= " ,c.JobName " & vbCrLf
        Else
            sqlstr &= " ,CASE WHEN c.JobID IS NULL THEN c.TrainID ELSE c.JobID END TrainID " & vbCrLf
            sqlstr &= " ,CASE WHEN c.JobID IS NULL THEN c.trainName ELSE c.JobName END TrainName " & vbCrLf
            sqlstr &= " ,CASE WHEN c.JobID IS NULL THEN c.TrainID ELSE c.JobID END JobID " & vbCrLf
            sqlstr &= " ,CASE WHEN c.JobID IS NULL THEN c.trainName ELSE c.JobName END JobName " & vbCrLf
        End If
        sqlstr &= " FROM ID_CLASS a " & vbCrLf
        sqlstr &= " JOIN KEY_PLAN b ON a.TPlanID = b.TPlanID " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If iPYNum >= 3 Then
                '產投類
                sqlstr &= " JOIN VIEW_TRAINTYPE c ON a.TMID = c.TMID AND c.BUSID = 'H' AND c.LEVELS = 2 " & vbCrLf
            Else
                '產投類
                sqlstr &= " JOIN VIEW_TRAINTYPE c ON a.TMID = c.TMID AND c.BUSID = 'G' AND c.LEVELS = 1 " & vbCrLf
            End If
        Else
            '非產投
            sqlstr &= " JOIN VIEW_TRAINTYPE c ON a.TMID = c.TMID AND c.BUSID < 'G' AND c.LEVELS = 2 " & vbCrLf
        End If
        sqlstr &= " WHERE 1=1 " & vbCrLf
        sqlstr &= " AND a.TPlanID =@TPlanID" & vbCrLf
        sqlstr &= " AND a.DistID =@DistID" & vbCrLf
        sqlstr &= " AND a.Years =@Years" & vbCrLf
        ClassNameVal = Trim(ClassNameVal)
        If ClassNameVal <> "" Then sqlstr &= " AND a.ClassName LIKE '%" & ClassNameVal & "%' " & vbCrLf
        If ClassIDVal <> "" Then sqlstr &= " AND a.ClassID LIKE '%" & ClassIDVal & "%' " & vbCrLf

        If TMID1 <> "" Then
            '97年前產業人才投資方案 TMID1使用  TrainID的tmid for (tims/產學訓)
            '97年後產業人才投資方案將 TMID1改為 JobID的tmid for (產學訓)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sqlstr &= " AND c.TMID IN ('" & TMID1 & "') " & vbCrLf
            Else
                sqlstr &= " AND c.TrainID IN ('" & TMID1 & "') " & vbCrLf
            End If
        End If
        sqlstr &= " ORDER BY a.ClassID " & vbCrLf
        course_info = DbAccess.GetDataTable(sqlstr, objconn, parms)

        'Common.RespWrite(Me, sqlstr)
        Me.search_tbl.Rows.Clear()

        For i = 0 To aryrow.Length - 1
            cell = New HtmlTableCell
            cell.InnerText = aryrow(i)
            row.Cells.Add(cell)
            'row.Style("Color") = "#FFFFFF"  'edit，by:20181024
            row.Style("Color") = "#FFFFFF"   'edit，by:20181024
        Next
        row.Align = "center"
        'row.BgColor = "#999900"  'edit，by:20181024
        row.Attributes.Add("class", "head_navy")  'edit，by:20181024
        Me.search_tbl.Rows.Add(row)

        If course_info.Rows.Count = 0 Then
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = "目前沒有班別可以轉入!!"
            cell.ColSpan = Me.search_tbl.Rows(0).Cells.Count
            row.Cells.Add(cell)
            row.Align = "center"
            Me.search_tbl.Rows.Add(row)
            Exit Sub
        End If

        i = 1
        For Each dr As DataRow In course_info.Rows
            row = New HtmlTableRow

            cell = New HtmlTableCell
            cell.InnerHtml = String.Format("<input type=radio id={0} name={0} value='{1}'>", "classid", dr("CLSID"))
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerHtml = dr("ClassID")
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerText = dr("ClassName")
            row.Cells.Add(cell)

            row.Align = "left"

            search_tbl.Rows.Add(row)
            If i Mod 2 = 0 Then row.BgColor = "#FFFDC7"
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