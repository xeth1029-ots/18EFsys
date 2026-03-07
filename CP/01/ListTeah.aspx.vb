Partial Class ListTeah
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        modifytype.Value = TIMS.ClearSQM(Request("type")) 'hidden (Add:新增到指定的欄位名稱)
        fieldname.Value = TIMS.ClearSQM(Request("fieldname")) 'hidden (fieldname:新增到指定的欄位名稱)

        If Not Page.IsPostBack Then
            Call CreateItem1()
        End If
    End Sub

    Sub CreateItem1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim objtable As DataTable
        Dim vRID As String = TIMS.ClearSQM(Request("RID"))
        If vRID = "" Then vRID = sm.UserInfo.RID

        'WorkStatus 1:在職(Default) 2:離職
        'Dim objstr As String = ""
        'objstr = "" & vbCrLf
        'objstr &= " select a.TechID" & vbCrLf
        'objstr &= " ,a.TeachCName+CASE WHEN WorkStatus!='1' THEN '(-)' ELSE '' END TeachCName" & vbCrLf
        'objstr &= " ,a.KindEngage" & vbCrLf
        'objstr &= " from Teach_TeacherInfo a WITH(NOLOCK)" & vbCrLf
        'objstr &= " where 1=1" & vbCrLf
        'objstr &= " and a.RID = '" & vRID & "'"

        'WorkStatus 1:在職(Default) 2:離職
        Dim objstr As String = ""
        objstr = ""
        objstr &= " select TechID,TeachCName,KindEngage "
        objstr &= " from Teach_TeacherInfo WITH(NOLOCK)"
        objstr &= " where 1=1"
        objstr &= " And WorkStatus = '1'"
        objstr &= " and RID = '" & vRID & "'"
        objtable = DbAccess.GetDataTable(objstr, objconn)

        Dim mydv1 As New DataView(objtable)
        Dim mydv2 As New DataView(objtable)
        mydv1.RowFilter = "KindEngage = '1'"
        mydv2.RowFilter = "KindEngage = '2'"

        Me.LessonTeahList1.DataSource = mydv1
        Me.LessonTeahList1.DataTextField = "TeachCName"
        Me.LessonTeahList1.DataValueField = "TechID"
        Me.LessonTeahList1.DataBind()

        Me.LessonTeahList2.DataSource = mydv2
        Me.LessonTeahList2.DataTextField = "TeachCName"
        Me.LessonTeahList2.DataValueField = "TechID"
        Me.LessonTeahList2.DataBind()
    End Sub

    Private Sub LessonTeahList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList1.SelectedIndexChanged
        Dim vLTL As String = TIMS.Get_TeacherDegree(Me.LessonTeahList1.SelectedValue, objconn)
        Dim vLTL1DegName As String = TIMS.Get_DegreeValue(vLTL, objconn)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('" & Me.LessonTeahList1.SelectedValue & "','" & Me.LessonTeahList1.Items(Me.LessonTeahList1.SelectedIndex).Text & "','" & vLTL & "','" & vLTL1DegName & "');" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub LessonTeahList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList2.SelectedIndexChanged
        Dim vLTL As String = TIMS.Get_TeacherDegree(Me.LessonTeahList2.SelectedValue, objconn)
        Dim vLTLDegName As String = TIMS.Get_DegreeValue(vLTL, objconn)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('" & Me.LessonTeahList2.SelectedValue & "','" & Me.LessonTeahList2.Items(Me.LessonTeahList2.SelectedIndex).Text & "','" & vLTL & "','" & vLTLDegName & "');" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "returnValue('','','','');" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

End Class
