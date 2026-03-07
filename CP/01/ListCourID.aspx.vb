Partial Class ListCourID
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        '' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Me.modifytype.Value = TIMS.ClearSQM(Request("type"))
        Me.fieldname.Value = TIMS.ClearSQM(Request("fieldname"))
        Dim OCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim ApplyDate As String = TIMS.ClearSQM(Request("ApplyDate"))

        If Not Page.IsPostBack Then

            Dim objtable As DataTable
            Dim objstr As String

            objstr = " SELECT cc.CourseName CourID, cs.class1 CourIDValue from Class_Schedule CS JOIN Course_CourseInfo CC on cs.class1= cc.courid where cs.OCID ='" & OCID & "' and cs.SchoolDate<= " & TIMS.To_date(ApplyDate) & vbCrLf
            For i As Integer = 2 To 12
                objstr += " UNION SELECT cc.CourseName CourID, cs.class" & CStr(i) & " CourIDValue from Class_Schedule CS JOIN Course_CourseInfo CC on cs.class" & CStr(i) & "= cc.courid where cs.OCID ='" & OCID & "' and cs.SchoolDate<= " & TIMS.To_date(ApplyDate) & vbCrLf '" & ApplyDate & "' " & vbCrLf
            Next
            objtable = DbAccess.GetDataTable(objstr, objconn)

            Dim mydv1 As New DataView(objtable)
            'mydv1.RowFilter = "KindEngage = '1'"

            Me.CourList1.DataSource = mydv1
            Me.CourList1.DataTextField = "CourID"
            Me.CourList1.DataValueField = "CourIDValue"
            Me.CourList1.DataBind()

        End If
    End Sub

    'Function GETDegreeID(ByVal TechID) As String
    '    Dim objstr As String
    '    objstr = "select DegreeID from Teach_TeacherInfo where TechID = '" & TechID & "'"
    '    Return DbAccess.ExecuteScalar(objstr)
    'End Function

    Private Sub CourList1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CourList1.SelectedIndexChanged
        Dim retValueMsg1 As String = $"returnValue('{CourList1.SelectedItem.Text}','{CourList1.SelectedValue}');"
        Dim strScript As String = $"<script language=""javascript"">{retValueMsg1}</script>{vbCrLf}"
        'strScript += "returnValue('" & Me.CourList1.SelectedValue & "','" & Me.CourList1.Items(Me.CourList1.SelectedIndex).Text & "');" + vbCrLf
        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        Dim strScript As String = "<script language=""javascript"">returnValue('','')</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

End Class
