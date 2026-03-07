'Imports Turbo
Partial Class SD_09_007_R
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

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
            syears = TIMS.GetSyear(syears)
            eyears = TIMS.GetSyear(eyears)
            TPlan = TIMS.Get_TPlan(TPlan)

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Button1.Attributes("onclick") = "javascript:return print();"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');showFrame('inline');"
            HistoryRID.Attributes("onclick") = "showFrame('none');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Private Sub syears_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles syears.SelectedIndexChanged, eyears.SelectedIndexChanged, TPlan.SelectedIndexChanged
        If syears.SelectedIndex <> 0 And eyears.SelectedIndex <> 0 And TPlan.SelectedIndex <> 0 Then
            Dim sqlstr As String = ""
            sqlstr = "" & vbCrLf
            sqlstr += " SELECT b.ClassCName, b.OCID,b.CyclType, b.LevelType  " & vbCrLf
            sqlstr += " FROM Class_ClassInfo b" & vbCrLf
            sqlstr += " join id_plan ip on ip.planid=b.planid and b.NotOpen='N' and b.IsSuccess='Y' " & vbCrLf
            sqlstr += " where 1=1" & vbCrLf
            sqlstr += " and ip.TPlanID='" & TPlan.SelectedValue & "'" & vbCrLf
            sqlstr += " and b.RID='" & RIDValue.Value & "'" & vbCrLf
            sqlstr += " and b.Years>='" & Right(syears.SelectedValue, 2) & "'" & vbCrLf
            sqlstr += " and b.Years<='" & Right(eyears.SelectedValue, 2) & "'" & vbCrLf
            Dim dt As DataTable
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            ClassName.Items.Clear()
            ClassName.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

            For i As Integer = 0 To dt.Rows.Count - 1
                Dim s_classname As String = TIMS.GET_CLASSNAME(Convert.ToString(dt.Rows(i).Item("ClassCName")), Convert.ToString(dt.Rows(i).Item("CyclType")))
                ClassName.Items.Add(New ListItem(s_classname, dt.Rows(i).Item("OCID")))
            Next
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim MyValue As String = ""
        MyValue = "syear=" & syears.SelectedValue & "&eyear=" & eyears.SelectedValue & "&RID=" & Me.RIDValue.Value & "&TPlanID=" & TPlan.SelectedValue & "&OCID=" & ClassName.SelectedValue & "&CJOB_UNKEY=" & cjobValue.Value & ""
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "MultiBlock", "Student_Rpt", MyValue)
    End Sub
End Class
