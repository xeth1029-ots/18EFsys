Partial Class SYS_06_001
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        PageControler1.PageDataGrid = DataGrid1

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'End If

        If Not IsPostBack Then
            cCreate1()
        End If

    End Sub

    Sub cCreate1()
        msg.Text = ""
        DataGridtable.Visible = False
        SDate.Text = TIMS.Cdate3(DateAdd(DateInterval.Month, -3, Now.Date))
        EDate.Text = TIMS.Cdate3(Now.Date)
        Dim s_FunIDs As String = "57,246,766,60,63,83"
        FunID = Get_FunIDUse3(FunID, objconn, s_FunIDs)

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Function Get_FunIDUse3(ByVal obj As ListControl, ByVal conn As SqlConnection, ByVal s_FunIDs As String) As ListControl
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select  a.Name +'('+cast(a.FunID as varchar)+')' Name1" & vbCrLf
        sql &= " ,a.FunID" & vbCrLf
        sql &= " FROM dbo.VIEW_FUNCTION a WITH(NOLOCK)" & vbCrLf
        sql &= " where a.FunID in (" & s_FunIDs & ")" & vbCrLf
        sql &= " ORDER BY a.FunID" & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, conn)
        With obj
            .DataSource = dt
            .DataTextField = "Name1"
            .DataValueField = "FunID"
            .DataBind()
        End With
        Return obj
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call search1()
    End Sub

    Sub search1()
        'A/M/D/S/P
        Dim v_sFunID As String = TIMS.GetListValue(FunID) '.SelectedValue)
        ModifyAcct.Text = TIMS.ClearSQM(ModifyAcct.Text)
        SDate.Text = TIMS.ClearSQM(SDate.Text)

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT a.ModifyAcct"
        sql &= " ,a.FunID"
        sql &= " ,a.DistID"
        sql &= " ,a.SDLID"
        sql &= " ,a.DelNote"
        sql &= " ,a.ModifyDate"
        sql &= " ,b.Name ACCNAME "
        sql &= " FROM dbo.SYS_DELLOG a WITH(NOLOCK)"
        sql &= " JOIN dbo.AUTH_ACCOUNT b WITH(NOLOCK) ON b.Account=a.ModifyAcct"
        sql &= " WHERE 1=1"
        sql &= " and a.FunID ='" & v_sFunID & "'"
        If SDate.Text <> "" Then
            sql &= " and a.ModifyDate >= " & TIMS.To_date(SDate.Text)
        End If
        If EDate.Text <> "" Then
            sql &= " and a.ModifyDate <= " & TIMS.To_date(DateAdd(DateInterval.Day, 1, CDate(EDate.Text)))
        End If
        If ModifyAcct.Text <> "" Then
            sql &= " and a.ModifyAcct like '%" & ModifyAcct.Text & "%'"
        End If
        If sm.UserInfo.RID <> "A" Then
            sql &= " and a.DistID='" & sm.UserInfo.DistID & "'"
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Text = "查無資料"
        DataGridtable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridtable.Visible = True
            'PageControler1.SqlPrimaryKeyDataCreate(sql, "SDLID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SDLID"
            PageControler1.ControlerLoad()
        End If

    End Sub
End Class
