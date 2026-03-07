Partial Class TC_01_006_add
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '20100208 按新增時代查詢之 師資別名稱
            If Request("ProcessType") = "Insert" Then
                TextBox1.Text = Convert.ToString(Request("KindName"))
            End If
        End If

        Button1.Attributes("onclick") = "return chkdata();"
        'Button2.Attributes("onclick") = "history.go(-1);"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '取得欄位上的數值
        Dim KindEngage As String = TIMS.ClearSQM(DropDownList1.SelectedValue)
        Dim KindName As String = TIMS.ClearSQM(TextBox1.Text)
        Dim CateKind As String = TIMS.ClearSQM(DropDownList2.SelectedValue)
        Dim BaseHours As String = TIMS.ClearSQM(TextBox2.Text)
        Dim HightHours As String = TIMS.ClearSQM(TextBox3.Text)
        'Dim SciCharge = TextBox4.Text
        'Dim TechCharge = TextBox5.Text
        Dim OverCharge As String = TIMS.ClearSQM(TextBox6.Text)

        Dim sql As String = ""
        sql = ""
        sql &= " INSERT INTO ID_KINDOFTEACHER "
        sql &= " (KINDID,KindEngage,KindName,CateKind,BaseHours,HightHours,SciCharge,TechCharge,OverCharge,ModifyAcct,ModifyDate)"
        sql &= " VALUES"
        sql &= " (@KINDID,@KindEngage,@KindName,@CateKind,@BaseHours,@HightHours,@SciCharge,@TechCharge,@OverCharge,@ModifyAcct,getdate())"

        Dim intSerial As Integer = DbAccess.GetNewId(objconn, "ID_KINDOFTEACHER_KINDID_SEQ,ID_KINDOFTEACHER,KINDID")
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("KindID", intSerial)
        parms.Add("KindEngage", KindEngage)
        parms.Add("KindName", KindName)
        parms.Add("CateKind", Val(CateKind))
        parms.Add("BaseHours", Val(BaseHours))
        parms.Add("HightHours", Val(HightHours))
        parms.Add("SciCharge", 0)
        parms.Add("TechCharge", 0)
        parms.Add("OverCharge", IIf(OverCharge <> "", Val(OverCharge), 0))
        parms.Add("ModifyAcct", sm.UserInfo.UserID)
        DbAccess.ExecuteNonQuery(sql, objconn, parms)
        Common.RespWrite(Me, "<script language=javascript>window.alert('資料新增成功!');")
        Common.RespWrite(Me, "window.location.href='TC_01_006.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Dim url1 As String = "TC_01_006.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class