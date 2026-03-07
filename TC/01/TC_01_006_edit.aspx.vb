Partial Class TC_01_006_edit
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

        If Not IsPostBack Then
            sCreate1()
        End If

        'Button2.Attributes("onclick") = "history.go(-1);"
    End Sub

    Sub sCreate1()
        'Button1.Attributes("onclick") = "return chk();"
        TextBox2.Attributes("onblur") = "return chk(this,'基本時數');"
        TextBox3.Attributes("onblur") = "return chk(this,'最高請領時數');"
        TextBox4.Attributes("onblur") = "return chk(this,'一般鐘點費');"
        TextBox5.Attributes("onblur") = "return chk(this,'一般鐘點費');"
        TextBox6.Attributes("onblur") = "return chk(this,'超時鐘點費');"
        Button1.Attributes("onclick") = "return chkdata();"

        Dim rqSerial As String = TIMS.ClearSQM(Request("serial"))
        Dim parms As New Hashtable From {{"KindID", rqSerial}}
        Dim sql As String = "select * from ID_KindOfTeacher where KindID=@KindID"
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr1 Is Nothing Then Exit Sub

        '判斷資料並填入表單

        '判斷內聘或外聘
        If dr1("KindEngage") = "1" Then
            DropDownList1.Items(1).Selected = True
        Else
            DropDownList1.Items(2).Selected = True
        End If
        TextBox1.Text = dr1("kindName")

        '判斷何種師資類型
        If dr1("CateKind") = "1" Then
            DropDownList2.Items(1).Selected = True
        ElseIf dr1("CateKind") = "2" Then
            DropDownList2.Items(2).Selected = True
        Else
            DropDownList2.Items(3).Selected = True
        End If

        TextBox2.Text = dr1("BaseHours")
        TextBox3.Text = dr1("HightHours")
        TextBox4.Text = dr1("SciCharge")
        TextBox5.Text = dr1("TechCharge")
        TextBox6.Text = dr1("OverCharge")

        CB_NOUSE.Checked = (Convert.ToString(dr1("NOUSE")) = "Y")
        'Dim cmd As New SqlCommand(sql, objconn)
        'Dim rs As SqlDataReader
        'rs = cmd.ExecuteReader
        'If rs.Read Then
        'End If
        'rs.Close()
        ''conn.Close()
        'rs = Nothing
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '取得欄位上的數值
        Dim KindEngage As String = TIMS.ClearSQM(DropDownList1.SelectedValue)
        Dim KindName As String = TIMS.ClearSQM(TextBox1.Text)
        Dim CateKind As String = TIMS.ClearSQM(DropDownList2.SelectedValue)
        Dim BaseHours As String = TIMS.ClearSQM(TextBox2.Text)
        Dim HightHours As String = TIMS.ClearSQM(TextBox3.Text)
        Dim SciCharge As String = TIMS.ClearSQM(TextBox4.Text)
        Dim TechCharge As String = TIMS.ClearSQM(TextBox5.Text)
        Dim OverCharge As String = TIMS.ClearSQM(TextBox6.Text)
        Dim rqSerial As String = TIMS.ClearSQM(Request("serial"))
        Dim V_CB_NOUSE As String = If(CB_NOUSE.Checked, "Y", "")
        Dim sql As String = ""
        sql &= " UPDATE ID_KindOfTeacher "
        sql &= " SET KindEngage = @KindEngage "
        sql &= " ,KindName = @KindName "
        sql &= " ,CateKind = @CateKind "
        sql &= " ,BaseHours = @BaseHours "
        sql &= " ,HightHours = @HightHours "
        sql &= " ,SciCharge = @SciCharge "
        sql &= " ,TechCharge = @TechCharge "
        sql &= " ,OverCharge = @OverCharge "
        sql &= " ,ModifyAcct = @ModifyAcct "
        sql &= " ,ModifyDate = GETDATE() "
        sql &= " ,NOUSE = @NOUSE"
        sql &= " WHERE KindID = @KindID "  ''" & rqSerial & "'

        'parms.Clear()
        Dim parms As New Hashtable From {
            {"KindEngage", KindEngage},
            {"KindName", KindName},
            {"CateKind", Val(CateKind)},
            {"BaseHours", Val(BaseHours)},
            {"HightHours", Val(HightHours)},
            {"SciCharge", If(SciCharge <> "", Val(SciCharge), 0)},
            {"TechCharge", If(TechCharge <> "", Val(TechCharge), 0)},
            {"OverCharge", If(OverCharge <> "", Val(OverCharge), 0)},
            {"ModifyAcct", sm.UserInfo.UserID},
            {"KindID", rqSerial},
            {"NOUSE", If(V_CB_NOUSE <> "", V_CB_NOUSE, Convert.DBNull)}
        }
        DbAccess.ExecuteNonQuery(sql, objconn, parms)

        Common.RespWrite(Me, "<script language=javascript>window.alert('資料修改成功!');")
        Common.RespWrite(Me, "window.location.href='TC_01_006.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Dim url1 As String = "TC_01_006.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class