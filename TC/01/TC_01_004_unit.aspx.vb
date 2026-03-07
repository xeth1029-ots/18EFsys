Partial Class TC_01_004_unit
    Inherits AuthBasePage

    'Dim CName_str, classunit_str As String
    'Dim objreader As SqlDataReader

    Const Cst_className1 As String = "短期電腦研習課程" '短期電腦研習課程
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

        Me.Label1.Text = Cst_className1 ' "短期電腦研習課程"
        bt_send.Attributes("onclick") = "javascript:return checkvalue();"

        Dim check1 As String = ""
        Dim check2 As String = ""
        Dim check3 As String = ""
        Dim check4 As String = ""

        If Not IsPostBack Then
            'Dim classunit_str As String = ""
            Dim sqlstr1 As String = "SELECT DGID,DGNAME FROM KEY_DGTHOUR ORDER BY DGID"
            Dim dt1 As DataTable = DbAccess.GetDataTable(sqlstr1, objconn)
            'objreader = DbAccess.GetReader(sqlstr_Key_DGTHour, objconn)
            With CheckBoxList1
                .DataSource = dt1
                .DataTextField = "DGNAME"
                .DataValueField = "DGID"
                .DataBind()
            End With

            Dim classunit_str As String = TIMS.ClearSQM(Request("classunit"))
            If classunit_str <> "" Then
                check1 = classunit_str.Substring(0, 1)
                check2 = classunit_str.Substring(1, 1)
                check3 = classunit_str.Substring(2, 1)
                check4 = classunit_str.Substring(3, 1)
            End If
            If check1 = "1" Then
                CheckBoxList1.Items(0).Selected = True
            Else
                CheckBoxList1.Items(0).Selected = False
            End If
            If check2 = "1" Then
                CheckBoxList1.Items(1).Selected = True
            Else
                CheckBoxList1.Items(1).Selected = False
            End If
            If check3 = "1" Then
                CheckBoxList1.Items(2).Selected = True
            Else
                CheckBoxList1.Items(2).Selected = False
            End If
            If check4 = "1" Then
                CheckBoxList1.Items(3).Selected = True
            Else
                CheckBoxList1.Items(3).Selected = False
            End If
        End If

    End Sub

    Private Sub bt_send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_send.Click
        Dim str1 As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim all As String = ""
        'Dim param, all, str1, str2, str3, str4 As String
        If CheckBoxList1.Items(0).Selected = True Then
            str1 = "1"
            all += "1"
        Else
            str1 = "0"
        End If
        If CheckBoxList1.Items(1).Selected = True Then
            str2 = "1"
            all += "2"
        Else
            str2 = "0"
        End If
        If CheckBoxList1.Items(2).Selected = True Then
            str3 = "1"
            all += "3"
        Else
            str3 = "0"
        End If
        If CheckBoxList1.Items(3).Selected = True Then
            str4 = "1"
            all += "4"
        Else
            str4 = "0"
        End If
        Dim param As String = Convert.ToString(str1) + Convert.ToString(str2) + Convert.ToString(str3) + Convert.ToString(str4) + "0"
        tb_class_unit.Value = param
        Dim CName_str As String = Cst_className1 + "-" + all + "單元"
        tb_class_name.Value = CName_str
        If tb_class_unit.Value <> "" And tb_class_name.Value <> "" Then '都有值才回傳,並關視窗
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "opener.document.form1." & Request("textField") & ".value='';" + vbCrLf
            strScript += "opener.document.form1." & Request("textField") & ".value=form1.tb_class_name.value;" + vbCrLf
            strScript += "opener.document.form1." & Request("valueField") & ".value='';" + vbCrLf
            strScript += "opener.document.form1." & Request("valueField") & ".value=form1.tb_class_unit.value;" + vbCrLf
            strScript += "window.close();" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("window_onload", strScript)
        End If
    End Sub


End Class
