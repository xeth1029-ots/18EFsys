Partial Class SD_13_003_Bligate
    Inherits AuthBasePage

    'Const cst_學號 = 0
    'Const cst_姓名 = 1
    'Const cst_身分證號碼 = 2
    'Const cst_是否取得結訓資格 = 3
    'Const cst_出席達2分之3 = 4
    'Const cst_是否補助 = 5
    'Const cst_總費用 = 6
    'Const cst_補助費用 = 7
    'Const cst_個人支付 = 8
    'Const cst_剩餘可用餘額 = 9
    'Const cst_其他申請中金額 = 10
    'Const cst_審核狀態 = 11
    'Const cst_審核備註 = 12
    'Const cst_保險證號 = 13
    'Const cst_預算別代碼 = 14
    '參考　SD_12_007.aspx

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
        '檢查Session是否存在 End

        Name.Text = Request("Name")
        IDNO.Text = Request("IDNO")
        'Me.ViewState("ActNo") = Request("ActNo")
        'Me.ViewState("STDate") = Request("STDate")
        HidActNo.Value = Request("ActNo")
        HidSTDate.Value = Request("STDate")
        HidSOCID.Value = Request("SOCID")

        Name.Text = TIMS.ClearSQM(Name.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        HidActNo.Value = TIMS.ClearSQM(HidActNo.Value)
        HidSTDate.Value = TIMS.ClearSQM(HidSTDate.Value)
        HidSOCID.Value = TIMS.ClearSQM(HidSOCID.Value)

        If Not IsPostBack Then
            Dim Errmsg As String = ""
            If Not chk_Value1(Errmsg) Then Exit Sub '異常離開
            If Errmsg <> "" Then
                '傳入參數有誤
                If Session("_Search") Is Nothing AndAlso Not Me.ViewState("_Search") Is Nothing Then
                    Session("_Search") = Me.ViewState("_Search")
                    Me.ViewState("_Search") = Nothing
                End If
                Common.RespWrite(Me, "<script>alert('" & Errmsg & "');</script>")
                Common.RespWrite(Me, "<script>location.href='SD_13_003.aspx?ID=" & Request("ID") & "'</script>")
                Exit Sub 'Return rst
            End If

            Call creaet1()
        End If
    End Sub

    Function chk_Value1(ByVal Errmsg As String) As Boolean
        Dim rst As Boolean = False 'true:無異常 false:異常(預設)

        If Not Session("_Search") Is Nothing Then
            Me.ViewState("_Search") = Session("_Search")
            'Session("_Search") = Nothing
        End If

        If IDNO.Text = "" Then '未傳入身分證號
            Errmsg += "身分證號碼為空白\n" & vbCrLf
        End If
        If HidActNo.Value = "" Then '未傳入保險證號
            Errmsg += "保險證號為空白\n"
        End If
        If HidSOCID.Value = "" Then '未傳入學員序號
            Errmsg += "查詢參數異常\n"
        End If
        If Errmsg <> "" Then Return False

        Dim drSS As DataRow = TIMS.Get_StudData(HidSOCID.Value, objconn)
        If drSS Is Nothing Then '查不到學員序號
            Errmsg += "查詢參數異常\n"
        End If
        If Errmsg <> "" Then Return False

        If CStr(drSS("IDNO")) <> IDNO.Text Then '身分證號比對有誤
            Errmsg += "查詢參數異常\n"
        End If
        If Errmsg <> "" Then Return False

        If Errmsg = "" Then rst = True
        Return rst 'true:無異常 false:異常
    End Function

    Sub creaet1()
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)

        Dim sql As String = ""
        sql &= " SELECT a.FType ,a.ActNo ,a.MDate" & vbCrLf
        sql &= " ,a.ChangeMode ,a.Salary ,isnull(a.comname,b.UName) UName" & vbCrLf
        sql &= " FROM Stud_BligateData a" & vbCrLf
        sql &= " LEFT JOIN Bus_BasicData b ON a.ActNo=b.Ubno" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND a.ACTNO=@ACTNO" & vbCrLf
        sql &= " UNION" & vbCrLf
        sql &= " SELECT  UTYPE FType ,a.ActNo ,a.MDate" & vbCrLf
        sql &= " ,a.ChangeMode ,a.Salary ,isnull(a.comname,b.UName) UName" & vbCrLf
        sql &= " FROM STUD_BLIGATEDATA28 a" & vbCrLf
        sql &= " LEFT JOIN BUS_BASICDATA b ON a.ActNo=b.Ubno" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND a.ACTNO=@ACTNO" & vbCrLf
        sql &= " Order By 3" & vbCrLf 'MDate

        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO.Text
            .Parameters.Add("ActNo", SqlDbType.VarChar).Value = HidActNo.Value 'Me.ViewState("ActNo")
            dt.Load(.ExecuteReader())
        End With

        msg3.Text = "查無資料"
        DataGrid3.Visible = False
        If dt.Rows.Count > 0 Then
            msg3.Text = ""
            DataGrid3.Visible = True

            DataGrid3.DataSource = dt
            DataGrid3.DataBind()
        End If
    End Sub

    '回上一頁
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Not Me.ViewState("_Search") Is Nothing Then
            Session("_Search") = Me.ViewState("_Search")
            Me.ViewState("_Search") = Nothing
        End If
        Common.RespWrite(Me, "<script>location.href='SD_13_003.aspx?ID=" & Request("ID") & "'</script>")
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                Select Case Left(drv("ActNo").ToString, 3)
                    Case "020"
                        e.Item.Cells(1).Text = "職業工會"
                    Case "030"
                        e.Item.Cells(1).Text = "漁會"
                    Case "079"
                        e.Item.Cells(1).Text = "外僱船員"
                    Case "090"
                        e.Item.Cells(1).Text = "訓練機構"
                    Case Else
                        e.Item.Cells(1).Text = "一般"
                End Select
                Select Case drv("ChangeMode").ToString
                    Case "1"
                        e.Item.Cells(4).Text = "工作部門或特殊身分異動"
                    Case "2"
                        e.Item.Cells(4).Text = "退保"
                    Case "3"
                        e.Item.Cells(4).Text = "調薪"
                    Case "4"
                        e.Item.Cells(4).Text = "加保"
                End Select
        End Select
    End Sub
End Class
