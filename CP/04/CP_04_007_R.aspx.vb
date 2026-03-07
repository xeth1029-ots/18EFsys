Partial Class CP_04_007_R
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

        '檢查日期格式
        Me.SSTDate.Attributes("onchange") = "check_date();"
        Me.ESTDate.Attributes("onchange") = "check_date();"
        Me.SFTDate.Attributes("onchange") = "check_date();"
        Me.EFTDate.Attributes("onchange") = "check_date();"
        print.Attributes("onclick") = "return search();"

        If Not Page.IsPostBack Then

            CreateItem()

        End If

        If sm.UserInfo.DistID = "000" Then
            DistrictList.Enabled = True
        Else
            DistrictList.SelectedValue = sm.UserInfo.DistID
            DistrictList.Enabled = False
        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        ''選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

    End Sub

    Sub CreateItem()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        'sqlstr = "SELECT Name,DistID FROM ID_District ORDER BY DistID"
        'dt = DbAccess.GetDataTable(sqlstr, objconn)
        'Me.DistrictList.DataSource = dt
        'Me.DistrictList.DataTextField = "Name"
        'Me.DistrictList.DataValueField = "DistID"
        'Me.DistrictList.DataBind()
        'Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        DistrictList = TIMS.Get_DistID(DistrictList)
        DistrictList.Items.Remove(DistrictList.Items.FindByValue(""))
        Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

    End Sub

    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click

        '選擇轄區
        'Dim Sqlstr As String
        'Dim objitem As ListItem
        'Dim itemstr As String
        'Dim DistID1, DistName, newDistID, newDistName As String
        'Dim TPlanID1, TPlanName, newTPlanID, newTPlanIDName As String
        'Dim i, j As Integer
        'Dim msg As String = ""
        'Dim sqlstr As String

        '報表要用的轄區參數
        '選擇轄區
        Dim DistID1, DistName As String
        DistID1 = ""
        DistName = ""
        If sm.UserInfo.DistID <> "000" Then
            DistID1 = sm.UserInfo.DistID

            For i As Integer = 0 To Me.DistrictList.Items.Count - 1
                If Me.DistrictList.Items(i).Value <> "" _
                    AndAlso Me.DistrictList.Items(i).Value = DistID1 Then

                    DistName = Me.DistrictList.Items(i).Text
                    Exit For
                End If
            Next
            'sqlstr = "SELECT Name FROM ID_District where DistID=" & sm.UserInfo.DistID
            'DistName = DbAccess.ExecuteScalar(sqlstr, objconn)

        Else
            For i As Integer = 1 To Me.DistrictList.Items.Count - 1
                If Me.DistrictList.Items(i).Selected Then
                    If DistID1 <> "" Then DistID1 &= ","
                    DistID1 &= Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")

                    If DistName <> "" Then DistName &= ","
                    DistName &= Convert.ToString(Me.DistrictList.Items(i).Text)
                End If
            Next

            'If DistID1 <> "" Then
            '    newDistID = Mid(DistID1, 1, DistID1.Length - 1)
            '    NewDistName = Mid(DistName, 1, DistName.Length - 1)
            'End If
        End If

        '報表要用的訓練計畫參數
        Dim j As Integer = 0
        Dim TPlanID1, TPlanName As String
        TPlanID1 = ""
        TPlanName = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")

                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Convert.ToString(Me.TPlanID.Items(i).Text)

                j = j + 1
            End If
        Next

        If TPlanID1 <> "" AndAlso j = (Me.TPlanID.Items.Count - 1) Then
            TPlanName = "全部"
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "SSTDate=" & SSTDate.Text
        MyValue &= "&ESTDate=" & ESTDate.Text
        MyValue &= "&SFTDate=" & SFTDate.Text
        MyValue &= "&EFTDate=" & EFTDate.Text
        MyValue &= "&DistID=" & DistID1
        MyValue &= "&TPlanID=" & TPlanID1
        MyValue &= "&DistName=" & Server.UrlEncode(DistName)
        MyValue &= "&PlanName=" & Server.UrlEncode(TPlanName)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CP_04_007_R", MyValue)

    End Sub

    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click

        Dim i As Short
        For i = 0 To Me.DistrictList.Items.Count - 1
            Me.DistrictList.Items(i).Selected = False
        Next

        Me.SSTDate.Text = ""
        Me.ESTDate.Text = ""
        Me.SFTDate.Text = ""
        Me.EFTDate.Text = ""

    End Sub
End Class
