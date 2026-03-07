Partial Class CM_03_009
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CM_03_009_1" '身分別
    Const cst_printFN2 As String = "CM_03_009_2" '年齡
    Const cst_printFN3 As String = "CM_03_009_3" '訓練職類
    Const cst_printFN4 As String = "CM_03_009_4" '教育程度
    Const cst_printFN5 As String = "CM_03_009_5" '性別
    Const cst_printFN6 As String = "CM_03_009_6" '通俗職類

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
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '年度
            list_Year = TIMS.GetSyear(list_Year)
            '轄區
            Show_ChkDistrict(True)
            '身分別
            Show_ChkIdentity(True)
            '計畫
            'Show_ChkTPlanID(True)
            chk_TPlanID = TIMS.Get_TPlan(chk_TPlanID, , 1, "Y")
            chk_TPlanID.Attributes.Add("style", "word-break:keep-all;word-wrap:normal;")
        End If

        '選擇全部轄區
        Dim tmpChk As CheckBox

        tmpChk = chk_District.Controls(0)
        tmpChk.Attributes.Add("onclick", "SelectAll(document.getElementById('" & chk_District.ClientID & "'));")
        '選擇全部身分別
        Dim tmpChk2 As CheckBox

        tmpChk2 = chk_Identity.Controls(0)
        tmpChk2.Attributes.Add("onclick", "SelectAll(document.getElementById('" & chk_Identity.ClientID & "'));")
        '選擇全部計畫
        Dim tmpChk3 As CheckBox

        tmpChk3 = chk_TPlanID.Controls(0)
        tmpChk3.Attributes.Add("onclick", "SelectAll(document.getElementById('" & chk_TPlanID.ClientID & "'));")
        '列印檢查
        Print.Attributes.Add("onclick", "return CheckPrint();")

        '如果統計項目的選項改變
        rdo_Mode.Attributes.Add("onclick", "ChangeMode(document.getElementById('" & rdo_Mode.ClientID & "'));")

        If rdo_Mode.SelectedIndex <= 0 Then
            RegisterStartupScript("", "<script language='javascript'>document.getElementById('" & IdentityTR.ClientID & "').style.display='none';</script>")
        End If
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Dim cGuid As String = ReportQuery.GetGuid(Page)
        Dim Url As String = ReportQuery.GetUrl(Page)
        'Dim urlStr As String = ""
        'Dim distStr As String = ""
        'Dim idStr As String = ""
        'Dim planStr As String = ""
        'Dim strScript As String

        '年度
        Dim urlStr As String = ""
        urlStr += "&Years=" & list_Year.SelectedValue
        '開訓日期起
        urlStr += "&SSTDate=" & txt_STDateS.Text
        '開訓日期迄
        urlStr += "&ESTDate=" & txt_STDateE.Text
        '結訓日期起
        urlStr += "&SFTDate=" & txt_FTDateS.Text
        '結訓日期迄
        urlStr += "&EFTDate=" & txt_FTDateE.Text

        '轄區
        Dim distStr As String = ""
        With chk_District.Items
            For j As Integer = 0 To .Count - 1
                If .Item(j).Selected AndAlso .Item(j).Value <> "" Then
                    If distStr <> "" Then distStr &= ","
                    distStr &= .Item(j).Value
                End If
            Next
        End With
        urlStr += "&DistID=" & distStr
        urlStr += "&DistID2=" & distStr

        '計畫
        Dim planStr As String = ""
        With chk_TPlanID.Items
            For j As Integer = 0 To .Count - 1
                If .Item(j).Selected AndAlso .Item(j).Value <> "" Then
                    If planStr <> "" Then planStr &= ","
                    planStr &= .Item(j).Value
                End If
            Next
        End With
        urlStr += "&TPlanID=" & planStr

        '身分別
        Dim idStr As String = ""
        If rdo_Mode.SelectedIndex <> 0 Then
            With chk_Identity.Items
                For j As Integer = 0 To .Count - 1
                    If .Item(j).Selected AndAlso .Item(j).Value <> "" Then
                        If idStr <> "" Then idStr &= ","
                        idStr &= "" & .Item(j).Value
                    End If
                Next
            End With
        End If
        urlStr += "&IdentityID=" & idStr

        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open(""" & Url & "GUID=" & cGuid & "&SQ_AutoLogout=true&path=TIMS" & If(SMPath <> "", "_" & SMPath, "") & "&sys=Report" & urlStr & """);"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        '<asp@ListItem Value="0">身分別</asp@ListItem>
        '<asp@ListItem Value="1">年齡</asp@ListItem>
        '<asp@ListItem Value="2">訓練職類</asp@ListItem>
        '<asp@ListItem Value="3">教育程度</asp@ListItem>
        '<asp@ListItem Value="4">性別</asp@ListItem>
        '<asp@ListItem Value="5">通俗職類</asp@ListItem>
        '</asp@RadioButtonList>

        Dim MyValue As String = ""
        Dim prtFName1 As String = ""
        '統計項目
        Select Case rdo_Mode.SelectedIndex
            Case 0 '身分別
                prtFName1 = cst_printFN1 '"CM_03_009_1"
            Case 1 '年齡
                prtFName1 = cst_printFN2 '"CM_03_009_2"
            Case 2 '訓練職類
                prtFName1 = cst_printFN3 '"CM_03_009_3"
            Case 3 '教育程度
                prtFName1 = cst_printFN4 '"CM_03_009_4"
            Case 4 '性別
                prtFName1 = cst_printFN5 '"CM_03_009_5"
            Case 5 '通俗職類
                prtFName1 = cst_printFN6 '"CM_03_009_6"
        End Select
        MyValue = urlStr
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFName1, MyValue)
    End Sub

#Region "NO USE"
    'Private Function Get_KeyPlan() As DataTable
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim objDS As New DataSet
    '    Dim sqlStr As String
    '    Dim rst As DataTable = Nothing

    '    Try
    '        sqlStr = "select TPlanID,PlanName from Key_Plan order by TPlanID asc "
    '        With sqlAdp
    '            .SelectCommand = New SqlCommand(sqlStr, objConn)
    '            .Fill(objDS, "Data")
    '        End With
    '        If Not objDS.Tables("Data") Is Nothing Then
    '            If objDS.Tables("Data").Rows.Count > 0 Then
    '                rst = objDS.Tables("Data")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        sqlAdp.Dispose()
    '        objDS.Clear()
    '        objDS.Dispose()
    '        objConn.Close()
    '        objConn.Dispose()
    '        Common.MessageBox(Me, ex.ToString)
    '    End Try
    '    Return rst
    'End Function

    'Private Sub Show_ChkTPlanID(Optional ByVal tmpState As Boolean = False)
    '    Dim dt As DataTable = Get_KeyPlan()

    '    If Not dt Is Nothing Then
    '        chk_TPlanID.DataSource = dt
    '        chk_TPlanID.DataTextField = "PlanName"
    '        chk_TPlanID.DataValueField = "TPlanID"
    '        chk_TPlanID.DataBind()
    '    End If
    '    If tmpState = True Then
    '        chk_TPlanID.Items.Insert(0, New ListItem("全部", ""))
    '    End If
    'End Sub
#End Region

#Region "Function"
    Function Get_IDDistrict() As DataTable
        Dim rst As DataTable = Nothing
        Dim sql As String = ""
        sql = "SELECT DISTID,NAME FROM ID_DISTRICT ORDER BY DISTID ASC "
        rst = DbAccess.GetDataTable(sql, objconn)
        Return rst
    End Function

    Function Get_KeyIdentity() As DataTable
        Dim rst As DataTable = Nothing
        Const cst_Iden1 As String = "'03','04','05','06','07','10','13','14','18','27','28'"
        Dim sql As String = ""
        sql = "SELECT * FROM KEY_IDENTITY WHERE IDENTITYID in (" & cst_Iden1 & ") ORDER BY IDENTITYID ASC "
        rst = DbAccess.GetDataTable(sql, objconn)
        Return rst
    End Function

    Private Sub Show_ChkDistrict(Optional ByVal tmpState As Boolean = False)
        Dim dt As DataTable = Get_IDDistrict()

        If Not dt Is Nothing Then
            chk_District.DataSource = dt
            chk_District.DataTextField = "Name"
            chk_District.DataValueField = "DistID"
            chk_District.DataBind()
        End If
        If tmpState = True Then
            chk_District.Items.Insert(0, New ListItem("全部", ""))
        End If
    End Sub

    Private Sub Show_ChkIdentity(Optional ByVal tmpState As Boolean = False)
        Dim dt As DataTable = Get_KeyIdentity()

        If Not dt Is Nothing Then
            chk_Identity.DataSource = dt
            chk_Identity.DataTextField = "Name"
            chk_Identity.DataValueField = "IdentityID"
            chk_Identity.DataBind()
        End If
        If tmpState = True Then
            chk_Identity.Items.Insert(0, New ListItem("全部", ""))
        End If
    End Sub
#End Region

End Class
