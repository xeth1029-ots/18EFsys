Partial Class CP_02_016_R
    Inherits System.Web.UI.Page

    '莊懷君…
    'CP_02_016_R
    'CP_02_016_R*.jrxml
    'Const cst_reportFN1 As String = "CP_02_016_R2" 'CP_02_016_R
    Const cst_reportFN1 As String = "CP_02_016_R3"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        'TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        '預算來源
        'BudgetList = TIMS.Get_Budget(BudgetList, 33, objconn)
        '身分別  '顯示的身分別
        '03"負擔家計婦女"併入28"獨立負擔家計者"計算,並把"負擔家計婦女"項目拿掉.
        'CM_03_007 (報表)
        'Identity = TIMS.Get_Identity(Identity, 66, objconn)
        '選擇全部身分別
        'Identity.Attributes("onclick") = "SelectAll('Identity','Identity_List');"
        Button1.Attributes("onclick") = "javascript:return search()"

        'Dim value2 As String = "" 'cst_Budget_無id
        'For i As Integer = 0 To BudgetList.Items.Count - 1
        '    If BudgetList.Items(i).Value <> "" Then
        '        If value2 <> "" Then value2 &= ","
        '        value2 &= BudgetList.Items(i).Value
        '    End If
        'Next
        'TIMS.SetCblValue(BudgetList, value2)
        'Dim value3 As String = "" 'cst_Budget_無id
        'For i As Integer = 0 To Identity.Items.Count - 1
        '    If Identity.Items(i).Value <> "" Then
        '        If value3 <> "" Then value3 &= ","
        '        value3 &= Identity.Items(i).Value
        '    End If
        'Next
        'TIMS.SetCblValue(Identity, value3)
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "結訓日期 起始日期格式有誤" & vbCrLf
            End If
        Else
            Errmsg += "結訓日期 起始日期 為必填" & vbCrLf
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "結訓日期 迄止日期格式有誤" & vbCrLf
            End If
        Else
            Errmsg += "結訓日期 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then
                    Errmsg += "【結訓日期】的起日不得大於【結訓日期】的迄日!!" & vbCrLf
                End If
            End If
        End If

        'If Convert.ToString(OCID.SelectedValue) = "" _
        '    OrElse Not IsNumeric(OCID.SelectedValue) Then
        '    Errmsg += "統計對象 為必選" & vbCrLf
        'End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        'Dim cGuid As String = SmartQuery.GetGuid(Page)
        'Dim Url As String = SmartQuery.GetUrl(Page)
        'Dim strScript As String
        'Dim Y1, Y2, SYMD, EYMD As String
        Dim Y1 As String = Convert.ToInt16(Convert.ToString(start_date.Text).Substring(0, 4)) - 1911
        Dim SYMD As String = Y1 & "年" & Convert.ToString(start_date.Text).Substring(5, 2) & "月"
        Dim Y2 As String = Convert.ToInt16(Convert.ToString(end_date.Text).Substring(0, 4)) - 1911
        Dim EYMD As String = Y2 & "年" & Convert.ToString(end_date.Text).Substring(5, 2) & "月"

        'Dim BudgetID As String = TIMS.GetCheckBoxListRptVal(BudgetList, 0)
        'Dim Identity1 As String = ""
        'For i As Integer = 1 To Identity.Items.Count - 1
        '    If Identity.Items(i).Selected Then
        '        If Identity1 <> "" Then Identity1 &= ","
        '        Identity1 &= "\'" & Me.Identity.Items(i).Value & "\'"
        '    End If
        'Next

        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "start_date", start_date.Text)
        TIMS.SetMyValue(sMyValue, "end_date", end_date.Text)
        TIMS.SetMyValue(sMyValue, "SYMD", SYMD)
        TIMS.SetMyValue(sMyValue, "EYMD", EYMD)
        'If BudgetID <> "" Then
        '    TIMS.SetMyValue(sMyValue, "BudgetID", BudgetID)
        'End If
        'If Identity1 <> "" Then
        '    TIMS.SetMyValue(sMyValue, "Identity", Identity1)
        'End If
        'MyValue = "start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&SYMD=" & SYMD & "&EYMD=" & EYMD
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN1, sMyValue)
    End Sub

End Class
