Partial Class CM_03_005
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CM_03_005_R"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        If Not IsPostBack Then
            CreateItem()
            FTDate2.Text = Now.Date
            'OCID.Style("display") = "none"  '把選擇機構及班級的選項去掉
            'msg.Text = cst_NODATAMsg11
            'Else
            '    msg.Text = ""
        End If

        DistID.Attributes("onclick") = "ClearData();"
        TPlanID.Attributes("onclick") = "ClearData();"

        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

        'Button4.Visible = False

        'Button2.Attributes("onclick") = "GetOrg();"
        'Button3.Style("display") = "none"
        'center.Visible = False
        'OCID.Visible = False
        'msg.Visible = False
    End Sub

    Sub CreateItem()
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        'DistID.Items(0).Selected = True
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub


    'Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
    '    DataGrid1.Visible = False
    '    Button4.Visible = False
    '    Button5.Visible = False
    '    Table2.Visible = True
    'End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        '選擇轄區
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "\'" & Me.DistID.Items(i).Value & "\'"
                If DistName <> "" Then DistName += ","
                DistName += Me.DistID.Items(i).Text
            End If
        Next

        'If DistID1 <> "" Then
        '    newDistID = Mid(DistID1, 1, DistID1.Length - 1)
        '    newDistName = Mid(DistName, 1, DistName.Length - 1)
        'End If

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += "\'" & Me.TPlanID.Items(i).Value & "\'"
                If TPlanName <> "" Then TPlanName += ","
                TPlanName += Me.TPlanID.Items(i).Text
            End If
        Next

        'If TPlanID1 <> "" Then
        '    newTPlanID = Mid(TPlanID1, 1, TPlanID1.Length - 1)
        '    newTPlanIDName = Mid(TPlanName, 1, TPlanName.Length - 1)
        'End If
        Dim myValue As String = ""
        myValue &= "&Years=" & sm.UserInfo.Years
        myValue &= "&FTDate1=" & FTDate1.Text
        myValue &= "&FTDate2=" & FTDate2.Text
        myValue &= "&DistID =" & DistID1
        myValue &= "&TPlanID=" & TPlanID1
        myValue &= "&DistName =" & DistName
        myValue &= "&PlanName=" & TPlanName
        ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub


End Class
