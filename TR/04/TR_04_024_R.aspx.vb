Public Class TR_04_024_R
    Inherits AuthBasePage

    'TR_04_024_R

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        'PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not Page.IsPostBack Then
            Call sCreate1()
        End If

        DistID.Attributes("onclick") = "ClearData();"
        'TPlanID.Attributes("onclick") = "ClearData();"
        Me.chkTPlanID0.Attributes("onclick") = "ClearData();"
        Me.chkTPlanID1.Attributes("onclick") = "ClearData();"
        Me.chkTPlanIDX.Attributes("onclick") = "ClearData();"

        btnPrint1.Attributes("OnClick") = "javascript:return chk()"

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        'Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"
        '選擇全部訓練計畫
        'TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"

        If sm.UserInfo.DistID <> "000" Then
            '分署只能選自己的轄區
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            'DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

        'If Not IsPostBack Then
        '    btnExport1.Visible = False
        '    btnExport2.Visible = False

        'End If
        'DataGrid1_Detail_1.Visible = False
        'DataGrid1_Detail_2.Visible = False
        'DataGrid1_Detail_3.Visible = False
        'DataGrid1_Detail_4.Visible = False
        'DataGrid1_Detail_5.Visible = False
        'DataGrid1_Detail_6.Visible = False
        'Button3.Style("display") = "none"

        'Me.ViewState("SVID") = ""
        'If TIMS.Server_Path() = "DEMO" Then
        '    If sm.UserInfo.Years >= "2009" Then '測試機mark
        '        Me.ViewState("SVID") = TIMS.GetSVID(sm.UserInfo.TPlanID)
        '    End If
        'End If

    End Sub

    Sub sCreate1()
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        End If

        'PageControler1.Visible = False

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        'yearlist.Items.Remove(yearlist.Items.FindByValue(""))
        DistID = TIMS.Get_DistID(DistID)
        If DistID.Items.FindByValue("") Is Nothing Then
            DistID.Items.Insert(0, New ListItem("全部", ""))
        End If
        'Tcitycode = TIMS.Get_CityName(Tcitycode, TIMS.dtNothing)
        Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn)

        'center.Text = sm.UserInfo.OrgName
        'RIDValue.Value = sm.UserInfo.RID
        'PlanID.Value = sm.UserInfo.PlanID
        'OCID.Style("display") = "none"
        'Print.Visible = False
        'btnExport1.Visible = False
        'PageControler1.Visible = False
        'msg.Text = cst_NODATAMsg11
        'Button3_Click(sender, e)

        If sm.UserInfo.LID = "2" Then   '2010/05/24 改成若是委訓單位登入下列欄位就不顯示
            Year_TR.Style("display") = "none"
            DistID_TR.Style("display") = "none"
            'PlanID_TR.Style("display") = "none"
            TPlanID0_TR.Style("display") = "none"
            TPlanID1_TR.Style("display") = "none"
            TPlanIDX_TR.Style("display") = "none"
            'Check_TR.Style("display") = "none"
            'Button2.Style("display") = "none"
        Else
            'LID: 0.1.
            Year_TR.Style("display") = "inline"
            DistID_TR.Style("display") = "inline"
            'PlanID_TR.Style("display") = "inline"
            TPlanID0_TR.Style("display") = "inline"
            TPlanID1_TR.Style("display") = "inline"
            TPlanIDX_TR.Style("display") = "inline"
            'Check_TR.Style("display") = "inline"
            'Button2.Style("display") = "inline"
        End If

    End Sub

    Sub CheckData1(ByRef Errmsg As String)
        Errmsg = ""
        '檢核必填輸入值
        Dim flag1 As Boolean = False
        Dim flag2 As Boolean = False
        Dim flag3 As Boolean = False
        If yearlist.SelectedValue <> "" Then
            flag1 = True '有輸入年度
        End If
        If STDate1.Text <> "" AndAlso FTDate1.Text <> "" Then
            flag2 = True '有輸入開訓日期範圍
        End If
        If STDate2.Text <> "" AndAlso FTDate2.Text <> "" Then
            flag3 = True '有輸入結訓日期範圍
        End If
        If Not flag1 AndAlso Not flag2 AndAlso Not flag3 Then
            Errmsg &= "至少要輸入年度或開結訓範圍!!"
        End If
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= "\'" & Me.DistID.Items(i).Value & "\'"
            End If
        Next

        Dim CYears As String = ""
        If CYears = "" AndAlso yearlist.SelectedValue <> "" Then
            CYears = Val(yearlist.SelectedValue) - 1911
        End If
        If CYears = "" AndAlso STDate1.Text <> "" Then
            CYears = CDate(STDate1.Text).Year - 1911
        End If
        If CYears = "" AndAlso FTDate1.Text <> "" Then
            CYears = CDate(FTDate1.Text).Year - 1911
        End If

        'Dim etitle As String = ""
        'If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
        '    etitle = FTDate1.Text & " ~ " & FTDate2.Text
        'End If

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)
        Const cst_DefTPlanID As String = "'02','14','62','17','20','21','61','26','34','37','47','50','51','53','55','58','64','65'"

        Dim myValue As String = ""
        myValue = "p=r"
        myValue += "&CYears=" & CYears 'sm.UserInfo.Years
        myValue += "&Years=" & yearlist.SelectedValue 'sm.UserInfo.Years
        myValue += "&DistID1=" & DistID1
        If TPlanID1 <> "" Then
            myValue += "&TPlanID=" & TPlanID1
        Else
            myValue += "&TPlanID=" & Replace(cst_DefTPlanID, "'", "\'")
        End If
        myValue += "&STDate1=" & Me.STDate1.Text
        myValue += "&STDate2=" & Me.STDate2.Text
        myValue += "&FTDate1=" & Me.FTDate1.Text
        myValue += "&FTDate2=" & Me.FTDate2.Text
        'myValue += "&etitle=" & etitle
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR_04_024_R", myValue)

    End Sub
End Class