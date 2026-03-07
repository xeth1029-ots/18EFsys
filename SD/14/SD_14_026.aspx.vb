Partial Class SD_14_026
    Inherits AuthBasePage

    'SD_14_026*.jrxml
    Const cst_printFN1_G As String = "SD_14_026G"
    Const cst_printFN1_W As String = "SD_14_026W"
    'Const cst_printFN2 As String = "SD_14_026_54" 
    '54:充電起飛計畫 SD_14_026_54*.jrxml
    Const cst_printFN54_G As String = "SD_14_026_54G"
    Const cst_printFN54_W As String = "SD_14_026_54W"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '充飛無【申請階段】，僅保留【訓練機構】選項
        tr_rbl_AppStage_TP28.Visible = If(TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, False, True)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
            If tr_rbl_AppStage_TP28.Visible Then
                rbl_AppStage = TIMS.Get_APPSTAGE2(rbl_AppStage)
                TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, rbl_AppStage) 'Common.SetListItem(rbl_AppStage, "3")
            End If

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '列印
        'btnPrint1.Attributes("onclick") = "return CheckPrint();"
    End Sub

    Sub Utl_Print1()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim v_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        If v_DistID = "" Then v_DistID = sm.UserInfo.DistID
        If v_DistID = "000" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim v_OrgID As String = TIMS.Get_OrgID(RIDValue.Value, objconn)
        If v_OrgID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim v_OrgKind As String = TIMS.Get_OrgKind2(v_OrgID, TIMS.c_ORGID, objconn)
        If v_OrgKind = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim prtstrVal As String = ""
        '新增【送件檢核表】 充電起飛計畫(補助在職勞工及自營作業者參訓)
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            prtstrVal = ""
            prtstrVal &= "&TPLANID=" & sm.UserInfo.TPlanID
            prtstrVal &= "&YEARS=" & sm.UserInfo.Years
            prtstrVal &= "&DISTID=" & v_DistID
            prtstrVal &= "&RID=" & RIDValue.Value
            Dim v_FileName54B As String = ""
            Select Case v_OrgKind
                Case "G"
                    v_FileName54B = cst_printFN54_G 'G/W
                Case "W"
                    v_FileName54B = cst_printFN54_W 'G/W
            End Select
            If v_FileName54B = "" Then
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return ' Exit Sub
            End If
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, v_FileName54B, prtstrVal)
            Return
        End If

        'check VAL_rbl_AppStage '有啟動才檢核/塞值
        Dim v_AppStage As String = TIMS.GetListValue(rbl_AppStage) 'TIMS.ClearSQM(rbl_AppStage.SelectedValue) '申請階段
        If v_AppStage = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return '  Exit Sub
        End If
        Select Case v_AppStage
            Case "1", "2", "3", "4"
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return ' Exit Sub
        End Select
        If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage

        Dim v_FileNameA As String = "" 'G/W
        Select Case v_OrgKind
            Case "G"
                v_FileNameA = cst_printFN1_G 'G/W
            Case "W"
                v_FileNameA = cst_printFN1_W 'G/W
        End Select
        If v_FileNameA = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return ' Exit Sub
        End If

        prtstrVal = ""
        prtstrVal &= "&TPLANID=" & sm.UserInfo.TPlanID
        prtstrVal &= "&YEARS=" & sm.UserInfo.Years
        prtstrVal &= "&DISTID=" & v_DistID
        prtstrVal &= "&RID=" & RIDValue.Value
        '依申請階段 
        prtstrVal &= "&APPSTAGE=" & v_AppStage

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, v_FileNameA, prtstrVal)
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Utl_Print1()
    End Sub
End Class