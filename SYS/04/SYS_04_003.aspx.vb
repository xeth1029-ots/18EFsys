Partial Class SYS_04_003
    Inherits AuthBasePage

    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("SYS_GLOBALVAR", objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", E1)
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", J1)
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", K1)
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", L1)
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", F1)
        Call TIMS.sUtl_SetMaxLen(dt, "ITEMVAR1", G1)

        Call TIMS.sUtl_SetMaxLen(dt, "MAWARDNO", MA1)
        Call TIMS.sUtl_SetMaxLen(dt, "MORALVARC", MVC)
        Call TIMS.sUtl_SetMaxLen(dt, "MORALVARE", MVE)
        Call TIMS.sUtl_SetMaxLen(dt, "PAWARDNO", PA1)
        Call TIMS.sUtl_SetMaxLen(dt, "PRESENTVARC", PVC)
        Call TIMS.sUtl_SetMaxLen(dt, "PRESENTVARE", SA1)
        Call TIMS.sUtl_SetMaxLen(dt, "SAFEVARC", SVC)
        Call TIMS.sUtl_SetMaxLen(dt, "SAFEVARE", SVE)
    End Sub

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    'UPDATE Sys_GlobalVar / Sys_OrgType
    '依轄區計畫
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線

        Call sUtl_PageInit1()

        If Not IsPostBack Then
            '依計畫、割區查詢可用資訊
            cCreate1()
            sSearch2()
        End If

    End Sub

    '依計畫、割區查詢可用資訊 SQL
    Sub cCreate1()
        'Button1.Enabled = False
        'If au.blnCanAdds Then Button1.Enabled = True
        'If Not au.blnCanAdds Then TIMS.Tooltip(Button1, "無新增權限", True)

        chkOpen.Attributes.Add("onclick", "chgOpen(this)")
        rblDD4.Attributes.Add("onclick", "ctrlrblDD4();")
        Button1.Attributes("onclick") = "javascript:return chkdata()"

        '(Sys_GlobalVar)
        'TPlan.SelectedValue
        'sm.UserInfo.DistID '依登入轄區
        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        Hid_DistID.Value = Convert.ToString(sm.UserInfo.DistID)
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
        ddlDISTID.Enabled = False
        Select Case sm.UserInfo.LID
            Case 0
                ddlDISTID.Enabled = True
        End Select

        TPlan = TIMS.Get_YearTPlan(TPlan, sm.UserInfo.Years, TIMS.cst_NO, objconn) '加入停用年度
        TPlan.Enabled = False
        '判斷使用權限
        If sm.UserInfo.RoleID <= 1 Then TPlan.Enabled = True
        If Not TPlan.Enabled Then TIMS.Tooltip(TPlan, "使用權限鎖定", True)
        '設定登入計劃
        Common.SetListItem(TPlan, sm.UserInfo.TPlanID)

        OrgType = TIMS.Get_OrgType(OrgType, objconn) '機構別

    End Sub

    Sub sSearch2()

        ClearData1()

        Select Case TPlan.SelectedValue
            Case "06" '在職進修訓練
                Me.checkboxList22a.Visible = True '有效
                Me.checkboxList22b.Visible = False '失效
                With Me.checkboxList22a
                    .Items.Clear()
                    .Items.Add(New ListItem("就職狀況", "JobStateType"))
                    .Items.Add(New ListItem("主要參訓身分別", "MIdentityID"))
                    .Items.Add(New ListItem("津貼類別", "SubsidyID"))
                    .Items.Add(New ListItem("津貼 身分別", "SubsidyIdentity"))
                    .Items.Add(New ListItem("受訓前任職 起迄日期", "SOfficeYM1"))
                    .Items.Add(New ListItem("受訓前薪資", "PriorWorkPay"))
                    .Items.Add(New ListItem("受訓前失業週數", "RealJobless"))
                End With
            Case Else
                '其餘計畫不顯示
                Me.checkboxList22a.Visible = False '有效
                Me.checkboxList22b.Visible = False '失效
        End Select

        Dim dt As DataTable = Search1dt()
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                'Case 1
                '    If dr("ItemVar1").ToString <> "0" Then
                '        Checkbox1.Checked = True
                '    End If
                '    A1.Text = Convert.ToString(dr("ItemVar1"))

                Select Case CInt(dr("GVID"))
                    Case 2 '成績
                        If dr("ItemVar1").ToString = "" Then
                            B1.Text = 0
                        Else
                            B1.Text = dr("ItemVar1").ToString
                        End If

                        If dr("ItemVar2").ToString = "" Then
                            B2.Text = 0
                        Else
                            B2.Text = dr("ItemVar2").ToString
                        End If
                    Case 3 '操行
                        C1.Text = Convert.ToString(dr("ItemVar1"))

                    Case 4 '出缺勤 -出缺勤警示  
                        Dim vItemVar1 As String = Convert.ToString(dr("ItemVar1"))
                        If vItemVar1 <> "0" AndAlso vItemVar1.IndexOf("/") > -1 Then
                            D1a.Text = Split(vItemVar1, "/")(0)
                            D1b.Text = Split(vItemVar1, "/")(1)
                        End If
                        Dim vItemVar2 As String = Convert.ToString(dr("ItemVar2"))
                        If vItemVar2 <> "" AndAlso vItemVar2 <> "0" AndAlso vItemVar2.IndexOf("/") > -1 Then
                            D2a.Text = Split(vItemVar2, "/")(0)
                            D2b.Text = Split(vItemVar2, "/")(1)
                        End If

                    Case 5 '在訓
                        E1.Text = Convert.ToString(dr("ItemVar1"))
                    Case 6 '甄試通知單
                        F1.Text = Convert.ToString(dr("ItemVar1"))
                    Case 7 '甄試結果
                        G1.Text = Convert.ToString(dr("ItemVar1"))
                    Case 8 '代扣所得稅
                        H1.Text = Convert.ToString(dr("ItemVar1"))
                        'Case 9 '核銷數--------------------------------------------2010/06/15 改成存另外一個table
                        '    Select Case Convert.ToString(TPlan.SelectedValue)
                        '        Case "23", "34", "41"
                        '            '設定核銷%數 核銷數
                        '            '23:訓用合一 
                        '            '34:與企業合作辦理職前訓練 
                        '            '41:推動營造業事業單位辦理職前培訓計畫
                        '            I1.Text = dr("ItemVar1").ToString
                        '            I2.Text = dr("ItemVar2").ToString
                        '    End Select---------------------------------
                        'If TPlan.SelectedValue = "23" Then
                        '    I1.Text = dr("ItemVar1").ToString
                        '    I2.Text = dr("ItemVar2").ToString
                        'End If
                    Case 10 '受訓
                        J1.Text = dr("ItemVar1").ToString
                    Case 11 '結訓
                        K1.Text = dr("ItemVar1").ToString
                    Case 12 '獎狀
                        L1.Text = dr("ItemVar1").ToString
                    Case 13 '計算成績
                        Common.SetListItem(M1, dr("ItemVar1").ToString)

                    Case 14 '操行字號
                        MA1.Text = Convert.ToString(dr("MAwardNo"))
                        MVC.Text = Convert.ToString(dr("MoralVarC")) '操行獎狀內容(中文)
                        MVE.Text = Convert.ToString(dr("MoralVarE")) '操行獎狀內容(英文)
                    Case 15 '全勤字號
                        PA1.Text = Convert.ToString(dr("PAwardNo"))
                        PVC.Text = Convert.ToString(dr("PresentVarC")) '全勤獎狀內容(中文)
                        PVE.Text = Convert.ToString(dr("PresentVarE")) '全勤獎狀內容(英文)
                    Case 16 '安全衛生教育
                        SA1.Text = Convert.ToString(dr("SAwardNo"))
                        SVC.Text = Convert.ToString(dr("SafeVarC")) '安全衛生教育獎狀內容(中文)
                        SVE.Text = Convert.ToString(dr("SafeVarE")) '安全衛生教育獎狀內容(英文)
                    Case 17 '學、術科百分比
                        If dr("ItemVar1").ToString <> "" Then
                            If IsNumeric(dr("ItemVar1").ToString) Then
                                N1.Text = CDbl(dr("ItemVar1") * 100).ToString
                            End If
                        End If
                        If dr("ItemVar2").ToString <> "" Then
                            If IsNumeric(dr("ItemVar2").ToString) Then
                                N2.Text = CDbl(dr("ItemVar2") * 100).ToString
                            End If
                        End If
                        'Case 18 'e網報名審核發送Email
                        '    If dr("ItemVar1").ToString <> "" Then
                        '        Common.SetListItem(R18, dr("ItemVar1").ToString)
                        '    Else
                        '        '預設為發送
                        '        Common.SetListItem(R18, "Y")
                        '    End If
                    Case 19 '訓練人數
                        TNum.Text = Convert.ToString(dr("ItemVar1"))

                    Case 20 '訓練小時數
                        If Convert.ToString(dr("ItemVar1")) <> "" Then
                            Thours1.Text = Convert.ToString(dr("ItemVar1"))
                        End If
                        If Convert.ToString(dr("ItemVar2")) <> "" Then
                            Thours2.Text = Convert.ToString(dr("ItemVar2"))
                        End If
                    Case 21 '報名表是否列印准考證號
                        'rdolist21.SelectedIndex = -1
                        If Convert.ToString(dr("ItemVar1")) <> "" Then
                            Common.SetListItem(rdolist21, Convert.ToString(dr("ItemVar1")))
                        End If
                    Case 22 '取消必填，學員資料維護 (SD_03_002_add.aspx)
                        If Convert.ToString(dr("ItemVar1")) <> "" Then
                            'Dim ItemA As String() = Split(Convert.ToString(dr("ItemVar1")), ",")
                            For xi As Integer = 0 To Me.checkboxList22a.Items.Count - 1
                                If Convert.ToString(dr("ItemVar1")).IndexOf(Me.checkboxList22a.Items(xi).Value) > -1 Then
                                    Me.checkboxList22a.Items(xi).Selected = True
                                End If
                                'For xj As Integer = 0 To ItemA.Length - 1
                                '    If Not ItemA(xj) Is Nothing AndAlso ItemA(xj) = Me.checkboxList22a.Items(xi).Value Then
                                '        ItemA(xj) = Nothing
                                '        Me.checkboxList22a.Items(xi).Selected = True
                                '        Exit For
                                '    End If
                                'Next
                            Next
                        End If
                    Case 23 '開放成績計算比例單位設定
                        If Convert.ToString(dr("ItemVar1")) = "Y" Then
                            chkOpen.Checked = True

                            B1.Text = ""
                            B2.Text = ""

                            B1.Enabled = False
                            B2.Enabled = False
                        Else
                            chkOpen.Checked = False
                        End If

                    Case 24 '是否可改備取名次設定
                        GVID24.Checked = False
                        If Convert.ToString(dr("ItemVar1")) = "Y" Then
                            GVID24.Checked = True
                        End If

                End Select
            Next
        End If

        Select Case Convert.ToString(TPlan.SelectedValue)
            Case "23", "34", "41"
                '設定核銷%數 核銷數
                '23:訓用合一 
                '34:與企業合作辦理職前訓練 
                '41:推動營造業事業單位辦理職前培訓計畫
                TPlan23.Visible = True
                BtnSave.Visible = False
                AddBtn.Visible = True
                OrgType.Enabled = True
                search() '取得 機構別
        End Select

        TPlan23.Visible = False '設定核銷%數 核銷數
        '設定 訓練人數設定 時數設定 (產業人才投資方案)
        Tplan28_1.Style("display") = "none"
        Tplan28_2.Style("display") = "none"

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(TPlan.SelectedValue) > -1 Then
            '設定 訓練人數設定 時數設定 (產業人才投資方案)
            Tplan28_1.Style("display") = ""
            Tplan28_2.Style("display") = ""
            TIMS.Tooltip(Tplan28_1, "產業人才投資方案 要設定")
            TIMS.Tooltip(Tplan28_2, "產業人才投資方案 要設定")
        Else
            Select Case Convert.ToString(TPlan.SelectedValue)
                Case "23", "34", "41"
                    '設定核銷%數 核銷數
                    '23:訓用合一 
                    '34:與企業合作辦理職前訓練 
                    '41:推動營造業事業單位辦理職前培訓計畫
                    TPlan23.Visible = True
                    'BtnSave.Visible = False
                    'OrgType.Enabled = True
            End Select
        End If

        Common.SetListItem(rblDD4, "1")
        If D1b.Text = "100" AndAlso D2b.Text = "100" Then
            Common.SetListItem(rblDD4, "2")
        End If
        Dim strScript As String = ""
        strScript = "<script>ctrlrblDD4();</script>"
        TIMS.RegisterStartupScript(Me, "", strScript)
    End Sub

    '清空資料列
    Sub ClearData1()
        chkOpen.Checked = False
        B1.Text = ""
        B2.Text = ""
        C1.Text = ""
        D1a.Text = ""
        D1b.Text = ""
        D2a.Text = ""
        D2b.Text = ""
        E1.Text = ""
        F1.Text = ""
        G1.Text = ""
        H1.Text = ""
        I1.Text = ""
        I2.Text = ""
        J1.Text = ""
        K1.Text = ""
        L1.Text = ""
        MA1.Text = ""
        MVC.Text = ""
        MVE.Text = ""
        PA1.Text = ""
        PVC.Text = ""
        PVE.Text = ""
        SA1.Text = ""
        SVC.Text = ""
        SVE.Text = ""
        N1.Text = ""
        N2.Text = ""
        TNum.Text = ""
        Thours1.Text = ""
        Thours2.Text = ""

        Me.M1.SelectedIndex = -1
        Me.rdolist21.SelectedIndex = -1
        'Common.SetListItem(R18, "Y")
    End Sub

    Function Search1dt() As DataTable
        Dim s_TPlanID As String = TIMS.ClearSQM(TPlan.SelectedValue)
        Dim s_DISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " SELECT * FROM SYS_GLOBALVAR"
        sql &= " WHERE TPlanID=@TPlanID AND DistID=@DistID"
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = s_TPlanID
            .Parameters.Add("DistID", SqlDbType.VarChar).Value = s_DISTID
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function


    '檢核儲存鈕( '驗証資料正確性)
    Function CheckData1(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = False '應該為FALSE:異常
        errMsg = ""
        If Convert.ToString(sm.UserInfo.DistID) = "" Then
            errMsg += "登入資訊消失，請重新登入" & vbCrLf
        End If
        If Convert.ToString(sm.UserInfo.UserID) = "" Then
            errMsg += "登入資訊消失，請重新登入" & vbCrLf
        End If
        If ddlDISTID.SelectedValue = "" Then
            errMsg += "請選擇 轄區分署" & vbCrLf
        End If
        If TPlan.SelectedValue = "" Then
            errMsg += "請選擇 訓練計畫" & vbCrLf
        End If
        If errMsg <> "" Then Return rst

        Dim update_flag17 As Boolean = True '學、術科百分比
        '***若有加號的需要請注意「存取 HEAD」 的設定
        Dim j As Integer = 24 'GVID參數序號_最大值
        Me.ViewState("N1_flag") = False
        Me.ViewState("N2_flag") = False

        '驗証資料正確性---- Start
        N1.Text = TIMS.ClearSQM(N1.Text)
        If N1.Text = "" Then
            Me.ViewState("N1_flag") = True
        End If
        If N1.Text <> "" AndAlso IsNumeric(N1.Text) Then
            Me.ViewState("N1_flag") = True
        End If

        N2.Text = TIMS.ClearSQM(N2.Text)
        If N2.Text = "" Then
            Me.ViewState("N2_flag") = True
        End If
        If N2.Text <> "" AndAlso IsNumeric(N2.Text) Then
            Me.ViewState("N2_flag") = True
        End If

        Select Case False
            Case Me.ViewState("N1_flag")
                errMsg &= "學、術科百分比，學科百分比只能填數字或空白" & vbCrLf
                Return rst
            Case Me.ViewState("N2_flag")
                errMsg &= "學、術科百分比，學科百分比只能填數字或空白" & vbCrLf
                Return rst
        End Select
        '驗証資料正確性---- End

        If N1.Text = "" OrElse N2.Text = "" Then
            update_flag17 = False
        End If

        B1.Text = TIMS.ClearSQM(B1.Text)
        B2.Text = TIMS.ClearSQM(B2.Text)
        C1.Text = TIMS.ClearSQM(C1.Text)
        D1a.Text = TIMS.ClearSQM(D1a.Text)
        D1b.Text = TIMS.ClearSQM(D1b.Text)
        D2a.Text = TIMS.ClearSQM(D2a.Text)
        D2b.Text = TIMS.ClearSQM(D2b.Text)
        E1.Text = TIMS.ClearSQM(E1.Text)
        F1.Text = TIMS.ClearSQM(F1.Text)
        G1.Text = TIMS.ClearSQM(G1.Text)
        H1.Text = TIMS.ClearSQM(H1.Text)
        J1.Text = TIMS.ClearSQM(J1.Text)
        K1.Text = TIMS.ClearSQM(K1.Text)
        L1.Text = TIMS.ClearSQM(L1.Text)
        Select Case rblDD4.SelectedValue
            Case "1" '依比例
                If D1a.Text <> "" Then
                    If Not TIMS.IsNumeric1(D1a.Text) OrElse Not TIMS.IsNumeric1(D1b.Text) Then
                        errMsg &= "出缺勤警示 第一次缺課警告：應輸入數字格式!" & vbCrLf
                        Return rst
                    End If
                End If
                If D2a.Text <> "" Then
                    If Not TIMS.IsNumeric1(D2a.Text) OrElse Not TIMS.IsNumeric1(D2b.Text) Then
                        errMsg &= "出缺勤警示 第二次缺課警告：應輸入數字格式!" & vbCrLf
                        Return rst
                    End If
                End If

            Case Else '"2"'依百分比
                If D1a.Text <> "" Then
                    If Not TIMS.IsNumeric1(D1a.Text) Then
                        errMsg &= "出缺勤警示 第一次缺課警告：應輸入數字格式!" & vbCrLf
                        Return rst
                    End If
                End If
                If D2a.Text <> "" Then
                    If Not TIMS.IsNumeric1(D2a.Text) Then
                        errMsg &= "出缺勤警示 第二次缺課警告：應輸入數字格式!" & vbCrLf
                        Return rst
                    End If
                End If

        End Select

        'M1.SelectedValue'計算成績方式 
        Select Case M1.SelectedValue
            Case "1", "2"
            Case Else
                errMsg &= "請選擇 計算成績方式!" & vbCrLf
                Return rst
        End Select
        '操行字號
        MA1.Text = TIMS.ClearSQM(MA1.Text)
        MVC.Text = TIMS.ClearSQM(MVC.Text)
        MVE.Text = TIMS.ClearSQM(MVE.Text)
        '全勤字號
        PA1.Text = TIMS.ClearSQM(PA1.Text)
        PVC.Text = TIMS.ClearSQM(PVC.Text)
        PVE.Text = TIMS.ClearSQM(PVE.Text)
        '安全衛生教育
        SA1.Text = TIMS.ClearSQM(SA1.Text)
        SVC.Text = TIMS.ClearSQM(SVC.Text)
        SVE.Text = TIMS.ClearSQM(SVE.Text)

        '20080617 Andy 
        'Case 17 '學、術科百分比
        If errMsg <> "" Then Return rst
        If N1.Text <> "" AndAlso N2.Text <> "" Then
            If (CDbl(N1.Text) + CDbl(N2.Text)) <> 100 Then
                errMsg &= "學科百分比 +術科百分比 不等於100!" & vbCrLf
                Return rst
            End If
        End If

        TNum.Text = TIMS.ClearSQM(TNum.Text)
        Thours1.Text = TIMS.ClearSQM(Thours1.Text)
        Thours2.Text = TIMS.ClearSQM(Thours2.Text)

        If errMsg = "" Then rst = True
        Return rst
    End Function

    '儲存
    Sub SaveData1()
        '***若有加號的需要請注意「存取 HEAD」 的設定
        Dim s_TPlanID As String = TIMS.ClearSQM(TPlan.SelectedValue)
        Dim s_DISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)

        Dim j As Integer = 24 'GVID參數序號_最大值
        Dim sql As String
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable
        Dim dr As DataRow
        sql = ""
        sql &= " SELECT * FROM Sys_GlobalVar"
        sql &= " WHERE 1=1"
        sql &= " AND TPlanID='" & s_TPlanID & "'"
        sql &= " AND DistID='" & s_DISTID & "'" 's_DISTID & "'"
        sql &= " ORDER BY GVID"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        Dim update_flag17 As Boolean = True '學、術科百分比
        If N1.Text = "" OrElse N2.Text = "" Then
            update_flag17 = False
        End If

        For i As Integer = 2 To j
            '存取 HEAD
            Select Case i
                Case 9 '核銷%
                    '排除9 :核銷%
                    dr = Nothing '不存取
                    'Exit For
                Case 17 '學、術科百分比
                    If Not update_flag17 Then
                        dr = Nothing '不存取
                        'Exit For
                    Else
                        '存取
                        If dt.Select("GVID=" & i).Length = 0 Then
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("GVID") = i
                            dr("TPlanID") = s_TPlanID
                            dr("DistID") = s_DISTID
                        Else
                            dr = dt.Select("GVID=" & i)(0)
                        End If
                    End If
                Case 18 'e網報名審核發送Email
                    '排除9
                    dr = Nothing '不存取
                    'Exit For

                Case 19, 20 '產投訓練人數  '產投時數
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(s_TPlanID) > -1 Then
                        '存取
                        If dt.Select("GVID=" & i).Length = 0 Then
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("GVID") = i
                            dr("TPlanID") = s_TPlanID
                            dr("DistID") = s_DISTID
                        Else
                            dr = dt.Select("GVID=" & i)(0)
                        End If
                    Else
                        dr = Nothing '不存取
                        'Exit For
                    End If
                Case Else
                    '存取
                    If dt.Select("GVID=" & i).Length = 0 Then
                        dr = dt.NewRow()
                        dt.Rows.Add(dr)
                        dr("GVID") = i
                        dr("TPlanID") = s_TPlanID
                        dr("DistID") = s_DISTID
                    Else
                        dr = dt.Select("GVID=" & i)(0)
                    End If

            End Select

            '存取 BODY
            'dr 有值表示可存取
            If Not dr Is Nothing Then
                Select Case i
                    Case 2 '成績
                        dr("ItemVar1") = If(B1.Text = "", "0", B1.Text)
                        dr("ItemVar2") = If(B2.Text = "", "0", B2.Text)
                    Case 3 '操行
                        dr("ItemVar1") = If(C1.Text = "", "0", C1.Text)
                        dr("ItemVar2") = Convert.DBNull
                    Case 4 '出缺勤
                        D1a.Text = TIMS.ClearSQM(D1a.Text)
                        D1b.Text = TIMS.ClearSQM(D1b.Text)
                        D2a.Text = TIMS.ClearSQM(D2a.Text)
                        D2b.Text = TIMS.ClearSQM(D2b.Text)
                        Select Case rblDD4.SelectedValue
                            Case "1" '依比例
                                dr("ItemVar1") = "0"
                                If D1a.Text <> "" OrElse D1b.Text <> "" Then
                                    dr("ItemVar1") = D1a.Text & "/" & D1b.Text
                                End If
                                dr("ItemVar2") = "0"
                                If D2a.Text <> "" OrElse D2b.Text <> "" Then
                                    dr("ItemVar2") = D2a.Text & "/" & D2b.Text
                                End If
                            Case Else '"2"'依百分比
                                dr("ItemVar1") = "0"
                                If D1a.Text <> "" Then
                                    dr("ItemVar1") = D1a.Text & "/100"
                                End If
                                dr("ItemVar2") = "0"
                                If D2a.Text <> "" Then
                                    dr("ItemVar2") = D2a.Text & "/100"
                                End If
                        End Select
                    Case 5 '在訓
                        dr("ItemVar1") = If(E1.Text = "", " ", E1.Text)
                        dr("ItemVar2") = Convert.DBNull
                    Case 6 '甄試通知單
                        dr("ItemVar1") = If(F1.Text = "", " ", F1.Text)
                        dr("ItemVar2") = Convert.DBNull
                    Case 7 '甄試結果
                        dr("ItemVar1") = If(G1.Text = "", " ", G1.Text)
                        dr("ItemVar2") = Convert.DBNull
                    Case 8 '代扣所得稅
                        dr("ItemVar1") = If(H1.Text = "", " ", H1.Text)
                        dr("ItemVar2") = Convert.DBNull
                        'Case 9                -----------------------2010/06/15 改成存在Sys_OrgType table 
                        '    Select Case Convert.ToString(s_TPlanID)
                        '        Case "23", "34", "41"
                        '            '設定核銷%數 核銷數
                        '            '23:訓用合一 
                        '            '34:與企業合作辦理職前訓練 
                        '            '41:推動營造業事業單位辦理職前培訓計畫
                        '            dr("ItemVar1") = I1.Text
                        '            dr("ItemVar2") = I2.Text
                        '    End Select
                        'If s_TPlanID = "23" Then
                        '    dr("ItemVar1") = I1.Text
                        '    dr("ItemVar2") = I2.Text
                        'End If
                    Case 10 '受訓
                        dr("ItemVar1") = If(J1.Text = "", " ", J1.Text) 'J1.Text
                        dr("ItemVar2") = Convert.DBNull
                    Case 11 '結訓
                        dr("ItemVar1") = If(K1.Text = "", " ", K1.Text) 'K1.Text
                        dr("ItemVar2") = Convert.DBNull
                    Case 12 '獎狀
                        dr("ItemVar1") = If(L1.Text = "", " ", L1.Text) 'L1.Text
                        dr("ItemVar2") = Convert.DBNull
                    Case 13 '計算成績
                        dr("ItemVar1") = " "
                        If M1.SelectedValue <> "" Then
                            dr("ItemVar1") = M1.SelectedValue
                        End If
                        dr("ItemVar2") = Convert.DBNull
                    Case 14  '操行字號
                        'dr("ItemVar1") = ""
                        dr("ItemVar1") = " "
                        dr("ItemVar2") = Convert.DBNull
                        dr("MAwardNo") = TIMS.GetValue1(MA1.Text)
                        dr("MoralVarC") = TIMS.GetValue1(MVC.Text)
                        dr("MoralVarE") = TIMS.GetValue1(MVE.Text)
                    Case 15 '全勤字號
                        'dr("ItemVar1") = ""
                        dr("ItemVar1") = " "
                        dr("ItemVar2") = Convert.DBNull
                        dr("PAwardNo") = TIMS.GetValue1(PA1.Text)
                        dr("PresentVarC") = TIMS.GetValue1(PVC.Text)
                        dr("PresentVarE") = TIMS.GetValue1(PVE.Text)
                    Case 16 '安全衛生教育
                        'dr("ItemVar1") = ""
                        dr("ItemVar1") = " "
                        dr("ItemVar2") = Convert.DBNull
                        dr("SAwardNo") = TIMS.GetValue1(SA1.Text)
                        dr("SafeVarC") = TIMS.GetValue1(SVC.Text)
                        dr("SafeVarE") = TIMS.GetValue1(SVE.Text)

                        '20080617 Andy 
                    Case 17 '學、術科百分比
                        '判斷後決定  update_flag17 
                        dr("ItemVar1") = "0" 'Convert.DBNull
                        If N1.Text <> "" Then
                            dr("ItemVar1") = (CDbl(N1.Text) / 100).ToString()
                        End If
                        dr("ItemVar2") = "0" ' Convert.DBNull
                        If N2.Text <> "" Then
                            dr("ItemVar2") = (CDbl(N2.Text) / 100).ToString()
                        End If
                        'Case 18 'e網報名審核發送Email
                        '    dr("ItemVar1") = R18.SelectedValue
                    Case 19 '產投訓練人數
                        TNum.Text = TIMS.ClearSQM(TNum.Text)
                        dr("ItemVar1") = If(TNum.Text <> "", TNum.Text, "0")
                        dr("ItemVar2") = Convert.DBNull

                    Case 20 '產投時數
                        'dr("ItemVar1") = ""
                        dr("ItemVar1") = TIMS.ClearSQM(Thours1.Text) '非學分班訓練時數上限
                        dr("ItemVar2") = TIMS.ClearSQM(Thours2.Text) '非學分班訓練時數下限

                    Case 21 '報名表是否列印准考證號
                        dr("ItemVar1") = rdolist21.SelectedValue
                        dr("ItemVar2") = Convert.DBNull

                    Case 22 '取消必填，學員資料維護 (SD_03_002_add.aspx)
                        Dim sItemA As String = ""
                        sItemA = ""
                        For xi As Integer = 0 To Me.checkboxList22a.Items.Count - 1
                            If Me.checkboxList22a.Items(xi).Selected = True Then
                                If sItemA <> "" Then sItemA += ","
                                sItemA += Me.checkboxList22a.Items(xi).Value
                            End If
                        Next
                        dr("ItemVar1") = If(sItemA = "", " ", sItemA)

                    Case 23 '開放成績計算比例單位設定
                        dr("ItemVar1") = If(chkOpen.Checked = True, "Y", "N")

                    Case 24 '是否可改備取名次設定
                        dr("ItemVar1") = If(GVID24.Checked = True, "Y", "N")
                End Select

                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
            End If
            'DbAccess.UpdateDataTable(dt, da)
        Next

        DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功!!")

        sSearch2()
    End Sub

    '儲存按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim errMsg As String = ""
        Select Case sm.UserInfo.LID '= lid  '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
            Case Else
                If Not CheckData1(errMsg) Then
                    Common.MessageBox(Me, errMsg)
                    Exit Sub
                End If
        End Select

        Call SaveData1()
    End Sub



#Region "Sys_OrgType"

    '取得 機構別
    Sub search()
        Dim s_TPlanID As String = TIMS.ClearSQM(TPlan.SelectedValue)
        Dim s_DISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)

        Dim dt As DataTable = Nothing
        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " select so.DistID,so.TPlanID,so.OrgTypeID,ItemVar1,ItemVar2" & vbCrLf
        Sql += " ,so.ItemVar1+' %' as ItemVar1B" & vbCrLf
        Sql += " ,so.ItemVar2+' %' as ItemVar2B" & vbCrLf
        Sql += " ,ko.*" & vbCrLf
        Sql += " from Sys_OrgType so " & vbCrLf
        Sql += " join Key_OrgType ko on so.OrgTypeID =ko.OrgTypeID " & vbCrLf
        Sql += " where 1=1" & vbCrLf
        Sql += " and so.TPlanid = '" & s_TPlanID & "'" & vbCrLf
        Sql += " and so.DistID ='" & s_DISTID & "' " & vbCrLf

        dt = DbAccess.GetDataTable(Sql, objconn)
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '機構別命令cmd
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sqlstr As String
        Dim Array1 As Array = Split(e.CommandArgument, ",")
        Select Case e.CommandName
            Case "Edit"
                OrgType.SelectedValue = Array1(2) '機構別
                OrgType.Enabled = False
                I1.Text = Array1(3) '一般身分
                I2.Text = Array1(4) '特殊對象
                BtnSave.Visible = True
                AddBtn.Visible = False
            Case "Del"
                Try
                    sqlstr = " delete Sys_OrgType "
                    sqlstr += " where DistID = '" & Array1(0) & "' and TPlanID = '" & Array1(1) & "' and OrgTypeID = '" & Array1(2) & "' "
                    DbAccess.ExecuteNonQuery(sqlstr, objconn)
                Catch ex As Exception
                    Common.MessageBox(Me, ex.ToString)
                End Try
                Common.MessageBox(Me, "刪除成功")
                search()
        End Select
    End Sub

    '顯示機構別List
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Dim btnEdit As Button = e.Item.FindControl("BtnEdit")
        Dim BtnDel As Button = e.Item.FindControl("BtnDel")

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                btnEdit.CommandArgument = drv("DistID") & "," & drv("TPlanID") & "," & drv("OrgTypeID") & "," & drv("ItemVar1") & "," & drv("ItemVar2")
                BtnDel.CommandArgument = drv("DistID") & "," & drv("TPlanID") & "," & drv("OrgTypeID")
                BtnDel.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    '核銷%數 新增  
    Private Sub AddBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddBtn.Click
        Dim errMsg As String = ""
        '輸入資料沒有錯誤
        If Not checkdata(errMsg) Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        Dim s_TPlanID As String = TIMS.ClearSQM(TPlan.SelectedValue)
        Dim s_DISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)

        Dim dr As DataRow = Nothing
        Dim sql2 As String = ""
        sql2 = "select * from Sys_OrgType where DistID ='" & s_DISTID & "' and TPlanID = '" & s_TPlanID & "' and OrgTypeID = '" & OrgType.SelectedValue & "' "
        dr = DbAccess.GetOneRow(sql2, objconn)
        If dr IsNot Nothing Then '若有資料
            Common.MessageBox(Me, "此機構別的設定己存在,請勿重複設定")
            Exit Sub
        End If

        '若沒有資料
        Dim i_sql As String = ""
        i_sql = " insert into Sys_OrgType (DistID,TPlanID,OrgTypeID,ItemVar1,ItemVar2,ModifyAcct,ModifyDate)"
        i_sql += " values (@DistID,@TPlanID,@OrgTypeID,@ItemVar1,@ItemVar2,@ModifyAcct,GETDATE())"

        I1.Text = TIMS.ClearSQM(I1.Text)
        I2.Text = TIMS.ClearSQM(I2.Text)
        Dim i_parms As New Hashtable
        i_parms.Add("DistID", s_DISTID)
        i_parms.Add("TPlanID", s_TPlanID)
        i_parms.Add("OrgTypeID", OrgType.SelectedValue)
        i_parms.Add("ItemVar1", I1.Text)
        i_parms.Add("ItemVar2", I2.Text)
        i_parms.Add("ModifyAcct", sm.UserInfo.UserID)
        DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)

        Common.MessageBox(Me, "新增成功")
        OrgType.SelectedIndex = 0
        OrgType.Enabled = True
        I1.Text = ""
        I2.Text = ""
        search() '取得 機構別
    End Sub

    '核銷%數 存檔
    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        Dim errMsg As String = ""
        '輸入資料沒有錯誤
        If Not checkdata(errMsg) Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        Dim s_TPlanID As String = TIMS.ClearSQM(TPlan.SelectedValue)
        Dim s_DISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        I1.Text = TIMS.ClearSQM(I1.Text)
        I2.Text = TIMS.ClearSQM(I2.Text)

        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " update Sys_OrgType"
        sqlstr &= " set ItemVar1 =@ItemVar1"
        sqlstr &= " ,ItemVar2 =@ItemVar2"
        sqlstr += " where 1=1"
        sqlstr &= " and TPlanID = '" & s_TPlanID & "'"
        sqlstr &= " and DistID = '" & s_DISTID & "'"
        sqlstr &= " and OrgTypeID = '" & OrgType.SelectedValue & "'"

        Dim parms As New Hashtable
        parms.Add("ItemVar1", I1.Text)
        parms.Add("ItemVar2", I2.Text)
        DbAccess.ExecuteNonQuery(sqlstr, objconn, parms)

        Common.MessageBox(Me, "修改成功")
        OrgType.SelectedIndex = 0
        OrgType.Enabled = True
        I1.Text = ""
        I2.Text = ""
        BtnSave.Visible = False
        AddBtn.Visible = True

        search()
    End Sub

    '核銷%數 檢查輸入資料
    Function checkdata(ByRef msg As String) As Boolean '檢查輸入資料
        Dim Rst As Boolean = True '輸入資料都沒錯
        'Dim msg As String = ""
        I1.Text = TIMS.ClearSQM(I1.Text)
        I2.Text = TIMS.ClearSQM(I2.Text)
        If Convert.ToString(sm.UserInfo.DistID) = "" Then
            msg += "登入資訊消失，請重新登入" & vbCrLf
        End If
        If Convert.ToString(sm.UserInfo.UserID) = "" Then
            msg += "登入資訊消失，請重新登入" & vbCrLf
        End If

        If OrgType.SelectedIndex = 0 Then
            msg += "請選擇機構別" & vbCrLf
        End If
        If I1.Text = "" Then
            msg += "請選擇一般身分" & vbCrLf
        Else
            If Not System.Text.RegularExpressions.Regex.IsMatch(I1.Text.ToString, "[0-9]") Then
                msg += "一般身分欄位請輸入數字" & vbCrLf
            End If
        End If
        If I2.Text = "" Then
            msg += "請選擇特定對象" & vbCrLf
        Else
            If Not System.Text.RegularExpressions.Regex.IsMatch(I2.Text.ToString, "[0-9]") Then
                msg += "特定對象欄位請輸入數字" & vbCrLf
            End If
        End If

        If msg <> "" Then
            Rst = False '輸入資料有誤
            'Common.MessageBox(Me, msg)
        End If
        Return Rst
    End Function
#End Region

    ''' <summary>
    ''' 選擇轄區
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub ddlDISTID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlDISTID.SelectedIndexChanged
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Hid_DistID.Value = If(v_ddlDISTID <> "", v_ddlDISTID, sm.UserInfo.DistID)
        If v_ddlDISTID = "" Then Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
        If TPlan.SelectedIndex <> 0 Then
            '有選擇計畫
            sSearch2()
        Else
            '沒有選 清除動作
            ClearData1()
        End If
    End Sub

    ''' <summary>
    ''' 選擇計畫
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TPlan.SelectedIndexChanged
        If TPlan.SelectedIndex <> 0 Then
            '有選擇計畫
            sSearch2()
        Else
            '沒有選 清除動作
            ClearData1()
        End If
    End Sub

End Class

