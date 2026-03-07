Partial Class SD_02_006_R
    Inherits AuthBasePage

    '列印、設定通知單內容， 正取內容、備取內容、未錄取內容儲存  
    'Maintest_result '非缺考
    'Maintest_result_1 '缺考
    Const cst_printFN1 As String = "Maintest_result"
    Const cst_printFN2 As String = "Maintest_result_1"

    Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

#Region "FUNC"

    Sub Get_SysOrgVar1(ByVal vRID As String, ByVal vTPlanID As String)
        Me.msg.Visible = True
        Me.button1.Enabled = False '系統參數未設定
        Dim parms As Hashtable = New Hashtable()
        Dim Str As String = " SELECT * FROM Sys_OrgVar WHERE RID = @RID AND TPlanID = @TPlanID "
        parms.Clear()
        parms.Add("RID", vRID)
        parms.Add("TPlanID", vTPlanID)
        Dim dt As DataTable = DbAccess.GetDataTable(Str, objconn, parms)

        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow = dt.Rows(0)
            If Not IsDBNull(dr("ItemVar_1")) Then
                Me.msg.Visible = False
                Me.button1.Enabled = True '系統參數設定
            End If
        End If
        If Not Me.button1.Enabled Then TIMS.Tooltip(button1, "系統參數未設定!!")
        btnExport1.Enabled = Me.button1.Enabled
    End Sub

#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        '(直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        'Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True

#Region "(No Use)"

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then Trwork2013a.Visible = True
        'End If

#End Region

        TIMS.ShowHistoryClass(Me, historytable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", True)
        If historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            TIMS.Tooltip(button1, "甄試結果通知單內容by機構，列印請先確認甄試結果通知單內容是否已儲存!!")
            SelResult = TIMS.Get_SelResult(SelResult, 1, objconn)
            button1.Attributes("onclick") = "javascript:return ReportPrint();"
            btnExport1.Attributes("onclick") = "javascript:return ReportPrint();"
            Me.table11.Style("display") = "none"
            Me.button5.Visible = False
            Call Get_SysOrgVar1(sm.UserInfo.RID, sm.UserInfo.TPlanID)
        End If
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim M1(5) As Integer
        For i As Integer = 0 To mailtype1.Items.Count - 1
            M1(i) = If(mailtype1.Items.Item(i).Selected, 1, 0)
        Next

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&OCID1=" & OCIDValue1.Value
        MyValue &= "&DistID=" & sm.UserInfo.DistID
        MyValue &= "&Mailtype1=" & M1(0).ToString
        MyValue &= "&Mailtype2=" & M1(1).ToString
        MyValue &= "&Mailtype3=" & M1(2).ToString
        MyValue &= "&Mailtype4=" & M1(3).ToString
        MyValue &= "&Mailtype5=" & M1(4).ToString

        '就服單位協助報名
        Select Case rblEnterPathW.SelectedValue
            Case "A"
            Case "Y", "N"
                MyValue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
        End Select

        'SelResultID Name 
        '01 正取 
        '02 備取 
        '03 未錄取 
        '04 缺考
        Dim v_SelResult As String = TIMS.GetListValue(SelResult)
        Select Case v_SelResult 'SelResult.SelectedValue
            Case "04"
                '缺考
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
            Case Else
                '非缺考
                MyValue &= "&SelResultID=" & v_SelResult 'SelResult.SelectedValue
                Select Case v_SelResult'SelResult.SelectedValue
                    Case "02", "03"
                        '不顯示報到日期
                        MyValue &= "&ChkDateList=N"
                    Case Else
                        '顯示報到日期
                        MyValue &= "&ChkDateList=Y"
                End Select
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
        End Select
    End Sub

    '設定通知單內容
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button2.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Me.table11.Style("display") = "inline"
        Me.button5.Visible = True
        Me.msg.Visible = False

        Dim Str1 As String = ""
        Dim dt1 As DataTable = Nothing
        Dim dr1 As DataRow = Nothing
        Dim parms As Hashtable = New Hashtable()

        '將參數設定-甄試結果內容代入----start
        Dim Str As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Str = " SELECT * FROM Sys_OrgVar WHERE RID = @RID AND TPlanID = @TPlanID "
        parms.Clear()
        parms.Add("RID", sm.UserInfo.RID)
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        dt = DbAccess.GetDataTable(Str, objconn, parms)

        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            If Not IsDBNull(dr("ItemVar_1")) Then
                Me.itemvar_1.Text = Convert.ToString(dr("ItemVar_1"))
            Else
                Str1 = " SELECT * FROM Sys_GlobalVar WHERE GVID = '7' AND DistID = @DistID AND TPlanID = @TPlanID "
                parms.Clear()
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
                dt1 = DbAccess.GetDataTable(Str1, objconn, parms)
                If dt1.Rows.Count <> 0 Then
                    dr1 = dt1.Rows(0)
                    If Not IsDBNull(dr1("ItemVar1")) Then Me.itemvar_1.Text = Convert.ToString(dr1("ItemVar1"))
                End If
            End If
            If Not IsDBNull(dr("ItemVar_2")) Then Me.itemvar_2.Text = Convert.ToString(dr("ItemVar_2"))
            If Not IsDBNull(dr("ItemVar_3")) Then Me.itemvar_3.Text = Convert.ToString(dr("ItemVar_3"))
        Else
            Str = " SELECT * FROM Sys_OrgVar WHERE RID = @RID AND TPlanID IS NULL "
            parms.Clear()
            parms.Add("RID", sm.UserInfo.RID)
            dt = DbAccess.GetDataTable(Str, objconn, parms)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                If Not IsDBNull(dr("ItemVar_1")) Then
                    Me.itemvar_1.Text = Convert.ToString(dr("ItemVar_1"))
                Else
                    Str1 = " SELECT * FROM Sys_GlobalVar WHERE GVID = '7' AND DistID = @DistID AND TPlanID = @TPlanID "
                    parms.Clear()
                    parms.Add("DistID", sm.UserInfo.DistID)
                    parms.Add("TPlanID", sm.UserInfo.TPlanID)
                    dt1 = DbAccess.GetDataTable(Str1, objconn, parms)
                    If dt1.Rows.Count <> 0 Then
                        dr1 = dt1.Rows(0)
                        If Not IsDBNull(dr1("ItemVar1")) Then Me.itemvar_1.Text = Convert.ToString(dr1("ItemVar1"))
                    End If
                End If
                If Not IsDBNull(dr("ItemVar_2")) Then Me.itemvar_2.Text = Convert.ToString(dr("ItemVar_2"))
                If Not IsDBNull(dr("ItemVar_3")) Then Me.itemvar_3.Text = Convert.ToString(dr("ItemVar_3"))
            Else
                Str1 = " SELECT * FROM Sys_GlobalVar WHERE GVID = '7' AND DistID = @DistID AND TPlanID = @TPlanID "
                parms.Clear()
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
                dt1 = DbAccess.GetDataTable(Str1, objconn, parms)
                If dt1.Rows.Count <> 0 Then
                    dr1 = dt1.Rows(0)
                    If Not IsDBNull(dr1("ItemVar1")) Then Me.itemvar_1.Text = Convert.ToString(dr1("ItemVar1"))
                End If
            End If
        End If
        '---end
    End Sub

    '儲存
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Me.itemvar_1.Text = TIMS.ClearSQM(Me.itemvar_1.Text)
        'Me.itemvar_2.Text = TIMS.ClearSQM(Me.itemvar_2.Text)
        'Me.itemvar_3.Text = TIMS.ClearSQM(Me.itemvar_3.Text)
        Dim sItemVar_1 As String = Me.itemvar_1.Text
        Dim sItemVar_2 As String = Me.itemvar_2.Text
        Dim sItemVar_3 As String = Me.itemvar_3.Text
        Dim sErrmsg As String = ""
        If sItemVar_1.Length > 512 Then sErrmsg &= "正取內容 資料太長超過系統範圍(512)!" & vbCrLf
        If sItemVar_2.Length > 512 Then sErrmsg &= "備取內容 資料太長超過系統範圍(512)!" & vbCrLf
        If sItemVar_3.Length > 512 Then sErrmsg &= "未錄取內容 資料太長超過系統範圍(512)!" & vbCrLf
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If
        'If sItemVar_1.Length > 512 Then sItemVar_1 = sItemVar_1.Substring(0, 512)
        'If sItemVar_2.Length > 512 Then sItemVar_2 = sItemVar_2.Substring(0, 512)
        'If sItemVar_3.Length > 512 Then sItemVar_3 = sItemVar_3.Substring(0, 512)
        'Dim oAd As New SqlDataAdapter
        'Dim dr As DataRow = Nothing
        'Dim dt As DataTable = Nothing

#Region "(No Use)"

        'Dim strSql1 As String = ""
        'strSql1 = "SELECT * FROM SYS_ORGVAR WHERE RID = '" & sm.UserInfo.RID & "' and TPlanID ='" & sm.UserInfo.TPlanID & "'"
        'Dim dr As DataRow = Nothing
        'Dim oAd As New SqlDataAdapter
        'Dim dt As DataTable = DbAccess.GetDataTable(strSql1, oAd, objconn)
        'If dt.Rows.Count > 0 Then '先判斷之前是否有存Planid 如果有就update
        '    dr = dt.Rows(0)
        'Else
        '    Dim iSOID As Integer = DbAccess.GetNewId(objconn, "SYS_ORGVAR_SOID_SEQ,SYS_ORGVAR,SOID")
        '    dr = dt.NewRow
        '    dt.Rows.Add(dr)
        '    dr("SOID") = iSOID
        '    dr("RID") = sm.UserInfo.RID
        '    dr("TPlanID") = sm.UserInfo.TPlanID
        'End If
        'If Me.itemvar_1.Text <> "" Then dr("ItemVar_1") = Me.itemvar_1.Text '(else NULL)
        'If Me.itemvar_2.Text <> "" Then dr("ItemVar_2") = Me.itemvar_2.Text '(else NULL)
        'If Me.itemvar_3.Text <> "" Then dr("ItemVar_3") = Me.itemvar_3.Text '(else NULL)
        ''dr("TPlanID") = sm.UserInfo.TPlanID
        'dr("ModifyAcct") = sm.UserInfo.UserID
        'dr("ModifyDate") = Now()
        'DbAccess.UpdateDataTable(dt, oAd)

#End Region

        Dim strSql1 As String = ""
        Dim dt As DataTable = Nothing

        '查詢之前是否有存在Planid
        strSql1 = " SELECT * FROM SYS_ORGVAR WHERE RID = @RID AND TPlanID = @TPlanID "
        Dim sCmd As New SqlCommand(strSql1, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
            dt = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
        End With

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '修改
            strSql1 = " UPDATE SYS_ORGVAR "
            strSql1 &= " SET ItemVar_1 = @ItemVar_1 "
            strSql1 &= "  ,ItemVar_2 = @ItemVar_2 "
            strSql1 &= "  ,ItemVar_3 = @ItemVar_3 "
            'strSql1 &= " ,TPlanID = @TPlanID "
            strSql1 &= "  ,ModifyAcct = @ModifyAcct "
            strSql1 &= "  ,ModifyDate = @ModifyDate "
            strSql1 &= " WHERE 1=1 "
            strSql1 &= "  AND RID = @RID "
            strSql1 &= "  AND TPlanID = @TPlanID "
            Dim uCmd As New SqlCommand(strSql1, objconn)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("ItemVar_1", SqlDbType.VarChar).Value = IIf(itemvar_1.Text = "", Convert.DBNull, itemvar_1.Text)
                .Parameters.Add("ItemVar_2", SqlDbType.VarChar).Value = IIf(itemvar_2.Text = "", Convert.DBNull, itemvar_2.Text)
                .Parameters.Add("ItemVar_3", SqlDbType.VarChar).Value = IIf(itemvar_3.Text = "", Convert.DBNull, itemvar_3.Text)
                .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ModifyDate", SqlDbType.DateTime).Value = Now()
                .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
                DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
            End With
        Else
            Dim iSOID As Integer = DbAccess.GetNewId(objconn, "SYS_ORGVAR_SOID_SEQ,SYS_ORGVAR,SOID")
            '新增
            strSql1 = " INSERT into SYS_ORGVAR (SOID,RID,ITEMVAR_1,ITEMVAR_2,ITEMVAR_3,TPLANID,ModifyAcct,ModifyDate) "
            strSql1 &= " VALUES (@SOID,@RID,@ITEMVAR_1,@ITEMVAR_2,@ITEMVAR_3,@TPLANID,@ModifyAcct,@ModifyDate) "
            Dim iCmd As New SqlCommand(strSql1, objconn)
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("SOID", SqlDbType.Decimal).Value = iSOID
                .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                .Parameters.Add("ItemVar_1", SqlDbType.VarChar).Value = IIf(itemvar_1.Text = "", Convert.DBNull, itemvar_1.Text)
                .Parameters.Add("ItemVar_2", SqlDbType.VarChar).Value = IIf(itemvar_2.Text = "", Convert.DBNull, itemvar_2.Text)
                .Parameters.Add("ItemVar_3", SqlDbType.VarChar).Value = IIf(itemvar_3.Text = "", Convert.DBNull, itemvar_3.Text)
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
                .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ModifyDate", SqlDbType.DateTime).Value = Now()
                DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
            End With
        End If

        Me.button5.Visible = False
        Me.msg.Visible = False
        Me.button1.Enabled = True
        btnExport1.Enabled = Me.button1.Enabled
        Common.MessageBox(Me, "儲存成功!!")
    End Sub

    '匯出
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        Dim sMyValue As String = ""
        sMyValue = ""
        sMyValue &= "&OCID1=" & OCIDValue1.Value
        sMyValue &= "&DistID=" & sm.UserInfo.DistID
        '就服單位協助報名
        Select Case rblEnterPathW.SelectedValue
            Case "A"
            Case "Y", "N"
                sMyValue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
        End Select

        'SelResultID Name 
        '01 正取 
        '02 備取 (缺考使用)
        '03 未錄取 
        '04 缺考
        Dim v_SelResult As String = TIMS.GetListValue(SelResult)
        Select Case v_SelResult' SelResult.SelectedValue
            Case "04"
                '02 備取 (缺考使用)
                '缺考 (使用備取)
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "Maintest_result_1", MyValue)
                sMyValue &= "&SelResultID=" & TIMS.cst_SelResultID_備取 '& SelResult.SelectedValue
                sMyValue &= "&WOResult=-1" '缺考
            Case Else
                '非缺考
                sMyValue &= "&SelResultID=" & v_SelResult 'SelResult.SelectedValue '(01/02/03)
                sMyValue &= "&WOResult=1" '非缺考
                'Select Case SelResult.SelectedValue
                '    Case "02", "03"
                '        '不顯示報到日期
                '        sMyValue &= "&ChkDateList=N"
                '    Case Else '01 正取
                '        '顯示報到日期
                '        sMyValue &= "&ChkDateList=Y"
                'End Select
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "Maintest_result", MyValue)
        End Select

        'planid,OCID,start_date,end_date,EnterPathY,EnterPathN
        'OCID1,DistID,SelResultID,WOResult,EnterPathY,EnterPathN
        Dim dt As DataTable = Nothing
        dt = TIMS.LoadDatab39(2, sMyValue, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        Call TIMS.ExpRptb39(Me, objconn, dt)
    End Sub
End Class