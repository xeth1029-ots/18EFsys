Partial Class SD_02_018_add
    Inherits AuthBasePage

    '職前計畫設計
    'labEXAMPLUS 'labEIdentity
    Const cst_printFN1 As String = "SD02003_RXM"
    Const cst_SD02018aspx As String = "SD_02_018.aspx?ID="
    Const cst_SD02018_addaspx As String = "SD_02_018_add.aspx?ID="
    'Dim au As New cAUTH
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objConn) '開啟連線
        '檢查Session是否存在 End
        'PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then Call Create1()
    End Sub

    '第1次載入
    Sub Create1()
        msg.Text = ""
        DataGridTable1.Visible = False

        'BtnSave1X.Visible = False
        'BtnSave2X.Visible = False
        'BtnPrint1.Visible = False
        'trODNUMBER.Visible = False

        BtnSave1X.Attributes.Add("onclick", "return savedataCHK1();")  '解鎖學員錄訓作業
        BtnSave2X.Attributes.Add("onclick", "return savedataCHK2();")  '審核確認

        Dim url1 As String = ""
        Dim ACT As String = TIMS.sUtl_GetRqValue(Me, "ACT")
        Dim OCID As String = TIMS.sUtl_GetRqValue(Me, "OCID")
        If OCID = "" Then
            url1 = cst_SD02018aspx & TIMS.Get_MRqID(Me)
            TIMS.Utl_Redirect(Me, objConn, url1)
        End If

        Dim CFGUID As String = TIMS.sUtl_GetRqValue(Me, "CFGUID")
        Dim CFSEQNO As String = TIMS.sUtl_GetRqValue(Me, "CFSEQNO")
        Hid_CFGUID.Value = CFGUID
        Hid_OCID.Value = OCID
        Hid_CFSEQNO.Value = CFSEQNO

        Select Case ACT
            Case "EDIT1" '(檢視)
                Call Search1()
            Case Else
                url1 = cst_SD02018aspx & TIMS.Get_MRqID(Me)
                TIMS.Utl_Redirect(Me, objConn, url1)
        End Select
    End Sub

    '錄訓名單審核－檢視
    Sub Search1()
        msg.Text = "查無資料!!"
        DataGridTable1.Visible = False

        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Dim drOCID As DataRow = TIMS.GetOCIDDate(Hid_OCID.Value, objConn)
        If drOCID Is Nothing Then Exit Sub
        LabClassName1.Text = drOCID("ClassCName2") & "，訓練人數" & drOCID("TNum") & "人"

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cf.OCID " & vbCrLf
        sql &= "  ,cf.CFSEQNO " & vbCrLf
        sql &= "  ,cf.ODDATE1 " & vbCrLf
        sql &= "  ,cf.ODNUMBER " & vbCrLf
        sql &= "  ,cf.ODNUMBER2 " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, cf.CONFIRDATE, 111) CONFIRDATE " & vbCrLf '(系統)公告日期(建檔日)
        sql &= "  ,cf.NOLOCK " & vbCrLf
        sql &= "  ,cf.ROVEDACCT " & vbCrLf
        sql &= "  ,cf.ROVEDDATE " & vbCrLf
        sql &= "  ,cf.ANNMENTACCT " & vbCrLf
        sql &= "  ,cf.ANNMENTDATE " & vbCrLf
        sql &= "  ,sf.SETID " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, sf.ENTERDATE, 111) ENTERDATE " & vbCrLf
        sql &= "  ,sf.SERNUM " & vbCrLf
        sql &= "  ,b.NAME STUDNAME " & vbCrLf
        sql &= "  ,b.IDNO " & vbCrLf
        'sql &= "  ,CASE WHEN b.EXAMPLUS = 1 THEN '是' END EXAMPLUS " & vbCrLf
        sql &= "  ,b.EIdentityID " & vbCrLf
        'sql &= "  ,kk.name EIDENTITY " & vbCrLf
        sql &= "  ,st.SUMOFGRAD " & vbCrLf
        sql &= "  ,st.SELRESULTID " & vbCrLf
        '01:正取、02:備取、03:未錄取
        sql &= " ,CASE st.SELRESULTID WHEN '01' THEN '正取' WHEN '02' THEN '備取' WHEN '03' THEN '未錄取' END SELRESULT_N" & vbCrLf
        sql &= "  ,st.SELSORT " & vbCrLf
        sql &= "  ,B.EXAMNO " & vbCrLf
        sql &= " FROM CLASS_CONFIRM cf " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc ON cf.OCID = cc.OCID " & vbCrLf
        sql &= " JOIN STUD_CONFIRM sf ON sf.OCID = cf.OCID AND sf.CFSEQNO = cf.CFSEQNO " & vbCrLf
        sql &= " JOIN VIEW_PLAN vp ON vp.PlanID = cc.PlanID " & vbCrLf
        sql &= " JOIN VIEW_RIDNAME rr ON rr.RID = cc.RID " & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo ON oo.comidno = cc.comidno " & vbCrLf
        sql &= " JOIN V_ENTERTYPE b ON sf.OCID = b.OCID1 AND sf.SETID = b.SETID AND sf.ENTERDATE = b.ENTERDATE AND sf.SERNUM = b.SERNUM " & vbCrLf
        sql &= " JOIN STUD_SELRESULT st ON st.SETID = b.SETID AND st.ENTERDATE = b.ENTERDATE AND st.SERNUM = b.SERNUM AND st.OCID = b.OCID1 " & vbCrLf
        'sql &= " LEFT JOIN key_Identity kk ON kk.IdentityID = b.EIdentityID " & vbCrLf
        'sql &= " LEFT JOIN MVIEW_RELSHIP23 r3 ON r3.RID3 = cc.RID AND r3.PlanID = cc.PlanID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "  AND cf.OCID = @OCID1 " & vbCrLf
        sql &= "  AND cf.CFGUID = @CFGUID " & vbCrLf
        sql &= "  AND cf.CFSEQNO = @CFSEQNO " & vbCrLf
        sql &= " ORDER BY st.SELRESULTID ,st.SELSORT ,B.EXAMNO " & vbCrLf
        Dim parms As New Hashtable
        parms.Add("OCID1", Hid_OCID.Value)
        parms.Add("CFGUID", Hid_CFGUID.Value)
        parms.Add("CFSEQNO", Hid_CFSEQNO.Value)
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)

        'Dim sCmd As New SqlCommand(sql, objConn)
        'Dim dt As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("OCID1", SqlDbType.VarChar).Value = Hid_OCID.Value
        '    .Parameters.Add("CFGUID", SqlDbType.VarChar).Value = Hid_CFGUID.Value
        '    .Parameters.Add("CFSEQNO", SqlDbType.VarChar).Value = Hid_CFSEQNO.Value
        '    dt.Load(.ExecuteReader())
        'End With
        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr1 As DataRow = dt.Rows(0)
        ODDATE1.Text = TIMS.Cdate3(dr1("ODDATE1"))
        ODNUMBER.Text = Convert.ToString(dr1("ODNUMBER"))
        ODNUMBER2.Text = Convert.ToString(dr1("ODNUMBER2"))
        Hid_NOLOCK.Value = Convert.ToString(dr1("NOLOCK"))
        Hid_ROVEDDATE.Value = TIMS.Cdate3(dr1("ROVEDDATE")) '審核
        Hid_ANNMENTDATE.Value = TIMS.Cdate3(dr1("ANNMENTDATE")) '公告日期
        Dim strFlagR2A3 As String = "" '2:審核 3:公告
        If Hid_ROVEDDATE.Value <> "" Then strFlagR2A3 = "2"
        If Hid_ANNMENTDATE.Value <> "" Then strFlagR2A3 = "3"
        BtnSave1X.Enabled = False '解鎖
        BtnSave2X.Enabled = False '審核
        BtnSave3X.Enabled = False '公告

        Dim s_tip As String = ""
        Select Case strFlagR2A3 '2:審核 3:公告
            Case "2"
                BtnSave1X.Enabled = True '解鎖
                BtnSave2X.Enabled = False '審核
                BtnSave3X.Enabled = True '公告
                s_tip = "已執行 審核"
                TIMS.Tooltip(BtnSave2X, s_tip, True)
            Case "3"
                BtnSave1X.Enabled = False '解鎖
                BtnSave2X.Enabled = False '審核
                BtnSave3X.Enabled = False '公告
                s_tip = "已執行 公告"
                TIMS.Tooltip(BtnSave1X, s_tip, True)
                TIMS.Tooltip(BtnSave2X, s_tip, True)
                TIMS.Tooltip(BtnSave3X, s_tip, True)
            Case Else
                BtnSave1X.Enabled = True '解鎖
                BtnSave2X.Enabled = True '審核
                BtnSave3X.Enabled = False '公告
                s_tip = "尚未審核"
                TIMS.Tooltip(BtnSave3X, s_tip, True)
        End Select

        'BtnPrint1.Enabled = False '列印
        If ODDATE1.Text <> "" Then
            imgODDATE1.Style.Item("display") = "none"
            ODDATE1.ReadOnly = True
            ODNUMBER.ReadOnly = True
            ODNUMBER2.ReadOnly = True
            'BtnPrint1.Enabled = True '列印
        End If

        msg.Text = ""
        DataGridTable1.Visible = True
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labEXAMNO As Label = e.Item.FindControl("labEXAMNO")
                Dim labStdName As Label = e.Item.FindControl("labStdName")
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")
                'Dim labEXAMPLUS As Label = e.Item.FindControl("labEXAMPLUS")
                'Dim labEIdentity As Label = e.Item.FindControl("labEIdentity")
                Dim labSUMOFGRAD As Label = e.Item.FindControl("labSUMOFGRAD")
                Dim labSELRESULT_N As Label = e.Item.FindControl("labSELRESULT_N") '甄試結果/錄訓結果

                Dim SETID As HtmlInputHidden = e.Item.FindControl("SETID")
                Dim ENTERDATE As HtmlInputHidden = e.Item.FindControl("ENTERDATE")
                Dim SERNUM As HtmlInputHidden = e.Item.FindControl("SERNUM")
                SETID.Value = Convert.ToString(drv("SETID"))
                ENTERDATE.Value = Convert.ToString(drv("ENTERDATE"))
                SERNUM.Value = Convert.ToString(drv("SERNUM"))

                labEXAMNO.Text = Convert.ToString(drv("EXAMNO")) 'TIMS.Get_DGSeqNo(sender, e) ' SignNo 'e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                labStdName.Text = Convert.ToString(drv("STUDNAME"))
                labIDNO.Text = Convert.ToString(drv("IDNO"))
                'labEXAMPLUS.Text = Convert.ToString(drv("EXAMPLUS"))
                'labEIdentity.Text = Convert.ToString(drv("EIdentity"))
                labSUMOFGRAD.Text = Convert.ToString(drv("SUMOFGRAD"))
                labSELRESULT_N.Text = Convert.ToString(drv("SELRESULT_N"))
        End Select
    End Sub

    '回上一頁
    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Dim url1 As String = ""
        url1 = ""
        url1 &= cst_SD02018aspx & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

    '儲存-解鎖學員錄訓作業
    Sub SaveData1X()
        Dim sCFGUID As String = Hid_CFGUID.Value
        Dim iCFSEQNO As Integer = Val(Hid_CFSEQNO.Value)
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim Trans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try
            Dim sql As String = ""
            sql &= " UPDATE CLASS_CONFIRM " & vbCrLf
            sql &= " SET NOLOCK = 'Y' " & vbCrLf
            sql &= "  ,MODIFYACCT = @MODIFYACCT " & vbCrLf
            sql &= "  ,MODIFYDATE = GETDATE() " & vbCrLf
            'sql &= " ,CONFIRACCT = @CONFIRACCT " & vbCrLf '公告日期(建檔日)
            'sql &= " ,CONFIRDATE = GETDATE() " & vbCrLf '公告日期(建檔日)
            sql &= " WHERE CFGUID = @CFGUID " & vbCrLf
            sql &= " AND OCID = @OCID " & vbCrLf
            sql &= " AND CFSEQNO = @CFSEQNO " & vbCrLf
            Dim uCmd As New SqlCommand(sql, oConn, Trans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("CONFIRACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("CFGUID", SqlDbType.VarChar).Value = sCFGUID
                .Parameters.Add("OCID", SqlDbType.Int).Value = Hid_OCID.Value
                .Parameters.Add("CFSEQNO", SqlDbType.Int).Value = iCFSEQNO
                '.ExecuteNonQuery()  'edit，by:20181017
                DbAccess.ExecuteNonQuery(uCmd.CommandText, Trans, uCmd.Parameters)  'edit，by:20181017
            End With
            Call DbAccess.CommitTrans(Trans)
        Catch ex As Exception
            Call DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(oConn)
            Common.MessageBox(Me, ex.ToString)
            Throw ex 'Exit Sub
        End Try
        Call TIMS.CloseDbConn(oConn)
        Call TIMS.CloseDbConn(objConn)
        Dim url1 As String = cst_SD02018aspx & TIMS.Get_MRqID(Me)
        'Common.MessageBox(Me, TIMS.cst_SAVEOKMsg1, url1)  'edit，by:20181017 (由於目前系統的轉頁功能仍有問題,所以先拿掉轉頁功能)
        Page.RegisterStartupScript("", "<script>alert('" & TIMS.cst_SAVEOKMsg1 + "'); window.location.href='" & url1 & "';</script>")  'edit，by:20181017
    End Sub

    '儲存-審核確認
    Sub SaveData2X()
        Dim sCFGUID As String = Hid_CFGUID.Value
        Dim iCFSEQNO As Integer = Val(Hid_CFSEQNO.Value)
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim Trans As SqlTransaction = DbAccess.BeginTrans(oConn)

        Try
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " UPDATE CLASS_CONFIRM " & vbCrLf
            sql &= " SET MODIFYACCT = @MODIFYACCT " & vbCrLf
            sql &= "  ,MODIFYDATE = GETDATE() " & vbCrLf
            sql &= "  ,ROVEDACCT = @ROVEDACCT " & vbCrLf '審核
            sql &= "  ,ROVEDDATE = GETDATE() " & vbCrLf '審核
            sql &= " WHERE 1=1 " & vbCrLf
            sql &= "  AND CFGUID = @CFGUID " & vbCrLf
            sql &= "  AND OCID = @OCID " & vbCrLf
            sql &= "  AND CFSEQNO = @CFSEQNO " & vbCrLf
            Dim uCmd As New SqlCommand(sql, oConn, Trans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ROVEDACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("CFGUID", SqlDbType.VarChar).Value = sCFGUID
                .Parameters.Add("OCID", SqlDbType.Int).Value = Hid_OCID.Value
                .Parameters.Add("CFSEQNO", SqlDbType.Int).Value = iCFSEQNO
                '.ExecuteNonQuery()  'edit，by:20181017
                DbAccess.ExecuteNonQuery(uCmd.CommandText, Trans, uCmd.Parameters)  'edit，by:20181017
            End With
            Call DbAccess.CommitTrans(Trans)
        Catch ex As Exception
            Call DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(oConn)
            Common.MessageBox(Me, ex.ToString)
            Throw ex 'Exit Sub
        End Try

        Call TIMS.CloseDbConn(oConn)
        Call TIMS.CloseDbConn(objConn)
        Dim sCmdArg As String = ""
        TIMS.SetMyValue(sCmdArg, "ACT", "EDIT1")
        TIMS.SetMyValue(sCmdArg, "OCID", Hid_OCID.Value)
        TIMS.SetMyValue(sCmdArg, "CFGUID", Hid_CFGUID.Value)
        TIMS.SetMyValue(sCmdArg, "CFSEQNO", Hid_CFSEQNO.Value)

        Dim url1 As String = cst_SD02018_addaspx & TIMS.Get_MRqID(Me) & sCmdArg
        'Common.MessageBox(Me, TIMS.cst_SAVEOKMsg2, url1)  'edit，by:20181017 (由於目前系統的轉頁功能仍有問題,所以先拿掉轉頁功能)
        Page.RegisterStartupScript("", "<script>alert('" & TIMS.cst_SAVEOKMsg2 + "'); window.location.href='" & url1 & "';</script>")  'edit，by:20181017
    End Sub

    '儲存-公告
    Sub SaveData3X()
        Dim sCFGUID As String = Hid_CFGUID.Value
        Dim iCFSEQNO As Integer = Val(Hid_CFSEQNO.Value)
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim Trans As SqlTransaction = DbAccess.BeginTrans(oConn)

        Try
            Dim sql As String = ""
            sql &= " UPDATE CLASS_CONFIRM " & vbCrLf
            sql &= " SET MODIFYACCT = @MODIFYACCT " & vbCrLf
            sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
            sql &= " ,ANNMENTACCT = @ANNMENTACCT " & vbCrLf '公告
            sql &= " ,ANNMENTDATE = GETDATE() " & vbCrLf '公告
            sql &= " ,ODDATE1 = dbo.fn_DATE(@ODDATE1) " & vbCrLf '公告文號：日期
            sql &= " ,ODNUMBER = @ODNUMBER " & vbCrLf '公告文號1
            sql &= " ,ODNUMBER2 = @ODNUMBER2 " & vbCrLf '公告文號2
            sql &= " WHERE CFGUID = @CFGUID " & vbCrLf
            sql &= " AND OCID = @OCID " & vbCrLf
            sql &= " AND CFSEQNO = @CFSEQNO " & vbCrLf
            Dim uCmd As New SqlCommand(sql, oConn, Trans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ANNMENTACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ODDATE1", SqlDbType.VarChar).Value = ODDATE1.Text 'TIMS.cdate3()
                .Parameters.Add("ODNUMBER", SqlDbType.VarChar).Value = ODNUMBER.Text
                .Parameters.Add("ODNUMBER2", SqlDbType.VarChar).Value = ODNUMBER2.Text
                .Parameters.Add("CFGUID", SqlDbType.VarChar).Value = sCFGUID
                .Parameters.Add("OCID", SqlDbType.Int).Value = Hid_OCID.Value
                .Parameters.Add("CFSEQNO", SqlDbType.Int).Value = iCFSEQNO
                '.ExecuteNonQuery()  'edit，by:20181017
                DbAccess.ExecuteNonQuery(uCmd.CommandText, Trans, uCmd.Parameters)  'edit，by:20181017
            End With
            Call DbAccess.CommitTrans(Trans)
        Catch ex As Exception
            Call DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(oConn)
            Common.MessageBox(Me, ex.ToString)
            Throw ex 'Exit Sub
        End Try

        Call TIMS.CloseDbConn(oConn)
        Call TIMS.CloseDbConn(objConn)

        Dim url1 As String = cst_SD02018_addaspx & TIMS.Get_MRqID(Me)
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg2, url1)  'edit，by:20181017 (由於目前系統的轉頁功能仍有問題,所以先拿掉轉頁功能)
        Page.RegisterStartupScript("", "<script>alert('" & TIMS.cst_SAVEOKMsg2 + "'); window.location.href='" & url1 & "';</script>")  'edit，by:20181017
    End Sub

    '檢核
    Function CheckData3X(ByRef Errmsg As String) As Boolean
        Errmsg = ""

        ODDATE1.Text = TIMS.ClearSQM(ODDATE1.Text)
        ODDATE1.Text = TIMS.Cdate3(ODDATE1.Text)
        ODNUMBER.Text = TIMS.ClearSQM(ODNUMBER.Text)
        ODNUMBER2.Text = TIMS.ClearSQM(ODNUMBER2.Text)

        If ODDATE1.Text = "" Then
            Errmsg &= "公文日期 為必填資料" & vbCrLf
            Return False
        End If
        If ODNUMBER.Text = "" Then
            Errmsg &= "公文文號(字) 為必填資料" & vbCrLf
            Return False
        End If
        If ODNUMBER2.Text = "" Then
            Errmsg &= "公文文號(函) 為必填資料" & vbCrLf
            Return False
        End If

        If Errmsg <> "" Then Return False
        Return True
    End Function

    '解鎖學員錄訓作業
    Protected Sub BtnSave1X_Click(sender As Object, e As EventArgs) Handles BtnSave1X.Click
        SaveData1X()
    End Sub

    '審核確認
    Protected Sub BtnSave2X_Click(sender As Object, e As EventArgs) Handles BtnSave2X.Click
        SaveData2X()
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Hid_CFGUID.Value = TIMS.ClearSQM(Hid_CFGUID.Value)
        Hid_CFSEQNO.Value = TIMS.ClearSQM(Hid_CFSEQNO.Value)
        Dim myValue As String = ""
        myValue = ""
        myValue &= "&OCID=" & Hid_OCID.Value
        myValue &= "&CFGUID=" & Hid_CFGUID.Value
        myValue &= "&CFSEQNO=" & Hid_CFSEQNO.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    '公告
    Protected Sub BtnSave3X_Click(sender As Object, e As EventArgs) Handles BtnSave3X.Click
        Dim Errmsg As String = ""
        Call CheckData3X(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        SaveData3X()
    End Sub
End Class