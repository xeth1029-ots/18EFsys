Partial Class RWB_01_004_edit
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then Call sCreate1() '頁面初始化
    End Sub

    '頁面初始化
    Sub sCreate1()
        txtSDATE1.Text = ""
        rblType.SelectedValue = "1"
        txtQ.Text = ""
        txtA.Text = ""
        txtCSORT1.Text = ""
        rblUse.SelectedValue = "Y"
        txtEDATE1.Text = ""

        ddlC_SDATE_hh1.Items.Clear()
        ddlC_EDATE_hh1.Items.Clear()
        For i As Integer = 0 To 23
            ddlC_SDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
            ddlC_EDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
        Next

        ddlC_SDATE_mm1.Items.Clear()
        ddlC_EDATE_mm1.Items.Clear()
        For j As Integer = 0 To 59
            ddlC_SDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
            ddlC_EDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
        Next

        Common.SetListItem(ddlC_SDATE_hh1, "00")
        Common.SetListItem(ddlC_SDATE_mm1, "00")
        Common.SetListItem(ddlC_EDATE_hh1, "23")
        Common.SetListItem(ddlC_EDATE_mm1, "59")

        If TIMS.ClearSQM(Request("A")) = "E" Then
            Dim rSEQNO_E As String = TIMS.DecryptAes(TIMS.ClearSQM(Request("SEQNO_E")))
            Dim rSEQNO As String = TIMS.ClearSQM(Request("QAID"))
            If rSEQNO_E <> "" AndAlso rSEQNO_E = rSEQNO Then hid_V.Value = rSEQNO
            If hid_V.Value <> "" Then LoadData1(Val(hid_V.Value))
        End If
    End Sub

    '資料讀取
    Private Sub LoadData1(ByVal iSEQNO As Integer)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.QAID DESC) ROWNUM " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, a.START_DATE, 111) CSDATE " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, a.END_DATE, 111) CEDATE " & vbCrLf
        sql &= "  ,FORMAT(a.START_DATE, 'HH') CSDATEHH " & vbCrLf
        sql &= "  ,FORMAT(a.END_DATE, 'HH') CEDATEHH " & vbCrLf
        sql &= "  ,FORMAT(a.START_DATE, 'mm') CSDATEMM " & vbCrLf
        sql &= "  ,FORMAT(a.END_DATE, 'mm') CEDATEMM " & vbCrLf
        sql &= "  ,a.QAID " & vbCrLf
        sql &= "  ,a.TYPEID " & vbCrLf
        sql &= "  ,CASE WHEN a.TYPEID = '1' THEN '產業人才投資方案' " & vbCrLf
        sql &= "   WHEN a.TYPEID = '2' THEN '自辦在職訓練' " & vbCrLf
        sql &= "   WHEN a.TYPEID = '3' THEN '企業委託訓練' " & vbCrLf
        sql &= "   WHEN a.TYPEID = '4' THEN '充電起飛' " & vbCrLf
        sql &= "   WHEN a.TYPEID = '5' THEN '網站操作問題' " & vbCrLf
        sql &= "   ELSE '' END C_TYPE " & vbCrLf
        sql &= "  ,a.START_DATE " & vbCrLf
        sql &= "  ,a.END_DATE " & vbCrLf
        sql &= "  ,a.QUESTION " & vbCrLf
        sql &= "  ,CASE WHEN LEN(a.QUESTION) > 15 THEN SUBSTRING(a.QUESTION, 1, 15) + '...' ELSE a.QUESTION END QUESTION1 " & vbCrLf
        sql &= "  ,a.ANSWER " & vbCrLf
        sql &= "  ,a.ISUSED " & vbCrLf
        sql &= "  ,CASE WHEN a.ISUSED = 'Y' THEN '啟用' " & vbCrLf
        sql &= "   WHEN a.ISUSED = 'N' THEN '停用' " & vbCrLf
        sql &= "   ELSE '' END C_ISUSED " & vbCrLf
        sql &= "  ,a.MODIFYACCT " & vbCrLf
        sql &= "  ,a.MODIFYDATE" & vbCrLf
        sql &= "  ,C_SORT1 " & vbCrLf
        sql &= " FROM TB_QA a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "  AND a.QAID = @QAID " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If Convert.ToString(iSEQNO) <> "" Then parms.Add("QAID", iSEQNO)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)
        txtSDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(Convert.ToString(dr("CSDATE"))), Convert.ToString(dr("CSDATE")))  'edit，by:20181019
        Common.SetListItem(ddlC_SDATE_hh1, Convert.ToString(dr("CSDATEHH")))
        Common.SetListItem(ddlC_SDATE_mm1, Convert.ToString(dr("CSDATEMM")))
        Common.SetListItem(rblType, Convert.ToString(dr("TYPEID")))
        txtQ.Text = Convert.ToString(dr("QUESTION"))
        txtA.Text = Convert.ToString(dr("ANSWER"))
        txtCSORT1.Text = Convert.ToString(dr("C_SORT1"))
        Common.SetListItem(rblUse, Convert.ToString(dr("ISUSED")))
        txtEDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(Convert.ToString(dr("CEDATE"))), Convert.ToString(dr("CEDATE")))  'edit，by:20181019
        Common.SetListItem(ddlC_EDATE_hh1, Convert.ToString(dr("CEDATEHH")))
        Common.SetListItem(ddlC_EDATE_mm1, Convert.ToString(dr("CEDATEMM")))
    End Sub

    '資料儲存
    Protected Sub bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
    End Sub

    '送出前檢核 ---> SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        Dim mySDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtQ.Text = TIMS.ClearSQM2(txtQ.Text)  'edit，by:20190102
        txtA.Text = TIMS.ClearSQM2(txtA.Text)  'edit，by:20190102
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        If txtSDATE1.Text = "" OrElse TIMS.CStr1(oC_SDATE) = "" Then Errmsg &= "上架日期 不可為空" & vbCrLf
        If txtQ.Text = "" Then Errmsg &= "問題內容 不可為空" & vbCrLf
        If txtA.Text = "" Then Errmsg &= "回答內容 不可為空" & vbCrLf
        txtCSORT1.Text = TIMS.ClearSQM(txtCSORT1.Text)
        txtCSORT1.Text = TIMS.ChangeIDNO(txtCSORT1.Text)
        If txtCSORT1.Text <> "" Then
            If Not TIMS.IsNumeric1(txtCSORT1.Text) Then Errmsg &= "序號 不可為非數字" & vbCrLf
        End If
        If txtEDATE1.Text = "" OrElse TIMS.CStr1(oC_EDATE) = "" Then Errmsg &= "停用日期 不可為空" & vbCrLf
        If rblType.SelectedValue = "" Then Errmsg &= "問題類型 不可為空" & vbCrLf
        If rblUse.SelectedValue = "" Then Errmsg &= "啟用狀態 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False

        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) = 0 Then Errmsg &= "上架日期與停用日期 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) < 0 Then Errmsg &= "上架日期與停用日期 順序異常!!" & vbCrLf
        If Errmsg <> "" Then rst = False

        Return rst
    End Function

    '儲存(part-1)
    Sub SaveData1()
        Dim flagSaveOK1 As Boolean = False

        Try
            flagSaveOK1 = SaveData2()
        Catch ex As Exception
            flagSaveOK1 = False
            Common.MessageBox(Me, ex.Message)
            Exit Sub
        End Try

        If flagSaveOK1 Then
            '儲存成功
            Dim url1 As String = "RWB_01_004.aspx?id1=" & TIMS.Get_MRqID(Me)
            'Common.MessageBox(Me, "儲存成功!", url1)
            Common.MessageBox(Me, "儲存成功!")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

    End Sub

    '儲存(part-2)
    Function SaveData2() As Boolean
        Dim rst As Boolean = False 'false:異常

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO TB_QA(QAID, TYPEID, QUESTION, ANSWER, START_DATE, END_DATE, ISUSED, MODIFYACCT, MODIFYDATE,C_SORT1) " & vbCrLf
        sql &= " VALUES (@QAID, @TYPEID, @QUESTION, @ANSWER, @START_DATE, @END_DATE, @ISUSED, @UACCT, GETDATE(),@C_SORT1) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        Dim iSql As String = sql
        sql = "" & vbCrLf
        sql &= " UPDATE TB_QA " & vbCrLf
        sql &= " SET TYPEID = @TYPEID " & vbCrLf
        sql &= " ,QUESTION = @QUESTION " & vbCrLf
        sql &= " ,ANSWER = @ANSWER " & vbCrLf
        sql &= " ,START_DATE = @START_DATE " & vbCrLf
        sql &= " ,END_DATE = @END_DATE " & vbCrLf
        sql &= " ,ISUSED = @ISUSED " & vbCrLf
        sql &= " ,MODIFYACCT = @UACCT " & vbCrLf
        sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " ,C_SORT1 = @C_SORT1 " & vbCrLf
        sql &= " WHERE QAID = @QAID " & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        Dim uSql As String = sql

        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Call TIMS.OpenDbConn(objconn)

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        txtQ.Text = TIMS.ClearSQM2(txtQ.Text)  'edit，by:20190102
        txtA.Text = TIMS.ClearSQM2(txtA.Text)  'edit，by:20190102

        Dim oC_SDATE As Object = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text) + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text) + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019

        Dim iRst As Integer = 0
        If hid_V.Value = "" Then
            '新增
            Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "TB_QA_QAID_SEQ,TB_QA,QAID")
            With iCmd
                Dim parms As Hashtable = New Hashtable()
                parms.Add("QAID", iSEQNO)
                parms.Add("TYPEID", rblType.SelectedValue)
                parms.Add("QUESTION", txtQ.Text)
                parms.Add("ANSWER", txtA.Text)
                parms.Add("START_DATE", oC_SDATE)
                parms.Add("END_DATE", oC_EDATE)
                parms.Add("ISUSED", rblUse.SelectedValue)
                parms.Add("UACCT", sm.UserInfo.UserID)
                parms.Add("C_SORT1", IIf(txtCSORT1.Text <> "", Val(txtCSORT1.Text), iSEQNO))
                iRst += DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            End With
            hid_V.Value = iSEQNO
        Else
            '修改
            With uCmd
                Dim parms As Hashtable = New Hashtable()
                parms.Add("QAID", hid_V.Value)
                parms.Add("TYPEID", rblType.SelectedValue)
                parms.Add("QUESTION", txtQ.Text)
                parms.Add("ANSWER", txtA.Text)
                parms.Add("START_DATE", oC_SDATE)
                parms.Add("END_DATE", oC_EDATE)
                parms.Add("ISUSED", rblUse.SelectedValue)
                parms.Add("UACCT", sm.UserInfo.UserID)
                parms.Add("C_SORT1", IIf(txtCSORT1.Text <> "", Val(txtCSORT1.Text), Val(hid_V.Value)))
                iRst += DbAccess.ExecuteNonQuery(uSql, objconn, parms)
            End With
        End If

        rst = True
        Return rst
    End Function

    '取消
    Protected Sub bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
        Dim url1 As String = ""
        url1 = "RWB_01_004.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class