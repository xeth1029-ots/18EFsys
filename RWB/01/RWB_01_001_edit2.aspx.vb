Partial Class RWB_01_001_edit2
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
        If flag_ROC Then
            lblCDATE1.Text = TIMS.Cdate17(DateTime.Now.ToString("yyyy/MM/dd"))  'edit，by:20181019
        Else
            lblCDATE1.Text = DateTime.Now.ToString("yyyy/MM/dd")
        End If

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
            Dim rSEQNO As String = TIMS.ClearSQM(Request("SEQNO"))
            If rSEQNO_E <> "" AndAlso rSEQNO_E = rSEQNO Then hid_V.Value = rSEQNO
            If hid_V.Value <> "" Then LoadData1(Val(hid_V.Value))
        End If
    End Sub

    '資料讀取
    Private Sub LoadData1(ByVal iSEQNO As Integer)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.C_UDATE DESC) AS ROWNUM " & vbCrLf
        sql &= "        ,FORMAT(a.C_SDATE, 'yyyy-MM-dd') CSDATE " & vbCrLf
        sql &= "        ,FORMAT(a.C_EDATE, 'yyyy-MM-dd') CEDATE " & vbCrLf
        sql &= "        ,FORMAT(a.C_CDATE, 'yyyy-MM-dd') CCDATE " & vbCrLf
        sql &= "        ,FORMAT(a.C_SDATE, 'HH') CSDATEHH " & vbCrLf
        sql &= "        ,FORMAT(a.C_EDATE, 'HH') CEDATEHH " & vbCrLf
        sql &= "        ,FORMAT(a.C_SDATE, 'mm') CSDATEMM " & vbCrLf
        sql &= "        ,FORMAT(a.C_EDATE, 'mm') CEDATEMM " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_SDATE, 111) CSDATED " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_EDATE, 111) CEDATED " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_CDATE, 111) CCDATED " & vbCrLf
        sql &= "        ,a.SEQNO " & vbCrLf
        sql &= "        ,a.FUNID " & vbCrLf
        sql &= "        ,a.SUB_FUNID " & vbCrLf
        sql &= "        ,a.C_SDATE " & vbCrLf
        sql &= "        ,a.C_EDATE " & vbCrLf
        sql &= "        ,a.C_TITLE " & vbCrLf
        sql &= "        ,a.C_CONTENT1 " & vbCrLf
        sql &= "        ,a.C_CONTENT2 " & vbCrLf
        sql &= "        ,a.C_CONTENT3 " & vbCrLf
        sql &= "        ,a.C_CDATE " & vbCrLf
        sql &= "        ,a.C_CACCT " & vbCrLf
        sql &= "        ,a.C_UDATE " & vbCrLf
        sql &= "        ,a.C_UACCT " & vbCrLf
        sql &= "        ,a.C_STATUS " & vbCrLf
        sql &= " FROM TB_CONTENT a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "       AND a.FUNID = '001' " & vbCrLf
        sql &= "       AND a.SEQNO = @SEQNO " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If Convert.ToString(iSEQNO) <> "" Then parms.Add("SEQNO", iSEQNO)
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)
        If flag_ROC Then
            lblCDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CCDATED")))  'edit，by:20181019
            txtSDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CSDATED")))  'edit，by:20181019
        Else
            lblCDATE1.Text = Convert.ToString(dr("CCDATED"))
            txtSDATE1.Text = Convert.ToString(dr("CSDATED"))
        End If
        Common.SetListItem(ddlC_SDATE_hh1, Convert.ToString(dr("CSDATEHH")))
        Common.SetListItem(ddlC_SDATE_mm1, Convert.ToString(dr("CSDATEMM")))
        lblTitle.Text = Convert.ToString(dr("C_TITLE"))
        If flag_ROC Then
            txtEDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CEDATED")))  'edit，by:20181019
        Else
            txtEDATE1.Text = Convert.ToString(dr("CEDATED"))
        End If
        Common.SetListItem(ddlC_EDATE_hh1, Convert.ToString(dr("CEDATEHH")))
        Common.SetListItem(ddlC_EDATE_mm1, Convert.ToString(dr("CEDATEMM")))
        '==========
        If Convert.ToString(dr("CEDATEHH")) = "" Then ddlC_EDATE_hh1.SelectedIndex = -1
        If Convert.ToString(dr("CEDATEMM")) = "" Then ddlC_EDATE_mm1.SelectedIndex = -1

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

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)  'edit，by:20181019
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)  'edit，by:20181019
        Dim mySDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019

        If txtSDATE1.Text = "" OrElse TIMS.CStr1(oC_SDATE) = "" Then Errmsg &= "上架日期 不可為空" & vbCrLf
        'If txtEDATE1.Text = "" OrElse TIMS.CStr1(oC_EDATE) = "" Then Errmsg &= "停用日期 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False
        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) = 0 Then Errmsg &= "上架日期與停用日期 不可相等!!" & vbCrLf
        'If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) < 0 Then Errmsg &= "上架日期與停用日期 順序異常!!" & vbCrLf
        If txtEDATE1.Text <> "" And TIMS.CStr1(oC_EDATE) <> "" Then
            If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) < 0 Then Errmsg &= "上架日期與停用日期 順序異常!!" & vbCrLf
        End If
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
            Dim url1 As String = "RWB_01_001.aspx?id1=" & TIMS.Get_MRqID(Me)
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
        sql &= " UPDATE TB_CONTENT " & vbCrLf
        sql &= " SET FUNID = '001' " & vbCrLf
        sql &= "     ,SUB_FUNID = '3' " & vbCrLf
        sql &= "     ,C_SDATE = @C_SDATE " & vbCrLf
        sql &= "     ,C_EDATE = @C_EDATE " & vbCrLf
        sql &= "     ,C_UDATE = GETDATE() " & vbCrLf
        sql &= "     ,C_UACCT = @C_UACCT " & vbCrLf
        sql &= "     ,C_STATUS = 'M' " & vbCrLf
        sql &= " WHERE SEQNO = @SEQNO " & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        Dim uSql As String = sql

        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Call TIMS.OpenDbConn(objconn)

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        Dim oC_SDATE As Object = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text) + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = Nothing
        If txtEDATE1.Text <> "" Then
            oC_EDATE = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text) + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Else
            oC_EDATE = Convert.DBNull
        End If

        Dim iRst As Integer = 0
        If hid_V.Value <> "" Then
            '修改
            With uCmd
                Dim parms As Hashtable = New Hashtable()
                parms.Add("SEQNO", hid_V.Value)
                parms.Add("C_SDATE", oC_SDATE)
                parms.Add("C_EDATE", oC_EDATE)
                parms.Add("C_UACCT", sm.UserInfo.UserID)
                iRst += DbAccess.ExecuteNonQuery(uSql, parms)
            End With
        End If

        rst = True
        Return rst
    End Function

    '取消
    Protected Sub bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
        Dim url1 As String = ""
        url1 = "RWB_01_001.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class