Partial Class SYS_06_012
    Inherits AuthBasePage

    '資料交換平台-KEY_WEBSITE
    Const cst_btnEdit As String = "btnEdit"
    Const cst_btnDel As String = "btnDel"
    Const cst_btnTest As String = "btnTest"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            create1()
            Call sSearch1()
        End If

    End Sub

    ''' <summary>
    ''' 通訊協定
    ''' </summary>
    ''' <param name="obj1"></param>
    ''' <returns></returns>
    Function Utl_GetMSPROTOCOL1(ByVal obj1 As DropDownList) As DropDownList
        Dim str_stv1 As String = "1:HTTPs GET/POST + XML,2:HTTP GET/POST + XML"
        With obj1
            .Items.Clear()
            For Each str1 As String In str_stv1.Split(",")
                Dim val1 As String = str1.Split(":")(0)
                Dim txt1 As String = str1.Split(":")(1)
                .Items.Add(New ListItem(txt1, val1))
            Next
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, ""))
            .AppendDataBoundItems = True
        End With
        Return obj1
    End Function

    Sub create1()
        divSch1.Visible = True
        divEdt1.Visible = False

        sddlMSPROTOCOL = Utl_GetMSPROTOCOL1(sddlMSPROTOCOL)
        ddlMSPROTOCOL = Utl_GetMSPROTOCOL1(ddlMSPROTOCOL)
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        'Dim MRQ_ID As String = TIMS.Get_MRqID(Me)
        'Dim url1 As String = "CO_01_002.aspx?ID=" & MRQ_ID
        'TIMS.Utl_Redirect(Me, objconn, url1)
        divSch1.Visible = True
        divEdt1.Visible = False
    End Sub

    Sub sSearch1()
        divSch1.Visible = True
        divEdt1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = TIMS.cst_NODATAMsg1

        sMSNAME.Text = TIMS.ClearSQM(sMSNAME.Text)
        sMSIP4.Text = TIMS.ClearSQM(sMSIP4.Text)
        sMSPORT.Text = TIMS.ClearSQM(sMSPORT.Text)
        '數字/'非數字
        If (sMSPORT.Text <> "") Then sMSPORT.Text = If(TIMS.IsNumeric1(sMSPORT.Text), Val(sMSPORT.Text).ToString(), "")
        sMSMODULE.Text = TIMS.ClearSQM(sMSMODULE.Text)
        Dim v_sddlMSPROTOCOL As String = TIMS.GetListValue(sddlMSPROTOCOL) '代碼

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.WSENO" & vbCrLf
        sql &= " ,a.MSNAME" & vbCrLf
        sql &= " ,a.MSIP4" & vbCrLf
        sql &= " ,a.MSPORT" & vbCrLf
        sql &= " ,a.MSMODULE" & vbCrLf
        sql &= " ,a.MSPROTOCOL" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM KEY_WEBSITE a" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If sMSNAME.Text <> "" Then sql &= " and a.MSNAME like '%'+@MSNAME+'%' " & vbCrLf
        If sMSIP4.Text <> "" Then sql &= " and a.MSIP4=@MSIP4 " & vbCrLf
        If sMSPORT.Text <> "" Then sql &= " and a.MSPORT=@MSPORT " & vbCrLf
        If sMSMODULE.Text <> "" Then sql &= " and a.MSMODULE like '%'+@MSMODULE+'%' " & vbCrLf
        If v_sddlMSPROTOCOL <> "" Then sql &= " AND a.MSPROTOCOL=@MSPROTOCOL" & vbCrLf
        sql &= " ORDER BY a.WSENO DESC" & vbCrLf

        Dim parms As New Hashtable
        parms.Clear()
        If sMSNAME.Text <> "" Then parms.Add("MSNAME", sMSNAME.Text)
        If sMSIP4.Text <> "" Then parms.Add("MSIP4", sMSIP4.Text)
        If sMSPORT.Text <> "" Then parms.Add("MSPORT", sMSPORT.Text)
        If sMSMODULE.Text <> "" Then parms.Add("MSMODULE", sMSMODULE.Text)
        If v_sddlMSPROTOCOL <> "" Then parms.Add("MSPROTOCOL", v_sddlMSPROTOCOL)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Sub sSaveData1()
        HID_WSENO.Value = TIMS.ClearSQM(HID_WSENO.Value)
        'If HID_WSENO.Value = "" Then Exit Sub

        Dim d_sql As String = ""
        Dim i_sql As String = ""
        Dim u_sql As String = ""
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO KEY_WEBSITE(" & vbCrLf
        sql &= " WSENO" & vbCrLf
        sql &= " ,MSNAME" & vbCrLf
        sql &= " ,MSIP4" & vbCrLf
        sql &= " ,MSPORT" & vbCrLf
        sql &= " ,MSMODULE" & vbCrLf
        sql &= " ,MSPROTOCOL" & vbCrLf
        sql &= " ,MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @WSENO" & vbCrLf
        sql &= " ,@MSNAME" & vbCrLf
        sql &= " ,@MSIP4" & vbCrLf
        sql &= " ,@MSPORT" & vbCrLf
        sql &= " ,@MSMODULE" & vbCrLf
        sql &= " ,@MSPROTOCOL" & vbCrLf
        sql &= " ,@MODIFYACCT" & vbCrLf
        sql &= " ,GETDATE()" & vbCrLf
        sql &= " )" & vbCrLf
        i_sql = sql

        sql = "" & vbCrLf
        sql &= " UPDATE KEY_WEBSITE" & vbCrLf
        sql &= " SET MSNAME=@MSNAME" & vbCrLf
        sql &= " ,MSIP4=@MSIP4" & vbCrLf
        sql &= " ,MSPORT=@MSPORT" & vbCrLf
        sql &= " ,MSMODULE=@MSMODULE" & vbCrLf
        sql &= " ,MSPROTOCOL=@MSPROTOCOL" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND WSENO=@WSENO" & vbCrLf
        u_sql = sql

        sql = "" & vbCrLf
        sql &= " DELETE KEY_WEBSITE" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND WSENO=@WSENO" & vbCrLf
        d_sql = sql

        Dim v_ddlMSPROTOCOL As String = TIMS.GetListValue(ddlMSPROTOCOL)
        Dim iWSENO As Integer = 0
        If HID_WSENO.Value = "" Then
            '新增
            iWSENO = DbAccess.GetNewId(objconn, "KEY_WEBSITE_WSENO_SEQ,KEY_WEBSITE,WSENO")
            Dim i_parms As New Hashtable
            With i_parms
                .Clear()
                .Add("WSENO", iWSENO)
                .Add("MSNAME", tMSNAME.Text)
                .Add("MSIP4", tMSIP4.Text)
                .Add("MSPORT", tMSPORT.Text)
                .Add("MSMODULE", tMSMODULE.Text)
                .Add("MSPROTOCOL", v_ddlMSPROTOCOL)
                .Add("MODIFYACCT", sm.UserInfo.UserID)
            End With
            DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        Else
            iWSENO = Val(HID_WSENO.Value)
            '修改
            Dim u_parms As New Hashtable
            With u_parms
                .Clear()
                .Add("MSNAME", tMSNAME.Text)
                .Add("MSIP4", tMSIP4.Text)
                .Add("MSPORT", tMSPORT.Text)
                .Add("MSMODULE", tMSMODULE.Text)
                .Add("MSPROTOCOL", v_ddlMSPROTOCOL)
                .Add("MODIFYACCT", sm.UserInfo.UserID)
                .Add("WSENO", iWSENO)
            End With
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If

    End Sub

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        sSearch1()
    End Sub

    Function CheckSaveData1(ByRef sReason As String) As Boolean
        Dim rst As Boolean = True
        Const cst_必須填寫 As String = "必須填寫"

        tMSNAME.Text = TIMS.ClearSQM(tMSNAME.Text)
        tMSIP4.Text = TIMS.ClearSQM(tMSIP4.Text)
        tMSPORT.Text = TIMS.ClearSQM(tMSPORT.Text)
        tMSMODULE.Text = TIMS.ClearSQM(tMSMODULE.Text)
        Dim v_ddlMSPROTOCOL As String = TIMS.GetListValue(ddlMSPROTOCOL)

        If tMSNAME.Text = "" Then
            sReason &= cst_必須填寫 & "掛載系統名稱" & vbCrLf
        End If
        If tMSIP4.Text = "" Then
            sReason &= cst_必須填寫 & "掛載系統IP" & vbCrLf
        Else
            '\\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b
            'Const cst_Pattern As String = "\\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\\b"
            'Dim strIn As String = tMSIP4.Text
            'Dim flag_chk1 As Boolean = RegularExpressions.Regex.IsMatch(strIn, cst_Pattern)
            If Not TIMS.CheckIP4_FMT(tMSIP4.Text) Then
                sReason &= "請檢查 掛載系統IP 格式有誤" & vbCrLf
            End If
        End If

        If tMSPORT.Text = "" Then
            sReason &= cst_必須填寫 & "掛載系統PORT" & vbCrLf
        Else
            If Not TIMS.IsNumeric2(tMSPORT.Text) Then
                '判斷是否為數字，正整數 '格式有誤，應為正整數數字格式('須大於0不可為0)
                sReason &= "掛載系統PORT-應為正整數數字格式(須大於0不可為0)" & vbCrLf
            End If
        End If
        If tMSMODULE.Text = "" Then
            sReason &= cst_必須填寫 & "掛載模組" & vbCrLf
        End If
        If v_ddlMSPROTOCOL = "" Then
            sReason &= cst_必須填寫 & "通訊協定" & vbCrLf
        End If

        If sReason <> "" Then rst = False
        Return rst
    End Function
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Dim sErrMsg1 As String = ""
        If Not CheckSaveData1(sErrMsg1) Then
            Common.MessageBox(Me, sErrMsg1)
            Exit Sub
        End If
        'divSch1.Visible = True
        'divEdt1.Visible = False
        sSaveData1()

        divSch1.Visible = True
        divEdt1.Visible = False
        sSearch1()
        sm.LastResultMessage = "儲存完畢"

    End Sub


    ''' <summary>
    ''' 單一資料
    ''' </summary>
    ''' <param name="sCmdArg"></param>
    Sub sLoadData1(ByRef sCmdArg As String)
        'Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        HID_WSENO.Value = TIMS.GetMyValue(sCmdArg, "WSENO")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.WSENO" & vbCrLf
        sql &= " ,a.MSNAME" & vbCrLf
        sql &= " ,a.MSIP4" & vbCrLf
        sql &= " ,a.MSPORT" & vbCrLf
        sql &= " ,a.MSMODULE" & vbCrLf
        sql &= " ,a.MSPROTOCOL" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM KEY_WEBSITE a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.WSENO=@WSENO" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("WSENO", Val(HID_WSENO.Value))
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        divSch1.Visible = True
        divEdt1.Visible = False
        If dt.Rows.Count = 0 Then
            sm.LastErrorMessage = "查無資料"
            Exit Sub
        End If

        divSch1.Visible = False
        divEdt1.Visible = True
        Dim dr1 As DataRow = dt.Rows(0)
        Call sClearlist1()
        Call sShowData1(dr1)
    End Sub

    Sub sShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub

        HID_WSENO.Value = Convert.ToString(dr("WSENO"))
        tMSNAME.Text = Convert.ToString(dr("MSNAME"))
        tMSIP4.Text = Convert.ToString(dr("MSIP4"))
        tMSPORT.Text = Convert.ToString(dr("MSPORT"))
        tMSMODULE.Text = Convert.ToString(dr("MSMODULE"))
        Dim vMSPROTOCOL As String = Convert.ToString(dr("MSPROTOCOL"))
        Common.SetListItem(ddlMSPROTOCOL, vMSPROTOCOL)
    End Sub

    Sub sClearlist1()
        HID_WSENO.Value = "" 'Convert.ToString(dr("WSENO"))
        tMSNAME.Text = "" 'Convert.ToString(dr("MSNAME"))
        tMSIP4.Text = "" 'Convert.ToString(dr("MSIP4"))
        tMSPORT.Text = "80" '443 'Convert.ToString(dr("MSPORT"))
        tMSMODULE.Text = "" 'Convert.ToString(dr("MSMODULE"))
        'Dim vMSPROTOCOL As String = Convert.ToString(dr("MSPROTOCOL"))
        Common.SetListItem(ddlMSPROTOCOL, "1")
    End Sub

    Sub sDeleteData1(ByVal sCmdArg As String)
        If sCmdArg = "" Then Exit Sub
        HID_WSENO.Value = TIMS.GetMyValue(sCmdArg, "WSENO")
        If HID_WSENO.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 1" & vbCrLf
        sql &= " FROM KEY_WEBSITE a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.WSENO=@WSENO" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("WSENO", Val(HID_WSENO.Value))
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        'PageControler1.Visible = False
        'DataGridTable.Visible = False
        'msg1.Text = "查無資料"
        'If dt.Rows.Count = 0 Then Exit Sub
        divSch1.Visible = True
        divEdt1.Visible = False
        If dt.Rows.Count = 1 Then
            Dim uSql As String = " UPDATE KEY_WEBSITE SET MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() WHERE WSENO=@WSENO"
            Dim iSql As String = " INSERT INTO KEY_WEBSITE_DEL SELECT * FROM KEY_WEBSITE WHERE WSENO=@WSENO"
            Dim dSql As String = " DELETE KEY_WEBSITE WHERE WSENO=@WSENO"

            Dim u_parms As New Hashtable
            u_parms.Clear()
            u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            u_parms.Add("WSENO", Val(HID_WSENO.Value))
            DbAccess.ExecuteNonQuery(uSql, objconn, u_parms)
            Dim id_parms As New Hashtable
            id_parms.Clear()
            id_parms.Add("WSENO", Val(HID_WSENO.Value))
            DbAccess.ExecuteNonQuery(iSql, objconn, id_parms)
            DbAccess.ExecuteNonQuery(dSql, objconn, id_parms)
            sm.LastResultMessage = "資料已刪除"
            Exit Sub
        End If

        'divSch1.Visible = False
        'divEdt1.Visible = True
        'Dim dr1 As DataRow = dt.Rows(0)
        'Call sClearlist1()
        'Call sShowData1(dr1)
    End Sub

    Sub sTestData1(ByVal sCmdArg As String)
        If sCmdArg = "" Then Exit Sub
        HID_WSENO.Value = TIMS.GetMyValue(sCmdArg, "WSENO")
        If HID_WSENO.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.WSENO" & vbCrLf
        sql &= " ,a.MSNAME" & vbCrLf
        sql &= " ,a.MSIP4" & vbCrLf
        sql &= " ,a.MSPORT" & vbCrLf
        sql &= " ,a.MSMODULE" & vbCrLf
        sql &= " ,a.MSPROTOCOL" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM KEY_WEBSITE a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.WSENO=@WSENO" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("WSENO", Val(HID_WSENO.Value))
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count <> 1 Then Exit Sub
        Dim dr1 As DataRow = dt.Rows(0)

        '"window.open('Test2.aspx','測試','menubar=no,status=no,scrollbars=yes,top=100,left=200,toolbar=no,width=450,height=300');"
        Dim s_MSPORT As String = Convert.ToString(dr1("MSPORT"))
        Dim s_Url As String = Convert.ToString(dr1("MSIP4"))
        Dim s_PROTOCOL As String = "http://"
        Dim s_MSPROTOCOL As String = Convert.ToString(dr1("MSPROTOCOL"))
        Select Case s_MSPROTOCOL
            Case "1" 'https
                s_PROTOCOL = "https://"
                If (s_MSPORT = "443") Then s_MSPORT = ""
            Case "2" 'http
                If (s_MSPORT = "80") Then s_MSPORT = ""
        End Select

        Dim strWinOpen As String = "" '組合連線字串。
        'strWinOpen = Url & "GUID=" & cGuid & "&RptID=" & Filename & NewStr
        strWinOpen = s_PROTOCOL & s_Url & If(s_MSPORT <> "", ":" & s_MSPORT, "")

        Dim strScript As String = ""
        strScript = "" & vbCrLf
        strScript &= "<script language=""javascript"">" & vbCrLf
        strScript &= ReportQuery.strWOScript(strWinOpen)
        strScript &= "</script>" & vbCrLf
        RegisterStartupScript(TIMS.xBlockName(), strScript)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Call sClearlist1()

        'Dim sCmdArg As String = ""
        Select Case e.CommandName
            Case cst_btnEdit
                'sCmdArg = e.CommandArgument
                sLoadData1(sCmdArg)
            Case cst_btnDel
                'sCmdArg = e.CommandArgument
                sDeleteData1(sCmdArg)
                sSearch1()
            Case cst_btnTest
                'sCmdArg = e.CommandArgument
                sTestData1(sCmdArg)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + dg1.PageSize * dg1.CurrentPageIndex
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtDel1 As LinkButton = e.Item.FindControl("lbtDel1")
                Dim lbtTest1 As LinkButton = e.Item.FindControl("lbtTest1")
                Dim labMSPROTOCOL As Label = e.Item.FindControl("labMSPROTOCOL")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "WSENO", Convert.ToString(drv("WSENO")))

                Common.SetListItem(ddlMSPROTOCOL, Convert.ToString(drv("MSPROTOCOL")))
                Dim v_Text1 As String = TIMS.GetListText(ddlMSPROTOCOL)
                If labMSPROTOCOL IsNot Nothing Then labMSPROTOCOL.Text = v_Text1

                lbtEdit.CommandArgument = sCmdArg
                lbtTest1.CommandArgument = sCmdArg
                'lbtDel1.Visible = False
                'If Convert.ToString(drv("CanDelete")) = "1" Then
                '    lbtDel1.Visible = True
                '    lbtDel1.CommandArgument = sCmdArg
                '    lbtDel1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                'End If
                lbtDel1.CommandArgument = sCmdArg
                lbtDel1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    Protected Sub btnAddNew1_Click(sender As Object, e As EventArgs) Handles btnAddNew1.Click
        sClearlist1()
        divSch1.Visible = False
        divEdt1.Visible = True

        tMSNAME.Text = If(sMSNAME.Text <> "", sMSNAME.Text, tMSNAME.Text)
        tMSIP4.Text = If(sMSIP4.Text <> "", sMSIP4.Text, tMSIP4.Text)
        tMSPORT.Text = If(sMSPORT.Text <> "", sMSPORT.Text, tMSPORT.Text)
        tMSMODULE.Text = If(sMSMODULE.Text <> "", sMSMODULE.Text, tMSMODULE.Text)
        Dim v_sddlMSPROTOCOL As String = TIMS.GetListValue(sddlMSPROTOCOL)
        'Dim vMSPROTOCOL As String = Convert.ToString(dr("MSPROTOCOL"))
        If v_sddlMSPROTOCOL <> "" Then Common.SetListItem(ddlMSPROTOCOL, v_sddlMSPROTOCOL)

    End Sub

    Protected Sub BtnReset2_Click(sender As Object, e As EventArgs) Handles BtnReset2.Click

        sMSNAME.Text = "" 'TIMS.ClearSQM(sMSNAME.Text)
        sMSIP4.Text = "" 'TIMS.ClearSQM(sMSIP4.Text)
        sMSPORT.Text = "" 'TIMS.ClearSQM(sMSPORT.Text)
        '數字/'非數字
        sMSPORT.Text = "" '
        'If (sMSPORT.Text <> "") Then sMSPORT.Text = "" 'If(TIMS.IsNumeric1(sMSPORT.Text), Val(sMSPORT.Text).ToString(), "")
        sMSMODULE.Text = "" 'TIMS.ClearSQM(sMSMODULE.Text)
        'Dim v_sddlMSPROTOCOL As String = TIMS.GetListValue(sddlMSPROTOCOL) '代碼
        Common.SetListItem(sddlMSPROTOCOL, "")

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
