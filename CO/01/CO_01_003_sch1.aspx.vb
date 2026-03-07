Partial Class CO_01_003_sch1
    Inherits AuthBasePage

    'ORG_TTQSVER-勾稽結果多筆

    Dim dtSENDVER As DataTable
    Dim dtRESULT As DataTable
    Const cst_SENDDATE As Integer = 0
    Const cst_SENDVER As Integer = 1
    Const cst_RESULT As Integer = 2
    Const cst_VALIDDATE As Integer = 3
    Const cst_EXTLICENS As Integer = 4
    Const cst_GOAL As Integer = 5
    Const cst_EXTEND As Integer = 6
    Const cst_ISSUEDATE As Integer = 7
    Const cst_MEMO As Integer = 8
    Const cst_EVALSCOPE As Integer = 9

    Dim strVTSID As String = ""
    Dim strCOMIDNO As String = ""
    Dim strORGID As String = ""
    Dim strOTTID As String = ""

    Dim strSENDDATE As String = "" '=aryTTQS(cst_SENDDATE  )
    Dim strSENDVER As String = "" '=aryTTQS(cst_SENDVER	)
    Dim strRESULT As String = "" '=aryTTQS(cst_RESULT	   )
    Dim strVALIDDATE As String = "" '=aryTTQS(cst_VALIDDATE	)
    Dim strEXTLICENS As String = "" '=aryTTQS(cst_EXTLICENS	)
    Dim strGOAL As String = "" '=aryTTQS(cst_GOAL		)
    Dim strEXTEND As String = "" '=aryTTQS(cst_EXTEND	   )
    Dim strISSUEDATE As String = "" '=aryTTQS(cst_ISSUEDATE	) 發文日期
    Dim strMEMO As String = "" '=aryTTQS(cst_MEMO		)
    Dim strEVALSCOPE As String = "" '=aryTTQS(cst_EVALSCOPE		)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''#Region "在這裡放置使用者程式碼以初始化網頁"
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            msg.Text = ""
            'Call create()
            Button2.Attributes("onclick") = "window.close();" '離開
            search1() '查詢／儲存 
            showDG1() '顯示
        End If

    End Sub

    ''' <summary>
    ''' webservice 取得資料
    ''' </summary>
    ''' <param name="str_comidno"></param>
    ''' <returns></returns>
    Function Get_ttqs_1(ByVal str_comidno As String) As String
        '"co_name=國立臺南大學&co_unit=69116104&ta_name=南分署
        'Dim str_co_name As String = "co_name=" & orgname
        'Dim cst_co_unit As String = "&co_unit="
        'Dim cst_ta_name As String = "&ta_name="
        'strBox1 = "co_name=" & TIMS.ClearSQM(orgname)
        'strBox1 &= "&co_unit=" & TIMS.ClearSQM(comidno)
        'strBox1 &= "&ta_name=" & TIMS.ClearSQM(distname)

        str_comidno = TIMS.ClearSQM(str_comidno)
        If str_comidno = "" Then Return String.Empty
        If Len(str_comidno) > 8 Then str_comidno = Left(str_comidno, 8) '取前8碼即可
        Dim strBox1 As String = ""
        strBox1 = ""
        strBox1 &= "co_unit=" & TIMS.ClearSQM(str_comidno) '統編
        Dim encBox1 As String = TIMS.EncryptAes(strBox1) '加密

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        'System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        System.Net.ServicePointManager.SecurityProtocol = 3072 'Net.SecurityProtocolType.Tls12

        'https://wltims.wda.gov.tw/GetJobMail3/ttqs_service_1.asmx
        Dim w As New ttqsservice1.ttqs_service_1
        Dim str_TTQS As String = String.Empty
        Try
            str_TTQS = w.ttqs(strBox1, encBox1)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            str_TTQS = ex.Message
        End Try
        'Restore SSL Certificate Validation Checking
        'System.Net.ServicePointManager.ServerCertificateValidationCallback = Nothing
        Return str_TTQS
        'Dim aryTTQS As String() = str_TTQS.Split(",")
        'If aryTTQS.Length >= 9 Then
        'End If
    End Function

    '修正轉代碼1
    Function GET_SENDVER_V1(ByVal VNAME As String) As String
        If VNAME = "" Then Return ""
        Dim rst As String = VNAME
        Dim fff As String = "VNAME='" & VNAME & "'"
        If dtSENDVER.Select(fff).Length > 0 Then
            rst = dtSENDVER.Select(fff)(0)("VID")
        End If
        Return rst
    End Function

    '修正轉代碼2
    Function GET_RESULT_V1(ByVal VNAME As String) As String
        If VNAME = "" Then Return ""
        Dim rst As String = VNAME
        Dim fff As String = "VNAME='" & VNAME & "'"
        If dtRESULT.Select(fff).Length > 0 Then
            rst = dtRESULT.Select(fff)(0)("VID")
        End If
        Return rst
    End Function

    '修正轉代碼3
    Function GET_EXTEND_V1(ByVal VNAME As String) As String
        If VNAME = "" Then Return ""
        Dim rst As String

        Select Case VNAME
            Case "是"
                rst = "Y"
            Case "否"
                rst = "N"
            Case Else
                rst = VNAME
        End Select
        Return rst
    End Function

    ''' <summary>
    ''' 整理 TTQS 取得的資料
    ''' </summary>
    ''' <param name="str_TTQS"></param>
    ''' <returns></returns>
    Function Get_ttqs_2(ByVal str_TTQS As String) As Boolean
        Dim rst As Boolean = False
        Dim aryTTQS As String() = str_TTQS.Split(",")

        strSENDDATE = "" 'aryTTQS(cst_SENDDATE  )
        strSENDVER = "" 'aryTTQS(cst_SENDVER	)
        strRESULT = "" 'aryTTQS(cst_RESULT	   )
        strVALIDDATE = "" 'aryTTQS(cst_VALIDDATE	)
        strEXTLICENS = "" 'aryTTQS(cst_EXTLICENS	)
        strGOAL = "" 'aryTTQS(cst_GOAL		)
        strEXTEND = "" 'aryTTQS(cst_EXTEND	   )
        strISSUEDATE = "" 'aryTTQS(cst_ISSUEDATE	)
        strMEMO = "" 'aryTTQS(cst_MEMO		)
        strEVALSCOPE = "" '=aryTTQS(cst_EVALSCOPE		)
        'strIMPORTDATE = "" 'aryTTQS(cst_ETDATE()  )
        'strMODIFYACCT = "" 'aryTTQS(cst_MODIFYACCT)
        'strMODIFYDATE = "" 'aryTTQS(cst_ETDATE())
        If aryTTQS.Length >= 9 Then
            strSENDDATE = aryTTQS(cst_SENDDATE)
            strSENDVER = GET_SENDVER_V1(aryTTQS(cst_SENDVER))
            strRESULT = GET_RESULT_V1(aryTTQS(cst_RESULT))
            strVALIDDATE = aryTTQS(cst_VALIDDATE)
            strEXTLICENS = aryTTQS(cst_EXTLICENS)
            strGOAL = aryTTQS(cst_GOAL)
            strEXTEND = GET_EXTEND_V1(aryTTQS(cst_EXTEND))
            strISSUEDATE = aryTTQS(cst_ISSUEDATE)
            strMEMO = aryTTQS(cst_MEMO)
            strEVALSCOPE = aryTTQS(cst_EVALSCOPE)
            'strMODIFYACCT = aryTTQS(cst_MODIFYACCT)
            rst = True
        End If
        Return rst
    End Function

    ''' <summary>
    ''' 新增-ORG_TTQSVER-確認
    ''' </summary>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Function INSERT_ORGTTQSREV(ByRef oConn As SqlConnection) As Integer
        Dim rst As Integer = -1

        Dim flag_DATA_NG As Boolean = False
        strSENDDATE = TIMS.ClearSQM(strSENDDATE)

        If strSENDDATE = "" Then flag_DATA_NG = True '若為空白不處理
        Dim oSENDDATE As Object = TIMS.Cdate2(strSENDDATE)
        If oSENDDATE Is Convert.DBNull Then flag_DATA_NG = True '若為空值不處理
        If flag_DATA_NG Then Return rst

        strISSUEDATE = TIMS.ClearSQM(strISSUEDATE)
        If strISSUEDATE = "" Then flag_DATA_NG = True '若為空白不處理
        Dim oISSUEDATE As Object = TIMS.Cdate2(strISSUEDATE)
        If oISSUEDATE Is Convert.DBNull Then flag_DATA_NG = True '若為空值不處理
        If flag_DATA_NG Then Return rst

        'ORG_TTQSVER
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO ORG_TTQSVER (" & vbCrLf
        sql &= " VTSID,ORGID,COMIDNO,SENDDATE,SENDVER,RESULT,VALIDDATE,EXTLICENS,GOAL,EXTEND,ISSUEDATE,MEMO,EVALSCOPE" & vbCrLf
        sql &= " ,IMPORTDATE,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @VTSID,@ORGID,@COMIDNO,@SENDDATE,@SENDVER,@RESULT,@VALIDDATE,@EXTLICENS,@GOAL,@EXTEND,@ISSUEDATE,@MEMO,@EVALSCOPE" & vbCrLf
        sql &= " ,GETDATE(),@MODIFYACCT,GETDATE()" & vbCrLf
        sql &= " )" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)
        '新增
        Dim iVTSID As Integer = 0
        iVTSID = DbAccess.GetNewId(oConn, "ORG_TTQSVER_VTSID_SEQ,ORG_TTQSVER,VTSID")
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("VTSID", SqlDbType.Int).Value = iVTSID
            .Parameters.Add("ORGID", SqlDbType.Int).Value = strORGID
            .Parameters.Add("COMIDNO", SqlDbType.NVarChar).Value = strCOMIDNO

            .Parameters.Add("SENDDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(strSENDDATE)
            .Parameters.Add("SENDVER", SqlDbType.NVarChar).Value = strSENDVER
            .Parameters.Add("RESULT", SqlDbType.NVarChar).Value = strRESULT
            .Parameters.Add("VALIDDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(strVALIDDATE)
            .Parameters.Add("EXTLICENS", SqlDbType.NVarChar).Value = If(strEXTLICENS <> "", strEXTLICENS, Convert.DBNull)
            .Parameters.Add("GOAL", SqlDbType.NVarChar).Value = If(strGOAL <> "", strGOAL, Convert.DBNull)
            .Parameters.Add("EXTEND", SqlDbType.NVarChar).Value = If(strEXTEND <> "", strEXTEND, Convert.DBNull)
            .Parameters.Add("ISSUEDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(strISSUEDATE)
            .Parameters.Add("MEMO", SqlDbType.NVarChar).Value = strMEMO
            .Parameters.Add("EVALSCOPE", SqlDbType.NVarChar).Value = strEVALSCOPE
            '.Parameters.Add(rIMPORTDATE", SqlDbType.NVarChar).Value = strIMPORTDATE
            .Parameters.Add("MODIFYACCT", SqlDbType.NVarChar).Value = sm.UserInfo.UserID
            '.Parameters.Add("strMODIFYDATE", SqlDbType.NVarChar).Value = strMODIFYDATE
            rst = .ExecuteNonQuery()
        End With
        Return rst
    End Function

    ''' <summary>
    ''' 檢核時限內是否有資料
    ''' </summary>
    ''' <param name="s_parms"></param>
    ''' <returns></returns>
    Function check15(ByRef s_parms As Hashtable) As Boolean
        Dim rst As Boolean = False
        Dim s_COMIDNO As String = TIMS.GetMyValue2(s_parms, "COMIDNO")
        Dim s_ORGID As String = TIMS.GetMyValue2(s_parms, "ORGID")
        If s_COMIDNO = "" Then Return rst
        If s_ORGID = "" Then Return rst

        Dim sql As String = ""
        sql &= " SELECT * FROM ORG_TTQSVER WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND MODIFYDATE >= DATEADD(n,-15,GETDATE())" & vbCrLf
        sql &= " AND COMIDNO = @COMIDNO" & vbCrLf
        sql &= " AND ORGID = @ORGID" & vbCrLf
        'sql &= " --https://www.fooish.com/sql/sql-server-dateadd-function.html" & vbCrLf
        'sql &= " --增加或減少指定的時間間隔" & vbCrLf

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("COMIDNO", s_COMIDNO)
        parms.Add("ORGID", s_ORGID)

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count > 0 Then rst = True '有資料
        Return rst
    End Function

    ''' <summary>
    ''' 刪除暫存資料
    ''' </summary>
    ''' <param name="s_parms"></param>
    Sub clear15(ByRef s_parms As Hashtable)

        Dim s_COMIDNO As String = TIMS.GetMyValue2(s_parms, "COMIDNO")
        Dim s_ORGID As String = TIMS.GetMyValue2(s_parms, "ORGID")
        If s_COMIDNO = "" Then Return
        If s_ORGID = "" Then Return

        Dim sql As String = ""
        sql &= " DELETE ORG_TTQSVER" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND MODIFYDATE >= DATEADD(n,-15,GETDATE())" & vbCrLf
        sql &= " AND COMIDNO = @COMIDNO" & vbCrLf
        sql &= " AND ORGID = @ORGID" & vbCrLf
        'sql &= " --https://www.fooish.com/sql/sql-server-dateadd-function.html" & vbCrLf
        'sql &= " --增加或減少指定的時間間隔" & vbCrLf

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("COMIDNO", s_COMIDNO)
        parms.Add("ORGID", s_ORGID)

        DbAccess.ExecuteNonQuery(sql, objconn, parms)
    End Sub

    Sub search1()
        'Request
        Hid_comidno.Value = TIMS.sUtl_GetRqValue(Me, "COMIDNO")
        Hid_ORGID.Value = TIMS.sUtl_GetRqValue(Me, "ORGID")
        Hid_OTTID.Value = TIMS.sUtl_GetRqValue(Me, "OTTID")
        strCOMIDNO = TIMS.ClearSQM(Hid_comidno.Value)
        strORGID = TIMS.ClearSQM(Hid_ORGID.Value)
        strOTTID = TIMS.ClearSQM(Hid_OTTID.Value)
        If strCOMIDNO = "" Then Return '統編為空，無法執行
        If strORGID = "" Then Return '機構id為空，不可執行

        Dim parms As New Hashtable
        parms.Add("COMIDNO", strCOMIDNO)
        parms.Add("ORGID", strORGID)

        '檢核資料15分鐘內
        Dim flag_check15 As String = check15(parms)
        If flag_check15 Then Return '有資料離開

        '先進行清理
        clear15(parms)

        '查無勾稽資料，進行勾稽
        Const cst_BR_TXT1 As String = "<br>"
        Dim str_TTQS_BR1 As String = Get_ttqs_1(strCOMIDNO)
        Dim flag_OK As Boolean = False '(false:異常程序)
        'Dim str_TTQS As String = Get_ttqs_1(strOrgname, strComidno, strDistname)
        If str_TTQS_BR1.Length > 0 Then
            Dim sql As String = "SELECT VID,VNAME FROM dbo.V_SENDVER ORDER BY VID"
            dtSENDVER = DbAccess.GetDataTable(sql, objconn)
            Dim sql2 As String = "SELECT VID,VNAME FROM dbo.V_RESULT ORDER BY VID"
            dtRESULT = DbAccess.GetDataTable(sql2, objconn)

            Dim ary_TTQS_BR() As String = Split(str_TTQS_BR1, cst_BR_TXT1)
            For Each str_TTQS As String In ary_TTQS_BR
                If str_TTQS = "" Then Exit For '查無有效字元
                If str_TTQS.Length < 9 Then Exit For '與基數不合

                flag_OK = Get_ttqs_2(str_TTQS)
                If flag_OK Then INSERT_ORGTTQSREV(objconn) '填入
            Next
        End If
        '(異常回傳錯誤資訊)
        If Not flag_OK Then Hid_RESULT1.Value = TIMS.ClearSQM(str_TTQS_BR1)

    End Sub

    Sub showDG1()
        ORGNAME.Text = TIMS.GET_OrgName(strORGID, objconn)

        '評核日期-SENDDATE	發文日期-ISSUEDATE	有效期限-VALIDDATE
        Dim parms As New Hashtable
        Dim sql As String = ""
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.VTSID ASC) AS ROWNUM" & vbCrLf
        sql &= " ,a.VTSID " & vbCrLf '/*PK*/
        sql &= " ,a.COMIDNO" & vbCrLf
        sql &= " ,format(a.SENDDATE,'yyyy/MM/dd') SENDDATE" & vbCrLf
        sql &= " ,a.SENDVER" & vbCrLf
        sql &= " ,a.RESULT" & vbCrLf
        sql &= " ,format(a.VALIDDATE,'yyyy/MM/dd') VALIDDATE" & vbCrLf
        sql &= " ,a.EXTLICENS" & vbCrLf
        sql &= " ,a.GOAL" & vbCrLf
        sql &= " ,a.EXTEND" & vbCrLf
        sql &= " ,format(a.ISSUEDATE,'yyyy/MM/dd') ISSUEDATE" & vbCrLf
        sql &= " ,a.MEMO" & vbCrLf
        sql &= " ,a.EVALSCOPE" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(a.EVALSCOPE,'-')='-' then a.MEMO ELSE concat(a.MEMO,'(',a.EVALSCOPE,')') END MEMO2" & vbCrLf
        sql &= " ,a.IMPORTDATE" & vbCrLf
        sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
        sql &= " ,v2.VNAME RESULT_N" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.ORG_TTQSVER a WITH(NOLOCK)" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=a.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=a.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " WHERE 0=0" & vbCrLf
        sql &= " AND a.COMIDNO=@COMIDNO" & vbCrLf
        sql &= " AND a.ORGID=@ORGID" & vbCrLf

        parms.Clear()
        parms.Add("COMIDNO", strCOMIDNO)
        parms.Add("ORGID", strORGID)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = TIMS.cst_NODATAMsg1
        If dt.Rows.Count = 0 Then Return

        msg.Text = ""
        DataGrid1.Visible = True
        Button1.Enabled = True

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Sub back2()
        Dim Script As String = ""
        Script &= "<script type=""text/javascript"" >" & vbCrLf
        'Script &= " var ActNo1 = opener.document.getElementById('ActNo1');" & vbCrLf
        'Script &= " if(ActNo1){ActNo1.value='" & drSB4("actno") & "';}" & vbCrLf
        'Script &= " var ActName = opener.document.getElementById('ActName');" & vbCrLf
        'Script &= " if(ActName){ActName.value='" & drSB4("COMNAME") & "';}" & vbCrLf
        'Script &= " var hidSB4ID = opener.document.getElementById('hidSB4ID');" & vbCrLf
        'Script &= " if(hidSB4ID){hidSB4ID.value='" & SB4ID & "';}" & vbCrLf
        Script &= " var btnQuery = opener.document.getElementById('btnQuery');" & vbCrLf
        Script &= " if(btnQuery){btnQuery.click();}" & vbCrLf

        Script &= " window.top.opener = null;" & vbCrLf
        Script &= " window.close();" & vbCrLf
        Script &= "</script>"
        Page.RegisterStartupScript(TIMS.xBlockName(), Script)
    End Sub

    Function CHECK_VALOK() As Boolean
        Dim rst As Boolean = False
        'Dim sql As String = ""
        'Dim strVTSID As String = ""
        'Dim SMDATE As String = ""
        'Dim FMDATE As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Radio1 As HtmlInputRadioButton = eItem.FindControl("Radio1")
            Dim Hid_VTSID As HiddenField = eItem.FindControl("Hid_VTSID")
            If Radio1.Checked AndAlso Hid_VTSID.Value <> "" Then
                strVTSID = Hid_VTSID.Value
                rst = True
                Return rst
            End If
        Next
        Dim s_errmsg1 As String = "查無資料，無法回傳值"
        If strVTSID = "" Then
            'Common.MessageBox2(Me, "查無資料，無法回傳值")
            msg.Text = s_errmsg1
            Return rst
        End If
        'Common.MessageBox2(Me, "查無資料，無法回傳值")
        msg.Text = s_errmsg1
        Return rst
    End Function

    '存一份log
    Sub savelog1(ByRef s_OTTID As String, ByRef oConn As SqlConnection)
        If s_OTTID = "" Then Exit Sub

        Dim parms As New Hashtable From {{"OTTID", s_OTTID}, {"MODIFYACCT", sm.UserInfo.UserID}}
        Dim sql As String = "UPDATE ORG_TTQS2 SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE() WHERE OTTID=@OTTID"
        DbAccess.ExecuteNonQuery(sql, oConn, parms)

        '備份資料 log
        Dim parms2 As New Hashtable From {{"OTTID", s_OTTID}}
        Dim sql2 As String = "SELECT * FROM ORG_TTQS2 WHERE OTTID=@OTTID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql2, oConn, parms2)
        Call TIMS.InsertDelTableLog("ORG_TTQS2_LOG", dt, oConn)
    End Sub

    Sub save1(ByRef s_parms As Hashtable)
        '取得 ORG_TTQSVER 流水號 VTSID
        '更新 ORG_TTQS2 流水號 OTTID
        Dim s_VTSID As String = TIMS.GetMyValue2(s_parms, "VTSID")
        Dim s_OTTID As String = TIMS.GetMyValue2(s_parms, "OTTID")
        Dim s_ORGID As String = TIMS.GetMyValue2(s_parms, "ORGID")
        Dim s_COMIDNO As String = TIMS.GetMyValue2(s_parms, "COMIDNO")

        Dim parms As New Hashtable
        Dim sql As String = ""
        sql &= " SELECT VTSID,ORGID,COMIDNO,SENDDATE,SENDVER,RESULT" & vbCrLf
        sql &= " ,VALIDDATE,EXTLICENS,GOAL,EXTEND,ISSUEDATE,MEMO,EVALSCOPE" & vbCrLf
        sql &= " ,IMPORTDATE,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ORG_TTQSVER" & vbCrLf
        sql &= " WHERE VTSID=@VTSID AND ORGID=@ORGID AND COMIDNO=@COMIDNO" & vbCrLf
        parms.Clear()
        parms.Add("VTSID", s_VTSID)
        parms.Add("ORGID", s_ORGID)
        parms.Add("COMIDNO", s_COMIDNO)
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr1 Is Nothing Then Return

        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT OTTID,ORGID,COMIDNO,TPLANID,DISTID,YEARS,MONTHS,SENDDATE,SENDVER,RESULT" & vbCrLf
        sql &= " ,VALIDDATE,EXTLICENS,GOAL,EXTEND,ISSUEDATE,MEMO,EVALSCOPE" & vbCrLf
        sql &= " ,IMPORTACCT,IMPORTDATE,CONFIRM,CONFIRMACCT,CONFIRMDATE,REASON1,APPLIEDRESULT" & vbCrLf
        sql &= " ,APPLIEDACCT,APPLIEDDATE,LOCKACCT,LOCKDATE" & vbCrLf
        sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ORG_TTQS2" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND OTTID=@OTTID" & vbCrLf
        sql &= " AND ORGID=@ORGID" & vbCrLf
        sql &= " AND COMIDNO=@COMIDNO" & vbCrLf
        parms.Clear()
        parms.Add("OTTID", s_OTTID)
        parms.Add("ORGID", s_ORGID)
        parms.Add("COMIDNO", s_COMIDNO)
        Dim dr2 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr2 Is Nothing Then Return '查無資料-異常

        '存一份log
        savelog1(s_OTTID, objconn)

        'Dim sql As String = ""
        'ORG_TTQS2
        sql = "" & vbCrLf
        sql &= " UPDATE ORG_TTQS2" & vbCrLf
        sql &= " SET SENDDATE=@SENDDATE ,SENDVER=@SENDVER" & vbCrLf
        sql &= " ,RESULT=@RESULT" & vbCrLf
        sql &= " ,VALIDDATE=@VALIDDATE" & vbCrLf
        sql &= " ,EXTLICENS=@EXTLICENS" & vbCrLf
        sql &= " ,GOAL=@GOAL" & vbCrLf
        sql &= " ,EXTEND=@EXTEND" & vbCrLf
        sql &= " ,ISSUEDATE=@ISSUEDATE" & vbCrLf
        sql &= " ,MEMO=@MEMO" & vbCrLf
        sql &= " ,EVALSCOPE=@EVALSCOPE" & vbCrLf
        '更新，清理審核狀況
        sql &= " ,APPLIEDRESULT=NULL" & vbCrLf
        sql &= " ,APPLIEDACCT=NULL" & vbCrLf
        sql &= " ,APPLIEDDATE=GETDATE()" & vbCrLf
        sql &= " ,IMPORTACCT=@IMPORTACCT" & vbCrLf
        sql &= " ,IMPORTDATE=GETDATE()" & vbCrLf
        sql &= " ,RENEWACCT=@RENEWACCT" & vbCrLf
        sql &= " ,RENEWDATE=GETDATE()" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND OTTID=@OTTID" & vbCrLf
        Dim u_sql As String = sql

        'UPDATE-ORG_TTQS2
        parms.Clear()
        parms.Add("SENDDATE", dr1("SENDDATE"))
        parms.Add("SENDVER", dr1("SENDVER"))
        parms.Add("RESULT", dr1("RESULT"))
        parms.Add("VALIDDATE", dr1("VALIDDATE"))
        parms.Add("EXTLICENS", dr1("EXTLICENS"))
        parms.Add("GOAL", dr1("GOAL"))
        parms.Add("EXTEND", dr1("EXTEND"))
        parms.Add("ISSUEDATE", dr1("ISSUEDATE"))
        parms.Add("MEMO", dr1("MEMO"))
        parms.Add("EVALSCOPE", dr1("EVALSCOPE"))
        parms.Add("IMPORTACCT", sm.UserInfo.UserID)
        parms.Add("RENEWACCT", sm.UserInfo.UserID)
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("OTTID", s_OTTID)
        DbAccess.ExecuteNonQuery(u_sql, objconn, parms)
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        strCOMIDNO = TIMS.ClearSQM(Hid_comidno.Value)
        strORGID = TIMS.ClearSQM(Hid_ORGID.Value)
        strOTTID = TIMS.ClearSQM(Hid_OTTID.Value)

        '檢核
        Dim flag_chk1 As Boolean = CHECK_VALOK()
        If Not flag_chk1 Then Return '檢核有誤離開

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("COMIDNO", strCOMIDNO)
        parms.Add("ORGID", strORGID)
        parms.Add("OTTID", strOTTID)
        parms.Add("VTSID", strVTSID)
        '存檔
        Call save1(parms)
        '回上層按一個search
        Call back2()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Hid_VTSID As HiddenField = e.Item.FindControl("Hid_VTSID")
                Hid_VTSID.Value = Convert.ToString(drv("VTSID"))
                'e.Item.Cells(1).Text = e.Item.ItemIndex + 1
                Radio1.Attributes("onclick") = "checkRadio(" & e.Item.ItemIndex + 1 & ");"
                Radio1.Value = Convert.ToString(drv("VTSID"))


                'If Convert.ToString(drv("NOUSE")) = "Y" Then
                '    '不可使用
                '    Radio1.Disabled = True
                '    TIMS.Tooltip(Radio1, "不可被點選")
                'Else
                '    '可使用
                '    'Radio1.Attributes("onclick") = "checkRadio(" & e.Item.ItemIndex + 1 & ");"
                '    'Radio1.Value = drv("SB4ID")
                '    Hid_VTSID.Value = Convert.ToString(drv("VTSID"))
                '    'Hid_SMDATE.Value = Convert.ToString(drv("SMDATE"))
                '    'Hid_FMDATE.Value = Convert.ToString(drv("FMDATE"))
                'End If

                'Dim flag_subEcfa As Boolean = False '不是ECFA
                'If sNoECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 _
                '    AndAlso sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                '    LabECFA.Text = ""
                '    If TIMS.CheckIsECFA(Me, Convert.ToString(drv("ACTNO")), "", Convert.ToString(drv("TODAY1")), objconn) = True Then
                '        flag_subEcfa = True '是ECFA
                '        LabECFA.Text = "是"
                '        Hid_ECFA_YES.Value = TIMS.cst_YES
                '    End If
                'End If
                'If sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) > -1 Then
                '    flag_subEcfa = True '是ECFA
                '    LabECFA.Text = "是"
                '    Hid_ECFA_YES.Value = TIMS.cst_YES
                'End If
                'If flag_subEcfa Then
                '    '是ECFA
                '    If sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                '        '是ECFA 不須要再檢核
                '        If sOkECFA_ACTNO <> "" Then sOkECFA_ACTNO &= ","
                '        sOkECFA_ACTNO &= Convert.ToString(drv("ACTNO"))
                '    End If
                'Else
                '    '不是ECFA
                '    If sNoECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                '        '不是ECFA 不須要再檢核
                '        If sNoECFA_ACTNO <> "" Then sNoECFA_ACTNO &= ","
                '        sNoECFA_ACTNO &= Convert.ToString(drv("ACTNO"))
                '    End If
                'End If
                'If drv("SB4ID").ToString = HidSBID.Value Then
                '    Radio1.Checked = True
                'End If

                'Dim sActNoType As String = TIMS.Get_ACTNOTYPE1(Convert.ToString(drv("ActNo")))
                'LabActNoType.Text = sActNoType

                'Dim sChangeMode As String = TIMS.Get_CHANGEMODE1(Convert.ToString(drv("ChangeMode")))
                'LabChangeMode.Text = sChangeMode

        End Select
    End Sub
End Class