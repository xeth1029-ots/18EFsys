Imports System.Xml

Public Class BliClass1

#Region "COMMON1"
    'Dim iCOUNT As Integer = 0 '取得有效資料筆數
    'Dim ErrorMsg As String = "" '錯誤暫存

    'Dim sGUID As String = "" '序號(System.Guid)
    'Dim cmdSYSNAME As String = "TRA001" 'TRA001 '訓練 'TRA002 '生活津貼 FOR001'外勞
    'Dim sIDNO As String = "" 'IDNO(投保人身分證號)
    'Dim sINAME As String = "" 'NAME(投保人姓名)  
    'Dim sBIRTH As String = "" 'BIRTH(出生年月日(YYYYMMDD))
    'Dim sUTYPE As String = "" 'FType'保險種類。 (A表示勞保+就保，L表示勞保，V農保)
    'Dim sBDATE As String = "" '投保起日(YYYYMMDD西元年月日)
    'Dim sEDATE As String = "" '投保迄日(YYYYMMDD西元年月日)

    'Dim gCmdStr As String = "" '讀取傳入的外部參數
    'Dim strDetail As String = "" '取得未整理的XML資訊

    Const Cst_Errmsg_1 As String = "查詢資料找不到" '不算錯誤
    Const Cst_Errmsg_2 As String = "查詢資料格式不符"
    Const Cst_Errmsg_3 As String = "不允許的查詢"
    Const Cst_Errmsg_4 As String = "不明的錯誤"
    Const Cst_Errmsg_5 As String = "程式內部錯誤"
    Const Cst_Errmsg_6 As String = "錯誤碼:6"
    Const Cst_Errmsg_7 As String = "身分證重號"
    Const Cst_Errmsg_8 As String = "資料超過100筆"

    ''' <summary>
    ''' 轉換日期 (yyyy/MM/dd->yyyyMMdd)
    ''' </summary>
    ''' <param name="strDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetReqDate(ByVal strDate As String) As String
        Dim rst As String = ""
        If strDate = "" Then Return rst
        rst = Replace(strDate, "/", "")
        Return rst
    End Function

    '傳 日期字元 (yyyyMMdd) 換成 西元年月日(yyyy/MM/dd)
    Public Shared Function GetADDate(ByVal StrDate As String) As String
        Dim Rst As String = ""
        If StrDate <> "" AndAlso StrDate.Length = 8 AndAlso IsNumeric(StrDate) Then
            Rst = Mid(StrDate, 1, 4) & "/" & Mid(StrDate, 5, 2) & "/" & Mid(StrDate, 7, 2)
            If Not IsDate(Rst) Then Rst = ""
        End If
        Return Rst
    End Function

    '傳 民國日期字元(YYYMMDD) 換成 西元年月日(yyyy/MM/dd)
    Public Shared Function GetRocADDate(ByVal StrDate As String) As String
        Dim Rst As String = ""
        If StrDate <> "" AndAlso StrDate.Length = 7 AndAlso IsNumeric(StrDate) Then
            Rst = CStr(Mid(StrDate, 1, 3) + 1911) & "/" & Mid(StrDate, 4, 2) & "/" & Mid(StrDate, 6, 2)
            If Not IsDate(Rst) Then Rst = ""
        End If
        Return Rst
    End Function
#End Region

#Region "勞保勾稽1"
    'ACTNObli
    Public Shared Function Get_SSValue(ByVal SETID As String, ByVal ENTERDATE As String, ByVal SERNUM As String, ByRef oConn As SqlConnection) As String
        Dim strSS As String = ""
        Dim drT As DataRow = TIMS.Get_ENTERTYPE1(SETID, ENTERDATE, SERNUM, oConn)
        If drT Is Nothing Then Return strSS
        'ACTNObli
        TIMS.SetMyValue(strSS, "IDNO", $"{drT("IDNO")}")
        TIMS.SetMyValue(strSS, "BIRTH", TIMS.Cdate3(drT("BIRTHDAY")))
        TIMS.SetMyValue(strSS, "CNAME", $"{drT("CNAME")}")
        TIMS.SetMyValue(strSS, "STDATE1", TIMS.Cdate3(drT("STDATE1")))
        TIMS.SetMyValue(strSS, "STDATE2", TIMS.Cdate3(drT("STDATE1")))
        TIMS.SetMyValue(strSS, "SETID", $"{drT("SETID")}")
        TIMS.SetMyValue(strSS, "ENTERDATE", TIMS.Cdate3(drT("ENTERDATE")))
        TIMS.SetMyValue(strSS, "SERNUM", $"{drT("SERNUM")}")
        TIMS.SetMyValue(strSS, "OCID", $"{drT("OCID1")}")
        Return strSS
    End Function

    '先判斷該員是否曾勾稽過，若無才需呼叫webservice 取得勞保勾稽(加保(4))
    Public Shared Sub Get_SELRESULTBNG(ByRef MyPage As Page,
                                       ByVal strSS As String,
                                       ByVal oConn As SqlConnection)
        If strSS = "" Then Exit Sub
        Dim rIDNO As String = TIMS.GetMyValue(strSS, "IDNO")
        Dim rBIRTH As String = TIMS.GetMyValue(strSS, "BIRTH") 'yyyy/MM/dd
        Dim rCNAME As String = TIMS.GetMyValue(strSS, "CNAME") '投保人姓名(in)
        Dim rSTDATE1 As String = TIMS.GetMyValue(strSS, "STDATE1") 'yyyy/MM/dd
        Dim rSTDATE2 As String = TIMS.GetMyValue(strSS, "STDATE2") 'yyyy/MM/dd
        Dim SETID As String = TIMS.GetMyValue(strSS, "SETID")
        Dim ENTERDATE As String = TIMS.GetMyValue(strSS, "ENTERDATE") 'yyyy/MM/dd
        Dim SERNUM As String = TIMS.GetMyValue(strSS, "SERNUM")
        Dim OCID As String = TIMS.GetMyValue(strSS, "OCID")
        If rIDNO = "" OrElse rSTDATE1 = "" OrElse rSTDATE2 = "" Then Exit Sub
        If SETID = "" OrElse ENTERDATE = "" OrElse SERNUM = "" Then Exit Sub
        If OCID = "" Then Exit Sub
        Dim blIsExists As Boolean = ChkSELRESULTBNGExists(strSS, oConn)
        If blIsExists Then Exit Sub

        'TRA001 '訓練 'TRA002 '生活津貼 FOR001'外勞 'TRA011產投
        'Const cst_gSysName As String = "TRA001"
        Dim strXML As String = String.Empty
        'Dim blIsExists As Boolean = False

        '勞保+就保A、勞保L、自願職災保險V、農保F、農民自願職災保險FV
        'Const cst_勞保就保 As String = "A" 'Const cst_自願職災保險 As String = "V" 'Const cst_農保 As String = "F" 'Const cst_農民自願職災保險 As String = "FV"
        Dim s_Bli_TYPE As String = "A,V,F,FV"
        Dim aBli_TYPE As String() = s_Bli_TYPE.Split(",")
        'For Each strUType As String In aBli_TYPE

        Dim strMsg As String = ""
        'Dim strUType As String = ""
        'Dim strIName As String = "" 'String.Empty '投保人姓名
        Dim strBIRTH As String = GetReqDate(rBIRTH) '(yyyy/MM/dd)->yyyyMMdd
        Dim strBDate As String = GetReqDate(rSTDATE1) '(yyyy/MM/dd)->yyyyMMdd
        Dim strEDate As String = GetReqDate(rSTDATE2) '(yyyy/MM/dd)->yyyyMMdd

        Dim htSS As New Hashtable From {
            {"rIDNO", rIDNO},
            {"rBIRTH", rBIRTH}, 'yyyy/MM/dd
            {"rCNAME", rCNAME}, '投保人姓名(in)
            {"rSTDATE1", rSTDATE1}, 'yyyy/MM/dd
            {"rSTDATE2", rSTDATE2}, 'yyyy/MM/dd
            {"SETID", SETID},
            {"ENTERDATE", ENTERDATE}, 'yyyy/MM/dd
            {"SERNUM", SERNUM},
            {"OCID", OCID},
            {"strBIRTH", strBIRTH}, 'yyyyMMdd
            {"strBDate", strBDate}, 'yyyyMMdd
            {"strEDate", strEDate}, 'yyyyMMdd
            {"strUType", ""},
            {"strXML", ""}
        } '= Nothing

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Try
            'https://wltims.wda.gov.tw/bli_wsv4/get_data2.asmx ,'https://wltims.wda.gov.tw/bli_wsv4/get_data3.asmx
            Dim bliWs As New bli_wsv4.get_data3
            For Each strUType As String In aBli_TYPE
                'strUType = s_UType 'cst_勞保就保 '查詢勞保+就保勾稽資料
                '呼叫勞保勾稽webservice
                strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
                htSS("strUType") = strUType
                htSS("strXML") = strXML
                '解析xml內容
                strMsg = parseXML2SAVE(MyPage, htSS, oConn)
                Select Case strMsg
                    Case "7"
                        If rCNAME = "" Then Exit Select
                        'strIName = strCNAME '(使用姓名重新勾稽)
                        'strUType = s_UType 'cst_勞保就保 '查詢勞保+就保勾稽資料
                        '呼叫勞保勾稽webservice
                        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
                        htSS("strUType") = strUType
                        htSS("strXML") = strXML
                        '解析xml內容
                        strMsg = parseXML2SAVE(MyPage, htSS, oConn)
                End Select
            Next

            'strUType = cst_自願職災保險 '查詢自願職災保險勾稽資料
            ''呼叫勞保勾稽webservice
            'strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
            'htSS("strUType") = strUType
            'htSS("strXML") = strXML

            ''解析xml內容
            'strMsg = parseXML2SAVE(MyPage, htSS, oConn)
            'Select Case strMsg
            '    Case "7"
            '        If rCNAME = "" Then Exit Select
            '        'strIName = strCNAME '(使用姓名重新勾稽)
            '        strUType = cst_自願職災保險 '查詢勞保+就保勾稽資料
            '        '呼叫勞保勾稽webservice
            '        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
            '        htSS("strUType") = strUType
            '        htSS("strXML") = strXML

            '        '解析xml內容
            '        strMsg = parseXML2SAVE(MyPage, htSS, oConn)
            'End Select

            'strUType = cst_農保 '查詢農保勾稽資料
            ''呼叫勞保勾稽webservice
            'strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
            'htSS("strUType") = strUType
            'htSS("strXML") = strXML

            ''解析xml內容
            'strMsg = parseXML2SAVE(MyPage, htSS, oConn)
            'Select Case strMsg
            '    Case "7"
            '        If rCNAME = "" Then Exit Select
            '        'strIName = strCNAME '(使用姓名重新勾稽)
            '        strUType = cst_農保 '查詢勞保+就保勾稽資料
            '        '呼叫勞保勾稽webservice
            '        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
            '        htSS("strUType") = strUType
            '        htSS("strXML") = strXML

            '        '解析xml內容
            '        strMsg = parseXML2SAVE(MyPage, htSS, oConn)
            'End Select

            bliWs = Nothing
        Catch ex As Exception
            'Common.MessageBox(MyPage, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg &= String.Concat("rIDNO:", rIDNO, vbCrLf, "rBIRTH:", rBIRTH, vbCrLf, "rCNAME:", rCNAME, vbCrLf)
            strErrmsg &= "rSTDATE1:" & rSTDATE1 & vbCrLf
            strErrmsg &= "rSTDATE2:" & rSTDATE2 & vbCrLf
            strErrmsg &= "SETID:" & SETID & vbCrLf
            strErrmsg &= "ENTERDATE:" & ENTERDATE & vbCrLf
            strErrmsg &= "SERNUM:" & SERNUM & vbCrLf
            strErrmsg &= "OCID:" & OCID & vbCrLf
            strErrmsg &= "strBIRTH:" & strBIRTH & vbCrLf
            strErrmsg &= "strBDate:" & strBDate & vbCrLf
            strErrmsg &= "strEDate:" & strEDate & vbCrLf

            strErrmsg &= "strXML:" & strXML & vbCrLf
            'strErrmsg &= "strUType:" & strUType & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg)
            'Throw ex
        End Try
    End Sub

    '解析XML 並儲存 STUD_SELRESULTBNG
    Public Shared Function parseXML2SAVE(ByRef MyPage As Page, ByRef htSS As Hashtable, ByVal oConn As SqlConnection) As String
        Dim strMsg As String = "" '回傳錯誤碼！
        Dim strXML As String = TIMS.GetMyValue2(htSS, "strXML")
        Dim strUType As String = TIMS.GetMyValue2(htSS, "strUType")

        Dim rIDNO As String = TIMS.GetMyValue2(htSS, "rIDNO")
        Dim SETID As String = TIMS.GetMyValue2(htSS, "SETID")
        Dim ENTERDATE As String = TIMS.GetMyValue2(htSS, "ENTERDATE")
        Dim SERNUM As String = TIMS.GetMyValue2(htSS, "SERNUM")
        Dim OCID As String = TIMS.GetMyValue2(htSS, "OCID")

        '檢核攻擊xml
        If Not TIMS.Check_xml(strXML) Then
            'Dim htSS As New Hashtable 'htSS Hashtable() 'htSS.Add("strSetId", strSetId)
            TIMS.SetMyValue2(htSS, "sTitle1", "檢核攻擊xml")
            Call SendMailTestErr1(MyPage, htSS)
            Return "-1"
        End If
        Dim xmlDoc As New XmlDocument
        Dim XmlNodes As XmlNodeList

        'Dim strRtnMsg As String = "0" '成功代碼 0:成功 
        'Dim iNodeCnt As Integer = 0
        'Const Cst_OKmsg As String = "0,8" '成功代碼 0:成功 8:資料超過100筆
        Const Cst_OKmsg As String = "0" '成功代碼 0:成功 

        Dim strIDNO As String = String.Empty
        Dim strBirth As String = String.Empty
        Dim strActNo As String = String.Empty
        Dim strMDate As String = String.Empty
        Dim strChgMode As String = String.Empty 'txcd
        'Dim strFType As String = String.Empty

        Try
            xmlDoc.LoadXml(strXML) '解析XML

            If Convert.ToString(xmlDoc) <> "" Then
                strMsg = xmlDoc.DocumentElement.FirstChild.InnerText
                If Cst_OKmsg.IndexOf(Convert.ToString(strMsg)) = -1 Then
                    '異常回傳錯誤碼！
                    Return strMsg
                End If

                TIMS.OpenDbConn(oConn)
                'oTrans = conn.BeginTransaction

                '查詢勾稽資料是否存在，存在就不再寫入
                Dim Sql_s As String = ""
                Sql_s &= " select 1 from STUD_SELRESULTBNG"
                Sql_s &= " where IDNO=@IDNO and SETID=@SETID and ENTERDATE=convert(datetime, @ENTERDATE, 111) and SERNUM=@SERNUM and OCID=@OCID"
                Dim sCmd As New SqlCommand(Sql_s, oConn)

                '新增勾稽資料
                Dim Sql_i As String = ""
                Sql_i &= " INSERT INTO STUD_SELRESULTBNG ( SB4ID,IDNO,NAME,BIRTHDAY,UTYPE,ACTNO,COMNAME" & vbCrLf
                Sql_i &= " ,CHANGEMODE,MDATE,SALARY,DEPARTMENT,MODIFYDATE ,SETID,ENTERDATE,SERNUM,OCID,CREATEDATE)" & vbCrLf
                Sql_i &= " VALUES ( @SB4ID,@IDNO,@NAME,@BIRTHDAY,@UTYPE,@ACTNO,@COMNAME" & vbCrLf
                Sql_i &= " ,@CHANGEMODE,@MDATE,@SALARY,@DEPARTMENT,getdate() ,@SETID,@ENTERDATE,@SERNUM,@OCID,getdate())" & vbCrLf
                Dim iCmd As New SqlCommand(Sql_i, oConn)

                '0:查詢成功
                If Cst_OKmsg.IndexOf(Convert.ToString(strMsg)) > -1 Then
                    XmlNodes = xmlDoc.SelectNodes("/result/record")

                    Dim iNodeCnt As Integer = XmlNodes.Count
                    If iNodeCnt > 0 Then
                        For Each itemNode As XmlNode In XmlNodes
                            strIDNO = Convert.ToString(itemNode.ChildNodes.Item(0).InnerText)
                            strBirth = GetADDate(Convert.ToString(itemNode.ChildNodes.Item(14).InnerText))
                            strActNo = Convert.ToString(itemNode.ChildNodes.Item(4).InnerText)
                            strMDate = GetRocADDate(Convert.ToString(itemNode.ChildNodes.Item(10).InnerText))
                            strChgMode = Convert.ToString(itemNode.ChildNodes.Item(8).InnerText)
                            'strFType = strUType

                            '只記錄加保(txcd=4)&退保(txcd=2)
                            'Dim flagCanSave As Boolean = False
                            'Select Case strChgMode
                            '    Case "2", "4"
                            '        flagCanSave = True
                            '        '1.勾稽資料，要排除證號為09、076、075、175、176的資料
                            '        Select Case strActNo.Substring(0, 2)
                            '            Case "09"
                            '                flagCanSave = False
                            '        End Select
                            '        Select Case strActNo.Substring(0, 3)
                            '            Case "076", "076", "175", "176"
                            '                flagCanSave = False
                            '        End Select
                            'End Select

                            Dim flagCanSave As Boolean = True '(全存，完全不排除)
                            If flagCanSave Then
                                Try
                                    'Dim ds As New DataSet
                                    Dim dtB4 As New DataTable
                                    With sCmd
                                        .Parameters.Clear()
                                        .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = rIDNO
                                        .Parameters.Add("@SETID", SqlDbType.Int).Value = Val(SETID)
                                        .Parameters.Add("@ENTERDATE", SqlDbType.VarChar).Value = TIMS.Cdate3(ENTERDATE)
                                        .Parameters.Add("@SERNUM", SqlDbType.Int).Value = Val(SERNUM)
                                        .Parameters.Add("@OCID", SqlDbType.Int).Value = Val(OCID)
                                        dtB4.Load(.ExecuteReader())
                                        .ExecuteReader.Close()
                                    End With

                                    '勾稽資料
                                    If dtB4.Rows.Count = 0 Then
                                        Dim iSB4ID As Integer = DbAccess.GetNewId(oConn, "STUD_SELRESULTBNG_SB4ID_SEQ,STUD_SELRESULTBNG,SB4ID")
                                        With iCmd
                                            .Parameters.Clear()
                                            .Parameters.Add("@SB4ID", SqlDbType.VarChar).Value = iSB4ID
                                            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = If(strIDNO = "", Convert.DBNull, strIDNO)
                                            .Parameters.Add("@NAME", SqlDbType.NVarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(1).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(1).InnerText)
                                            .Parameters.Add("@BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(strBirth) 'If(strBirth = "", Convert.DBNull, strBirth)
                                            .Parameters.Add("@UTYPE", SqlDbType.VarChar).Value = If(strUType = "", " ", strUType)
                                            .Parameters.Add("@ACTNO", SqlDbType.VarChar).Value = If(strActNo = "", Convert.DBNull, strActNo)
                                            .Parameters.Add("@COMNAME", SqlDbType.NVarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(5).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(5).InnerText)

                                            .Parameters.Add("@CHANGEMODE", SqlDbType.VarChar).Value = If(strChgMode = "", Convert.DBNull, strChgMode)
                                            .Parameters.Add("@MDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(strMDate) 'If(strMDate = "", Convert.DBNull, strMDate)
                                            .Parameters.Add("@SALARY", SqlDbType.VarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(11).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(11).InnerText)
                                            .Parameters.Add("@DEPARTMENT", SqlDbType.VarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(15).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(15).InnerText)

                                            .Parameters.Add("@SETID", SqlDbType.Int).Value = Val(SETID)
                                            .Parameters.Add("@ENTERDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(ENTERDATE)
                                            .Parameters.Add("@SERNUM", SqlDbType.Int).Value = Val(SERNUM)
                                            .Parameters.Add("@OCID", SqlDbType.Int).Value = Val(OCID)
                                            .ExecuteNonQuery()
                                        End With
                                    End If
                                    dtB4 = Nothing

                                    'ds.Tables("stud_bligatedata4").Clear()
                                Catch ex As Exception
                                    Dim sErrMsg1 As String = ""
                                    sErrMsg1 &= TIMS.EncryptAes(String.Concat("&strIDNO=", strIDNO, "&strBirth=", strBirth))
                                    sErrMsg1 &= String.Concat("&strActNo=", strActNo, "&strMDate=", strMDate, "&strChgMode=", strChgMode)
                                    sErrMsg1 &= String.Concat("&SETID=", SETID, "&ENTERDATE=", ENTERDATE, "&SERNUM=", SERNUM, "&OCID=", OCID)
                                    Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
                                End Try
                            End If
                        Next
                    End If

                End If
            End If

            'oTrans.Commit()
        Catch ex As Exception
            'oTrans.Rollback()
            Dim sErrMsg1 As String = ""
            sErrMsg1 &= TIMS.EncryptAes(String.Concat("&strIDNO=", strIDNO, "&strBirth=", strBirth))
            sErrMsg1 &= String.Concat("&strActNo=", strActNo, "&strMDate=", strMDate, "&strChgMode=", strChgMode)
            sErrMsg1 &= String.Concat("&SETID=", SETID, "&ENTERDATE=", ENTERDATE, "&SERNUM=", SERNUM, "&OCID=", OCID)
            Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
            TIMS.CloseDbConn(oConn)
            Throw ex
        Finally
            'If Not insOda Is Nothing Then insOda.Dispose()
            'If Not qryOda Is Nothing Then qryOda.Dispose()
            'If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
        Return strMsg
    End Function

    '檢核是否已勾稽過 STUD_SELRESULTBNG
    Public Shared Function ChkSELRESULTBNGExists(ByVal strSS As String, ByVal oConn As SqlConnection) As Boolean
        If strSS = "" Then Return False
        Dim IDNO As String = TIMS.GetMyValue(strSS, "IDNO")
        Dim SETID As String = TIMS.GetMyValue(strSS, "SETID")
        Dim ENTERDATE As String = TIMS.GetMyValue(strSS, "ENTERDATE")
        Dim SERNUM As String = TIMS.GetMyValue(strSS, "SERNUM")
        Dim OCID As String = TIMS.GetMyValue(strSS, "OCID")
        If IDNO = "" OrElse OCID = "" Then Return False
        If SETID = "" OrElse ENTERDATE = "" OrElse SERNUM = "" Then Return False

        Dim bolFlag As Boolean = False
        Dim strSql As String = ""
        strSql &= " SELECT 'X' FROM STUD_SELRESULTBNG WHERE IDNO=@IDNO and OCID=@OCID"
        strSql &= " and SETID=@SETID and ENTERDATE=convert(datetime, @ENTERDATE, 111) and SERNUM=@SERNUM"
        Dim sCmd As New SqlCommand(strSql, oConn)
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = IDNO
            .Parameters.Add("@SETID", SqlDbType.VarChar).Value = SETID
            .Parameters.Add("@ENTERDATE", SqlDbType.VarChar).Value = TIMS.Cdate3(ENTERDATE) 'yyyy/MM/dd
            .Parameters.Add("@SERNUM", SqlDbType.VarChar).Value = SERNUM
            .Parameters.Add("@OCID", SqlDbType.VarChar).Value = OCID
            dt.Load(.ExecuteReader())
            .ExecuteReader.Close()
        End With
        If TIMS.dtHaveDATA(dt) Then bolFlag = True
        Return bolFlag
    End Function

    '錯誤資訊檢查
    Public Shared Sub SendMailTestErr1(ByRef MyPage As Page, ByRef htSS As Hashtable)
        Dim sTitle1 As String = TIMS.GetMyValue2(htSS, "sTitle1")
        Dim rIDNO As String = TIMS.GetMyValue2(htSS, "rIDNO")
        Dim rBIRTH As String = TIMS.GetMyValue2(htSS, "rBIRTH")
        Dim rCNAME As String = TIMS.GetMyValue2(htSS, "rCNAME")

        Dim strBDate As String = TIMS.GetMyValue2(htSS, "strBDate")
        Dim strEDate As String = TIMS.GetMyValue2(htSS, "strEDate")
        Dim strXML As String = TIMS.GetMyValue2(htSS, "strXML")
        Dim strUType As String = TIMS.GetMyValue2(htSS, "strUType")

        Dim strErrmsg As String = ""
        strErrmsg &= "sTitle1:" & sTitle1 & vbCrLf
        strErrmsg &= "rIDNO:" & rIDNO & vbCrLf
        strErrmsg &= "rBIRTH:" & rBIRTH & vbCrLf
        strErrmsg &= "rCNAME:" & rCNAME & vbCrLf

        strErrmsg &= "strBDate:" & strBDate & vbCrLf
        strErrmsg &= "strEDate:" & strEDate & vbCrLf
        strErrmsg &= "strXML:" & strXML & vbCrLf
        strErrmsg &= "strUType:" & strUType & vbCrLf
        strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
        'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        Call TIMS.WriteTraceLog(strErrmsg)
    End Sub
#End Region

#Region "勞保勾稽06"
    'ACTNObli
    Public Shared Function Get_SSValue06(ByVal ESERNUM As String, ByRef oConn As SqlConnection) As String
        Dim strSS As String = ""
        Dim drT As DataRow = TIMS.Get_ENTERTYPE2(ESERNUM, oConn)
        If drT Is Nothing Then Return strSS
        'ACTNObli
        TIMS.SetMyValue(strSS, "IDNO", Convert.ToString(drT("IDNO")))
        TIMS.SetMyValue(strSS, "BIRTH", TIMS.Cdate3(drT("BIRTHDAY")))
        TIMS.SetMyValue(strSS, "CNAME", Convert.ToString(drT("CNAME")))
        TIMS.SetMyValue(strSS, "STDATE1", TIMS.Cdate3(drT("STDATE1")))
        TIMS.SetMyValue(strSS, "STDATE2", TIMS.Cdate3(drT("STDATE1")))
        TIMS.SetMyValue(strSS, "ESETID", Convert.ToString(drT("ESETID")))
        TIMS.SetMyValue(strSS, "ESERNUM", Convert.ToString(drT("ESERNUM")))
        TIMS.SetMyValue(strSS, "OCID", Convert.ToString(drT("OCID1")))
        Return strSS
    End Function

    '先判斷該員是否曾勾稽過，若無才需呼叫webservice 取得勞保勾稽(加保(4))
    Public Shared Sub Get_BLIGATEDATA06(ByRef MyPage As Page,
                                       ByVal strSS As String,
                                       ByVal oConn As SqlConnection)
        If strSS = "" Then Exit Sub
        Dim rIDNO As String = TIMS.GetMyValue(strSS, "IDNO")
        Dim rBIRTH As String = TIMS.GetMyValue(strSS, "BIRTH") 'yyyy/MM/dd
        Dim rCNAME As String = TIMS.GetMyValue(strSS, "CNAME") '投保人姓名(in)
        Dim rSTDATE1 As String = TIMS.GetMyValue(strSS, "STDATE1") 'yyyy/MM/dd
        Dim rSTDATE2 As String = TIMS.GetMyValue(strSS, "STDATE2") 'yyyy/MM/dd
        Dim ESETID As String = TIMS.GetMyValue(strSS, "ESETID")
        Dim ESERNUM As String = TIMS.GetMyValue(strSS, "ESERNUM")
        Dim OCID As String = TIMS.GetMyValue(strSS, "OCID")
        If rIDNO = "" OrElse rSTDATE1 = "" OrElse rSTDATE2 = "" Then Exit Sub
        If ESETID = "" OrElse ESERNUM = "" OrElse OCID = "" Then Exit Sub
        Dim blIsExists As Boolean = ChkBLIGATEDATA06Exists(strSS, oConn)
        If blIsExists Then Exit Sub

        'TRA001 '訓練 'TRA002 '生活津貼 FOR001'外勞 'TRA011產投
        'Const cst_gSysName As String = "TRA001"
        Dim strXML As String = String.Empty
        'Dim blIsExists As Boolean = False

        '勞保+就保A、勞保L、自願職災保險V、農保F、農民自願職災保險FV
        'Const cst_勞保就保 As String = "A"
        'Const cst_自願職災保險 As String = "V"
        'Const cst_農保 As String = "F"
        'Const cst_農民自願職災保險 As String = "FV"
        Dim s_Bli_TYPE As String = "A,V,F,FV"
        Dim aBli_TYPE As String() = s_Bli_TYPE.Split(",")
        'For Each strUType As String In aBli_TYPE

        Dim strMsg As String = ""
        'Dim strUType As String = ""
        'Dim strIName As String = "" 'String.Empty '投保人姓名
        Dim strBIRTH As String = GetReqDate(rBIRTH) '(yyyy/MM/dd)->yyyyMMdd
        Dim strBDate As String = GetReqDate(rSTDATE1) '(yyyy/MM/dd)->yyyyMMdd
        Dim strEDate As String = GetReqDate(rSTDATE2) '(yyyy/MM/dd)->yyyyMMdd

        Dim htSS As New Hashtable From {
            {"rIDNO", rIDNO},
            {"rBIRTH", rBIRTH}, 'yyyy/MM/dd
            {"rCNAME", rCNAME}, '投保人姓名(in)
            {"rSTDATE1", rSTDATE1}, 'yyyy/MM/dd
            {"rSTDATE2", rSTDATE2}, 'yyyy/MM/dd
            {"ESETID", ESETID},
            {"ESERNUM", ESERNUM},
            {"OCID", OCID},
            {"strBIRTH", strBIRTH}, 'yyyyMMdd
            {"strBDate", strBDate}, 'yyyyMMdd
            {"strEDate", strEDate}, 'yyyyMMdd
            {"strUType", ""},
            {"strXML", ""}
        } '= Nothing

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12

        Try
            'https://wltims.wda.gov.tw/bli_wsv4/get_data2.asmx
            Dim bliWs As New bli_wsv4.get_data3

            For Each strUType As String In aBli_TYPE
                'strUType = s_UType 'cst_勞保就保 '查詢勞保+就保勾稽資料
                '呼叫勞保勾稽webservice
                strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
                htSS("strUType") = strUType
                htSS("strXML") = strXML
                '解析xml內容
                strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
                Select Case strMsg
                    Case "7"
                        If rCNAME = "" Then Exit Select
                        'strIName = strCNAME '(使用姓名重新勾稽)
                        'strUType = cst_勞保就保 '查詢勞保+就保勾稽資料
                        '呼叫勞保勾稽webservice
                        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
                        htSS("strUType") = strUType
                        htSS("strXML") = strXML

                        '解析xml內容
                        strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
                End Select
            Next

            'strUType = cst_勞保就保 '查詢勞保+就保勾稽資料
            ''呼叫勞保勾稽webservice
            'strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
            'htSS("strUType") = strUType
            'htSS("strXML") = strXML

            ''解析xml內容
            'strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'Select Case strMsg
            '    Case "7"
            '        If rCNAME = "" Then Exit Select
            '        'strIName = strCNAME '(使用姓名重新勾稽)
            '        strUType = cst_勞保就保 '查詢勞保+就保勾稽資料
            '        '呼叫勞保勾稽webservice
            '        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
            '        htSS("strUType") = strUType
            '        htSS("strXML") = strXML

            '        '解析xml內容
            '        strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'End Select

            'strUType = cst_自願職災保險 '查詢自願職災保險勾稽資料
            ''呼叫勞保勾稽webservice
            'strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
            'htSS("strUType") = strUType
            'htSS("strXML") = strXML

            ''解析xml內容
            'strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'Select Case strMsg
            '    Case "7"
            '        If rCNAME = "" Then Exit Select
            '        'strIName = strCNAME '(使用姓名重新勾稽)
            '        strUType = cst_自願職災保險 '查詢勞保+就保勾稽資料
            '        '呼叫勞保勾稽webservice
            '        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
            '        htSS("strUType") = strUType
            '        htSS("strXML") = strXML

            '        '解析xml內容
            '        strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'End Select

            'strUType = cst_農保 '查詢農保勾稽資料
            ''呼叫勞保勾稽webservice
            'strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, "", strBIRTH, strUType, strBDate, strEDate)
            'htSS("strUType") = strUType
            'htSS("strXML") = strXML

            ''解析xml內容
            'strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'Select Case strMsg
            '    Case "7"
            '        If rCNAME = "" Then Exit Select
            '        'strIName = strCNAME '(使用姓名重新勾稽)
            '        strUType = cst_農保 '查詢勞保+就保勾稽資料
            '        '呼叫勞保勾稽webservice
            '        strXML = bliWs.get_detail(TIMS.cst_gSysName, rIDNO, rCNAME, strBIRTH, strUType, strBDate, strEDate)
            '        htSS("strUType") = strUType
            '        htSS("strXML") = strXML

            '        '解析xml內容
            '        strMsg = parseXML2SAVE06(MyPage, htSS, oConn)
            'End Select

            bliWs = Nothing
        Catch ex As Exception
            'Common.MessageBox(MyPage, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg &= "rIDNO:" & rIDNO & vbCrLf
            strErrmsg &= "rBIRTH:" & rBIRTH & vbCrLf
            strErrmsg &= "rCNAME:" & rCNAME & vbCrLf
            strErrmsg &= "rSTDATE1:" & rSTDATE1 & vbCrLf
            strErrmsg &= "rSTDATE2:" & rSTDATE2 & vbCrLf
            strErrmsg &= "ESETID:" & ESETID & vbCrLf
            strErrmsg &= "ESERNUM:" & ESERNUM & vbCrLf
            strErrmsg &= "OCID:" & OCID & vbCrLf
            strErrmsg &= "strBIRTH:" & strBIRTH & vbCrLf
            strErrmsg &= "strBDate:" & strBDate & vbCrLf
            strErrmsg &= "strEDate:" & strEDate & vbCrLf

            strErrmsg &= "strXML:" & strXML & vbCrLf
            'strErrmsg &= "strUType:" & strUType & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            'Throw ex
        End Try
    End Sub

    '解析XML 並儲存 STUD_BLIGATEDATA06
    Public Shared Function parseXML2SAVE06(ByRef MyPage As Page, ByRef htSS As Hashtable, ByVal oConn As SqlConnection) As String
        Dim strMsg As String = "" '回傳錯誤碼！
        Dim strXML As String = TIMS.GetMyValue2(htSS, "strXML")
        Dim strUType As String = TIMS.GetMyValue2(htSS, "strUType")

        Dim rIDNO As String = TIMS.GetMyValue2(htSS, "rIDNO")
        Dim ESETID As String = TIMS.GetMyValue2(htSS, "ESETID")
        Dim ESERNUM As String = TIMS.GetMyValue2(htSS, "ESERNUM")
        Dim OCID As String = TIMS.GetMyValue2(htSS, "OCID")

        '檢核攻擊xml
        If Not TIMS.Check_xml(strXML) Then
            'Dim htSS As New Hashtable 'htSS Hashtable() 'htSS.Add("strSetId", strSetId)
            Call SendMailTestErr2(MyPage, htSS)
            Return "-1"
        End If
        Dim xmlDoc As New XmlDocument
        Dim XmlNodes As XmlNodeList

        'Dim strRtnMsg As String = "0" '成功代碼 0:成功 
        'Dim iNodeCnt As Integer = 0
        'Const Cst_OKmsg As String = "0,8" '成功代碼 0:成功 8:資料超過100筆
        Const Cst_OKmsg As String = "0" '成功代碼 0:成功 

        Dim strIDNO As String = String.Empty
        Dim strBirth As String = String.Empty
        Dim strActNo As String = String.Empty
        Dim strMDate As String = String.Empty
        Dim strChgMode As String = String.Empty 'txcd
        'Dim strFType As String = String.Empty

        Try
            xmlDoc.LoadXml(strXML) '解析XML

            If Convert.ToString(xmlDoc) <> "" Then
                strMsg = xmlDoc.DocumentElement.FirstChild.InnerText
                If Cst_OKmsg.IndexOf(Convert.ToString(strMsg)) = -1 Then
                    '異常回傳錯誤碼！
                    Return strMsg
                End If

                TIMS.OpenDbConn(oConn)
                'oTrans = conn.BeginTransaction

                '查詢勾稽資料是否存在，存在就不再寫入
                Dim strSql As String = ""
                strSql &= " SELECT 1 FROM STUD_BLIGATEDATA06"
                strSql &= " WHERE IDNO=@IDNO AND ESETID=@ESETID AND ESERNUM=@ESERNUM AND OCID1=@OCID1"
                Dim sCmd As New SqlCommand(strSql, oConn)

                '新增勾稽資料
                Dim strSqli As String = ""
                strSqli &= " INSERT INTO STUD_BLIGATEDATA06 ( SBEID,IDNO,NAME,BIRTHDAY,UTYPE,ACTNO,COMNAME" & vbCrLf
                strSqli &= " ,CHANGEMODE,MDATE,SALARY,DEPARTMENT,BIEF,MODIFYDATE ,ESETID,ESERNUM,OCID1 )" & vbCrLf
                strSqli &= " VALUES ( @SBEID,@IDNO,@NAME,@BIRTHDAY,@UTYPE,@ACTNO,@COMNAME" & vbCrLf
                strSqli &= " ,@CHANGEMODE,@MDATE,@SALARY,@DEPARTMENT,@BIEF,getdate() ,@ESETID,@ESERNUM,@OCID1 ) "
                Dim iCmd As New SqlCommand(strSqli, oConn)

                '0:查詢成功
                If Cst_OKmsg.IndexOf(Convert.ToString(strMsg)) > -1 Then
                    XmlNodes = xmlDoc.SelectNodes("/result/record")

                    Dim iNodeCnt As Integer = XmlNodes.Count
                    If iNodeCnt > 0 Then
                        For Each itemNode As XmlNode In XmlNodes
                            strIDNO = Convert.ToString(itemNode.ChildNodes.Item(0).InnerText)
                            strBirth = GetADDate(Convert.ToString(itemNode.ChildNodes.Item(14).InnerText))
                            strActNo = Convert.ToString(itemNode.ChildNodes.Item(4).InnerText)
                            strMDate = GetRocADDate(Convert.ToString(itemNode.ChildNodes.Item(10).InnerText))
                            strChgMode = Convert.ToString(itemNode.ChildNodes.Item(8).InnerText)
                            'strFType = strUType

                            '只記錄加保(txcd=4)&退保(txcd=2)
                            'Dim flagCanSave As Boolean = False
                            'Select Case strChgMode
                            '    Case "2", "4"
                            '        flagCanSave = True
                            '        '1.勾稽資料，要排除證號為09、076、075、175、176的資料
                            '        Select Case strActNo.Substring(0, 2)
                            '            Case "09"
                            '                flagCanSave = False
                            '        End Select
                            '        Select Case strActNo.Substring(0, 3)
                            '            Case "076", "076", "175", "176"
                            '                flagCanSave = False
                            '        End Select
                            'End Select

                            Dim flagCanSave As Boolean = True '(全存，完全不排除)
                            If flagCanSave Then
                                Try
                                    'Dim ds As New DataSet
                                    Dim dtB4 As New DataTable
                                    With sCmd
                                        .Parameters.Clear()
                                        .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = rIDNO
                                        .Parameters.Add("@ESETID", SqlDbType.Int).Value = Val(ESETID)
                                        .Parameters.Add("@ESERNUM", SqlDbType.Int).Value = Val(ESERNUM)
                                        .Parameters.Add("@OCID1", SqlDbType.Int).Value = Val(OCID)
                                        dtB4.Load(.ExecuteReader())
                                        .ExecuteReader.Close()
                                    End With

                                    '勾稽資料
                                    If dtB4.Rows.Count = 0 Then
                                        Dim iSBEID As Integer = DbAccess.GetNewId(oConn, "STUD_BLIGATEDATA06_SBEID_SEQ,STUD_BLIGATEDATA06,SBEID")
                                        With iCmd
                                            .Parameters.Clear()
                                            .Parameters.Add("@SBEID", SqlDbType.VarChar).Value = iSBEID
                                            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = If(strIDNO = "", Convert.DBNull, strIDNO)
                                            .Parameters.Add("@NAME", SqlDbType.NVarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(1).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(1).InnerText)
                                            .Parameters.Add("@BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(strBirth) 'If(strBirth = "", Convert.DBNull, strBirth)
                                            .Parameters.Add("@UTYPE", SqlDbType.VarChar).Value = If(strUType = "", " ", strUType)
                                            .Parameters.Add("@ACTNO", SqlDbType.VarChar).Value = If(strActNo = "", Convert.DBNull, strActNo)
                                            .Parameters.Add("@COMNAME", SqlDbType.NVarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(5).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(5).InnerText)

                                            .Parameters.Add("@CHANGEMODE", SqlDbType.VarChar).Value = If(strChgMode = "", Convert.DBNull, strChgMode)
                                            .Parameters.Add("@MDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(strMDate) 'If(strMDate = "", Convert.DBNull, strMDate)
                                            .Parameters.Add("@SALARY", SqlDbType.VarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(11).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(11).InnerText)
                                            .Parameters.Add("@DEPARTMENT", SqlDbType.VarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(15).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(15).InnerText)
                                            .Parameters.Add("@BIEF", SqlDbType.VarChar).Value = If(Convert.ToString(itemNode.ChildNodes.Item(15).InnerText) = "", Convert.DBNull, itemNode.ChildNodes.Item(12).InnerText)

                                            .Parameters.Add("@ESETID", SqlDbType.Int).Value = Val(ESETID)
                                            .Parameters.Add("@ESERNUM", SqlDbType.Int).Value = Val(ESERNUM)
                                            .Parameters.Add("@OCID1", SqlDbType.Int).Value = Val(OCID)
                                            .ExecuteNonQuery()
                                        End With
                                    End If
                                    dtB4 = Nothing

                                    'ds.Tables("stud_bligatedata4").Clear()
                                Catch ex As Exception
                                    Dim sErrMsg1 As String = ""
                                    sErrMsg1 &= TIMS.EncryptAes(String.Concat("&strIDNO=", strIDNO, "&strBirth=", strBirth))
                                    sErrMsg1 &= String.Concat("&strActNo=", strActNo, "&strMDate=", strMDate, "&strChgMode=", strChgMode)
                                    sErrMsg1 &= String.Concat("&ESETID=", ESETID, "&ESERNUM=", ESERNUM, "&OCID=", OCID)
                                    Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
                                End Try
                            End If
                        Next
                    End If

                End If
            End If

            'oTrans.Commit()
        Catch ex As Exception
            'oTrans.Rollback()
            Dim sErrMsg1 As String = ""
            sErrMsg1 &= TIMS.EncryptAes(String.Concat("&strIDNO=", strIDNO, "&strBirth=", strBirth))
            sErrMsg1 &= String.Concat("&strActNo=", strActNo, "&strMDate=", strMDate, "&strChgMode=", strChgMode)
            sErrMsg1 &= String.Concat("&ESETID=", ESETID, "&ESERNUM=", ESERNUM, "&OCID=", OCID)
            Call TIMS.WriteTraceLog(MyPage, ex, sErrMsg1)
            TIMS.CloseDbConn(oConn)
            Throw ex
        Finally
            'If Not insOda Is Nothing Then insOda.Dispose()
            'If Not qryOda Is Nothing Then qryOda.Dispose()
            'If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
        Return strMsg
    End Function

    '檢核是否已勾稽過 STUD_BLIGATEDATA06
    Public Shared Function ChkBLIGATEDATA06Exists(ByVal strSS As String, ByVal oConn As SqlConnection) As Boolean
        If strSS = "" Then Return False
        Dim IDNO As String = TIMS.GetMyValue(strSS, "IDNO")
        Dim ESETID As String = TIMS.GetMyValue(strSS, "ESETID")
        Dim ESERNUM As String = TIMS.GetMyValue(strSS, "ESERNUM")
        Dim OCID As String = TIMS.GetMyValue(strSS, "OCID")
        If IDNO = "" OrElse ESETID = "" OrElse ESERNUM = "" OrElse OCID = "" Then Return False

        Dim bolFlag As Boolean = False
        Dim strSql As String = ""
        strSql &= " SELECT 'X' FROM STUD_BLIGATEDATA06"
        strSql &= " WHERE IDNO=@IDNO AND ESETID=@ESETID AND ESERNUM=@ESERNUM AND OCID1=@OCID1"
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        Using sCmd As New SqlCommand(strSql, oConn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = IDNO
                .Parameters.Add("@ESETID", SqlDbType.VarChar).Value = ESETID
                .Parameters.Add("@ESERNUM", SqlDbType.VarChar).Value = ESERNUM
                .Parameters.Add("@OCID1", SqlDbType.VarChar).Value = OCID
                dt.Load(.ExecuteReader())
                .ExecuteReader.Close()
            End With
        End Using
        If TIMS.dtHaveDATA(dt) Then bolFlag = True
        Return bolFlag
    End Function

    '錯誤資訊檢查06
    Public Shared Sub SendMailTestErr2(ByRef MyPage As Page, ByRef htSS As Hashtable)
        Dim rIDNO As String = TIMS.GetMyValue2(htSS, "rIDNO")
        Dim rBIRTH As String = TIMS.GetMyValue2(htSS, "rBIRTH")
        Dim rCNAME As String = TIMS.GetMyValue2(htSS, "rCNAME")

        Dim strBDate As String = TIMS.GetMyValue2(htSS, "strBDate")
        Dim strEDate As String = TIMS.GetMyValue2(htSS, "strEDate")
        Dim strXML As String = TIMS.GetMyValue2(htSS, "strXML")
        Dim strUType As String = TIMS.GetMyValue2(htSS, "strUType")

        Dim strErrmsg As String = ""
        strErrmsg &= TIMS.EncryptAes(String.Concat(",rIDNO:", rIDNO, vbCrLf, ",rBIRTH:", rBIRTH, vbCrLf, ",rCNAME:", rCNAME, vbCrLf))
        strErrmsg &= ",strBDate:" & strBDate & vbCrLf
        strErrmsg &= ",strEDate:" & strEDate & vbCrLf
        strErrmsg &= ",strXML:" & strXML & vbCrLf
        strErrmsg &= ",strUType:" & strUType & vbCrLf
        strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
        'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        Call TIMS.WriteTraceLog(strErrmsg)
    End Sub

#End Region

End Class
