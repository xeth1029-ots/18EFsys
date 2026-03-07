
Partial Class SD_15_034
    Inherits AuthBasePage

    'SD_15_034
    Const cst_lab_NODATA_TXT As String = "(查無資料)"
    Const CST_SCOPE1_報名 As String = "01"
    Const CST_SCOPE1_參訓 As String = "02"
    Const CST_ETENTER_USE_系統預設區間 As String = "1"
    Const CST_ETENTER_USE_自選日期區間 As String = "2"
    Const CST_ETENTER_TXT_系統預設區間 As String = "(系統預設區間)"
    Const CST_ETENTER_TXT_自選日期區間 As String = "(自選日期區間)"
    '1.系統預設區間：
    '報名人數：【開訓日】：(災害起始日+1個月)~(今日+30天)，【報名日】>=災害起始日
    '參訓人數：【開訓日】： 災害起始日~今日

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            Call CCREATE1()
        End If

        LabTPLANNAME.Visible = False
        CBL_TPLANID_S1.Visible = False
        Select Case sm.UserInfo.LID
            Case 0
                CBL_TPLANID_S1.Visible = True
            Case Else
                LabTPLANNAME.Visible = True
        End Select
    End Sub

    Sub CCREATE1()
        UTL_DIVSHOW(0)
        'div_search1.Visible = True 'div_detail1.Visible = False
        TIMS.OpenDbConn(objconn)

        '重大災害名稱
        DDL_DISASTER_S1 = TIMS.GET_DISASTER(objconn, DDL_DISASTER_S1, 2)

        '轄區代碼
        DDL_DISTID_S1 = TIMS.Get_DistID(DDL_DISTID_S1, TIMS.Get_DISTIDT2(objconn))
        Common.SetListItem(DDL_DISTID_S1, sm.UserInfo.DistID)
        If sm.UserInfo.LID <> 0 Then DDL_DISTID_S1.Enabled = False

        '取出鍵詞-訓練計畫代碼
        Select Case sm.UserInfo.LID
            Case 0
                CBL_TPLANID_S1 = TIMS.Get_TPlan(CBL_TPLANID_S1,, 1, "Y",, objconn)
                TIMS.SetCblValue(CBL_TPLANID_S1, sm.UserInfo.TPlanID)
                '選擇全部訓練計畫
                CBL_TPLANID_S1.Attributes("onclick") = "SelectAll('CBL_TPLANID_S1','CBL_TPLANID_S1Hidden');"
            Case Else
                LabTPLANNAME.Text = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        End Select

        'ETENTERDATE1.Text = TIMS.Cdate3(DateAdd(DateInterval.Day, -30, Now))
        'ETENTERDATE2.Text = TIMS.Cdate3(DateAdd(DateInterval.Day, 0, Now))
        TIMS.Tooltip(BTN_SEARCH1, "基底限制2年內的資料", True)
    End Sub

    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click
        Dim ERRMSG1 As String = CHECKSEARCH1()
        If ERRMSG1 <> "" Then
            Common.MessageBox(Me, ERRMSG1)
            Return
        End If

        Call SCH_DISASTER_1(0)

        '資料篩選方式
        Dim V_RBL_ETENTER_USE As String = TIMS.GetListValue(RBL_ETENTER_USE)
        '報名人數
        Dim DT2 As DataTable = SCH_DATA_2(V_RBL_ETENTER_USE)
        If TIMS.dtHaveDATA(DT2) Then
            Lab_NumberAPP.Text = $"{DT2.Rows.Count} 人次"
        Else
            Lab_NumberAPP.Text = cst_lab_NODATA_TXT
        End If
        TIMS.Tooltip(Lab_NumberAPP, "基底限制2年內的資料", True)

        '參訓人數
        Dim DT3 As DataTable = SCH_DATA_3(V_RBL_ETENTER_USE)
        If TIMS.dtHaveDATA(DT3) Then
            Lab_NumberPart.Text = $"{DT3.Rows.Count} 人次"
        Else
            Lab_NumberPart.Text = cst_lab_NODATA_TXT
        End If
        TIMS.Tooltip(Lab_NumberPart, "基底限制2年內的資料", True)

    End Sub

    ''' <summary>
    ''' 報名人數-資料篩選方式
    ''' </summary>
    ''' <param name="V_ETENTER_USE"></param>
    ''' <returns></returns>
    Function SCH_DATA_2(V_ETENTER_USE As String) As DataTable
        '重大災害名稱 'Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
        Dim V_DDL_DISTID_S1 As String = TIMS.GetListValue(DDL_DISTID_S1)
        If V_DDL_DISTID_S1 = "" Then V_DDL_DISTID_S1 = sm.UserInfo.DistID
        '訓練計畫
        Dim V_CBL_TPLANID_S1 As String = sm.UserInfo.TPlanID
        If sm.UserInfo.LID = 0 Then V_CBL_TPLANID_S1 = TIMS.GetCblValue(CBL_TPLANID_S1)
        ETENTERDATE1.Text = TIMS.Cdate3(ETENTERDATE1.Text)
        ETENTERDATE2.Text = TIMS.Cdate3(ETENTERDATE2.Text)
        STDATE1.Text = TIMS.Cdate3(STDATE1.Text)
        STDATE2.Text = TIMS.Cdate3(STDATE2.Text)
        Dim vETENTERDATE1 As String = ETENTERDATE1.Text
        Dim vETENTERDATE2 As String = ETENTERDATE2.Text
        Dim vSTDATE1 As String = STDATE1.Text
        Dim vSTDATE2 As String = STDATE2.Text

        hid_ADID_ZIPCODES.Value = TIMS.ClearSQM(hid_ADID_ZIPCODES.Value)

        Dim PMSX1 As New Hashtable From {{"DISTID", V_DDL_DISTID_S1}}
        Select Case V_ETENTER_USE
            Case CST_ETENTER_USE_系統預設區間
                '報名人數日期區間
                Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
                Dim dt1 As DataTable = SCH_DATA_DISASTER(V_DDL_DISASTER_S1)
                If TIMS.dtHaveDATA(dt1) Then
                    Dim dr1 As DataRow = dt1.Rows(0)
                    vETENTERDATE1 = TIMS.Cdate3(dr1("BEGDATE"))
                    vETENTERDATE2 = TIMS.Cdate3(dr1("RETURNDATE1"))
                    vSTDATE1 = TIMS.Cdate3(DateAdd(DateInterval.Month, 1, dr1("BEGDATE")))
                    vSTDATE2 = TIMS.Cdate3(DateAdd(DateInterval.Day, 30, dr1("RETURNDATE1")))
                    PMSX1.Add("ETENTERDATE1", TIMS.Cdate2(vETENTERDATE1))
                    PMSX1.Add("ETENTERDATE2", TIMS.Cdate2(vETENTERDATE2))
                    PMSX1.Add("STDATE1", TIMS.Cdate2(vSTDATE1))
                    PMSX1.Add("STDATE2", TIMS.Cdate2(vSTDATE2))
                    Dim ETENTERDATE1_ROC As String = TIMS.Cdate17(vETENTERDATE1)
                    Dim ETENTERDATE2_ROC As String = TIMS.Cdate17(vETENTERDATE2)
                    Dim STDATE1_ROC As String = TIMS.Cdate17(vSTDATE1)
                    Dim STDATE2_ROC As String = TIMS.Cdate17(vSTDATE2)
                    Lab_APPMSG.Text = $"(報名人數條件說明)：報名日期區間：{ETENTERDATE1_ROC}~{ETENTERDATE2_ROC}，開訓日期區間：{STDATE1_ROC}~{STDATE2_ROC}"
                End If
            Case CST_ETENTER_USE_自選日期區間
                If vETENTERDATE1 <> "" AndAlso vETENTERDATE2 <> "" Then
                    PMSX1.Add("ETENTERDATE1", TIMS.Cdate2(vETENTERDATE1))
                    PMSX1.Add("ETENTERDATE2", TIMS.Cdate2(vETENTERDATE2))
                End If
                If vSTDATE1 <> "" AndAlso vSTDATE2 <> "" Then
                    PMSX1.Add("STDATE1", TIMS.Cdate2(vSTDATE1))
                    PMSX1.Add("STDATE2", TIMS.Cdate2(vSTDATE2))
                End If
        End Select

        Dim V_TPLANID_IN As String = TIMS.CombiSQLIN(V_CBL_TPLANID_S1)
        Dim V_ZIPCODE_IN As String = TIMS.CombiSQLIN(hid_ADID_ZIPCODES.Value)

        Dim SSQLX1 As String = ""
        SSQLX1 &= " WITH WIZ1 AS (SELECT ZIPCODE,ZIPNAME,CTNAME,ZNAME FROM VIEW_ZIPNAME)" & vbCrLf
        If V_ZIPCODE_IN <> "" Then
            SSQLX1 &= $" ,WIZ2 AS ( SELECT ZIPCODE,ZIPNAME FROM WIZ1 WHERE ZIPCODE IN ({V_ZIPCODE_IN}) )" & vbCrLf
        Else
            SSQLX1 &= $" ,WIZ2 AS ( SELECT ZIPCODE,ZIPNAME FROM WIZ1 WHERE 1!=1 )" & vbCrLf
        End If
        'WC1
        SSQLX1 &= " ,WC1 AS ( SELECT cc.OCID,cc.CTNAME,cc.ORGNAME,cc.CLASSCNAME,cc.STDATE,cc.FTDATE FROM VIEW2 cc" & vbCrLf
        SSQLX1 &= " WHERE cc.YEARS>=YEAR(GETDATE())-2 AND cc.STDATE>=DATEADD(YEAR,-2,GETDATE())" & vbCrLf
        SSQLX1 &= " AND cc.DISTID=@DISTID" & vbCrLf
        SSQLX1 &= If(V_TPLANID_IN <> "", $" AND cc.TPLANID IN ({V_TPLANID_IN})", " AND 1!=1") & vbCrLf
        SSQLX1 &= If(vSTDATE1 <> "" AndAlso vSTDATE2 <> "", " AND cc.STDATE>=@STDATE1 AND cc.STDATE<=@STDATE2", "") & " )" & vbCrLf
        'WT2
        SSQLX1 &= " ,WT2 AS ( SELECT a.ENTERDATE,a.ESERNUM,a.NAME STDNAME,a.IDNO,a.OCID1,a.ORGNAME,a.CLASSCNAME,a.STDATE,a.FTDATE,a.ZIPCODE,a.ZIPCODE2,ISNULL(a.MIDENTITYID,substring(a.IDENTITYID,1,2)) IDENTITYID" & vbCrLf
        SSQLX1 &= " FROM dbo.V_ENTERTYPE2 a JOIN WC1 cc on cc.OCID=a.OCID1" & vbCrLf
        SSQLX1 &= " WHERE a.ENTERDATE>=DATEADD(YEAR,-2,GETDATE())" & vbCrLf
        SSQLX1 &= " AND (a.ZIPCODE IN (SELECT ZIPCODE FROM WIZ2) OR a.ZIPCODE2 IN (SELECT ZIPCODE FROM WIZ2))" & vbCrLf
        If vETENTERDATE1 <> "" AndAlso vETENTERDATE2 <> "" Then
            SSQLX1 &= " AND a.ENTERDATE>=@ETENTERDATE1 AND convert(date,a.ENTERDATE)<=convert(date,@ETENTERDATE2)" & vbCrLf
        End If
        SSQLX1 &= " )" & vbCrLf
        'WT1+WT2
        SSQLX1 &= " ,WT1 AS ( SELECT a.ENTERDATE,a.SENID ESERNUM,a.STDNAME,a.IDNO,a.OCID1,a.ORGNAME,a.CLASSCNAME,cc.STDATE,cc.FTDATE,a.ZIPCODE,a.ZIPCODE2,ISNULL(a.MIDENTITYID,substring(a.IDENTITYID,1,2)) IDENTITYID" & vbCrLf
        SSQLX1 &= " FROM dbo.V_ENTERTYPET1 a JOIN WC1 cc on cc.OCID=a.OCID1 AND A.SENID IS NOT NULL" & vbCrLf
        SSQLX1 &= " WHERE a.ENTERDATE>=DATEADD(YEAR,-2,GETDATE())" & vbCrLf
        SSQLX1 &= " AND (a.ZIPCODE IN (SELECT ZIPCODE FROM WIZ2) OR a.ZIPCODE2 IN (SELECT ZIPCODE FROM WIZ2))" & vbCrLf
        If vETENTERDATE1 <> "" AndAlso vETENTERDATE2 <> "" Then
            SSQLX1 &= " AND a.ENTERDATE>=@ETENTERDATE1 AND convert(date,a.ENTERDATE)<=convert(date,@ETENTERDATE2)" & vbCrLf
        End If
        SSQLX1 &= " AND NOT EXISTS (SELECT 1 FROM WT2 x WHERE x.OCID1=a.OCID1 AND x.IDNO=a.IDNO)" & vbCrLf
        SSQLX1 &= " UNION" & vbCrLf
        SSQLX1 &= " SELECT a.ENTERDATE,a.ESERNUM,a.STDNAME,a.IDNO,a.OCID1,a.ORGNAME,a.CLASSCNAME,a.STDATE,a.FTDATE,a.ZIPCODE,a.ZIPCODE2,a.IDENTITYID" & vbCrLf
        SSQLX1 &= " FROM WT2 a" & vbCrLf
        SSQLX1 &= " )" & vbCrLf
        'CROSS 
        SSQLX1 &= " SELECT '報名人數' 類別,format(a.ENTERDATE,'yyyy/MM/dd') 報名日期" & vbCrLf
        SSQLX1 &= " ,a.ESERNUM 序號,a.STDNAME 姓名,a.OCID1 班級代碼, format(GETDATE(),'yyyy/MM/dd') 今天" & vbCrLf
        SSQLX1 &= " ,(select CTNAME FROM WC1 WHERE OCID=a.OCID1) 班級位置" & vbCrLf
        SSQLX1 &= " ,a.ORGNAME 機構,a.CLASSCNAME 班名,format(a.STDATE,'yyyy/MM/dd') 開訓日,format(a.FTDATE,'yyyy/MM/dd') 結訓日" & vbCrLf
        SSQLX1 &= " ,iz.ZIPNAME 住址縣市,iz2.ZIPNAME 戶籍縣市" & vbCrLf
        SSQLX1 &= " ,kd.NAME 參訓身分別" & vbCrLf
        SSQLX1 &= " FROM WT1 a" & vbCrLf
        SSQLX1 &= " LEFT JOIN KEY_IDENTITY kd on kd.IDENTITYID=a.IDENTITYID" & vbCrLf
        SSQLX1 &= " LEFT JOIN WIZ1 iz on iz.ZIPCODE=a.ZIPCODE" & vbCrLf
        SSQLX1 &= " LEFT JOIN WIZ1 iz2 on iz2.ZIPCODE=a.ZIPCODE2" & vbCrLf
        'SSQLX1 &= " /*0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取*/" & vbCrLf 'A.SIGNUPSTATUS NOT IN (2,5) 
        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--PMS: ", TIMS.GetMyValue5(PMSX1), vbCrLf, "--#SD_15_034:", vbCrLf, SSQLX1))
        End If
        Return DbAccess.GetDataTable(SSQLX1, objconn, PMSX1)
    End Function

    ''' <summary>
    ''' 參訓人數-資料篩選方式
    ''' </summary>
    ''' <param name="V_ETENTER_USE"></param>
    ''' <returns></returns>
    Function SCH_DATA_3(V_ETENTER_USE As String) As DataTable
        '重大災害名稱 'Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
        Dim V_DDL_DISTID_S1 As String = TIMS.GetListValue(DDL_DISTID_S1)
        If V_DDL_DISTID_S1 = "" Then V_DDL_DISTID_S1 = sm.UserInfo.DistID
        '訓練計畫
        Dim V_CBL_TPLANID_S1 As String = sm.UserInfo.TPlanID
        If sm.UserInfo.LID = 0 Then V_CBL_TPLANID_S1 = TIMS.GetCblValue(CBL_TPLANID_S1)

        ETENTERDATE1.Text = TIMS.Cdate3(ETENTERDATE1.Text)
        ETENTERDATE2.Text = TIMS.Cdate3(ETENTERDATE2.Text)
        STDATE1.Text = TIMS.Cdate3(STDATE1.Text)
        STDATE2.Text = TIMS.Cdate3(STDATE2.Text)

        Dim vETENTERDATE1 As String = ETENTERDATE1.Text
        Dim vETENTERDATE2 As String = ETENTERDATE2.Text
        Dim vSTDATE1 As String = STDATE1.Text
        Dim vSTDATE2 As String = STDATE2.Text

        hid_ADID_ZIPCODES.Value = TIMS.ClearSQM(hid_ADID_ZIPCODES.Value)
        Dim PMSX1 As New Hashtable From {{"DISTID", V_DDL_DISTID_S1}}
        Select Case V_ETENTER_USE
            Case CST_ETENTER_USE_系統預設區間
                '參訓人數日期區間
                Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
                Dim dt1 As DataTable = SCH_DATA_DISASTER(V_DDL_DISASTER_S1)
                If TIMS.dtHaveDATA(dt1) Then
                    Dim dr1 As DataRow = dt1.Rows(0)
                    vETENTERDATE1 = ""
                    vETENTERDATE2 = ""
                    vSTDATE1 = TIMS.Cdate3(dr1("BEGDATE"))
                    vSTDATE2 = TIMS.Cdate3(dr1("RETURNDATE1"))
                    PMSX1.Add("ETENTERDATE1", TIMS.Cdate2(vETENTERDATE1))
                    PMSX1.Add("ETENTERDATE2", TIMS.Cdate2(vETENTERDATE2))
                    PMSX1.Add("STDATE1", TIMS.Cdate2(vSTDATE1))
                    PMSX1.Add("STDATE2", TIMS.Cdate2(vSTDATE2))
                    Dim STDATE1_ROC As String = TIMS.Cdate17(vSTDATE1)
                    Dim STDATE2_ROC As String = TIMS.Cdate17(vSTDATE2)
                    Lab_PartMSG.Text = $"(參訓人數條件說明)：開訓日期區間：{STDATE1_ROC}~{STDATE2_ROC}"
                End If
            Case CST_ETENTER_USE_自選日期區間
                If vETENTERDATE1 <> "" AndAlso vETENTERDATE2 <> "" Then
                    PMSX1.Add("ETENTERDATE1", TIMS.Cdate2(vETENTERDATE1))
                    PMSX1.Add("ETENTERDATE2", TIMS.Cdate2(vETENTERDATE2))
                End If
                If vSTDATE1 <> "" AndAlso vSTDATE2 <> "" Then
                    PMSX1.Add("STDATE1", TIMS.Cdate2(vSTDATE1))
                    PMSX1.Add("STDATE2", TIMS.Cdate2(vSTDATE2))
                End If
        End Select

        Dim V_TPLANID_IN As String = TIMS.CombiSQLIN(V_CBL_TPLANID_S1)
        Dim V_ZIPCODE_IN As String = TIMS.CombiSQLIN(hid_ADID_ZIPCODES.Value)

        Dim SSQLX1 As String = ""
        SSQLX1 &= " WITH WIZ1 AS (SELECT ZIPCODE,ZIPNAME,CTNAME,ZNAME FROM VIEW_ZIPNAME)" & vbCrLf
        If V_ZIPCODE_IN <> "" Then
            SSQLX1 &= $" ,WIZ2 AS ( SELECT ZIPCODE,ZIPNAME FROM WIZ1 WHERE ZIPCODE IN ({V_ZIPCODE_IN}) )" & vbCrLf
        Else
            SSQLX1 &= $" ,WIZ2 AS ( SELECT ZIPCODE,ZIPNAME FROM WIZ1 WHERE 1!=1 )" & vbCrLf
        End If
        SSQLX1 &= " SELECT '參訓人數' 類別,format(a.ETENTERDATE,'yyyy/MM/dd') 報名日期" & vbCrLf
        SSQLX1 &= " ,a.SOCID 序號,a.NAME 姓名,a.OCID 班級代碼, format(GETDATE(),'yyyy/MM/dd') 今天" & vbCrLf
        SSQLX1 &= " ,(select CTNAME FROM VIEW2 WHERE OCID=a.OCID) 班級位置" & vbCrLf
        SSQLX1 &= " ,a.ORGNAME 機構,a.CLASSCNAME 班名,format(a.STDATE,'yyyy/MM/dd') 開訓日,format(a.FTDATE,'yyyy/MM/dd') 結訓日" & vbCrLf
        'SSQLX1 &= " /*,a.ZIPCODE,a.ZIPCODE2*/" & vbCrLf
        SSQLX1 &= " ,iz.ZIPNAME 住址縣市,iz2.ZIPNAME 戶籍縣市" & vbCrLf
        SSQLX1 &= " ,A.MINAME 參訓身分別" & vbCrLf
        SSQLX1 &= " FROM dbo.V_STUDENTINFO a" & vbCrLf
        SSQLX1 &= " LEFT JOIN WIZ1 iz on iz.ZIPCODE=a.ZIPCODE1" & vbCrLf
        SSQLX1 &= " LEFT JOIN WIZ1 iz2 on iz2.ZIPCODE=a.ZIPCODE2" & vbCrLf
        SSQLX1 &= " WHERE a.YEARS>=YEAR(GETDATE())-2 AND a.STDATE>=DATEADD(YEAR,-2,GETDATE()) AND a.ETENTERDATE>=DATEADD(YEAR,-2,GETDATE())" & vbCrLf
        SSQLX1 &= $" AND a.DISTID=@DISTID" & vbCrLf
        If V_TPLANID_IN <> "" Then
            SSQLX1 &= $" AND a.TPLANID IN ({V_TPLANID_IN})" & vbCrLf
        Else
            SSQLX1 &= $" AND 1!=1" & vbCrLf
        End If
        If vETENTERDATE1 <> "" AndAlso vETENTERDATE2 <> "" Then
            SSQLX1 &= " AND a.ETENTERDATE>=@ETENTERDATE1 AND convert(date,a.ETENTERDATE)<=convert(date,@ETENTERDATE2) " & vbCrLf
        End If
        If vSTDATE1 <> "" AndAlso vSTDATE2 <> "" Then
            SSQLX1 &= " AND a.STDATE>=@STDATE1 AND a.STDATE<=@STDATE2" & vbCrLf
        End If
        SSQLX1 &= " AND a.STUDSTATUS NOT IN (2,3)" & vbCrLf '/*在訓*/
        SSQLX1 &= " AND (a.ZIPCODE1 IN (SELECT ZIPCODE FROM WIZ2) OR a.ZIPCODE2 IN (SELECT ZIPCODE FROM WIZ2))" & vbCrLf
        'Dim DT1 As DataTable = DbAccess.GetDataTable(SSQLX1, objconn, PMSX1)
        Return DbAccess.GetDataTable(SSQLX1, objconn, PMSX1)
    End Function

    Function SCH_DATA_DISASTER(V_ADID As String) As DataTable
        If V_ADID = "" Then Return Nothing
        Dim pms1 As New Hashtable From {{"ADID", V_ADID}}
        Dim sql As String = ""
        sql &= " SELECT a.ADID ,a.CNAME" & vbCrLf
        sql &= " ,format(a.BEGDATE,'yyyy/MM/dd') BEGDATE" & vbCrLf
        sql &= " ,format(a.ENDDATE,'yyyy/MM/dd') ENDDATE" & vbCrLf
        sql &= " ,a.ALARMMSG1 ,a.MEMO1, a.FUNC1, a.FUNC2" & vbCrLf
        sql &= " ,CONCAT(a.CNAME,'(自',dbo.FN_CDATE1B(a.BEGDATE),'~',dbo.FN_CDATE1B(a.ENDDATE),'止)',CASE WHEN DATEDIFF(DAY,a.ENDDATE,GETDATE())>0 THEN'(已結束)'END ) CNAME_N" & vbCrLf
        sql &= " ,a.CREATEACCT ,a.CREATEDATE" & vbCrLf
        sql &= " ,format(GETDATE(),'yyyy/MM/dd') RETURNDATE1" & vbCrLf
        sql &= " ,dbo.FN_CDATE1B(GETDATE()) RETURNDATE1_ROC" & vbCrLf
        sql &= " ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.ADP_DISASTER a" & vbCrLf
        sql &= " WHERE a.ADID=@ADID" & vbCrLf
        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
        Return DbAccess.GetDataTable(sql, objconn, pms1)
    End Function

    ''' <summary>查詢資料，並顯示相關欄位</summary>
    ''' <param name="iEXP">iEXP = 1(匯出資料使用)</param>
    Sub SCH_DISASTER_1(iEXP As Integer)
        hid_ADID_ZIPCODES.Value = ""
        '重大災害名稱
        Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
        '訓練計畫 ' Dim V_CBL_TPLANID_S1 As String = TIMS.GetCblValue(CBL_TPLANID_S1)
        Dim TXT_CBL_TPLANID_S1 As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        If sm.UserInfo.LID = 0 Then TXT_CBL_TPLANID_S1 = TIMS.GetCblText(CBL_TPLANID_S1)
        '資料篩選方式
        Dim V_RBL_ETENTER_USE As String = TIMS.GetListValue(RBL_ETENTER_USE)
        '轄區分署
        Dim TXT_DDL_DISTID_S1 As String = TIMS.GetListText(DDL_DISTID_S1)

        Lab_RegDateRange.Text = cst_lab_NODATA_TXT
        Lab_TrainDateRange.Text = cst_lab_NODATA_TXT
        Lab_APPMSG.Text = ""
        Lab_PartMSG.Text = ""
        tr_SHOWMSG_1a.Visible = True
        tr_SHOWMSG_1b.Visible = True
        Select Case V_RBL_ETENTER_USE '資料篩選方式
            Case CST_ETENTER_USE_系統預設區間
                tr_SHOWMSG_1a.Visible = False
                tr_SHOWMSG_1b.Visible = False
            Case CST_ETENTER_USE_自選日期區間
                Dim ETENTERDATE1_ROC As String = TIMS.Cdate17(ETENTERDATE1.Text)
                Dim ETENTERDATE2_ROC As String = TIMS.Cdate17(ETENTERDATE2.Text)
                Dim STDATE1_ROC As String = TIMS.Cdate17(STDATE1.Text)
                Dim STDATE2_ROC As String = TIMS.Cdate17(STDATE2.Text)
                If ETENTERDATE1_ROC <> "" AndAlso ETENTERDATE2_ROC <> "" Then
                    Lab_RegDateRange.Text = $"{ETENTERDATE1_ROC}~{ETENTERDATE2_ROC}"
                End If
                If STDATE1_ROC <> "" AndAlso STDATE2_ROC <> "" Then
                    Lab_TrainDateRange.Text = $"{STDATE1_ROC}~{STDATE2_ROC}"
                End If
        End Select

        Dim dt As DataTable = SCH_DATA_DISASTER(V_DDL_DISASTER_S1)
        If TIMS.dtNODATA(dt) Then Return
        Dim dr1 As DataRow = dt.Rows(0)

        hid_ADID.Value = $"{dr1("ADID")}"
        LAB_TITLE1.Text = $"重大災害名稱：{dr1("CNAME_N")}"
        LAB_TPLANNAME.Text = TXT_CBL_TPLANID_S1
        LAB_DISTNAME.Text = TXT_DDL_DISTID_S1
        LAB_AREAS.Text = TIMS.GetDISASTER2_N(hid_ADID.Value, objconn, 1)
        If LAB_AREAS.Text = "" Then LAB_AREAS.Text = cst_lab_NODATA_TXT
        Lab_ReturnDate.Text = $"{dr1("RETURNDATE1_ROC")}"
        Dim dt2 As DataTable = TIMS.GetDISASTER2dt(hid_ADID.Value, objconn)
        Dim v_ZIPCODES As String = ""
        If TIMS.dtHaveDATA(dt2) Then
            For Each dr2 As DataRow In dt2.Rows
                v_ZIPCODES &= $"{If(v_ZIPCODES <> "", ",", "")}{dr2("ZIPCODE")}"
            Next
        End If
        hid_ADID_ZIPCODES.Value = v_ZIPCODES

        'iEXP = 1(匯出資料使用)
        If iEXP = 1 Then Return

        UTL_DIVSHOW(1)
    End Sub
    Sub UTL_DIVSHOW(ITYPE As Integer)
        div_search1.Visible = True
        div_detail1.Visible = False
        If ITYPE = 1 Then
            div_search1.Visible = False
            div_detail1.Visible = True
        End If
    End Sub

    Function CHECKSEARCH1() As String
        Dim RST As String = ""
        Dim V_DDL_DISASTER_S1 As String = TIMS.GetListValue(DDL_DISASTER_S1)
        If V_DDL_DISASTER_S1 = "" Then RST &= "請選擇，重大災害名稱" & vbCrLf

        If sm.UserInfo.LID = 0 Then
            Dim V_CBL_TPLANID_S1 As String = TIMS.GetCblValue(CBL_TPLANID_S1)
            If V_CBL_TPLANID_S1 = "" Then RST &= "請選擇，訓練計畫" & vbCrLf
        End If

        Dim V_DDL_DISTID_S1 As String = TIMS.GetListValue(DDL_DISTID_S1)
        'If V_DDL_DISTID_S1 = "" Then V_DDL_DISTID_S1 = sm.UserInfo.DistID
        If V_DDL_DISTID_S1 = "" Then RST &= "請選擇，轄區分署" & vbCrLf

        ETENTERDATE1.Text = TIMS.Cdate3(ETENTERDATE1.Text)
        ETENTERDATE2.Text = TIMS.Cdate3(ETENTERDATE2.Text)

        If RST <> "" Then Return RST

        LabETENTER.Text = CST_ETENTER_TXT_自選日期區間
        '資料篩選方式
        Dim V_RBL_ETENTER_USE As String = TIMS.GetListValue(RBL_ETENTER_USE)
        '資料篩選方式
        Select Case V_RBL_ETENTER_USE
            Case CST_ETENTER_USE_系統預設區間
                LabETENTER.Text = CST_ETENTER_TXT_系統預設區間
            Case CST_ETENTER_USE_自選日期區間
                If ETENTERDATE1.Text <> "" AndAlso ETENTERDATE2.Text = "" Then
                    RST &= "請選擇或輸入，報名日期區間-迄止,起始有值，不可為空" & vbCrLf
                ElseIf ETENTERDATE1.Text = "" AndAlso ETENTERDATE2.Text <> "" Then
                    RST &= "請選擇或輸入，報名日期區間-起始,迄止有值，不可為空" & vbCrLf
                End If
        End Select

        STDATE1.Text = TIMS.Cdate3(STDATE1.Text)
        STDATE2.Text = TIMS.Cdate3(STDATE2.Text)
        If STDATE1.Text <> "" AndAlso STDATE2.Text = "" Then
            RST &= "請選擇或輸入，開訓日期期間-迄止,起始有值，不可為空" & vbCrLf
        ElseIf STDATE1.Text = "" AndAlso STDATE2.Text <> "" Then
            RST &= "請選擇或輸入，開訓日期期間-起始,迄止有值，不可為空" & vbCrLf
        End If
        If RST <> "" Then Return RST

        '(置換前後)
        If ETENTERDATE1.Text <> "" AndAlso ETENTERDATE2.Text <> "" Then
            Dim dateETENTER1 As Date = TIMS.Cdate2(ETENTERDATE1.Text)
            Dim dateETENTER2 As Date = TIMS.Cdate2(ETENTERDATE2.Text)
            If DateDiff(DateInterval.Day, dateETENTER2, dateETENTER1) > 0 Then
                ETENTERDATE1.Text = TIMS.Cdate3(dateETENTER2)
                ETENTERDATE2.Text = TIMS.Cdate3(dateETENTER1)
            End If
        End If
        '(置換前後)
        If STDATE1.Text <> "" AndAlso STDATE2.Text <> "" Then
            Dim dateSTDATE1 As Date = TIMS.Cdate2(STDATE1.Text)
            Dim dateSTDATE2 As Date = TIMS.Cdate2(STDATE2.Text)
            If DateDiff(DateInterval.Day, dateSTDATE2, dateSTDATE1) > 0 Then
                STDATE1.Text = TIMS.Cdate3(dateSTDATE2)
                STDATE2.Text = TIMS.Cdate3(dateSTDATE1)
            End If
        End If

        Return RST
    End Function

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        UTL_DIVSHOW(0)
        'div_search1.Visible = True 'div_detail1.Visible = False
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Dim ERRMSG1 As String = CHECKSEARCH1()
        If ERRMSG1 <> "" Then
            Common.MessageBox(Me, ERRMSG1)
            Return
        End If

        Call SCH_DISASTER_1(1)
        '資料篩選方式
        Dim V_RBL_ETENTER_USE As String = TIMS.GetListValue(RBL_ETENTER_USE)

        Dim dtH As DataTable = Nothing
        Dim S_EXPT1 As String = "人數"
        Dim V_RBL_SCOPE1 As String = TIMS.GetListValue(RBL_SCOPE1)
        Select Case V_RBL_SCOPE1
            Case CST_SCOPE1_報名
                '報名人數
                S_EXPT1 = "報名人數"
                dtH = SCH_DATA_2(V_RBL_ETENTER_USE)
            Case CST_SCOPE1_參訓
                '參訓人數
                S_EXPT1 = "參訓人數"
                dtH = SCH_DATA_3(V_RBL_ETENTER_USE)
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return
        End Select

        If TIMS.dtNODATA(dtH) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim dsH As New DataSet
        dsH.Tables.Add(dtH)

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        Select Case v_ExpType
            Case "EXCEL"
                Dim s_fileName1 As String = $"{S_EXPT1}-{TIMS.GetDateNo()}.xlsx"
                DbAccess.CloseDbConn(objconn) : ExpClass1.Utl_Export2_XLSX_Direct(Me, dsH, s_fileName1)
            Case "ODS"
                Dim tk As New TurboOdfUtil.ExcelToODS() '讀取參考【TurboOdfUtil】。
                Dim s_fileName1 As String = $"{S_EXPT1}-{TIMS.GetDateNo()}.ods"
                DbAccess.CloseDbConn(objconn) : ExpClass1.Utl_Export2_ODS_Direct(Me, dsH, s_fileName1)
        End Select
    End Sub
End Class
