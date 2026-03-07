Imports System.Threading
Imports System.Threading.Tasks

Partial Class main2
    Inherits AuthBasePage

    Private Shared ReadOnly main2_lock As New Object
    'Private logger As ILog=LogManager.GetLogger(GetType(AppError))
    Dim objconn As SqlConnection
    Dim strSS As String = ""

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '#Region "在這裡放置使用者程式碼以初始化網頁" '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Session("main2") = "xxx" Then
            SyncLock main2_lock
                'Dim iSendMailCount As Integer = TIMS.GlobalMailCount '目前寄信總數量
                If TIMS.GlobalMailCount < 4 Then Call TIMS.DEL_ADP_ZIPFILE(objconn) '錯誤小於4(無錯誤才可繼續)
            End SyncLock
        End If

        If Not IsPostBack Then
            Call SCreate1() '頁面初始化
        End If
    End Sub

    ''' <summary>
    ''' 讀取[視覺化圖表]區塊內容
    ''' </summary>
    Sub Show_VISUALCHART()
        '#Region "讀取[視覺化圖表]區塊內容"
        Dim chartJs As String = ""

        Call TIMS.OpenDbConn(objconn)

        Dim myPram1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"ACCOUNT", sm.UserInfo.UserID}}
        Dim chartSql1 As String = ""
        chartSql1 &= " SELECT DISTINCT TOP(3) b.CRTID, b.TYPEID, b.SORT" & vbCrLf
        chartSql1 &= " FROM PS_VISUALCHART a WITH(NOLOCK)" & vbCrLf
        chartSql1 &= " JOIN VISUALCHART b WITH(NOLOCK) ON a.CRTID=b.CRTID" & vbCrLf
        chartSql1 &= " WHERE b.ISUSED='Y' AND a.SHOW='Y' AND b.TPLANID=@TPLANID AND a.ACCOUNT=@ACCOUNT" & vbCrLf
        chartSql1 &= " ORDER BY b.SORT ASC" & vbCrLf
        Dim chartDT1 As DataTable = DbAccess.GetDataTable(chartSql1, objconn, myPram1)

        If TIMS.dtHaveDATA(chartDT1) Then
            'Dim c1 As Integer=0
            For c1 As Integer = 0 To chartDT1.Rows.Count - 1
                If c1 = 0 Then
                    Dim pic1Path As String = String.Concat("VisualChart/", chartDT1.Rows(0)("TYPEID"), ".aspx")
                    If System.IO.File.Exists(Server.MapPath("~/" + pic1Path)) Then chartJs &= "$('#ifrm1').attr('src', '" & pic1Path & "'); "
                ElseIf c1 = 1 Then
                    Dim pic2Path As String = String.Concat("VisualChart/", chartDT1.Rows(1)("TYPEID"), ".aspx")
                    If System.IO.File.Exists(Server.MapPath("~/" + pic2Path)) Then chartJs &= "$('#ifrm2').attr('src', '" & pic2Path & "'); "
                ElseIf c1 = 2 Then
                    Dim pic3Path As String = String.Concat("VisualChart/", chartDT1.Rows(2)("TYPEID"), ".aspx")
                    If System.IO.File.Exists(Server.MapPath("~/" + pic3Path)) Then chartJs &= "$('#ifrm3').attr('src', '" & pic3Path & "'); "
                End If
            Next
        Else
            Dim myPram2 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
            Dim chartSql2 As String = ""
            chartSql2 &= " SELECT TOP 3 a.TYPEID" & vbCrLf
            chartSql2 &= " FROM VISUALCHART a WITH(NOLOCK)" & vbCrLf
            chartSql2 &= " WHERE a.ISUSED='Y' AND a.DEFAULTSHOW='Y' AND a.TPLANID=@TPLANID" & vbCrLf
            chartSql2 &= " ORDER BY a.SORT ASC" & vbCrLf
            Dim chartDT2 As DataTable = DbAccess.GetDataTable(chartSql2, objconn, myPram2)

            If TIMS.dtHaveDATA(chartDT2) Then
                'Dim c2 As Integer=0
                For c2 As Integer = 0 To chartDT2.Rows.Count - 1
                    If c2 = 0 Then
                        Dim pic1Path As String = String.Concat("VisualChart/", chartDT2.Rows(0)("TYPEID"), ".aspx")
                        If System.IO.File.Exists(Server.MapPath("~/" + pic1Path)) Then chartJs &= "$('#ifrm1').attr('src', '" & pic1Path & "'); "
                    ElseIf c2 = 1 Then
                        Dim pic2Path As String = String.Concat("VisualChart/", chartDT2.Rows(1)("TYPEID"), ".aspx")
                        If System.IO.File.Exists(Server.MapPath("~/" + pic2Path)) Then chartJs &= "$('#ifrm2').attr('src', '" & pic2Path & "'); "
                    ElseIf c2 = 2 Then
                        Dim pic3Path As String = String.Concat("VisualChart/", chartDT2.Rows(2)("TYPEID"), ".aspx")
                        If System.IO.File.Exists(Server.MapPath("~/" + pic3Path)) Then chartJs &= "$('#ifrm3').attr('src', '" & pic3Path & "'); "
                    End If
                Next
            End If
        End If

        If chartJs.Length > 0 Then
            Dim cs As ClientScriptManager = Page.ClientScript
            cs.RegisterStartupScript(Me.GetType(), "doChartData", chartJs, True)
        End If

    End Sub

    '頁面初始化
    Sub SCreate1()
        '#Region "頁面初始化" '#Region "判斷是否要讀取[視覺化圖表]區塊內容"
        Session("main2") = "xxx" '防止session id跳動

        Dim myLID As String = sm.UserInfo.LID.ToString
        Dim flag_can_show_divChart As Boolean = True
        If myLID.Equals("0") OrElse myLID.Equals("1") Then flag_can_show_divChart = True
        If myLID.Equals("2") Then flag_can_show_divChart = False
        '接受企業委託訓練
        If sm.UserInfo.TPlanID = "07" Then flag_can_show_divChart = False
        '區域產業據點職業訓練計畫(在職)
        If sm.UserInfo.TPlanID = "70" Then flag_can_show_divChart = False

        If Not flag_can_show_divChart Then
            divChart.Visible = False
        Else
            Show_VISUALCHART()
        End If

        '#Region "讀取[作業提醒]區塊內容"
        'gv1
        Dim tDt As DataTable = Nothing
        If tDt Is Nothing Then
            tDt = New DataTable
            tDt.Columns.Add("PostDate")
            tDt.Columns.Add("Subject")
            tDt.Columns.Add("msg1")
            Dim tDr As DataRow = tDt.NewRow
            tDr("PostDate") = ""
            tDr("Subject") = "本日無系統作業提醒。"
            tDr("msg1") = ""
            tDt.Rows.Add(tDr)
            gv1.DataSource = tDt
            gv1.DataBind()
        End If

        'gv2
        Dim tDt2 As DataTable = Nothing
        If tDt2 Is Nothing Then
            tDt2 = New DataTable
            tDt2.Columns.Add("PostDate")
            tDt2.Columns.Add("Subject")
            tDt2.Columns.Add("msg1")
            Dim tDr2 As DataRow = tDt2.NewRow
            tDr2("PostDate") = TIMS.Cdate3(Now)
            tDr2("Subject") = "(暫無最新消息)"
            tDr2("msg1") = ""
            tDt2.Rows.Add(tDr2)
            gv2.DataSource = tDt2
            gv2.DataBind()
        End If

        '作業提醒(資料)
        Dim Dt1 As New DataTable
        Dt1.Columns.Add("Subject")
        Dt1.Columns.Add("Msg1")
        Dt1.Columns.Add("Status1")
        Call Warning(Me, sm, objconn, Dt1)

        '作業提醒(顯示)
        Dim Dt2 As New DataTable
        Dt2.Columns.Add("PostDate")
        Dt2.Columns.Add("Subject")
        Dt2.Columns.Add("msg1")
        Dt2.Columns.Add("Status1")

        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        If (flag_test) Then SUtl_AddoneMsg(Dt2, "test msg 1..")

        If TIMS.dtHaveDATA(Dt1) Then
            For Each dr As DataRow In Dt1.Rows
                Dim blnDelete As Boolean = False
                If $"{dr("Subject")}" = "" OrElse Len(dr("Subject")) <= 1 Then blnDelete = True
                If Not blnDelete Then SUtl_AddoneMsg(Dt2, $"{dr("Subject")}")
            Next
        End If
        'Dt2.AcceptChanges()

        If TIMS.dtHaveDATA(Dt2) Then
            'gv1.DataSource=Dt2 '作業提醒。'gv1.DataBind()
            'If Dt2.Rows.Count > 5 Then btnMore1.Visible=(Dt2.Rows.Count > 5) Else btnMore1.Visible=False
            btnMore1.Visible = (Dt2.Rows.Count > 5)
            Dim dtTop5 As DataTable = GetDtTopX(Dt2, 5)
            gv1.DataSource = dtTop5
            gv1.DataBind()
        End If

        Dim tDv As New DataView
        Try
            tDt = TIMS.Get_SelHomeNewsS1(objconn, 0)
            tDt.TableName = "Result"
            tDv.Table = tDt
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(Me)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try

        If tDv.Count > 0 Then
            '抓第一筆訊息區資料 by nick 060306  
            Dim ff As String = "type In (1,2) And msg='Y' AND msg2='Y'" '合理的顯示週數。('未到結束日期。)
            For Each drv As DataRow In tDt.Select(ff)
                Dim SUBJECT As String = ""
                SUBJECT = Convert.ToString(drv("SUBJECT"))
                SUBJECT = Replace(SUBJECT, "！", "！" & vbCrLf)
                SUBJECT = Replace(SUBJECT, "。", "。" & vbCrLf)

                '含提示訊息
                Dim script As String = "<script>window.alert('" & Common.GetJsString(SUBJECT) & "');</script>"
                Me.RegisterStartupScript("", script)
                Exit For
            Next
        End If

        '#Region "讀取[最新消息]、[功能增修說明]、[文件下載]區塊內容"

        'todo: 1:最新消息 2: 功能增修說明 3: 文件下載 4:影音教學專區

        Dim dtNewsAll As DataTable = TIMS.Get_SelHomeNewsS1(objconn, 0)
        dtNewsAll.TableName = "Result"

        Dim iMax1 As Integer = 4
        For i_todo As Integer = 1 To iMax1
            'Dim dt As DataTable=TIMS.Get_SelHomeNewsS1(i_todo, objconn)
            'dt.TableName="Result"
            Dim dv As New DataView With {
                .Table = dtNewsAll 'dt
                }
            Call ShowGVData(i_todo, dv)
        Next
    End Sub

    Public Sub ShowGVData(ByRef i As Integer, ByRef dv As DataView)
        Select Case i.ToString
            Case "1"
                dv.RowFilter = "Type=1"
                btnMore2.Visible = False
                If dv.Count > 0 Then
                    btnMore2.Visible = If(dv.Count > 5, True, False)
                    gv2.DataSource = GetDtTopX(dv.ToTable, 5)
                    gv2.DataBind()
                End If
            Case "2"
                dv.RowFilter = "Type=2"
                btnMore3.Visible = False
                If dv.Count > 0 Then
                    btnMore3.Visible = If(dv.Count > 5, True, False)
                    gv3.DataSource = GetDtTopX(dv.ToTable, 5)
                    gv3.DataBind()
                End If
            Case "3"
                dv.RowFilter = "Type=3"
                btnMore4.Visible = False
                If dv.Count > 0 Then
                    btnMore4.Visible = If(dv.Count > 5, True, False)
                    gv4.DataSource = GetDtTopX(dv.ToTable, 5)
                    gv4.DataBind()
                End If
            Case "4"
                dv.RowFilter = "Type=4"
                btnMore5.Visible = False
                If dv.Count > 0 Then
                    btnMore5.Visible = If(dv.Count > 5, True, False)
                    gv5.DataSource = GetDtTopX(dv.ToTable, 5)
                    gv5.DataBind()
                End If
        End Select
    End Sub

    '#Region "取得[作業提醒]資料內容"
    Public Shared Sub Warning(ByRef MyPage As Page, ByRef sm As SessionModel, ByRef objconn As SqlConnection, ByRef oDt As DataTable)
        'sErrmsg 用換行常切換資料。'提示訊息，不跳離系統'確認登入帳號之機構是否在黑名單中 20090724 by AMU
        Dim vSubject As String = Check_AccoutBlackList(MyPage, sm, objconn)
        Call SUtl_AddoneMsg(oDt, vSubject)

        Dim strSS As String = ""
        Dim sHref As String = ""
        Const cst_maindetail5aspx As String = "main2_detail.aspx?todo=5"
        'Const cst_maindetail6aspx As String = "main2_detail.aspx?todo=6"

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        Call TIMS.OpenDbConn(objconn)

        '+作業提醒,'查詢審查計分等級開關機制,審查計分表開關機制
        Dim dtTQSQUERY_1 As DataTable = TIMS.Get_TTQSQUERY_TB(objconn, 1)
        If TIMS.dtHaveDATA(dtTQSQUERY_1) Then
            For Each drTQSQUERY_1 As DataRow In dtTQSQUERY_1.Rows
                Dim s_QEXPLAIN As String = $"{drTQSQUERY_1("QEXPLAIN")}"
                If s_QEXPLAIN <> "" Then SUtl_AddoneMsg(oDt, s_QEXPLAIN)
            Next
        End If
        '+作業提醒,'查詢審查計分等級開關機制,審查計分表開關機制
        Dim dtTQSQUERY_2 As DataTable = TIMS.Get_TTQSQUERY_TB(objconn, 2)
        If TIMS.dtHaveDATA(dtTQSQUERY_2) Then
            For Each drTQSQUERY_2 As DataRow In dtTQSQUERY_2.Rows
                Dim s_REMIND1 As String = $"{drTQSQUERY_2("REMIND1")}"
                If s_REMIND1 <> "" Then SUtl_AddoneMsg(oDt, s_REMIND1)
            Next
        End If

        '+作業提醒 '即日起開放訓練單位確認TTQS評核結果，請於YYYY/MM/DD【控制結束日】前至：【訓練機構管理>>最近一次TTQS評核結果確認】功能，完成確認作業。"
        Dim s_TTQSLOCKMsg1 As String = TIMS.Get_TTQSLOCKMsg1(objconn)
        If s_TTQSLOCKMsg1 <> "" Then SUtl_AddoneMsg(oDt, s_TTQSLOCKMsg1)

        '顯示的格式為「ＸＸＸ班有不具失、待業身分者」
        If TIMS.Cst_TPlanID_BliDet201605.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql = "" & vbCrLf
            sql &= " SELECT DISTINCT cc.RID, cc.PLANID, cc.OCID" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) + ' 有不具失、待業身分者' XMSG" & vbCrLf
            sql &= " FROM STUD_SELRESULTBLIDET bd WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN STUD_SELRESULTBLI bi WITH(NOLOCK) ON bi.SB3ID=bd.SB3ID" & vbCrLf
            sql &= " JOIN CLASS_CLASSINFO cc WITH(NOLOCK) ON cc.ocid=bd.ocid" & vbCrLf
            sql &= " JOIN VIEW_PLAN ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
            sql &= " LEFT JOIN Stud_SelResult t1 WITH(NOLOCK) ON t1.setid=bi.setid AND t1.enterdate=bi.enterdate AND t1.sernum=bi.sernum" & vbCrLf
            sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
            'sql &= " AND cc.STDATE > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
            sql &= " AND cc.STDATE > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
            sql &= " AND ip.tplanid IN (" & TIMS.Cst_TPlanID_BliDet201605 & ")" & vbCrLf
            sql &= " AND ip.tplanid NOT IN (" & TIMS.Cst_TPlanID_BliDet201605_NG & ")" & vbCrLf
            sql &= " AND ISNULL(t1.Admission,'Y')='Y'" & vbCrLf '錄取或尚未錄取者 
            sql &= " AND bd.STATUSPT IN ('ES3','ES2')" & vbCrLf '甄試日前
            sql &= " AND bi.ACTNO IS NOT NULL" & vbCrLf '有加保事實
            sql &= " AND ISNULL(bd.STATUSNC1,' ')!='1'" & vbCrLf '尚未處理轉知
            sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

            Try
                Dim ss As String = "XMSG"
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        'sErrmsg += Convert.ToString(dr("xMsg")) & vbCrLf
                        '甄試前檢核/開訓前檢核，ＯＯＯ不具失、待業身分，請轉知ＯＯＯ於開訓日前須符合參訓資格，否則不得參訓。
                        strSS = "" 'Dim strSS As String=""
                        TIMS.SetMyValue(strSS, "RID", dr("RID"))
                        TIMS.SetMyValue(strSS, "PlanID", dr("PLANID"))
                        TIMS.SetMyValue(strSS, "OCID", dr("OCID"))
                        TIMS.SetMyValue(strSS, "STATUSPT", "ES")
                        sHref = $"{cst_maindetail5aspx}{strSS}"
                        vSubject = $"<a href='{sHref}' style='color:#1f336b;text-decoration:underline'>{dr("XMSG")}{vbCrLf}</a>{vbCrLf}"
                        Call SUtl_AddoneMsg(oDt, vSubject)
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try

            'STATUSPT:AS
            sql = "" & vbCrLf
            sql &= " SELECT DISTINCT cc.RID, cc.PLANID, cc.OCID" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) + ' 有不具失、待業身分者' XMSG" & vbCrLf
            sql &= " FROM STUD_SELRESULTBLIDET bd WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN STUD_SELRESULTBLI bi WITH(NOLOCK) ON bi.SB3ID=bd.SB3ID" & vbCrLf
            sql &= " JOIN CLASS_CLASSINFO cc WITH(NOLOCK) ON cc.ocid=bd.ocid" & vbCrLf
            sql &= " JOIN VIEW_PLAN ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
            sql &= " JOIN STUD_SELRESULT t1 WITH(NOLOCK) ON t1.setid=bi.setid AND t1.enterdate=bi.enterdate AND t1.sernum=bi.sernum AND ISNULL(t1.Admission,'Y')='Y'" & vbCrLf
            'sql &= " AND cc.STDATE > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
            sql &= " WHERE t1.Admission='Y'" & vbCrLf '錄取者 
            sql &= " AND cc.STDATE > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
            sql &= " AND bd.STATUSPT IN ('AS2')" & vbCrLf '開訓日前
            sql &= " AND bi.ACTNO IS NOT NULL" & vbCrLf '有加保事實
            sql &= " AND ISNULL(bd.STATUSNC1,' ') != '1'" & vbCrLf '尚未處理轉知
            sql &= " AND ip.TPLANID IN (" & TIMS.Cst_TPlanID_BliDet201605 & ")" & vbCrLf
            sql &= " AND ip.TPLANID NOT IN (" & TIMS.Cst_TPlanID_BliDet201605_NG & ")" & vbCrLf
            sql &= " AND cc.IsSuccess='Y'" & vbCrLf
            sql &= " AND cc.NotOpen='N'" & vbCrLf
            sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

            Try
                Dim ss As String = "XMSG"
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        strSS = "" 'Dim strSS As String=""
                        TIMS.SetMyValue(strSS, "RID", dr("RID"))
                        TIMS.SetMyValue(strSS, "PlanID", dr("PLANID"))
                        TIMS.SetMyValue(strSS, "OCID", dr("OCID"))
                        TIMS.SetMyValue(strSS, "STATUSPT", "AS")
                        sHref = $"{cst_maindetail5aspx}{strSS}"
                        vSubject = $"<a href='{sHref}' style='color:#1f336b;text-decoration:underline'>{dr("XMSG")}{vbCrLf}</a>{vbCrLf}"
                        Call SUtl_AddoneMsg(oDt, vSubject)
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg)
                'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try

            '查詢目前計畫所有學員
            'Dim strSS As String=""
            strSS = ""
            TIMS.SetMyValue(strSS, "PlanID", sm.UserInfo.PlanID)
            TIMS.SetMyValue(strSS, "RID", sm.UserInfo.RID)
            Dim dtStud As DataTable = TIMS.Get_STUDENTINFO(strSS, objconn)
            '現有"就保非自願離職者"及"公司負責人身分"的提醒，還有"不具失、待業身分者"及"屆退官兵
            '"的提醒(詳如上一項)，由於資料過於繁雜詳細，恐反易造成各承辦人麻痺疏失，
            '故均改為「ＸＸＸ班有就保非自願離職者」、「ＸＸＸ班有公司負責人身分者」、
            '「ＸＸＸ班有不具失、待業身分者」及「ＸＸＸ班有屆退官兵身分者」。
            '其中"不具失、待業身分者"與"屆退官兵"均需開啟下一個頁面，顯示該班次被查核到的名單。
            If TIMS.dtHaveDATA(dtStud) Then
                'STATUSPT:ST
                'ＯＯＯ於訓中加保，請查明是否有工作事實，並予以離退訓。
                sql = "" & vbCrLf
                sql &= " SELECT DISTINCT bd.IDNO, bd.OCID, bi.NAME + '於訓中加保，請查明是否有工作事實，並予以離退訓。' XMSG" & vbCrLf
                sql &= " FROM STUD_SELRESULTBLIDET bd WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN STUD_SELRESULTBLI bi WITH(NOLOCK) ON bi.SB3ID=bd.SB3ID" & vbCrLf
                sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=bd.ocid" & vbCrLf
                sql &= " JOIN view_plan ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
                sql &= " JOIN Stud_SelResult t1 WITH(NOLOCK) ON t1.setid=bi.setid AND t1.enterdate=bi.enterdate AND t1.sernum=bi.sernum AND ISNULL(t1.Admission,'Y')='Y'" & vbCrLf
                'sql &= " AND cc.STDATE < CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf '開訓日後
                'sql &= " AND cc.FTDATE > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf '結訓日前
                sql &= " WHERE t1.Admission='Y'" & vbCrLf '錄取者 
                sql &= " AND cc.STDATE < dbo.TRUNC_DATETIME(GETDATE()) AND cc.FTDATE > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
                sql &= " AND bd.STATUSPT IN ('ST1','ST2')" & vbCrLf '開訓日後
                sql &= " AND bi.ACTNO IS NOT NULL" & vbCrLf '有加保事實
                'STATUSNC2: 1:未有工作事實
                '2.已轉知項目-使用者點選「未有工作事實（請說明）」選項，點選此項目時「處理說明」欄位設為必填，儲存後不需加做離退訓作業，也取消首頁告警
                sql &= " AND ISNULL(bd.STATUSNC2,' ') != '1'" & vbCrLf '1:未有工作事實
                sql &= " AND ip.tplanid IN (" & TIMS.Cst_TPlanID_BliDet201605 & ")" & vbCrLf
                sql &= " AND ip.tplanid NOT IN (" & TIMS.Cst_TPlanID_BliDet201605_NG & ")" & vbCrLf
                sql &= " AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " AND cc.NotOpen='N'" & vbCrLf
                sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

                Try
                    Dim ss As String = "XMSG"
                    dt = DbAccess.GetDataTable(sql, objconn)
                    If TIMS.dtHaveDATA(dt) Then
                        For Each dr As DataRow In dt.Select(Nothing, ss)
                            Dim ff As String = $"IDNO='{dr("IDNO")}' AND OCID='{dr("OCID")}'"
                            If dtStud.Select(ff).Length > 0 Then
                                Call SUtl_AddoneMsg(oDt, $"{dr("XMSG")}")
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                    Call TIMS.WriteTraceLog(strErrmsg)
                    'Call TIMS.CloseDbConn(objconn) 'Throw ex
                End Try

            End If

        End If

        '本作業僅限自辦職前訓練計畫
        '依據屆退官兵荐訓名冊資料，未開訓之班級，於甄試前2日檢核報名資料，
        '檢核報名資料之對象是否有存在屆退官兵荐訓名冊資料之中，
        '且檢核當日檢核"預定退伍日"是否已過該日，
        '若未過，則於作業提醒顯示"XXXX班報名資料有屆退官兵身分者"，前述資訊，在班級開訓日當日起，則不再顯示。
        'ex：甄試日：104/09/11，勾稽日：104/09/09
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    sql = "" & vbCrLf
        '    sql &= " WITH T1 AS (" & vbCrLf
        '    sql &= " SELECT DISTINCT b.OCID1" & vbCrLf
        '    sql &= " FROM STUD_ENTERTEMP a WITH(NOLOCK)" & vbCrLf
        '    sql &= " JOIN STUD_ENTERTYPE b WITH(NOLOCK) ON b.SETID=a.SETID" & vbCrLf
        '    sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.OCID1" & vbCrLf
        '    sql &= " JOIN id_plan ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
        '    'sql &= " AND CAST(CONVERT(VARCHAR, b.PREEXDATE, 111) AS DATETIME) > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
        '    sql &= " AND dbo.TRUNC_DATETIME(b.PREEXDATE) > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        '    '未開訓之班級，於甄試前2日檢核報名資料，
        '    sql &= " AND dbo.TRUNC_DATETIME(cc.ExamDate - 2) <=  dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    '未開訓之班級，於甄試前2日檢核報名資料，
        '    sql &= " AND cc.STDATE > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    '限定200.100天的系統範圍
        '    sql &= " AND cc.STDATE >= GETDATE() - 200" & vbCrLf
        '    sql &= " AND cc.STDATE <= GETDATE() + 100" & vbCrLf
        '    sql &= " AND b.PREEXDATE IS NOT NULL" & vbCrLf
        '    sql &= " AND b.MODIFYDATE >= GETDATE() - 200" & vbCrLf
        '    sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        '    sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        '    sql &= " )" & vbCrLf
        '    sql &= " ,T2 AS (" & vbCrLf
        '    sql &= " SELECT DISTINCT b.OCID1" & vbCrLf
        '    sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
        '    sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.ESETID=a.ESETID" & vbCrLf
        '    sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.OCID1" & vbCrLf
        '    sql &= " JOIN id_plan ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
        '    'sql &= " AND CAST(CONVERT(VARCHAR, b.PREEXDATE, 111) AS DATETIME) > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
        '    sql &= " AND dbo.TRUNC_DATETIME(b.PREEXDATE) > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        '    '未開訓之班級，於甄試前2日檢核報名資料，
        '    'sql &= " AND CAST(CONVERT(VARCHAR, cc.ExamDate - 2, 111) AS DATETIME) <= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
        '    sql &= " AND dbo.TRUNC_DATETIME(cc.ExamDate - 2) <= dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    '未開訓之班級，於甄試前2日檢核報名資料，
        '    'sql &= " AND cc.STDATE > CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
        '    sql &= " AND cc.STDATE > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        '    '限定200.100天的系統範圍
        '    sql &= " AND cc.STDATE >= GETDATE() - 200" & vbCrLf
        '    sql &= " AND cc.STDATE <= GETDATE() + 100" & vbCrLf
        '    sql &= " AND b.PREEXDATE IS NOT NULL" & vbCrLf
        '    sql &= " AND b.modifydate >= GETDATE() - 200" & vbCrLf
        '    sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        '    sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        '    sql &= " )" & vbCrLf

        '    sql &= " SELECT DISTINCT cc.RID, cc.PLANID, cc.OCID" & vbCrLf
        '    sql &= " ,dbo.FN_CLSNAME(cc.classcname) + '報名資料有屆退官兵身分者' XMSG" & vbCrLf
        '    sql &= " FROM class_classinfo cc WITH(NOLOCK)" & vbCrLf
        '    sql &= " JOIN (SELECT * FROM T1 UNION SELECT * FROM T2) b ON b.OCID1=cc.OCID" & vbCrLf
        '    sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf

        '    Try
        '        Dim ss As String = "XMSG"
        '        dt = DbAccess.GetDataTable(sql, objconn)
        '        For Each dr As DataRow In dt.Select(Nothing, ss)
        '            TIMS.SetMyValue(strSS, "RID", dr("RID"))
        '            TIMS.SetMyValue(strSS, "PlanID", dr("PLANID"))
        '            TIMS.SetMyValue(strSS, "OCID", dr("OCID"))
        '            sHref = cst_maindetail6aspx & strSS
        '            vSubject = ""
        '            vSubject &= "<a href='" & sHref & "' style='color:#1f336b;text-decoration:underline'>"
        '            vSubject &= Convert.ToString(dr("XMSG")) & vbCrLf
        '            vSubject &= "</a>" & vbCrLf
        '            Call sUtl_AddoneMsg(oDt, vSubject)
        '        Next
        '    Catch ex As Exception
        '        Dim strErrmsg As String = ""
        '        strErrmsg &= TIMS.GetErrorMsg(MyPage) & vbCrLf
        '        strErrmsg &= "ex.ToString:" & vbCrLf
        '        strErrmsg &= ex.ToString & vbCrLf
        '        'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '        Call TIMS.WriteTraceLog(strErrmsg)
        '        'Call TIMS.CloseDbConn(objconn)
        '        'Throw ex
        '    End Try
        'End If

        '檢查是否為負責人 (Y:是負責人 N:不是 ERROR:異常) '限定計畫執行
        '(分署/單位限定)
        '公司負責人身分
        If TIMS.Cst_NotTPlanID5.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            sql = "" & vbCrLf
            sql &= " SELECT DISTINCT b.OCID1, dbo.FN_CLSNAME(cc.classcname) + '有公司負責人身分者' XMSG" & vbCrLf
            sql &= " FROM STUD_ENTERTEMP a WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN STUD_ENTERTYPE b WITH(NOLOCK) ON b.SETID=a.SETID" & vbCrLf
            sql &= " JOIN CLASS_CLASSINFO cc WITH(NOLOCK) ON cc.ocid=b.OCID1 AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
            'sql &= " AND cc.STDATE <= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) + 1" & vbCrLf '開訓日期前1天
            'sql &= " AND cc.FTDATE >= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf '直至結訓日過
            sql &= " WHERE cc.STDATE<=dbo.TRUNC_DATETIME(GETDATE()+1)" & vbCrLf '開訓日期前1天
            sql &= " AND cc.FTDATE>=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf '直至結訓日過
            sql &= " AND b.CMASTER1='Y'" & vbCrLf '認定為公司負責人
            sql &= " AND b.CMASTER1NS IS NULL" & vbCrLf '未轉知
            sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
            sql &= " AND NOT EXISTS (" & vbCrLf
            sql &= "   SELECT 'X'" & vbCrLf
            sql &= "   FROM STUD_ENTERTEMP2 xa WITH(NOLOCK)" & vbCrLf
            sql &= "   JOIN STUD_ENTERTYPE2 xb WITH(NOLOCK) ON xb.eSETID=xa.eSETID" & vbCrLf
            sql &= "   JOIN class_classinfo xcc WITH(NOLOCK) ON xcc.ocid=xb.OCID1 AND xcc.IsSuccess='Y' AND xcc.NotOpen='N'" & vbCrLf
            'sql &= "   AND xcc.STDATE <= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) + 1" & vbCrLf '開訓日期前1天
            'sql &= "   AND xcc.FTDATE >= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf '直至結訓日過
            sql &= "   WHERE xcc.STDATE<=dbo.TRUNC_DATETIME(GETDATE()+1)" & vbCrLf '開訓日期前1天
            sql &= "   AND xcc.FTDATE>=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf '直至結訓日過
            sql &= "   AND xcc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= "   AND xcc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
            sql &= "   AND xb.CMASTER1='Y'" & vbCrLf '認定為公司負責人
            sql &= "   AND xb.CMASTER1NS IS NULL" & vbCrLf '未轉知
            sql &= "   AND xb.OCID1=b.OCID1 AND xa.IDNO=a.IDNO" & vbCrLf '依學員
            sql &= " )" & vbCrLf
            sql &= " UNION" & vbCrLf
            sql &= " SELECT DISTINCT b.OCID1, dbo.FN_CLSNAME(cc.CLASSCNAME) + '有公司負責人身分者' XMSG" & vbCrLf
            sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.ESETID=a.ESETID" & vbCrLf
            sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.OCID1 AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
            'sql &= " AND cc.STDATE <= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) + 1" & vbCrLf '開訓日期前1天
            'sql &= " AND cc.FTDATE >= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf '直至結訓日過
            sql &= " WHERE cc.STDATE<=dbo.TRUNC_DATETIME(GETDATE()+1)" & vbCrLf '開訓日期前1天
            sql &= " AND cc.FTDATE>=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf '直至結訓日過
            sql &= " AND b.CMASTER1='Y'" & vbCrLf '認定為公司負責人
            sql &= " AND b.CMASTER1NS IS NULL" & vbCrLf '未轉知
            sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

            Try
                Dim ss As String = "XMSG"
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        Call SUtl_AddoneMsg(oDt, Convert.ToString(dr("XMSG")))
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try
        End If

        '尚有○○○○(訓練單位名稱)○○○○課(課程名稱)須填寫不預告(電話)抽訪學員紀錄表
        '產投(分署)提醒
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso Val(sm.UserInfo.LID) = 1 Then
            'Dim sql As String=""
            sql = "" & vbCrLf
            sql &= " WITH WC1 AS (" & " SELECT DISTINCT cc.OCID,oo.ORGNAME,cc.CLASSCNAME,cc.CYCLTYPE" & vbCrLf 'DISTINCT 有多次抽訪紀錄
            sql &= " FROM CLASS_UNEXPECTVISITOR U1 WITH(NOLOCK)" & vbCrLf
            sql &= " LEFT JOIN CLASS_CLASSINFO cc WITH(NOLOCK) ON U1.OCID=cc.OCID" & vbCrLf
            sql &= " LEFT JOIN ORG_ORGINFO oo WITH(NOLOCK) on oo.COMIDNO=cc.COMIDNO" & vbCrLf
            sql &= " LEFT JOIN ID_PLAN ip on ip.PLANID=cc.PLANID" & vbCrLf
            sql &= " WHERE cc.ISSUCCESS='Y' AND cc.NOTOPEN='N'" & vbCrLf
            sql &= " AND cc.STDATE <= dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
            sql &= " AND (U1.SITEM1='5' OR U1.SITEM1B='5')" & vbCrLf
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & " )" & vbCrLf

            sql &= " ,WC2 AS (" & " SELECT cc.OCID" & vbCrLf
            sql &= " ,COUNT(CASE WHEN U2.OCID IS NOT NULL THEN 1 END) U2CNT1" & vbCrLf
            sql &= " ,COUNT(CASE WHEN U2.TELVISITREASON='2' THEN 1 END) U2CNT2" & vbCrLf
            sql &= " FROM WC1 cc" & vbCrLf
            sql &= " LEFT JOIN CLASS_UNEXPECTTEL U2 WITH(NOLOCK) ON U2.OCID=cc.OCID" & vbCrLf
            'sql &= " WHERE 1=1" & vbCrLf
            sql &= " GROUP BY cc.OCID" & " )" & vbCrLf

            sql &= " ,WC3 AS (" & " SELECT cc.OCID,cc.ORGNAME,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
            sql &= " FROM WC1 cc WITH(NOLOCK)" & vbCrLf
            sql &= " LEFT JOIN WC2 U2 WITH(NOLOCK) ON cc.OCID=U2.OCID" & vbCrLf
            sql &= " WHERE (U2.U2CNT1=0 OR U2.U2CNT2=0)" & " )" & vbCrLf

            sql &= " SELECT '尚有('+ORGNAME+')('+CLASSCNAME+')須填寫不預告(電話)抽訪學員紀錄表' XMSG" & vbCrLf
            sql &= " FROM WC3" & vbCrLf
            Try
                Dim ss1 As String = "XMSG"
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss1)
                        Call SUtl_AddoneMsg(oDt, Convert.ToString(dr("XMSG")))
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try

            sql = "" & vbCrLf
            sql &= " SELECT oo.orgname + '-' + cc.classcname + '，班級變更申請作業已逾20天尚未給予審核，請再確認！' XMSG" & vbCrLf
            sql &= " FROM PLAN_REVISE rr WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN id_plan ip WITH(NOLOCK) ON ip.planid=rr.planid" & vbCrLf
            sql &= " JOIN plan_planinfo pp WITH(NOLOCK) ON pp.planid=rr.planid AND pp.comidno=rr.comidno AND pp.seqno=rr.seqno" & vbCrLf
            sql &= " JOIN org_orginfo oo WITH(NOLOCK) ON oo.comidno=pp.comidno" & vbCrLf
            sql &= " LEFT JOIN class_classinfo cc WITH(NOLOCK) ON pp.planid=cc.planid AND pp.comidno=cc.comidno AND pp.seqno=cc.seqno" & vbCrLf
            sql &= " WHERE rr.REVISESTATUS IS NULL" & vbCrLf
            'sql &= " and trunc(rr.CDATE)>= trunc(sysdate-30)" & vbCrLf
            'sql &= " and trunc(rr.cdate)<= trunc(sysdate-20)" & vbCrLf
            'sql &= " AND CAST(CONVERT(VARCHAR, rr.cdate, 111) AS DATETIME) >= CAST(CONVERT(VARCHAR, GETDATE() - 30, 111) AS DATETIME)" & vbCrLf
            'sql &= " and CAST(CONVERT(VARCHAR, rr.cdate, 111) AS DATETIME) <= CAST(CONVERT(VARCHAR, GETDATE() - 20, 111) AS DATETIME)" & vbCrLf
            sql &= " AND rr.CDATE >=dbo.TRUNC_DATETIME(GETDATE()-30)" & vbCrLf '前30天
            sql &= " AND rr.CDATE <=dbo.TRUNC_DATETIME(GETDATE()-20)" & vbCrLf '前20天
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf

            Try
                Dim ss As String = "XMSG" 'sort
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        Call SUtl_AddoneMsg(oDt, Convert.ToString(dr("XMSG")))
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try
        End If

        '產投(單位)提醒
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso Val(sm.UserInfo.LID) = 2 Then
            '產投：14天開訓資料提醒---部分辨訓單位（尤其是新單位）
            '，因初次執行產投方案，故在課程期程上，有時疏於管控
            '，導致開訓已到14天，尚未在系統完成學員錄取
            '，甚至維護的缺失，造成常又要請系統從後台去重新維護
            '，所以是否可請系統未來在單位端部分
            '，若尚未完成錄取相關作業，能增設警示功能。
            sql = ""
            sql &= " SELECT cc.OCID,oo.ORGNAME,cc.CLASSCNAME" & vbCrLf
            sql &= " ,cc.RID,cc.TNUM" & vbCrLf
            sql &= " ,ISNULL(se.cnt1,0) CNT1" & vbCrLf
            sql &= " ,ISNULL(se.cnt3,0) CNT3" & vbCrLf
            sql &= " ,CONVERT(varchar, cc.stdate, 111) STDATE" & vbCrLf
            sql &= " ,CONVERT(varchar, cc.ftdate, 111) FTDATE" & vbCrLf
            sql &= " FROM CLASS_CLASSINFO cc WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN ID_PLAN ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
            sql &= " JOIN ORG_ORGINFO oo WITH(NOLOCK) ON oo.comidno=cc.comidno" & vbCrLf
            sql &= " JOIN (" & vbCrLf
            sql &= "   SELECT b.OCID1" & vbCrLf
            sql &= "   ,COUNT(1) CNT1" & vbCrLf
            'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
            sql &= "   ,COUNT(CASE WHEN b.signupstatus=0 THEN 1 END) CNT3" & vbCrLf
            sql &= "   FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
            sql &= "   JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON a.ESETID=b.ESETID" & vbCrLf
            sql &= "   JOIN class_classinfo cx WITH(NOLOCK) ON cx.ocid=b.ocid1" & vbCrLf
            sql &= "   JOIN id_plan ipx WITH(NOLOCK) ON ipx.planid=cx.planid" & vbCrLf
            sql &= "   WHERE cx.IsSuccess='Y' AND cx.NotOpen='N'" & vbCrLf
            sql &= "   AND ipx.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= "   AND cx.RID='" & sm.UserInfo.RID & "'" & vbCrLf
            sql &= "   GROUP BY b.ocid1" & vbCrLf
            '尚未在系統完成學員錄取
            'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
            sql &= "   HAVING COUNT(CASE WHEN b.signupstatus=0 THEN 1 END) > 0" & vbCrLf
            sql &= " ) se ON se.ocid1=cc.ocid" & vbCrLf
            sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
            'sql &= " AND CAST(CONVERT(VARCHAR, cc.stdate - 14, 111) AS DATETIME) <= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
            'sql &= " and CAST(CONVERT(VARCHAR, cc.stdate + 14, 111) AS DATETIME) >= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
            sql &= " AND cc.STDATE-14 <=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf '前14天
            sql &= " AND cc.STDATE+14 >=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf '後14天
            sql &= " and ip.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

            Try
                Dim ss As String = "CLASSCNAME" 'sort
                dt = DbAccess.GetDataTable(sql, objconn)

                Dim tClasscnames As String = ""
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        tClasscnames &= String.Concat(If(tClasscnames <> "", "、", ""), dr("CLASSCNAME"))
                    Next
                End If
                If tClasscnames <> "" Then
                    vSubject = $"貴單位 班級「{tClasscnames}」尚有未錄取作業。"
                    Call SUtl_AddoneMsg(oDt, vSubject)
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn)'Throw ex      
            End Try
        End If

        '訓後動態調查表
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso Val(sm.UserInfo.LID) = 2 Then
            '要填寫訓後動態調查表
            'To fill out the post-training dynamic questionnaire
            Call Utl_TRAIN_DYNAMIC_QUESTION(MyPage, sm, objconn, oDt)
        End If


        '產投(單位)提醒
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso Val(sm.UserInfo.LID) = 2 Then
            '新增訊息：首頁
            '如有民眾自行取消報名的課程，於訓練單位登入時出現(浮出)提示訊息
            '貴單位 班級名稱有民眾自行取消報名，請至「民眾自行取消報名」查詢。
            '此訊息於該班開訓後，就不顯示。
            sql = ""
            sql &= " SELECT DISTINCT cc.CLASSCNAME" & vbCrLf
            sql &= " FROM stud_enterTemp2 a WITH(NOLOCK)" & vbCrLf
            sql &= " JOIN stud_enterType2DelData b WITH(NOLOCK) ON a.esetid=b.esetid AND a.idno=b.modifyAcct" & vbCrLf
            sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.ocid1" & vbCrLf
            sql &= " JOIN id_plan ip WITH(NOLOCK) ON ip.planid=cc.planid" & vbCrLf
            sql &= " WHERE cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
            'sql &= " and cc.STDATE >= CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME)" & vbCrLf
            sql &= " and cc.STDATE >= GETDATE()" & vbCrLf
            sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

            Try
                Dim ss As String = "CLASSCNAME" 'sort
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    For Each dr As DataRow In dt.Select(Nothing, ss)
                        vSubject = $"貴單位 班級「{dr("CLASSCNAME")}」有民眾自行取消報名，請至「民眾自行取消報名」查詢。"
                        Call SUtl_AddoneMsg(oDt, vSubject)
                    Next
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try
        End If

        '可登入者所有人 
        '就保非自願離職者 '檢查排程 DBT_2335
        'Dim sql As String=""
        sql = ""
        sql &= " WITH WCF1 AS (" & " SELECT DISTINCT cc.PLANID, cc.RID, b.OCID1, dbo.fn_clsname(cc.classcname) + '有就保非自願離職者' XMSG" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.ESETID=a.ESETID" & vbCrLf
        sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.OCID1 AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        sql &= " WHERE b.CFIRE1='Y' AND b.CFIRE1NS IS NULL" & vbCrLf
        sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & " )" & vbCrLf
        sql &= " SELECT DISTINCT b.OCID1, dbo.FN_CLSNAME(cc.CLASSCNAME) + '有就保非自願離職者' XMSG" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b WITH(NOLOCK) ON b.SETID=a.SETID" & vbCrLf
        sql &= " JOIN class_classinfo cc WITH(NOLOCK) ON cc.ocid=b.OCID1 AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        sql &= " WHERE b.ENTERPATH NOT IN ('W','R')" & vbCrLf
        sql &= " AND b.CFIRE1='Y' AND b.CFIRE1NS IS NULL" & vbCrLf
        sql &= " AND cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= " AND cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql &= " AND NOT EXISTS (SELECT 'X' FROM WCF1 X WHERE X.PlanID=cc.PlanID AND X.RID=cc.RID)" & vbCrLf
        sql &= " UNION" & vbCrLf
        sql &= " SELECT DISTINCT b.OCID1, b.XMSG FROM WCF1 b" & vbCrLf

        Try
            Dim ss As String = "XMSG"
            dt = DbAccess.GetDataTable(sql, objconn)
            If TIMS.dtHaveDATA(dt) Then
                For Each dr As DataRow In dt.Select(Nothing, ss)
                    Call SUtl_AddoneMsg(oDt, Convert.ToString(dr("XMSG")))
                Next
            End If
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
        End Try

        '辦訓機構 訊息。
        sql = "" & vbCrLf
        sql &= " SELECT a.CLASSCNAME, a.CYCLTYPE, a.STDATE, b.TNUM" & vbCrLf
        sql &= " ,ISNULL(c.EnterCount,0) ENTERCOUNT" & vbCrLf '報名數
        sql &= " ,ISNULL(d.TrainCount,0) TRAINCOUNT" & vbCrLf '訓練數
        sql &= " FROM Class_ClassInfo a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN Plan_PlanInfo b WITH(NOLOCK) ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo" & vbCrLf
        sql &= " AND a.PlanID='" & sm.UserInfo.PlanID & "' AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql &= " AND a.NotOpen='N' AND a.IsSuccess='Y'" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "   SELECT m.OCID1, count(1) ENTERCOUNT" & vbCrLf '報名數
        sql &= "   FROM Stud_EnterType m WITH(NOLOCK)" & vbCrLf
        sql &= "   JOIN Class_ClassInfo a WITH(NOLOCK) ON a.OCID=m.OCID1" & vbCrLf
        sql &= "   WHERE a.PlanID='" & sm.UserInfo.PlanID & "' AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql &= "   AND a.NotOpen='N' AND a.IsSuccess='Y'" & vbCrLf
        sql &= "   GROUP BY m.OCID1 ) c ON a.OCID=c.OCID1" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "   SELECT m.OCID, count(1) TRAINCOUNT" & vbCrLf '訓練數
        sql &= "   FROM Class_StudentsOfClass m WITH(NOLOCK)" & vbCrLf
        sql &= "   JOIN Class_ClassInfo a WITH(NOLOCK) ON a.OCID=m.OCID" & vbCrLf
        sql &= "   WHERE a.PlanID='" & sm.UserInfo.PlanID & "' AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql &= "   AND a.NotOpen='N' AND a.IsSuccess='Y'" & vbCrLf
        sql &= "   GROUP BY m.OCID ) d ON a.OCID=d.OCID" & vbCrLf
        sql &= " WHERE b.TNum > ISNULL(d.TrainCount,0)" & vbCrLf '訓練數
        'sql &= " AND CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) - CAST(CONVERT(VARCHAR, a.STDate, 111) AS DATETIME) > 14" & vbCrLf
        'sql &= " AND CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) - CAST(CONVERT(VARCHAR, a.STDate, 111) AS DATETIME) < 22" & vbCrLf
        sql &= " and DATEDIFF(DAY, a.STDATE, GETDATE() ) > 14" & vbCrLf
        sql &= " and DATEDIFF(DAY, a.STDATE, GETDATE() ) < 22" & vbCrLf
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
            If TIMS.dtHaveDATA(dt) Then
                vSubject = $"目前有 <font color='Red'>{dt.Rows.Count}</font> 個班已開訓兩週"
                Call SUtl_AddoneMsg(oDt, vSubject)
            End If
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
        End Try

        '辦訓機構 訊息。
        sql = "" & vbCrLf
        sql &= " SELECT a.CLASSCNAME, a.CYCLTYPE, a.FTDATE" & vbCrLf
        sql &= " ,ISNULL(b.FinCount,0) AS FINCOUNT" & vbCrLf '結訓人員數
        sql &= " ,ISNULL(b.ResultCount,0) AS RESULTCOUNT" & vbCrLf '結訓學員資料卡
        sql &= " ,ISNULL(b.GradeCount,0) AS GRADECOUNT" & vbCrLf '結訓成績檔
        sql &= " ,ISNULL(b.InsCount,0) AS INSCOUNT" & vbCrLf '加退保檔
        sql &= " ,ISNULL(b.SubCount,0) AS SUBCOUNT" & vbCrLf '津貼結果檔
        sql &= " ,ISNULL(b.JobCount,0) AS JOBCOUNT" & vbCrLf '求職基本資料檔
        sql &= " ,ISNULL(CASE WHEN b6.OCID IS NOT NULL THEN 1 END ,0) QUESTCOUNT" & vbCrLf '期末學員滿意度調查檔
        sql &= " FROM Class_ClassInfo a" & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= "   SELECT b.OCID" & vbCrLf
        sql &= "   ,SUM(CASE WHEN b.StudStatus Not IN (2,3) THEN 1 END) FINCOUNT" & vbCrLf
        sql &= "   ,COUNT(b1.socid) RESULTCOUNT" & vbCrLf
        sql &= "   ,COUNT(b2.socid) GRADECOUNT" & vbCrLf
        sql &= "   ,COUNT(b3.socid) INSCOUNT" & vbCrLf
        sql &= "   ,COUNT(b4.socid) SUBCOUNT" & vbCrLf
        sql &= "   ,COUNT(b5.socid) JOBCOUNT" & vbCrLf
        sql &= "   FROM Class_StudentsOfClass b WITH(NOLOCK)" & vbCrLf
        sql &= "   JOIN Class_ClassInfo a WITH(NOLOCK) ON a.OCID=b.OCID" & vbCrLf '班級。
        sql &= "   LEFT JOIN Stud_ResultStudData b1 WITH(NOLOCK) ON b1.socid=b.socid" & vbCrLf
        sql &= "   LEFT JOIN view_TrainingResults b2 WITH(NOLOCK) ON b2.socid=b.socid" & vbCrLf
        sql &= "   LEFT JOIN Stud_Insurance b3 WITH(NOLOCK) ON b3.socid=b.socid" & vbCrLf
        sql &= "   LEFT JOIN Stud_SubsidyResult b4 WITH(NOLOCK) ON b4.socid=b.socid" & vbCrLf
        sql &= "   LEFT JOIN Jobseeker_BaseData b5 WITH(NOLOCK) ON b5.socid=b.socid" & vbCrLf
        sql &= "   WHERE a.NotOpen='N' AND a.IsSuccess='Y'" & vbCrLf
        sql &= "   AND a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= "   AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        '距離結訓前15天顯示此資訊
        'sql &= "   AND CAST(CONVERT(VARCHAR, a.FTDate, 111) AS DATETIME) - CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) >= 0" & vbCrLf
        'sql &= "   AND CAST(CONVERT(VARCHAR, a.FTDate + 14, 111) AS DATETIME) - CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) <= 0" & vbCrLf
        sql &= "   AND DATEDIFF(DAY,GETDATE(),a.FTDate)>=0" & vbCrLf '未結訓
        sql &= "   AND DATEDIFF(DAY,a.FTDate-14,GETDATE())>=0" & vbCrLf '距離結訓前15天
        sql &= "   GROUP BY b.OCID" & vbCrLf
        sql &= "  ) b ON b.OCID=a.OCID" & vbCrLf
        sql &= "  LEFT JOIN VIEW_QUESTIONARY2 b6 ON b6.OCID=a.OCID" & vbCrLf
        sql &= "  WHERE a.NotOpen='N' AND a.IsSuccess='Y'" & vbCrLf
        sql &= "  AND a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= "  AND a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        '距離結訓前15天顯示此資訊
        'sql &= "  AND CAST(CONVERT(VARCHAR, a.FTDate, 111) AS DATETIME) - CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) >= 0" & vbCrLf
        'sql &= "  AND CAST(CONVERT(VARCHAR, a.FTDate + 14, 111) AS DATETIME) - CAST(CONVERT(VARCHAR, GETDATE(), 111) AS DATETIME) <= 0" & vbCrLf
        sql &= "  AND DATEDIFF(DAY,GETDATE(),a.FTDate)>=0" & vbCrLf '未結訓
        sql &= "  AND DATEDIFF(DAY,a.FTDate-14,GETDATE())>=0" & vbCrLf '距離結訓前15天
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg) 'TIMS.WriteTraceLog(Me, ex) 
        End Try

        'TIMS.dtHaveDATA(dt)
        Dim flag_show_data As Boolean = If(TIMS.dtHaveDATA(dt), True, False)

        If flag_show_data Then
            Dim i_FinCount As Integer = 0
            Dim i_GradeCount As Integer = 0
            Dim i_QuestCount As Integer = 0
            Dim i_ResultCount As Integer = 0
            For Each dr As DataRow In dt.Rows
                '結訓人員數
                i_FinCount += dr("FINCOUNT")
                '結訓學員資料卡
                i_ResultCount += dr("RESULTCOUNT")
                '結訓成績
                i_GradeCount += dr("GRADECOUNT")
                '期末學員滿意度調查檔
                i_QuestCount += dr("QUESTCOUNT")
            Next
            i_ResultCount = i_FinCount - i_ResultCount
            i_GradeCount = i_FinCount - i_GradeCount
            i_QuestCount = i_FinCount - i_QuestCount

            vSubject = "有 <font color='Red'>" & dt.Rows.Count & "</font> 個班將在兩週內結訓"
            Call SUtl_AddoneMsg(oDt, vSubject)
            '產業人才投資方案(28):不顯示'結訓學員資料卡
            If i_ResultCount > 0 AndAlso Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                vSubject = "尚有 <font color='Red'>" & i_ResultCount & "</font> 人未填寫結訓學員資料卡"
                Call SUtl_AddoneMsg(oDt, vSubject)
            End If
            If i_GradeCount > 0 Then
                vSubject = "尚有 <font color='Red'>" & i_GradeCount & "</font> 人未填寫結訓成績"
                Call SUtl_AddoneMsg(oDt, vSubject)
            End If
            '產業人才投資方案(28):不顯示'期末學員滿意度調查
            If i_QuestCount > 0 AndAlso Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                vSubject = "尚有 <font color='Red'>" & i_QuestCount & "</font> 人未填寫期末學員滿意度調查"
                Call SUtl_AddoneMsg(oDt, vSubject)
            End If
        End If

        'Try
        '    If dt.Rows.Count=0 Then
        '        'msg5.Text="查無資料"
        '    End If
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    'Throw ex
        'End Try

        '辦訓機構 訊息。
        sql = ""
        sql &= " SELECT a.CLASSNAME, a.CYCLTYPE, b.VERDATE, a.STDATE, a.FDDATE "
        sql &= " FROM Plan_PlanInfo a WITH(NOLOCK) "
        sql &= " JOIN Plan_VerRecord b WITH(NOLOCK) ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo "
        sql &= " WHERE a.AppliedResult='Y' AND a.TransFlag='N' "
        sql &= " AND a.RID='" & sm.UserInfo.RID & "' "
        sql &= " AND a.PlanID='" & sm.UserInfo.PlanID & "' "

        Try
            dt = DbAccess.GetDataTable(sql, objconn)
            If TIMS.dtHaveDATA(dt) Then
                vSubject = "尚有 <font color='Red'>" & dt.Rows.Count & "</font> 個計畫尚未轉入班級" & vbCrLf
                Call SUtl_AddoneMsg(oDt, vSubject)
            End If
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
        End Try

        '署(局)、分署(中心) 或 補助地方政府 :【署(局)、分署(中心)、補助地方】
        If sm.UserInfo.LID <= 1 OrElse (sm.UserInfo.TPlanID = "17" AndAlso sm.UserInfo.LID <= 2) Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '產投類。 '與 TC_04_002.btnQuery_Click 查詢條件一致
                sql = ""
                sql &= " SELECT P1.CLASSNAME, P1.CYCLTYPE, P1.STDATE, P1.FDDATE, O1.ORGNAME" & vbCrLf
                'sql &= " ,P2.YEARS,a1.RELSHIP,P1.PLANID,p1.RID,pvr.SecResult,P1.AppliedResult,p1.RESULTBUTTON,p1.DataNotSent" & vbCrLf
                sql &= " FROM PLAN_PLANINFO P1 WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN ID_PLAN P2 WITH(NOLOCK) ON P1.PlanID=P2.PlanID" & vbCrLf
                sql &= " JOIN AUTH_RELSHIP A1 WITH(NOLOCK) ON P1.RID=A1.RID" & vbCrLf
                sql &= " JOIN ORG_ORGINFO O1 WITH(NOLOCK) ON A1.OrgID=O1.OrgID" & vbCrLf
                sql &= " JOIN ORG_ORGPLANINFO O2 WITH(NOLOCK) ON A1.RSID=O2.RSID" & vbCrLf
                sql &= " LEFT JOIN PLAN_VERREPORT pvr WITH(NOLOCK) ON P1.PlanID=pvr.PlanID AND P1.ComIDNO=pvr.ComIDNO AND P1.SeqNO=pvr.SeqNo" & vbCrLf
                sql &= " LEFT JOIN PLAN_VERRECORD pvrc1 WITH(NOLOCK) ON pvrc1.PlanID=pvr.PlanID AND pvrc1.ComIDNO=pvr.ComIDNO AND pvrc1.SeqNO=pvr.SeqNo AND (pvrc1.VerSeqNo=1)" & vbCrLf
                sql &= " WHERE A1.RelShip LIKE '" & sm.UserInfo.RelShip & "%'" & vbCrLf
                sql &= " AND P1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND P2.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                sql &= " AND P2.PlanKind=2" & vbCrLf
                'AND P1.AppliedResult IS NULL'(無審核結果) AND P1.ResultButton IS NULL'(已送出) AND P1.DataNotSent IS NULL '(檢送資料-已檢送資料)
                sql &= " AND P1.AppliedResult IS NULL AND P1.ResultButton IS NULL AND P1.DataNotSent IS NULL" & vbCrLf
                sql &= " AND P1.IsApprPaper='Y' AND pvr.IsApprPaper='Y'" & vbCrLf
                sql &= " AND pvr.SecResult IS NULL" & vbCrLf
            Else
                '一般TIMS
                sql = ""
                sql &= " SELECT P1.CLASSNAME, P1.CYCLTYPE, P1.STDATE, P1.FDDATE, O1.ORGNAME" & vbCrLf
                sql &= " FROM PLAN_PLANINFO P1 WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN AUTH_RELSHIP A1 WITH(NOLOCK) ON P1.RID=A1.RID" & vbCrLf
                sql &= " JOIN ORG_ORGINFO O1 WITH(NOLOCK) ON A1.OrgID=O1.OrgID" & vbCrLf
                sql &= " JOIN ORG_ORGPLANINFO O2 WITH(NOLOCK) ON A1.RSID=O2.RSID" & vbCrLf
                sql &= " WHERE (P1.AppliedResult IS NULL OR P1.AppliedResult='O')" & vbCrLf
                sql &= " AND P1.IsApprPaper='Y' AND P1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND A1.RelShip LIKE '" & sm.UserInfo.RelShip & "%'" & vbCrLf
            End If

            Try
                dt = DbAccess.GetDataTable(sql, objconn)
                If TIMS.dtHaveDATA(dt) Then
                    vSubject = "尚有 <font color='Red'>" & dt.Rows.Count & "</font> 個班級待審核"
                    Call SUtl_AddoneMsg(oDt, vSubject)
                End If
            Catch ex As Exception
                Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
                Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
            End Try
        End If

        '委訓單位加提醒e網報名待審核人數。
        '若有調整 SD_01_004之e網報名者待審核，煩請一併調整位置 主頁搜尋條件。
        '所有單位【含署(局)、分署(中心)】
        'Dim sql As String=""
        sql = ""
        sql &= " WITH WC1 AS (" & " SELECT DISTINCT c.OCID" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO c WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP A1 WITH(NOLOCK) ON A1.RID=c.RID" & vbCrLf
        sql &= " WHERE c.NotOpen='N' AND c.IsSuccess='Y'" & vbCrLf
        sql &= " AND c.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= " AND A1.RelShip LIKE '" & sm.UserInfo.RelShip & "%'" & " )" & vbCrLf
        sql &= " SELECT COUNT(1) NEEDCHECK" & vbCrLf
        sql &= " FROM WC1 c" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2 s WITH(NOLOCK) ON s.OCID1=c.OCID AND s.signUpStatus=0" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP2 e WITH(NOLOCK) ON e.eSETID=s.eSETID" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn)

        If dr1 IsNot Nothing AndAlso dr1(0) <> 0 Then
            vSubject = String.Concat("尚有 <font color='Red'>", dr1(0), "</font> 名e網報名者待審核")
            Call SUtl_AddoneMsg(oDt, vSubject)
        End If

        'Dim QCnt As Integer=Count_FAQ_Question()
        'If QCnt > 0 Then
        '    vSubject="<a href='./FAQ/01/FAQ_01_004.aspx' style='color:#1f336b;text-decoration:underline'>目前有 <font color='red'>" & QCnt & "</font> 筆問題單等待回覆</a>"
        '    Call sUtl_AddoneMsg(oDt, vSubject)
        'End If
        'Dim QCCnt As Integer=Count_ClosedQuestion()
        'If QCCnt > 0 Then
        '    vSubject="<a href='./FAQ/01/FAQ_01_005.aspx' style='color:#1f336b;text-decoration:underline'>尚有 <font color='red'>" & QCCnt & "</font> 筆問題單需填寫滿意度</a>" & vbCrLf
        '    Call sUtl_AddoneMsg(oDt, vSubject)
        'End If

    End Sub

    '#Region "確認登入帳號之機構是否在黑名單中"
    Public Shared Function Check_AccoutBlackList(ByRef MyPage As Page, ByRef sm As SessionModel, ByRef objconn As SqlConnection) As String
        Dim rst As String = ""
        '提示訊息，不跳離系統
        '確認登入帳號之機構是否在黑名單中 20090724 by AMU
        Call TIMS.Check_AccoutBlackList(MyPage, sm.UserInfo.UserID, rst, objconn)
        Return rst
    End Function

    '#Region "基本問題單的搜尋條件，含登入時的計畫別、問題單不為刪除單 回傳數量"
    Function Count_FAQ_Question() As Integer
        Dim rst As Integer = 0
        Dim sqlStr As String = ""
        sqlStr &= " SELECT COUNT(1) CNTQ"
        sqlStr &= " FROM FAQ_Question a "
        sqlStr &= " LEFT JOIN code_mood b ON a.cod_id=b.cod_id "
        sqlStr &= " JOIN Auth_Account c ON c.Account=a.QAccount "
        sqlStr &= " JOIN Org_OrgInfo e ON e.OrgID=a.OrgID "
        sqlStr &= " JOIN ID_FAQLevel f ON f.FAQID=a.FAQID "
        sqlStr &= " LEFT JOIN (SELECT QID,COUNT(1) ACOUNT FROM FAQ_Answer GROUP BY QID) d ON d.QID=a.QID "
        sqlStr &= " WHERE a.TPlanID=@TPlanID AND a.QStatus <> 'D' "
        sqlStr &= " AND (CASE WHEN d.QID IS NULL THEN 'N' ELSE 'Y' END)=@ansType "
        sqlStr &= " AND a.Closed=@Closed AND a.DistID=@DistID AND a.QAccount != @QAccount "
        sqlStr &= " AND ISNULL(f.ParentFAQID, '4')=@FAQID"
        Call TIMS.OpenDbConn(objconn)
        'myParams.Clear()
        '依搜尋條件增加回覆狀態、提問日區間、問題內容關鍵字
        Dim myParams As New Hashtable From {
            {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
            {"ansType", "N"},
            {"Closed", "N"},
            {"DistID", Convert.ToString(sm.UserInfo.DistID)},
            {"QAccount", Convert.ToString(sm.UserInfo.UserID).Trim.Replace("'", "''")},
            {"FAQID", TIMS.Get_FAQID(sm.UserInfo.LID, sm.UserInfo.UserID, objconn)}
        }
        Dim myDr As DataRow = DbAccess.GetOneRow(sqlStr, objconn, myParams)

        If (Convert.ToString(myDr(0))) Then rst = Convert.ToInt32(myDr(0)) Else rst = 0
        Return rst
    End Function

    '#Region "基本問題單的搜尋條件，含登入時的計畫別、問題單不為刪除單 回傳數量"
    Private Function Count_ClosedQuestion() As Integer
        Dim rst As Integer = 0
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT a.QID, a.COD_ID, a.QACCOUNT, a.OrgID" & vbCrLf
        sql &= " FROM FAQ_Question a" & vbCrLf
        sql &= " JOIN Auth_Account c ON c.Account=a.QAccount" & vbCrLf
        sql &= " JOIN Org_OrgInfo e ON e.OrgID=a.OrgID" & vbCrLf
        sql &= " LEFT JOIN CODE_MOOD b ON b.cod_id=a.cod_id" & vbCrLf
        sql &= " WHERE a.QStatus <> 'D'" & vbCrLf
        sql &= " AND a.Score IS NULL" & vbCrLf
        sql &= " AND a.Closed=@Closed" & vbCrLf
        sql &= " AND a.TPlanID=@TPlanID" & vbCrLf
        sql &= " AND a.QAccount=@QAccount )" & vbCrLf

        sql &= " SELECT count(1) CNTQ" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " LEFT JOIN (" & " SELECT dx.QID, COUNT(1) ACOUNT" & vbCrLf
        sql &= " FROM FAQ_Answer dx" & vbCrLf
        sql &= " JOIN WC1 ax ON ax.QID=dx.QID" & vbCrLf
        sql &= " GROUP BY dx.QID" & " ) d ON d.QID=a.QID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        'myParams.Clear()
        Dim myParams As New Hashtable From {
            {"Closed", "Y"},
            {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
            {"QAccount", Convert.ToString(sm.UserInfo.UserID)}
        }
        Dim myDr As DataRow = DbAccess.GetOneRow(sql, objconn, myParams)

        If (Convert.ToString(myDr(0))) Then rst = Convert.ToInt32(myDr(0)) Else rst = 0
        Return rst
    End Function

    '從DataTable取前x筆內容"
    Function GetDtTopX(ByVal sourceDT As DataTable, ByVal topCount As Integer) As DataTable
        Dim oldDTCount As Integer = sourceDT.Rows.Count
        Dim myNewDT As DataTable = New DataTable
        Dim myCol1 As New DataColumn("PostDate", GetType(String))
        myNewDT.Columns.Add(myCol1)
        Dim myCol2 As New DataColumn("Subject", GetType(String))
        myNewDT.Columns.Add(myCol2)

        For i As Integer = 1 To topCount
            If i <= oldDTCount Then
                Dim myRow As DataRow = myNewDT.NewRow()
                myRow("PostDate") = sourceDT.Rows(i - 1)("PostDate").ToString()
                myRow("Subject") = sourceDT.Rows(i - 1)("Subject").ToString()
                myNewDT.Rows.Add(myRow)
            End If
        Next

        Return myNewDT
    End Function

    '按下[more...]按鈕,導向明細頁"
    Sub ViewDetailData(ByVal i_Todo As Integer)

        Dim myNextPage As String = "main2_detail.aspx" & "?todo=" & CStr(i_Todo)
        TIMS.Utl_Redirect(Me, objconn, myNextPage)

    End Sub

    ''' <summary> 訓後動態調查表 </summary>
    ''' <param name="sm"></param>
    ''' <param name="objconn"></param>
    ''' <param name="oDt"></param>
    Public Shared Sub Utl_TRAIN_DYNAMIC_QUESTION(ByRef MyPage As Page, ByRef sm As SessionModel, ByRef objconn As SqlConnection, ByRef oDt As DataTable)
        Dim v_ComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        v_ComIDNO = TIMS.ClearSQM(v_ComIDNO)
        If v_ComIDNO = "" Then Exit Sub '統編資訊異常 以下不處理!

        'Dim sql As String="" sql="" & vbCrLf
        'sql &= " SELECT TOP 10  c.COMIDNO,c.CLASSCNAME" & vbCrLf
        Dim sql As String = ""
        sql &= " SELECT IP.YEARS" & vbCrLf
        sql &= " ,c.CLASSCNAME" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO c WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO d WITH(NOLOCK) ON c.PlanID=d.PlanID AND c.ComIDNO=d.ComIDNO AND c.SeqNO=d.SeqNO" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip WITH(NOLOCK) ON ip.planid=c.planid" & vbCrLf
        sql &= " WHERE c.IsSuccess='Y' AND c.NotOpen='N'" & vbCrLf
        sql &= " AND c.STDATE <= GETDATE()" & vbCrLf
        sql &= " AND dbo.TRUNC_DATETIME(GETDATE()) >= DATEADD(MONTH,3,c.FTDate)+1" & vbCrLf
        sql &= " AND dbo.TRUNC_DATETIME(GETDATE()) <= DATEADD(MONTH,4,c.FTDate)" & vbCrLf
        sql &= " AND ip.TPlanID='28'" & vbCrLf '產投:28
        sql &= " and ip.DISTID='" & sm.UserInfo.DistID & "'" & vbCrLf
        sql &= " and c.COMIDNO ='" & v_ComIDNO & "'" & vbCrLf
        Try
            Dim ss As String = "CLASSCNAME" 'sort
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
            If TIMS.dtHaveDATA(dt) Then
                For Each dr As DataRow In dt.Select(Nothing, ss)
                    Dim tClasscnames As String = Convert.ToString(dr("CLASSCNAME"))
                    Dim tYEARS As String = CStr(CInt(dr("YEARS")) - 1911)
                    If tClasscnames <> "" Then
                        Dim vSubject As String = String.Concat("尚有", tYEARS, "年度「", tClasscnames, "」要填寫訓後動態調查表。")
                        Call SUtl_AddoneMsg(oDt, vSubject)
                    End If
                Next
            End If
        Catch ex As Exception
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(MyPage)}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg) 'Call TIMS.CloseDbConn(objconn) 'Throw ex
        End Try
    End Sub

    ''' <summary>增加作業提醒2</summary>
    ''' <param name="odt"></param>
    ''' <param name="vSubject"></param>
    Public Shared Sub SUtl_AddoneMsg(ByRef odt As DataTable, ByVal vSubject As String)
        SUtl_AddoneMsg(odt, vSubject, "")
    End Sub

    ''' <summary>增加作業提醒3</summary>
    ''' <param name="odt"></param>
    ''' <param name="vSubject"></param>
    ''' <param name="vMsg1"></param>
    Public Shared Sub SUtl_AddoneMsg(ByRef odt As DataTable, ByVal vSubject As String, ByVal vMsg1 As String)
        'Optional ByVal vMsg1 As String=""
        'vSubject=TIMS.GetValue1(vSubject)
        'vMsg1=TIMS.GetValue1(vMsg1)
        If vSubject = "" Then Exit Sub
        Dim dr As DataRow = odt.NewRow
        dr("Subject") = vSubject
        dr("Msg1") = vMsg1 '"Y"
        odt.Rows.Add(dr)
    End Sub

    Protected Sub BtnMore1_Click(sender As Object, e As EventArgs) Handles btnMore1.Click
        ViewDetailData(5)
    End Sub

    Protected Sub BtnMore2_Click(sender As Object, e As EventArgs) Handles btnMore2.Click
        ViewDetailData(1)
    End Sub

    Protected Sub BtnMore3_Click(sender As Object, e As EventArgs) Handles btnMore3.Click
        ViewDetailData(2)
    End Sub

    Protected Sub BtnMore4_Click(sender As Object, e As EventArgs) Handles btnMore4.Click
        ViewDetailData(3)
    End Sub

    Protected Sub BtnMore5_Click(sender As Object, e As EventArgs) Handles btnMore5.Click
        ViewDetailData(4)
    End Sub
End Class