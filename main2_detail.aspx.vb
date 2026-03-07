Partial Class main2_detail
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim strSS As String = ""

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call sCreate1() '頁面初始化
        End If

    End Sub

    '頁面初始化
    Sub sCreate1()
        '#Region "頁面初始化"
        '#Region "讀取[作業提醒]區塊內容"

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

        '作業提醒(資料)
        Dim Dt1 As New DataTable
        Dt1.Columns.Add("Subject")
        Dt1.Columns.Add("Msg1")
        Dt1.Columns.Add("Status1")
        'Call Warning(Dt1)
        Call main2.Warning(Me, sm, objconn, Dt1)

        '作業提醒(顯示)
        Dim Dt2 As New DataTable
        Dt2.Columns.Add("Subject")
        Dt2.Columns.Add("msg1")
        Dt2.Columns.Add("Status1")

        If Dt1.Rows.Count > 0 Then
            For Each dr As DataRow In Dt1.Rows
                Dim blnDelete As Boolean = False
                If Convert.ToString(dr("Subject")) = "" Then blnDelete = True
                If Len(dr("Subject")) <= 1 Then blnDelete = True
                If Not blnDelete Then sUtl_AddoneMsg(Dt2, Convert.ToString(dr("Subject")))
            Next
        End If
        'Dt2.AcceptChanges()

        If Dt2 IsNot Nothing AndAlso Dt2.Rows.Count > 0 Then
            gv1.DataSource = Dt2 '作業提醒。
            gv1.DataBind()
        End If

        Dim tDv As New DataView
        Try
            tDt = TIMS.Get_SelHomeNewsS1(objconn, 0)
            tDt.TableName = "Result"
            tDv.Table = tDt
        Catch ex As Exception
            Exit Sub
        End Try

        If tDv.Count > 0 Then
            '抓第一筆訊息區資料 by nick 060306  
            Dim script As String

            Dim ff As String = "type in (1,2) AND msg='Y' AND msg2='Y'" '合理的顯示週數。('未到結束日期。)
            For Each drv As DataRow In tDt.Select(ff)
                Dim SUBJECT As String = ""
                SUBJECT = Convert.ToString(drv("SUBJECT"))
                SUBJECT = Replace(SUBJECT, "！", "！" & vbCrLf)
                SUBJECT = Replace(SUBJECT, "。", "。" & vbCrLf)

                '含提示訊息
                script = String.Concat("<script>", "window.alert('", Common.GetJsString(SUBJECT), "');", "</script>")
                Me.RegisterStartupScript("", script)
                Exit For
            Next
        End If

        '#Region "讀取[最新消息]、[功能增修說明]、[文件下載]區塊內容"

        'todo: 1:最新消息 2: 功能增修說明 3: 文件下載 5/6
        'todo: 1:最新消息 2: 功能增修說明 3: 文件下載 4:影音教學專區
        For i As Integer = 1 To 4
            Dim dt As DataTable = TIMS.Get_SelHomeNewsS1(i, objconn)
            Dim dv As New DataView
            dt.TableName = "Result"
            dv.Table = dt

            Select Case i.ToString
                Case "1"
                    dv.RowFilter = "Type=1"
                    If dv.Count > 0 Then
                        gv2.DataSource = dv
                        gv2.DataBind()
                    End If
                Case "2"
                    dv.RowFilter = "Type=2"
                    If dv.Count > 0 Then
                        gv3.DataSource = dv
                        gv3.DataBind()
                    End If
                Case "3"
                    dv.RowFilter = "Type=3"
                    If dv.Count > 0 Then
                        gv4.DataSource = dv
                        gv4.DataBind()
                    End If
                Case "4"
                    dv.RowFilter = "Type=4"
                    If dv.Count > 0 Then
                        gv5.DataSource = dv
                        gv5.DataBind()
                    End If
            End Select
        Next

        '===================
        Dim myTodo As String = Convert.ToString(Request("todo"))
        Select Case myTodo
            Case "1"
                divArea2.Visible = True
            Case "2"
                divArea3.Visible = True
            Case "3"
                divArea4.Visible = True
            Case "4"
                divArea5.Visible = True
            Case "5", "6"
                divArea1.Visible = True
        End Select


    End Sub

    Protected Sub btnMore1_Click(sender As Object, e As EventArgs) Handles btnMore1.Click
        backIndexPage()
    End Sub

    Protected Sub btnMore2_Click(sender As Object, e As EventArgs) Handles btnMore2.Click
        backIndexPage()
    End Sub

    Protected Sub btnMore3_Click(sender As Object, e As EventArgs) Handles btnMore3.Click
        backIndexPage()
    End Sub

    Protected Sub btnMore4_Click(sender As Object, e As EventArgs) Handles btnMore4.Click
        backIndexPage()
    End Sub

    Protected Sub btnMore5_Click(sender As Object, e As EventArgs) Handles btnMore5.Click
        backIndexPage()
    End Sub

    '#Region "確認登入帳號之機構是否在黑名單中"
    Function Check_AccoutBlackList() As String
        Dim rst As String = ""
        '提示訊息，不跳離系統
        '確認登入帳號之機構是否在黑名單中 20090724 by AMU
        Call TIMS.Check_AccoutBlackList(Me, sm.UserInfo.UserID, rst, objconn)
        Return rst
    End Function

    '#Region "基本問題單的搜尋條件，含登入時的計畫別、問題單不為刪除單 回傳數量"
    Function Count_FAQ_Question() As Integer
        Dim rst As Integer = 0
        Dim sqlStr As String
        sqlStr = ""
        sqlStr &= " SELECT COUNT(1) CNTQ "
        sqlStr &= " FROM FAQ_Question a "
        sqlStr &= " LEFT JOIN code_mood b ON a.cod_id = b.cod_id "
        sqlStr &= " JOIN Auth_Account c ON c.Account = a.QAccount "
        sqlStr &= " JOIN Org_OrgInfo e ON e.OrgID = a.OrgID "
        sqlStr &= " JOIN ID_FAQLevel f ON f.FAQID = a.FAQID "
        sqlStr &= " LEFT JOIN (SELECT QID,COUNT(1) ACOUNT FROM FAQ_Answer GROUP BY QID) d ON d.QID = a.QID "
        sqlStr &= " WHERE a.TPlanID = @TPlanID AND a.QStatus <> 'D' "
        sqlStr &= " AND (CASE WHEN d.QID IS NULL THEN 'N' ELSE 'Y' END) = @ansType "
        sqlStr &= " AND a.Closed = @Closed "
        sqlStr &= " AND a.DistID = @DistID "
        sqlStr &= " AND a.QAccount != @QAccount "
        sqlStr &= " AND ISNULL(f.ParentFAQID, '4') = @FAQID "
        Call TIMS.OpenDbConn(objconn)
        Dim myParams As Hashtable = New Hashtable
        myParams.Clear()
        myParams.Add("TPlanID", Convert.ToString(sm.UserInfo.TPlanID))
        '依搜尋條件增加回覆狀態、提問日區間、問題內容關鍵字
        myParams.Add("ansType", "N")
        myParams.Add("Closed", "N")
        myParams.Add("DistID", Convert.ToString(sm.UserInfo.DistID))
        myParams.Add("QAccount", Convert.ToString(sm.UserInfo.UserID).Trim.Replace("'", "''"))
        myParams.Add("FAQID", TIMS.Get_FAQID(sm.UserInfo.LID, sm.UserInfo.UserID, objconn))
        Dim myDr As DataRow = DbAccess.GetOneRow(sqlStr, objconn, myParams)
        If (Convert.ToString(myDr(0))) Then rst = Convert.ToInt32(myDr(0)) Else rst = 0
        Return rst
    End Function

    '#Region "基本問題單的搜尋條件，含登入時的計畫別、問題單不為刪除單 回傳數量"
    Private Function Count_ClosedQuestion() As Integer
        Dim rst As Integer = 0
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS ( " & vbCrLf
        sql &= " SELECT a.QID, a.COD_ID, a.QACCOUNT, a.OrgID " & vbCrLf
        sql &= " FROM FAQ_Question a " & vbCrLf
        sql &= " JOIN Auth_Account c ON c.Account = a.QAccount " & vbCrLf
        sql &= " JOIN Org_OrgInfo e ON e.OrgID = a.OrgID " & vbCrLf
        sql &= " LEFT JOIN CODE_MOOD b ON b.cod_id = a.cod_id " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.QStatus <> 'D' " & vbCrLf
        sql &= " AND a.Score IS NULL " & vbCrLf
        sql &= " AND a.Closed = @Closed " & vbCrLf
        sql &= " AND a.TPlanID = @TPlanID " & vbCrLf
        sql &= " AND a.QAccount = @QAccount " & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT count(1) CNTQ " & vbCrLf
        sql &= " FROM WC1 a " & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " SELECT dx.QID, COUNT(1) ACOUNT " & vbCrLf
        sql &= " FROM FAQ_Answer dx " & vbCrLf
        sql &= " JOIN WC1 ax ON ax.QID = dx.QID " & vbCrLf
        sql &= " GROUP BY dx.QID " & vbCrLf
        sql &= " ) d ON d.QID = a.QID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim myParams As Hashtable = New Hashtable
        myParams.Clear()
        myParams.Add("Closed", "Y")
        myParams.Add("TPlanID", Convert.ToString(sm.UserInfo.TPlanID))
        myParams.Add("QAccount", Convert.ToString(sm.UserInfo.UserID))
        Dim myDr As DataRow = DbAccess.GetOneRow(sql, objconn, myParams)
        If (Convert.ToString(myDr(0))) Then rst = Convert.ToInt32(myDr(0)) Else rst = 0
        Return rst
    End Function

    '#Region "按下[back]按鈕,導向首頁"
    Sub backIndexPage()

        Dim myNextPage As String = "main2"
        TIMS.Utl_Redirect(Me, objconn, myNextPage)

    End Sub

    ''' <summary>
    ''' 增加作業提醒2
    ''' </summary>
    ''' <param name="odt"></param>
    ''' <param name="vSubject"></param>
    Public Shared Sub sUtl_AddoneMsg(ByRef odt As DataTable, ByVal vSubject As String)
        sUtl_AddoneMsg(odt, vSubject, "")
    End Sub

    ''' <summary>
    ''' 增加作業提醒3
    ''' </summary>
    ''' <param name="odt"></param>
    ''' <param name="vSubject"></param>
    ''' <param name="vMsg1"></param>
    Public Shared Sub sUtl_AddoneMsg(ByRef odt As DataTable, ByVal vSubject As String, ByVal vMsg1 As String)
        'Optional ByVal vMsg1 As String = ""
        'vSubject = TIMS.GetValue1(vSubject)
        'vMsg1 = TIMS.GetValue1(vMsg1)
        If vSubject = "" Then Exit Sub
        Dim dr As DataRow = odt.NewRow
        dr("Subject") = vSubject
        dr("Msg1") = vMsg1 '"Y"
        odt.Rows.Add(dr)
    End Sub

End Class