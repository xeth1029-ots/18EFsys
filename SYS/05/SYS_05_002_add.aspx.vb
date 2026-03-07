Partial Class SYS_05_002_add
    Inherits AuthBasePage

    Const cst_gptodo_update As String = "update"
    Const cst_gptodo_add As String = "add"

    Const cst_procType_Insert As String = "Insert"
    Const cst_procType_Update As String = "Update"

    'Dim all_TypeList As String = "1,2,3"
    Const cst_TypeList_1_News As String = "1" 'News
    Const cst_TypeList_2_新功能 As String = "2" '新功能
    Const cst_TypeList_3_文件下載 As String = "3" '文件下載
    Const cst_TypeList_4_影音教學 As String = "4" '影音教學

    Const cst_doc_rep_1 As String = "<SPAN class=newlink><A class=newlink href=""Doc/{0}"">{1}。</A></SPAN>"
    Const cst_doc_rep_4 As String = "<SPAN class=newlink><A class=newlink href=""media/{0}"">{1}。</A></SPAN>"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        hid_gptodo.Value = TIMS.ClearSQM(Request("gptodo"))

        If Not Page.IsPostBack Then
            TypeList.Attributes("onclick") = "change_TypeList();"
            'TypeList.Attributes("onchange") = "change_TypeList();"

            '檢查日期格式
            Me.PostDate.Attributes("onchange") = "check_date();"
            '檢查日期格式
            Me.PostFDate.Attributes("onchange") = "check_dateF();"

            Me.isShow.Attributes("onclick") = "set_lab();"

            bt_addrow.Attributes("onclick") = "return js1_close();"
            Button1.Attributes("onclick") = "return js1_close();"

            '保留查詢字串
            If Session("_search") IsNot Nothing Then
                'Me.ViewState("_search") = Session("_search")
                hid_AcceptSearch.Value = Convert.ToString(Session("_search")) 'ViewState("_search")
                'Session("_search") = Nothing
            End If

            Call Create1()

            'Dim strScript1 As String = ""
            'strScript1 = "<script>change_TypeList();</script>" & vbCrLf
            'TIMS.RegisterClientScriptBlock(Me, TIMS.xBlockName, strScript1)
        End If
        ' 無效，若是用.Net 驗証功能
        ' Button1.Attributes("onclick") = "return save_check();"
    End Sub

    Sub Create1()
        '新增時日期預設為系統日期
        PostDate.Text = Common.FormatDate(Now())

        TIMS.SUB_SET_HR_MI(HR2, MM2)
        Common.SetListItem(HR2, "18")
        Common.SetListItem(MM2, "00")

        '修改時帶出資料
        Select Case hid_gptodo.Value
            Case cst_gptodo_update '"update"
                Show_EditData()
        End Select

    End Sub

    Sub Show_EditData()
        hid_HNID.Value = TIMS.ClearSQM(Request("HNID"))
        If hid_HNID.Value = "" Then Exit Sub '(異常)

        'Dim iHNID As Integer = Val(Request("HNID"))
        Dim parms As Hashtable = New Hashtable()
        Dim sqlstr As String = ""
        sqlstr = "SELECT * FROM HOME_NEWS WHERE HNID=@HNID "
        parms.Clear()
        parms.Add("HNID", hid_HNID.Value)
        Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objconn, parms)
        If dr Is Nothing Then Exit Sub '(異常)

        Common.SetListItem(TypeList, Convert.ToString(dr("Type")))
        Common.SetListItem(isShow, dr("isShow").ToString)
        Me.PostDate.Text = Common.FormatDate(dr("PostDate"))
        If Convert.ToString(dr("PostFDate")) <> "" Then
            Me.PostFDate.Text = Common.FormatDate(dr("PostFDate"))
            TIMS.SET_DateHM(CDate(dr("PostFDate")), HR2, MM2)
        End If

        Subject.Text = TIMS.ClearSQM(dr("Subject"))
        Dim s_Subject As String = Subject.Text
        If (s_Subject <> "" AndAlso s_Subject.Contains(">") AndAlso s_Subject.Contains("<")) Then
            Hid_context_decode1.Value = ""
            Subject.Text = HttpUtility.HtmlEncode(s_Subject)
        End If
        Common.SetListItem(msgweek, Convert.ToString(dr("msgweek")))
        txtDoc0.Text = TIMS.ClearSQM(dr("Doc0")) 'dr("Doc0")
        txtDoc1.Text = TIMS.ClearSQM(dr("Doc1")) 'dr("Doc0")

        Select Case Convert.ToString(dr("Type"))
            Case cst_TypeList_3_文件下載
                '新式狀況
                If txtDoc0.Text <> "" AndAlso txtDoc1.Text <> "" Then Subject.Text = ""
            Case cst_TypeList_4_影音教學
                '新式狀況
                If txtDoc0.Text <> "" AndAlso txtDoc1.Text <> "" Then Subject.Text = ""
        End Select

        If Convert.ToString(dr("PostFDate")) = "" Then
            Page.RegisterStartupScript("load", "<script>set_lab();</script>")
        End If

        trSubject1.Style("display") = ""
        trHtmlx.Style("display") = ""
        trDoc1.Style("display") = "none"
        trDoc2.Style("display") = "none"
        Select Case Convert.ToString(dr("Type"))
            Case "3", "4"
                trSubject1.Style("display") = ""
                trHtmlx.Style("display") = ""
                trDoc1.Style("display") = ""
                trDoc2.Style("display") = ""
                If Subject.Text = "" Then
                    trSubject1.Style("display") = "none"
                    trHtmlx.Style("display") = "none"
                    trDoc1.Style("display") = ""
                    trDoc2.Style("display") = ""
                End If
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session("_search") Is Nothing Then
            Session("_search") = hid_AcceptSearch.Value
        End If

        '回上一頁
        Dim url1 As String = "SYS_05_002.aspx?ID=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect1(Me, url1)
    End Sub

    '儲存 (檢核後儲存)   
    Function sUtl_Save1() As Boolean
        Dim rst As Boolean = False '儲存異常
        Dim errMsg As String = ""
        Dim v_TypeList As String = TIMS.GetListValue(TypeList)

        Dim v_Subject As String = TIMS.ClearSQM(Subject.Text)
        Dim vDoc0 As String = TIMS.ClearSQM(txtDoc0.Text)
        Dim vDoc1 As String = TIMS.ClearSQM(txtDoc1.Text)
        Dim v_isShow As String = TIMS.GetListValue(isShow)
        Dim v_msgweek As String = TIMS.GetListValue(msgweek)

        Dim str_Doc_R As String = ""
        Select Case v_TypeList 'TypeList.SelectedValue
            Case cst_TypeList_1_News
            Case cst_TypeList_2_新功能
            Case cst_TypeList_3_文件下載, cst_TypeList_4_影音教學
                If vDoc0 = "" Then errMsg &= "請輸入 文件下載-檔名!" & vbCrLf
                If vDoc1 = "" Then errMsg &= "請輸入 文件下載-說明!" & vbCrLf

            Case Else
                If Subject.Text = "" Then
                    '請發布主題
                    errMsg &= "請發布主題!!" & vbCrLf
                End If
                'Case "1", "2", "3"
                'Common.MessageBox(Me, "請選擇項目代碼!!")
                errMsg &= "項目代碼有誤或未選擇!!" & vbCrLf
                'Return rst
        End Select
        Me.PostDate.Text = TIMS.ClearSQM(Me.PostDate.Text)
        Me.PostFDate.Text = TIMS.ClearSQM(Me.PostFDate.Text)
        If Me.PostDate.Text = "" Then
            errMsg &= "請輸入或選擇發布日期起始"
        Else
            If Not IsDate(Me.PostDate.Text) Then
                errMsg &= "發布日期，請輸入正確的日期格式,YYYY/MM/DD!!" & vbCrLf
            End If
        End If
        Select Case v_isShow 'Me.isShow.SelectedValue
            Case "N"
                If Not IsDate(Me.PostFDate.Text) Then
                    errMsg &= "發布日期迄止，請輸入正確的日期格式,YYYY/MM/DD!!" & vbCrLf
                    'Return rst 'Exit Function
                End If
        End Select
        If errMsg <> "" Then
            Dim s_Subject As String = Subject.Text
            If (s_Subject <> "" AndAlso s_Subject.Contains(">") AndAlso s_Subject.Contains("<")) Then
                Hid_context_decode1.Value = ""
                Subject.Text = HttpUtility.HtmlEncode(s_Subject)
            End If
            Common.MessageBox(Me, errMsg)
            Return rst
        End If

        'Dim Search() As String = Split(Me.ViewState("_search"), "&")
        hid_gptodo.Value = TIMS.ClearSQM(Request("gptodo"))

        Dim sqlDA As New SqlDataAdapter
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sqlstr As String = ""
        Dim iHNID As Integer = 0
        If hid_HNID.Value <> "" Then iHNID = Val(hid_HNID.Value)

        '新增、修改實作
        Select Case hid_gptodo.Value
            Case cst_gptodo_add '"add"
                iHNID = DbAccess.GetNewId(objconn, "HOME_NEWS_HNID_SEQ,HOME_NEWS,HNID")
                sqlstr = "SELECT * FROM HOME_NEWS WHERE 1<>1"
                dt = DbAccess.GetDataTable(sqlstr, sqlDA, objconn)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("HNID") = iHNID

            Case Else '"update" 
                sqlstr = "SELECT * FROM HOME_NEWS WHERE HNID=" & hid_HNID.Value
                'Dim parms As Hashtable = New Hashtable()
                dt = DbAccess.GetDataTable(sqlstr, sqlDA, objconn)
                dr = dt.Rows(0)
        End Select
        dr("Type") = v_TypeList 'Me.TypeList.SelectedValue

        Select Case v_TypeList
            Case cst_TypeList_3_文件下載
                v_Subject = String.Format(cst_doc_rep_1, vDoc0, vDoc1)
            Case cst_TypeList_4_影音教學
                v_Subject = String.Format(cst_doc_rep_4, vDoc0, vDoc1)
            Case Else
                v_Subject = TIMS.ClearSQM(Subject.Text)
        End Select
        dr("Subject") = v_Subject ' Me.Subject.Text
        dr("isShow") = v_isShow 'Me.isShow.SelectedValue
        Select Case v_isShow'Me.isShow.SelectedValue
            Case "Y"
                dr("PostDate") = CDate(Me.PostDate.Text)
                dr("PostFDate") = Convert.DBNull
            Case Else
                'https://msdn.microsoft.com/zh-tw/library/system.datetime.aspx
                dr("PostDate") = CDate(Me.PostDate.Text)
                Dim vsPostFDate As String = TIMS.GET_DateHM(PostFDate, HR2, MM2)
                dr("PostFDate") = CDate(vsPostFDate) ' CDate(Me.PostFDate.Text)
        End Select
        dr("msgweek") = v_msgweek 'Me.msgweek.SelectedValue
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
        dr("Doc0") = If(vDoc0 <> "", vDoc0, Convert.DBNull)
        dr("Doc1") = If(vDoc1 <> "", vDoc1, Convert.DBNull)

        '2018 add sql 交易記錄
        Call SaveHomeNewsTransLog(dr, If(hid_gptodo.Value = cst_gptodo_add, cst_procType_Insert, cst_procType_Update))

        DbAccess.UpdateDataTable(dt, sqlDA)

        rst = True '儲存正常
        Return rst
    End Function

    '儲存    
    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click
        If sUtl_Save1() Then
            If Session("_search") Is Nothing Then
                Session("_search") = hid_AcceptSearch.Value
            End If
            Dim url1 As String = "SYS_05_002.aspx?ID=" & TIMS.Get_MRqID(Me)
            TIMS.Utl_Redirect1(Me, url1)
        End If
    End Sub

    Private Sub SaveHomeNewsTransLog(ByVal dr As DataRow, ByVal s_procType As String)
        Dim myParam As Hashtable = New Hashtable()
        Select Case s_procType
            Case cst_procType_Insert '"Insert"
                Dim beforeValues As String = ""
                beforeValues = ""
                beforeValues += "HNID=" + Convert.ToString(dr("HNID"))
                beforeValues += ",TYPE=" + Convert.ToString(dr("Type"))
                beforeValues += ",SUBJECT=" + Convert.ToString(dr("Subject"))
                beforeValues += ",ISSHOW=" + Convert.ToString(dr("isShow"))
                beforeValues += ",POSTDATE=" & If(IsDBNull(dr("PostDate")), "", Convert.ToDateTime(dr("PostDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                beforeValues += ",POSTFDATE=" & If(IsDBNull(dr("PostFDate")), "", Convert.ToDateTime(dr("PostFDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                beforeValues += ",MSGWEEK=" + Convert.ToString(dr("msgweek"))
                beforeValues += ",MODIFYACCT=" + Convert.ToString(dr("ModifyAcct"))
                beforeValues += ",MODIFYDATE=" + Convert.ToDateTime(dr("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")

                Dim t_iSql As String = ""
                t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql &= " VALUES(@SessionID, @TransTime, '/SYS/05/SYS_05_002_add.aspx', @UserID, 'Insert', 'HOME_NEWS', '', @BeforeValues, '') "
                myParam.Clear()
                myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("BeforeValues", beforeValues)
                'DbAccess.ExecuteNonQuery(iSql, objconn, parms)
                Dim i_tCmd As New SqlCommand(t_iSql, objconn)
                DbAccess.HashParmsChange(i_tCmd, myParam)
                Dim i_rst As Integer = i_tCmd.ExecuteNonQuery()

            Case cst_procType_Update '"Update"
                Dim beforeDt As DataTable = Nothing
                Dim sSql As String = ""
                sSql = "SELECT * FROM HOME_NEWS WHERE HNID=@HNID "
                myParam.Clear()
                myParam.Add("HNID", Convert.ToInt64(dr("HNID")))
                beforeDt = DbAccess.GetDataTable(sSql, objconn, myParam)

                Dim beforeValues As String = ""
                Dim afterValues As String = ""
                Dim conditions As String = ""

                If beforeDt IsNot Nothing Then
                    Dim beforeDr As DataRow = Nothing
                    beforeDr = beforeDt.Rows(0)
                    beforeValues += "TYPE=" + Convert.ToString(beforeDr("Type"))
                    beforeValues += ",SUBJECT=" + Convert.ToString(beforeDr("Subject"))
                    beforeValues += ",ISSHOW=" + Convert.ToString(beforeDr("isShow"))
                    beforeValues += ",POSTDATE=" & If(IsDBNull(dr("PostDate")), "", Convert.ToDateTime(dr("PostDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    beforeValues += ",POSTFDATE=" & If(IsDBNull(dr("PostFDate")), "", Convert.ToDateTime(dr("PostFDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    beforeValues += ",MSGWEEK=" + Convert.ToString(beforeDr("msgweek"))
                    beforeValues += ",MODIFYACCT=" + Convert.ToString(beforeDr("ModifyAcct"))
                    beforeValues += ",MODIFYDATE=" + Convert.ToDateTime(beforeDr("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")
                End If

                afterValues = ""
                afterValues += "TYPE=" + Convert.ToString(dr("Type"))
                afterValues += ",SUBJECT=" + Convert.ToString(dr("Subject"))
                afterValues += ",ISSHOW=" + Convert.ToString(dr("isShow"))
                afterValues += ",POSTDATE=" & If(IsDBNull(dr("PostDate")), "", Convert.ToDateTime(dr("PostDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                afterValues += ",POSTFDATE=" & If(IsDBNull(dr("PostFDate")), "", Convert.ToDateTime(dr("PostFDate")).ToString("yyyy-MM-dd HH:mm:ss.fff"))
                afterValues += ",MSGWEEK=" + Convert.ToString(dr("msgweek"))
                afterValues += ",MODIFYACCT=" + Convert.ToString(dr("ModifyAcct"))
                afterValues += ",MODIFYDATE=" + Convert.ToDateTime(dr("ModifyDate")).ToString("yyyy-MM-dd HH:mm:ss.fff")

                conditions = "HNID=" + Convert.ToString(dr("HNID"))
                '==========
                Dim t_iSql As String = ""
                t_iSql = ""
                t_iSql += " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql += " VALUES(@SessionID, @TransTime, '/SYS/05/SYS_05_002_add.aspx', @UserID, 'Update', 'HOME_NEWS', @CONDITIONS, @BeforeValues, @AfterValues) "
                myParam.Clear()
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("CONDITIONS", conditions)
                myParam.Add("BeforeValues", beforeValues)
                myParam.Add("AfterValues", afterValues)
                'DbAccess.ExecuteNonQuery(iSql, objconn, parms)
                Dim i_tCmd As New SqlCommand(t_iSql, objconn)
                DbAccess.HashParmsChange(i_tCmd, myParam)
                Dim i_rst As Integer = i_tCmd.ExecuteNonQuery()
        End Select

    End Sub

    Protected Sub BtnEncode1_Click(sender As Object, e As EventArgs) Handles BtnEncode1.Click
        Subject.Text = HttpUtility.HtmlEncode(Subject.Text)
    End Sub

    Protected Sub BtnDecode1_Click(sender As Object, e As EventArgs) Handles BtnDecode1.Click
        Subject.Text = HttpUtility.HtmlDecode(Subject.Text)
    End Sub
End Class
