Partial Class SD_07_003
    Inherits AuthBasePage

    Dim gAction As String = ""
    Dim gID As String = ""

    Const cst_SD_07_003_search As String = "SD_07_003_search"
    Const cst_search As String = "search" 'gAction
    Const cst_search2 As String = "search2" 'gAction
    Const cst_search3 As String = "search3" 'gAction

    Const cst_titleMsg1 As String = "點選可以觀看詳細資料"
    Const cst_分母為零提示字 As String = "0"
    Const cst_Format1 As String = "##0.0#%"

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        If Not IsPostBack Then
            TPlanID = TIMS.Get_TPlan(TPlanID)
            ddlYears = TIMS.Get_Years(ddlYears)
            ddlYears.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, "")) '2009/06/03加上請選擇
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            'Common.SetListItem(ddlYears, sm.UserInfo.Years)
            txtYearsOldS.Text = "15"
            txtYearsOldE.Text = "99"
        End If

        If Not Request("Action") Is Nothing Then gAction = Request("Action").ToString Else gAction = ""
        If Not Request("ID") Is Nothing Then gID = Request("ID").ToString Else gID = ""
        gAction = TIMS.ClearSQM(gAction)

        If Not IsPostBack Then
            'rblType1.Attributes("onclick") = "ShowType1();"
            rdoType1.Attributes("onclick") = "return ShowType1();"
            rdoType2.Attributes("onclick") = "return ShowType1();"

            btnSend.Attributes("onclick") = "return check();"

            Div1.Visible = False
            Div2.Visible = False
            Div3.Visible = False
            Div4.Visible = False
            Select Case gAction
                Case ""
                    Dim JScript As String = "<script>ShowType1();</script>" & vbCrLf
                    Page.RegisterStartupScript(TIMS.xBlockName, JScript)

                    Div1.Visible = True
                Case cst_search
                    Div2.Visible = True
                    Call GetSearch()
                    Call Search1()
                Case cst_search2
                    Div3.Visible = True
                    Call GetSearch()
                    Call Search2()
                Case cst_search3
                    Div4.Visible = True
                    Call GetSearch()
                    Call Search3()
            End Select
            'Session(Cst_SD_07_003_search) = Nothing
        End If

    End Sub

    Function Sch_WG1_x() As String
        Dim sTPlanID As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "TPlanID")
        Dim stxtExamDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateS")
        Dim stxtExamDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateE")
        Dim ssType1 As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "sType1")
        Dim stxtSTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateS")
        Dim stxtSTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateE")
        Dim stxtFTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateS")
        Dim stxtFTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateE")
        Dim sddlYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "ddlYears")
        Dim stxtYearsOldS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldS")
        Dim stxtYearsOldE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldE")
        Dim sPlanYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "PlanYears")
        Dim sComIDNO As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "ComIDNO")

        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " select a.Years PlanYears" & vbCrLf
        sqlstr &= " ,a.Years+a.PlanName YearPlan" & vbCrLf
        sqlstr &= " ,a.COMIDNO" & vbCrLf
        sqlstr &= " ,a.ORGNAME" & vbCrLf
        sqlstr &= " ,a.OCID" & vbCrLf
        'sqlstr &= " ,a.ClassCName+'(第'+a.CyclType+'期)' CLASSCNAME" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sqlstr &= " ,a.Examkind" & vbCrLf
        sqlstr &= " ,a.ExamName" & vbCrLf
        sqlstr &= " ,a.ExamLvName" & vbCrLf
        sqlstr &= " ,a.cnt1" & vbCrLf
        sqlstr &= " ,a.cnt2" & vbCrLf
        sqlstr &= " ,a.cnt3" & vbCrLf
        sqlstr &= " ,a.cnt4" & vbCrLf
        sqlstr &= " FROM V_STUDTECHEXAM2 a" & vbCrLf
        sqlstr &= " where 1=1" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sqlstr &= " and a.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        If sTPlanID <> "" Then
            sqlstr &= " and a.TPlanID='" & sTPlanID & "'" & vbCrLf
        End If
        If stxtExamDateS <> "" Then
            sqlstr &= " and a.ExamDate>=" & TIMS.To_date(stxtExamDateS) & vbCrLf
        End If
        If stxtExamDateE <> "" Then
            sqlstr &= " and a.ExamDate<=" & TIMS.To_date(stxtExamDateE) & vbCrLf
        End If
        If sPlanYears <> "" Then sqlstr &= " and a.YEARS='" & sPlanYears & "'" & vbCrLf
        If sComIDNO <> "" Then sqlstr &= " and a.COMIDNO='" & sComIDNO & "'" & vbCrLf
        Select Case ssType1
            Case "1"
                If stxtSTDateS <> "" Then
                    sqlstr &= " and a.STDate >=" & TIMS.To_date(stxtSTDateS) & vbCrLf
                End If
                If stxtSTDateE <> "" Then
                    sqlstr &= " and a.STDate <=" & TIMS.To_date(stxtSTDateE) & vbCrLf
                End If
                If stxtFTDateS <> "" Then
                    sqlstr &= " and a.FTDate >=" & TIMS.To_date(stxtFTDateS) & vbCrLf
                End If
                If stxtFTDateE <> "" Then
                    sqlstr &= " and a.FTDate <=" & TIMS.To_date(stxtFTDateE) & vbCrLf
                End If
                If sddlYears <> "" Then
                    sqlstr &= " and a.Years='" & sddlYears & "'" & vbCrLf
                End If
            Case "2"
                If sddlYears <> "" Then
                    '考試年度改成檢定日的年
                    sqlstr &= " and DATEPART(YEAR, a.ExamDate)='" & sddlYears & "'" & vbCrLf
                End If
            Case Else
                'sqlstr &= " and a.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        End Select
        If stxtYearsOldS <> "" Then
            sqlstr &= " and a.YearsOld >= " & Val(stxtYearsOldS) & vbCrLf
        End If
        If stxtYearsOldE <> "" Then
            sqlstr &= " and a.YearsOld <= " & Val(stxtYearsOldE) & vbCrLf
        End If
        Return sqlstr
    End Function

    Sub Search1()
        Dim sqlxWG1 As String = Sch_WG1_x()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WG1 AS (" & sqlxWG1 & ")" & vbCrLf
        sql &= " ,WG2 AS (" & vbCrLf
        sql &= " SELECT PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,IsNull(count(cnt1),0) StudCount" & vbCrLf
        sql &= " ,IsNull(count(cnt2),0) ExamCount" & vbCrLf
        sql &= " ,IsNull(count(cnt3),0) okPassCount" & vbCrLf
        sql &= " ,IsNull(count(cnt4),0) noPassCount" & vbCrLf
        sql &= " FROM WG1 g" & vbCrLf
        sql &= " GROUP BY PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,STUDCOUNT" & vbCrLf
        sql &= " ,EXAMCOUNT" & vbCrLf
        sql &= " ,OKPASSCOUNT" & vbCrLf
        sql &= " ,NOPASSCOUNT" & vbCrLf
        'sql &= " ,case when EXAMCOUNT > 0 then CONVERT(VARCHAR,round((OKPASSCOUNT/EXAMCOUNT*100),2))+'%'" & vbCrLf
        'sql &= " else '0' end passrate" & vbCrLf
        sql &= " ,CASE WHEN EXAMCOUNT > 0 THEN CONVERT(varchar,ROUND((convert(float,OKPASSCOUNT)/EXAMCOUNT*100),2))+'%'" & vbCrLf
        sql &= " ELSE '0' END PASSRATE" & vbCrLf
        sql &= " FROM WG2" & vbCrLf
        sql &= " ORDER BY PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Label1.Text = "<FONT color=Red>查無資料</FONT>"
            'Me.DataGridTable1.Visible = False
            DataGrid1.Visible = False
            Me.DataGrid1.DataSource = Nothing
            Me.DataGrid1.DataBind()
            Exit Sub
        End If

        Label1.Text = TIMS.GET_DISTNAME(objconn, sm.UserInfo.DistID)
        'Me.DataGridTable1.Visible = True
        DataGrid1.Visible = True
        Me.DataGrid1.DataSource = dt
        Me.DataGrid1.DataBind()

    End Sub

    Sub Search2()
        Dim sTPlanID As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "TPlanID")
        Dim stxtExamDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateS")
        Dim stxtExamDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateE")
        Dim srblType1 As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "rblType1")
        Dim stxtSTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateS")
        Dim stxtSTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateE")
        Dim stxtFTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateS")
        Dim stxtFTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateE")
        Dim sddlYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "ddlYears")
        Dim stxtYearsOldS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldS")
        Dim stxtYearsOldE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldE")
        Dim sPlanYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "PlanYears")

        Dim sqlxWG1 As String = Sch_WG1_x()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WG1 AS (" & sqlxWG1 & ")" & vbCrLf
        sql &= " ,WG2 AS (" & vbCrLf
        sql &= " SELECT PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,COMIDNO" & vbCrLf
        sql &= " ,ORGNAME" & vbCrLf
        sql &= " ,IsNull(count(cnt1),0) StudCount" & vbCrLf
        sql &= " ,IsNull(count(cnt2),0) ExamCount" & vbCrLf
        sql &= " ,IsNull(count(cnt3),0) okPassCount" & vbCrLf
        sql &= " ,IsNull(count(cnt4),0) noPassCount" & vbCrLf
        sql &= " FROM WG1 g" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        'If sPlanYears <> "" Then sql &= " AND PLANYEARS='" & sPlanYears & "'" & vbCrLf
        sql &= " GROUP BY PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,COMIDNO" & vbCrLf
        sql &= " ,ORGNAME" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,COMIDNO" & vbCrLf
        sql &= " ,ORGNAME" & vbCrLf
        sql &= " ,STUDCOUNT" & vbCrLf
        sql &= " ,EXAMCOUNT" & vbCrLf
        sql &= " ,OKPASSCOUNT" & vbCrLf
        sql &= " ,NOPASSCOUNT" & vbCrLf
        'sql &= " ,case when EXAMCOUNT > 0 then round((OKPASSCOUNT/EXAMCOUNT*100),2)+'%'" & vbCrLf
        'sql &= " else '0' end passrate" & vbCrLf
        sql &= " ,CASE WHEN EXAMCOUNT > 0 THEN CONVERT(varchar,ROUND((convert(float,OKPASSCOUNT)/EXAMCOUNT*100),2))+'%'" & vbCrLf
        sql &= " ELSE '0' END PASSRATE" & vbCrLf
        sql &= " FROM WG2" & vbCrLf
        sql &= " ORDER BY PLANYEARS" & vbCrLf
        sql &= " ,YEARPLAN" & vbCrLf
        sql &= " ,COMIDNO" & vbCrLf
        sql &= " ,ORGNAME" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            Label2A.Text = TIMS.GET_DISTNAME(objconn, sm.UserInfo.DistID)
            Label2B.Text = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "YearPlan")
            'Me.DataGridTable1.Visible = True
            DataGrid2.Visible = True
            Me.DataGrid2.DataSource = dt
            Me.DataGrid2.DataBind()
        Else
            Label2A.Text = "<FONT color=Red>查無資料</FONT>"
            Label2B.Text = "<FONT color=Red>查無資料</FONT>"
            'Me.DataGridTable1.Visible = False
            DataGrid2.Visible = False
            Me.DataGrid2.DataSource = Nothing
            Me.DataGrid2.DataBind()
        End If

    End Sub

    '統計項目若是以班別
    Sub Search3()
        Dim sTPlanID As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "TPlanID")
        Dim stxtExamDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateS")
        Dim stxtExamDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtExamDateE")
        Dim srblType1 As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "rblType1")
        Dim stxtSTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateS")
        Dim stxtSTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtSTDateE")
        Dim stxtFTDateS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateS")
        Dim stxtFTDateE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtFTDateE")
        Dim sddlYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "ddlYears")
        Dim stxtYearsOldS As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldS")
        Dim stxtYearsOldE As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "txtYearsOldE")
        Dim sPlanYears As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "PlanYears")
        Dim sComIDNO As String = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "ComIDNO")

        'Dim sql As String
        'Dim sqlstr As String = ""
        'Dim dt As DataTable


        Dim sql As String = ""
        Select Case rbCounttype.SelectedValue
            Case "1" '統計項目若是以班別
            Case "2" '依檢定類別
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        Select Case rbCounttype.SelectedValue
            Case "1" '統計項目若是以班別
                'Dim sqlx1 As String = Search1x()
                Dim sqlxWG1 As String = Sch_WG1_x()

                sql = "" & vbCrLf
                sql &= " WITH WG1 AS (" & sqlxWG1 & ")" & vbCrLf
                sql &= " ,WG2 AS (" & vbCrLf
                sql &= " SELECT PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,OCID,CLASSCNAME" & vbCrLf
                sql &= " ,EXAMLVNAME" & vbCrLf
                sql &= " ,IsNull(count(cnt1),0) StudCount" & vbCrLf
                sql &= " ,IsNull(count(cnt2),0) ExamCount" & vbCrLf
                sql &= " ,IsNull(count(cnt3),0) okPassCount" & vbCrLf
                sql &= " ,IsNull(count(cnt4),0) noPassCount" & vbCrLf
                sql &= " FROM WG1 g" & vbCrLf
                'sql &= " WHERE 1=1" & vbCrLf
                'If sPlanYears <> "" Then sql &= " AND PLANYEARS='" & sPlanYears & "'" & vbCrLf
                'If sComIDNO <> "" Then sql &= " AND COMIDNO='" & sComIDNO & "'" & vbCrLf
                sql &= " GROUP BY PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,OCID,CLASSCNAME" & vbCrLf
                sql &= " ,EXAMLVNAME" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " SELECT PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,OCID" & vbCrLf
                sql &= " ,CLASSCNAME+'-'+EXAMLVNAME CLASSCNAME" & vbCrLf
                sql &= " ,STUDCOUNT" & vbCrLf
                sql &= " ,EXAMCOUNT" & vbCrLf
                sql &= " ,OKPASSCOUNT" & vbCrLf
                sql &= " ,NOPASSCOUNT" & vbCrLf
                'sql &= " ,case when EXAMCOUNT > 0 then round((OKPASSCOUNT/EXAMCOUNT*100),2)+'%'" & vbCrLf
                'sql &= " else '0' end passrate" & vbCrLf
                sql &= " ,CASE WHEN EXAMCOUNT > 0 THEN CONVERT(varchar,ROUND((convert(float,OKPASSCOUNT)/EXAMCOUNT*100),2))+'%'" & vbCrLf
                sql &= " ELSE '0' END PASSRATE" & vbCrLf
                sql &= " FROM WG2" & vbCrLf
                sql &= " ORDER BY PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,OCID,CLASSCNAME" & vbCrLf
                sql &= " ,EXAMLVNAME" & vbCrLf

            Case "2" '依檢定類別
                'Dim sqlx1 As String = Search1x()
                Dim sqlxWG1 As String = Sch_WG1_x()

                sql = "" & vbCrLf
                sql &= " WITH WG1 AS (" & sqlxWG1 & ")" & vbCrLf
                sql &= " ,WG2 AS (" & vbCrLf
                sql &= " SELECT PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                'sql &= " ,OCID,CLASSCNAME" & vbCrLf
                sql &= " ,EXAMKIND,EXAMLVNAME" & vbCrLf
                sql &= " ,IsNull(count(cnt1),0) StudCount" & vbCrLf
                sql &= " ,IsNull(count(cnt2),0) ExamCount" & vbCrLf
                sql &= " ,IsNull(count(cnt3),0) okPassCount" & vbCrLf
                sql &= " ,IsNull(count(cnt4),0) noPassCount" & vbCrLf
                sql &= " FROM WG1 g" & vbCrLf
                'sql &= " WHERE 1=1" & vbCrLf
                'If sPlanYears <> "" Then sql &= " AND PLANYEARS='" & sPlanYears & "'" & vbCrLf
                'If sComIDNO <> "" Then sql &= " AND COMIDNO='" & sComIDNO & "'" & vbCrLf
                sql &= " GROUP BY PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,EXAMKIND,EXAMLVNAME" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " SELECT PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,EXAMKIND,EXAMLVNAME" & vbCrLf
                sql &= " ,STUDCOUNT" & vbCrLf
                sql &= " ,EXAMCOUNT" & vbCrLf
                sql &= " ,OKPASSCOUNT" & vbCrLf
                sql &= " ,NOPASSCOUNT" & vbCrLf
                'sql &= " ,case when EXAMCOUNT > 0 then round((OKPASSCOUNT/EXAMCOUNT*100),2)+'%'" & vbCrLf
                'sql &= " else '0' end passrate" & vbCrLf
                sql &= " ,CASE WHEN EXAMCOUNT > 0 THEN CONVERT(varchar,ROUND((convert(float,OKPASSCOUNT)/EXAMCOUNT*100),2))+'%'" & vbCrLf
                sql &= " ELSE '0' END PASSRATE" & vbCrLf
                sql &= " FROM WG2" & vbCrLf
                sql &= " ORDER BY PLANYEARS,YEARPLAN" & vbCrLf
                sql &= " ,COMIDNO,ORGNAME" & vbCrLf
                sql &= " ,EXAMKIND,EXAMLVNAME" & vbCrLf

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            Label3A.Text = "<FONT color=Red>查無資料</FONT>"
            Label3B.Text = "<FONT color=Red>查無資料</FONT>"
            Label3C.Text = "<FONT color=Red>查無資料</FONT>"
            'Me.DataGridTable1.Visible = False
            DataGrid3.Visible = False
            Datagrid4.Visible = False
            Me.DataGrid3.DataSource = Nothing
            Me.DataGrid3.DataBind()
            Me.Datagrid4.DataSource = Nothing
            Me.Datagrid4.DataBind()
            Exit Sub
        End If

        Label3A.Text = TIMS.GET_DISTNAME(objconn, sm.UserInfo.DistID)
        Label3B.Text = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "YearPlan")
        Label3C.Text = TIMS.GetMyValueSQM(ViewState(cst_SD_07_003_search), "OrgName")
        'Me.DataGridTable1.Visible = True
        Select Case rbCounttype.SelectedValue
            Case "1" '統計項目若是以班別
                DataGrid3.Visible = True
                DataGrid3.DataSource = dt
                DataGrid3.DataBind()

            Case "2" '依檢定類別
                Datagrid4.Visible = True
                Datagrid4.DataSource = dt
                Datagrid4.DataBind()

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click
        Call KeepSearch()
        'Response.Redirect("SD_07_003.aspx?ID=" & gID & "&Action=" & cst_search)
        Dim url1 As String = "SD_07_003.aspx?ID=" & gID & "&Action=" & cst_search
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Sub GetSearch()
        If Not Session(cst_SD_07_003_search) Is Nothing Then
            ViewState(cst_SD_07_003_search) = Session(cst_SD_07_003_search)
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "TPlanID") <> "" Then
                Common.SetListItem(TPlanID, TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "TPlanID"))
            End If
            Dim sType1 As String = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "sType1")
            rdoType1.Checked = False
            rdoType2.Checked = False
            Select Case sType1
                Case "1"
                    rdoType1.Checked = True
                Case "2"
                    rdoType2.Checked = True
            End Select

            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "ddlYears") <> "" Then
                Common.SetListItem(ddlYears, TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "ddlYears"))
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtSTDateS") <> "" Then
                txtSTDateS.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtSTDateS")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtSTDateE") <> "" Then
                txtSTDateE.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtSTDateE")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtFTDateS") <> "" Then
                txtFTDateS.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtFTDateS")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtFTDateE") <> "" Then
                txtFTDateE.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtFTDateE")
            End If
            'If TIMS.GetMyValueSQM(Session(Cst_SD_07_003_search), "txtSendoutCertDateS") <> "" Then       '2006/06/03 拿掉發證日期
            '    txtSendoutCertDateS.Text = TIMS.GetMyValueSQM(Session(Cst_SD_07_003_search), "txtSendoutCertDateS")
            'End If
            'If TIMS.GetMyValueSQM(Session(Cst_SD_07_003_search), "txtSendoutCertDateE") <> "" Then
            '    txtSendoutCertDateE.Text = TIMS.GetMyValueSQM(Session(Cst_SD_07_003_search), "txtSendoutCertDateE")
            'End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtExamDateS") <> "" Then
                txtExamDateS.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtExamDateS")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtExamDateE") <> "" Then
                txtExamDateE.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtExamDateE")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtYearsOldS") <> "" Then
                txtYearsOldS.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtYearsOldS")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtYearsOldE") <> "" Then
                txtYearsOldE.Text = TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "txtYearsOldE")
            End If
            If TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "rbCounttype") <> "" Then
                Common.SetListItem(rbCounttype, TIMS.GetMyValueSQM(Session(cst_SD_07_003_search), "rbCounttype"))
            End If

            Dim JScript As String = "<script>ShowType1();</script>" & vbCrLf
            Page.RegisterStartupScript(TIMS.xBlockName, JScript)

        End If
        'Session(Cst_SD_07_003_search) = Nothing
    End Sub

    Sub KeepSearch()
        Dim strsearch As String = ""

        strsearch = "ID=" & gID
        strsearch += "&TPlanID=" & TPlanID.SelectedValue
        Dim sType1 As String = ""
        If rdoType1.Checked Then sType1 = "1"
        If rdoType2.Checked Then sType1 = "2"
        strsearch += "&sType1=" & sType1

        strsearch += "&ddlYears=" & ddlYears.SelectedValue
        strsearch += "&txtSTDateS=" & txtSTDateS.Text.Trim
        strsearch += "&txtSTDateE=" & txtSTDateE.Text.Trim
        strsearch += "&txtFTDateS=" & txtFTDateS.Text.Trim
        strsearch += "&txtFTDateE=" & txtFTDateE.Text.Trim
        'strsearch += "&txtSendoutCertDateS=" & txtSendoutCertDateS.Text.Trim '2006/06/03 拿掉發證日期
        'strsearch += "&txtSendoutCertDateE=" & txtSendoutCertDateE.Text.Trim
        strsearch += "&txtExamDateS=" & txtExamDateS.Text.Trim
        strsearch += "&txtExamDateE=" & txtExamDateE.Text.Trim
        strsearch += "&txtYearsOldS=" & txtYearsOldS.Text.Trim
        strsearch += "&txtYearsOldE=" & txtYearsOldE.Text.Trim
        strsearch += "&rbCounttype=" & rbCounttype.SelectedValue

        Session(cst_SD_07_003_search) = strsearch

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Call KeepSearch()
        Session(cst_SD_07_003_search) &= e.CommandArgument

        '班別連結
        Dim url1 As String = "SD_07_003.aspx?ID=" & gID & "&Action=" & cst_search2 & e.CommandArgument
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LinkButton1 As LinkButton = e.Item.FindControl("LinkButton1")
                LinkButton1.Text = Convert.ToString(drv("YearPlan"))
                LinkButton1.ForeColor = Color.Blue

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "YearPlan", LinkButton1.Text)
                TIMS.SetMyValue(sCmdArg, "PlanYears", Convert.ToString(drv("PlanYears")))
                LinkButton1.CommandArgument = sCmdArg '"&YearPlan=" & LinkButton1.Text & "&PlanYears=" & drv("PlanYears").ToString
                TIMS.Tooltip(LinkButton1, cst_titleMsg1)
                Dim passrate As Label = e.Item.FindControl("passrate")
                passrate.Text = Convert.ToString(drv("passrate"))
                'LinkButton1.ToolTip = "點選可以觀看詳細資料"
                'passrate.ForeColor = Color.Blue
                'passrate.ToolTip = "點選可以觀看詳細資料"

            Case ListItemType.Footer
                For i As Integer = 1 To DataGrid1.Columns.Count - 1 - 1
                    e.Item.Cells(i).Text = 0
                    For Each item As DataGridItem In DataGrid1.Items
                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(item.Cells(i).Text)
                    Next
                Next

                ''及格率(%)
                If Int(e.Item.Cells(2).Text) <> 0 Then
                    e.Item.Cells(5).Text = Format(Int(e.Item.Cells(3).Text) / Int(e.Item.Cells(2).Text), cst_Format1)
                End If

        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        KeepSearch()
        Session(cst_SD_07_003_search) &= e.CommandArgument

        '班別連結
        Dim url1 As String = "SD_07_003.aspx?ID=" & gID & "&Action=" & cst_search3 & e.CommandArgument
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim LinkButton2 As LinkButton = e.Item.FindControl("LinkButton2")
                LinkButton2.Text = Convert.ToString(drv("OrgName"))
                LinkButton2.ForeColor = Color.Blue
                LinkButton2.CommandArgument = "&ComIDNO=" & drv("ComIDNO").ToString & "&OrgName=" & LinkButton2.Text & "&PlanYears=" & drv("PlanYears").ToString
                TIMS.Tooltip(LinkButton2, cst_titleMsg1)
                Dim passrate2 As Label = e.Item.FindControl("passrate2")
                passrate2.Text = Convert.ToString(drv("passrate"))
                'passrate.ForeColor = Color.Blue
                'passrate.ToolTip = "點選可以觀看詳細資料"

            Case ListItemType.Footer
                For i As Integer = 1 To DataGrid2.Columns.Count - 1 - 1
                    e.Item.Cells(i).Text = 0
                    For Each item As DataGridItem In DataGrid2.Items
                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(item.Cells(i).Text)
                    Next
                Next

                ''及格率(%)
                If Int(e.Item.Cells(2).Text) <> 0 Then
                    e.Item.Cells(5).Text = Format(Int(e.Item.Cells(3).Text) / Int(e.Item.Cells(2).Text), cst_Format1)
                End If

        End Select
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labClassName As Label = e.Item.FindControl("labClassName")
                labClassName.Text = Convert.ToString(drv("ClassCName"))

                Dim passrate3 As Label = e.Item.FindControl("passrate3")
                passrate3.Text = Convert.ToString(drv("passrate"))

            Case ListItemType.Footer
                For i As Integer = 1 To DataGrid3.Columns.Count - 1 - 1
                    e.Item.Cells(i).Text = 0
                    For Each item As DataGridItem In DataGrid3.Items
                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(item.Cells(i).Text)
                    Next
                Next

                ''及格率(%)
                If Int(e.Item.Cells(2).Text) <> 0 Then
                    e.Item.Cells(5).Text = Format(Int(e.Item.Cells(3).Text) / Int(e.Item.Cells(2).Text), cst_Format1)
                End If

        End Select
    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Examkind As Label = e.Item.FindControl("Examkind")
                'Examkind.Text = Convert.ToString(drv("ExamName"))
                Examkind.Text = Convert.ToString(drv("ExamLvName"))
                Dim passrate4 As Label = e.Item.FindControl("passrate4")
                passrate4.Text = Convert.ToString(drv("passrate"))

            Case ListItemType.Footer
                For i As Integer = 1 To Datagrid4.Columns.Count - 1 - 1
                    e.Item.Cells(i).Text = 0
                    For Each item As DataGridItem In Datagrid4.Items
                        e.Item.Cells(i).Text = Int(e.Item.Cells(i).Text) + Int(item.Cells(i).Text)
                    Next
                Next

                ''及格率(%)
                If Int(e.Item.Cells(2).Text) <> 0 Then
                    e.Item.Cells(5).Text = Format(Int(e.Item.Cells(3).Text) / Int(e.Item.Cells(2).Text), cst_Format1)
                End If

        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Div1.Visible = True
        Div2.Visible = False
        Div3.Visible = False
        Div4.Visible = False

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Div1.Visible = False
        Div2.Visible = True
        GetSearch()
        Search1()
        Div3.Visible = False
        Div4.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Div1.Visible = False
        Div2.Visible = False
        Div3.Visible = True
        GetSearch()
        Search2()
        Div4.Visible = False
    End Sub

    'Private Sub rblType1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblType1.SelectedIndexChanged
    '    txtSTDateS.Text = ""
    '    txtSTDateE.Text = ""
    '    txtFTDateS.Text = ""
    '    txtFTDateE.Text = ""
    '    'txtSendoutCertDateS.Text = "" '2006/06/03 拿掉發證日期
    '    'txtSendoutCertDateE.Text = ""
    '    txtYearsOldS.Text = ""
    '    txtYearsOldE.Text = ""
    '    txtExamDateS.Text = ""
    '    txtExamDateE.Text = ""
    'End Sub
End Class
