Partial Class TC_01_003_import
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Dim aTPlanid As String = TIMS.ClearSQM(Request("TPlanid"))
            Year = getYear(Year, aTPlanid, sm.UserInfo.DistID, "")
            Year2 = getYear(Year2, aTPlanid, sm.UserInfo.DistID, sm.UserInfo.Years)
        End If
    End Sub

    Function getYear(ByVal Obj As DropDownList, ByVal TplanID As String, ByVal DistID As String, ByVal SessionYears As String) As DropDownList
        If TplanID = "" Then Return Obj
        If DistID = "" Then Return Obj

        Obj.Items.Clear()
        If SessionYears <> "" Then
            Dim txt_Years As String = SessionYears
            If flag_ROC Then txt_Years = CStr(CInt(SessionYears) - 1911) 'edit，by:20181018
            '目的年度
            Obj.Items.Insert(0, New ListItem(txt_Years, SessionYears))
            Return Obj
        End If

        '來源年度
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        'sql += " SELECT DISTINCT Years FROM ID_Class " & vbCrLf
        sql &= " SELECT DISTINCT YEARS " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, CONVERT(INT, years)-1911) ROC_YEAR" & vbCrLf
        sql &= "  FROM ID_Class " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and TplanID = '" & TplanID & "' " & vbCrLf
        sql &= " AND DistID = '" & DistID & "' " & vbCrLf
        sql &= " AND Years IS NOT NULL " & vbCrLf
        sql &= " ORDER BY YEARS DESC" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Obj.Items.Insert(0, New ListItem("此功能尚未分年度前的資料", "00"))
            Return Obj
        End If

        With Obj
            .DataSource = dt
            Dim TXT_Years As String = "YEARS"
            If flag_ROC Then TXT_Years = "ROC_YEAR"   'edit，by:20181018

            .DataTextField = TXT_Years
            .DataValueField = "Years"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        Return Obj
    End Function


    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Convert.ToString(Year.SelectedValue) <> "" Then
            If Not IsNumeric(Year.SelectedValue) Then
                Errmsg += "來源年度 資料有誤" & vbCrLf
            Else
                If Year.SelectedValue = "0" Then Errmsg += "請選擇 來源年度" & vbCrLf
            End If
        Else
            Errmsg += "請選擇 來源年度" & vbCrLf
        End If

        If Convert.ToString(Year2.SelectedValue) <> "" Then
            If Not IsNumeric(Year2.SelectedValue) Then
                Errmsg += "目的年度 資料有誤" & vbCrLf
            Else
                If Year2.SelectedValue = "0" Then
                    Errmsg += "請選擇 目的年度" & vbCrLf
                End If
            End If
        Else
            Errmsg += "請選擇 目的年度" & vbCrLf
        End If

        If Errmsg = "" Then
            If Year2.SelectedValue <> sm.UserInfo.Years Then Errmsg += "登入計畫年度 與 要匯入的目的年度 不同" & vbCrLf
        End If
        If Errmsg = "" Then
            If Year2.SelectedValue = Year.SelectedValue Then Errmsg += "來源年度 與 要匯入的目的年度 相同，必須不同" & vbCrLf
        End If
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim aYearSel As String = TIMS.ClearSQM(Year.SelectedValue)
        If aYearSel = "" Then Exit Sub
        Dim aTPlanid As String = TIMS.ClearSQM(Request("TPlanid"))
        If aTPlanid = "" Then Exit Sub
        If sm Is Nothing Then Exit Sub
        If sm.UserInfo Is Nothing Then Exit Sub
        If sm.UserInfo.UserID Is Nothing Then Exit Sub
        If sm.UserInfo.UserID = "" Then Exit Sub

        Call TIMS.OpenDbConn(objconn)
        Dim iSql As String = ""
        iSql = "" & vbCrLf
        iSql &= " INSERT INTO ID_CLASS (CLSID,ClassID,ClassName,ClassEName,TPlanID,Content,TMID,DistID,ModifyAcct,ModifyDate,CJOB_UNKEY,Years,CLASSID2) " & vbCrLf
        iSql &= " VALUES (@CLSID,@ClassID,@ClassName,@ClassEName,@TPlanID,@Content,@TMID,@DistID,@ModifyAcct,GETDATE(),@CJOB_UNKEY,@Years,@CLASSID2) " & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT p.ClassID,p.ClassName,p.ClassEName,p.TplanID,p.Content,p.TMID,p.DistID" & vbCrLf
        'sql &= " ,'" & sm.UserInfo.UserID & "' ModifyAcct " & vbCrLf
        'sql &= " ,GETDATE() ModifyDate " & vbCrLf
        sql &= " ,p.CJOB_UNKEY " & vbCrLf
        sql &= " ,p.CLASSID2 " & vbCrLf
        'sql &= " ,'" & sm.UserInfo.Years & "' Years " & vbCrLf
        sql &= " FROM ID_Class p " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND p.TPlanID = '" & aTPlanid & "' " & vbCrLf
        sql &= " AND p.Distid = " & sm.UserInfo.DistID & " " & vbCrLf
        If Year.SelectedValue <> "00" Then sql &= " AND p.Years = '" & aYearSel & "'" & vbCrLf
        sql &= " AND NOT EXISTS ( " & vbCrLf
        sql &= "   SELECT 'x'" & vbCrLf
        sql &= "   FROM ID_Class x" & vbCrLf
        sql &= "   WHERE 1=1" & vbCrLf
        sql &= "   AND x.TPlanID = '" & aTPlanid & "'" & vbCrLf
        sql &= "   AND x.DistID ='" & sm.UserInfo.DistID & "' " & vbCrLf
        sql &= "   AND x.Years = '" & sm.UserInfo.Years & "'" & vbCrLf
        sql &= "   AND x.CLassID = p.CLassID " & vbCrLf
        sql &= " ) " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Common.MessageBox2(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim i_ADD As Integer = 0
        For Each dr1 As DataRow In dt.Rows
            Dim iCLSID As Integer = DbAccess.GetNewId(objconn, "ID_CLASS_CLSID_SEQ,ID_CLASS,CLSID")
            Dim myParam As New Hashtable
            myParam.Clear()
            myParam.Add("CLSID", iCLSID)
            myParam.Add("ClassID", dr1("ClassID"))
            myParam.Add("ClassName", dr1("ClassName"))
            myParam.Add("ClassEName", dr1("ClassEName"))
            myParam.Add("TPlanID", dr1("TPlanID"))
            myParam.Add("Content", dr1("Content"))
            myParam.Add("TMID", dr1("TMID"))
            myParam.Add("DistID", dr1("DistID"))
            myParam.Add("ModifyAcct", sm.UserInfo.UserID)
            myParam.Add("CJOB_UNKEY", dr1("CJOB_UNKEY"))
            myParam.Add("Years", sm.UserInfo.Years)
            myParam.Add("CLASSID2", dr1("CLASSID2"))
            DbAccess.ExecuteNonQuery(iSql, objconn, myParam)
            myParam = Nothing
            i_ADD += 1
        Next

        Dim v_Msg As String = "執行匯入完成"
        If i_ADD = 0 Then v_Msg = TIMS.cst_NODATAMsg1
        Common.MessageBox(Me, v_Msg)

        'Common.RespWrite(Me, "<Script>alert('匯入完成!!');</Script>")
        'If Request("bt_search") <> "" Then
        '    Common.RespWrite(Me, "<Script>opener.document.getElementById('bt_search').click();</Script>")
        '    Common.RespWrite(Me, "<Script>window.close();</Script>")
        'End If
    End Sub

    Protected Sub Btn_Back1_Click(sender As Object, e As EventArgs) Handles Btn_Back1.Click
        Dim uUrl1 As String = ""
        uUrl1 = "TC_01_003.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, uUrl1)
    End Sub
End Class