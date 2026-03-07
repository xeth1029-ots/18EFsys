Partial Class TR_04_011_R
    Inherits AuthBasePage

    'TR_04_011_R_c.jrxml 'TR_04_011_R_b.jrxml
    'Const cst_RptName As String = "TR_04_011_R"

    'Const cst_printFN1 As String = "TR_04_011_R_b"
    Const cst_printFN1 As String = "TR_04_011_R_d"
    Const cst_ttipMsg1 As String = "尚有就業資料未填寫「就業關聯性」，請全部填寫完成後 始可執行!!"

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            If sm.UserInfo.RID = "A" Then
                DistID = TIMS.Get_DistID(DistID)
                TPlanID = TIMS.Get_TPlan(TPlanID, , 1)
            Else
                DistID = TIMS.Get_DistID(DistID)
                TPlanID = TIMS.Get_TPlan(TPlanID, , 1)
                DistID.Enabled = False
                TPlanID.Enabled = False
            End If
            Syear = TIMS.GetSyear(Syear)
            'Common.SetListItem(Syear, Now.Year)
            Common.SetListItem(Syear, sm.UserInfo.Years)
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            'DistID.SelectedValue = sm.UserInfo.DistID
            'TPlanID.SelectedValue = sm.UserInfo.TPlanID
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PlanID.Value = sm.UserInfo.PlanID
            'Button2_Click(sender, e) 

            '依查詢條件，顯示班級資料。
            Call Search_OCID1()

        End If

        OCID.Attributes("onchange") = "if(this.selectedIndex!=0){document.form1.OCIDValue.value=this.value;}else{document.form1.OCIDValue.value='';}"
        DistID.Attributes("onchange") = "GetMode();"
        TPlanID.Attributes("onchange") = "GetMode();"
        Button1.Attributes("onclick") = "javascript:return print();"
        Button2.Style("display") = "none"
        If sm.UserInfo.LID <= 1 Then
            Button3.Attributes("onclick") = "wopen('../../Common/MainOrg.aspx?DistID='+document.form1.DistID.value+'&amp;TPlanID='+document.form1.TPlanID.value,'訓練機構',400,400,1);"
        Else
            Button3.Attributes("onclick") = "openOrg('../../Common/LevOrg1.aspx?btnName=Button2');"
        End If

    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim MyValue As String = ""
        '取得查詢參數。outValue@MyValue
        If Not GetSearchValue(MyValue, 1) Then
            Common.MessageBox(Me.Page, "請選擇班別!")
            Exit Sub
        End If

        '檢核 就業關聯性 是否全部填寫 false:有部份未填寫 true:全部已填寫
        Dim flag2 As Boolean = TIMS.CHK_JOBRELATE_OCID(OCIDValue.Value, 1, objconn)
        hid_CHKJOBRELATE_NG.Value = ""
        If Not flag2 Then hid_CHKJOBRELATE_NG.Value = TIMS.cst_YES '就業關聯性(有部份未填寫)

        '假設該按鈕可新增，但確有資料未填寫 就業關聯性 CHK_JOBRELATE_OCID
        If hid_CHKJOBRELATE_NG.Value = TIMS.cst_YES Then
            Common.MessageBox(Me, cst_ttipMsg1)
            Exit Sub
        End If

        'TR_04_011_R_b
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    '取得查詢參數。
    Function GetSearchValue(ByRef outValue As String, ByVal sType As Integer) As Boolean
        Dim rst As Boolean = True '沒有異常為True
        'sType :1 報表用參數功能。
        'sType :2 查詢後匯出參數功能。

        '報名管道
        Dim sEnterChanl As String = ""
        sEnterChanl = ""
        For i As Integer = 0 To CkEnterChannel.Items.Count - 1
            If CkEnterChannel.Items.Item(i).Selected = True Then
                If sEnterChanl <> "" Then sEnterChanl &= ","
                sEnterChanl &= CkEnterChannel.Items.Item(i).Value
            End If
        Next

        OCIDValue.Value = "" '輸出用。
        Select Case sType
            Case 1
                OCIDValue.Value = OCID.SelectedValue
                OCIDValue.Value = TIMS.ClearSQM(OCIDValue.Value)
                If OCIDValue.Value = "" Then Return False '異常
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue.Value, objconn)
                If drCC Is Nothing Then Return False '異常

            Case 2
                OCIDValue.Value = OCID.SelectedValue
                OCIDValue.Value = TIMS.ClearSQM(OCIDValue.Value)
                If OCIDValue.Value = "" Then Return False '異常
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue.Value, objconn)
                If drCC Is Nothing Then Return False '異常

        End Select
        'Dim MyValue As String = ""

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        outValue = ""
        outValue = "kjs=kjs"
        outValue += "&OCID=" & OCIDValue.Value
        outValue += "&Years=" & Syear.SelectedValue
        outValue += "&DistID=" & DistID.SelectedValue
        outValue += "&RID=" & RIDValue.Value
        outValue += "&TPlanID=" & TPlanID.SelectedValue
        outValue += "&EnterChanl=" & sEnterChanl
        outValue += "&STDate1=" & STDate1.Text
        outValue += "&STDate2=" & STDate2.Text
        Return rst
    End Function

    '依查詢條件，顯示班級資料。(SQL)
    Sub Search_OCID1()
        OCID.Items.Clear()

        If RIDValue.Value <> "" Then
            'Dim sql As String = ""
            'Dim dt As DataTable
            'Dim dr As DataRow
            'Dim strSelected As String = ""
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " SELECT cc.ClassCName" & vbCrLf
            sql &= " ,cc.CyclType" & vbCrLf
            sql &= " ,cc.LevelType" & vbCrLf
            sql &= " ,cc.OCID" & vbCrLf
            sql &= " FROM Class_ClassInfo cc " & vbCrLf
            sql &= " join id_plan ip on ip.planid=cc.planid" & vbCrLf
            sql &= " WHERE 1=1" & vbCrLf
            sql &= " and cc.IsSuccess='Y' " & vbCrLf
            sql &= " and cc.NotOpen='N'" & vbCrLf
            sql &= " and cc.RID='" & RIDValue.Value & "' " & vbCrLf
            sql &= " and ip.TPlanID='" & TPlanID.SelectedValue & "'" & vbCrLf
            sql &= " and ip.DistID='" & DistID.SelectedValue & "'" & vbCrLf
            sql &= " and ip.PlanID='" & PlanID.Value & "'" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

            If dt.Rows.Count = 0 Then
                OCID.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
            Else
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objconn)
                Dim strSelected As String = ""
                For Each dr As DataRow In dt.Rows
                    Dim ClassName As String = dr("ClassCName").ToString
                    If Int(dr("CyclType")) <> 0 Then
                        ClassName += "第" & Int(dr("CyclType")) & "期"
                    End If
                    If Not IsDBNull(dr("LevelType")) Then
                        If Int(dr("LevelType")) <> 0 Then
                            ClassName += "第" & Int(dr("LevelType")) & "階段"
                        End If
                    End If

                    OCID.Items.Add(New ListItem(ClassName, dr("OCID")))
                    If Convert.ToString(dr("OCID")) = OCIDValue1.Value Then
                        strSelected = Convert.ToString(dr("OCID"))
                    End If
                Next

                OCID.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                If strSelected <> "" Then
                    Common.SetListItem(OCID, strSelected)
                    'OCID.SelectedValue = strSelected
                End If
            End If
        End If
    End Sub

    '查詢
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '依查詢條件，顯示班級資料。
        Call Search_OCID1()
    End Sub

    'SQL ('匯出EXCEL)
    Function LoadData1(ByVal MyValue As String) As DataTable
        Dim dt As New DataTable

        Dim vOCID As String = TIMS.GetMyValue(MyValue, "OCID")
        Dim vYears As String = TIMS.GetMyValue(MyValue, "Years")
        Dim vDistID As String = TIMS.GetMyValue(MyValue, "DistID")
        Dim vRID As String = TIMS.GetMyValue(MyValue, "RID")

        Dim vTPlanID As String = TIMS.GetMyValue(MyValue, "TPlanID")
        Dim vEnterChanl As String = TIMS.GetMyValue(MyValue, "EnterChanl")
        Dim vSTDate1 As String = TIMS.GetMyValue(MyValue, "STDate1")
        Dim vSTDate2 As String = TIMS.GetMyValue(MyValue, "STDate2")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select cc.ocid" & vbCrLf
        sql &= " ,cc.orgname" & vbCrLf
        sql &= " ,cc.classcname2" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and cc.ocid = 71801" & vbCrLf
        '可能為空。
        If vOCID <> "" Then
            sql &= " AND cc.OCID='" & vOCID & "'" & vbCrLf
        End If
        If vYears <> "" Then
            sql &= " AND cc.Years='" & vYears & "'" & vbCrLf
        End If
        If vDistID <> "" Then
            sql &= " AND cc.DistID='" & vDistID & "'" & vbCrLf
        End If
        Select Case sm.UserInfo.LID
            Case "0", "1" '署'分署
            Case Else '其他單位。
                sql &= " AND cc.RID='" & vRID & "'" & vbCrLf
                'If vRID <> "" Then sql &= " AND cc.RID='" & vRID & "'" & vbCrLf
        End Select
        If vTPlanID <> "" Then
            sql &= " AND cc.TPlanID='" & vTPlanID & "'" & vbCrLf
        End If
        If vSTDate1 <> "" Then
            sql &= " AND cc.STDate >= " & TIMS.To_date(vSTDate1) & vbCrLf
        End If
        If vSTDate2 <> "" Then
            sql &= " AND cc.STDate <= " & TIMS.To_date(vSTDate2) & vbCrLf
        End If
        sql &= " )" & vbCrLf
        sql &= " ,WJ1 AS (" & vbCrLf
        sql &= " select cc.ocid,ss.socid" & vbCrLf
        sql &= " ,ss.Studstatus,ss.StudID,CONVERT(numeric, ss.StudID) StudID2,ss.name,ss.idno" & vbCrLf
        sql &= " ,ss.phoneD,ss.CellPhone" & vbCrLf
        sql &= " ,j1.BusName,j1.zipName,j1.BusAddr" & vbCrLf
        sql &= " ,j1.BusTel,CONVERT(varchar, j1.MDate, 111) MDate" & vbCrLf
        sql &= " ,j1.SalID,k1.SalName" & vbCrLf
        sql &= " ,j1.NGJobDesc" & vbCrLf
        sql &= " ,j1.NGJobDesc2" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN V_STUDENTINFO ss on ss.ocid =cc.ocid" & vbCrLf
        sql &= " JOIN V_GETJOBC1 j1 on j1.socid =ss.socid" & vbCrLf
        sql &= " LEFT JOIN KEY_SALARY k1 ON k1.SalID=j1.SalID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and ss.Studstatus NOT IN (2,3)" & vbCrLf
        If vEnterChanl <> "" Then
            sql &= " AND ss.EnterChannel IN (" & vEnterChanl & ")" & vbCrLf
        End If
        sql &= " )" & vbCrLf
        sql &= " ,WJ9 AS (" & vbCrLf
        sql &= " select cc.ocid,ss.socid" & vbCrLf
        sql &= " ,ss.Studstatus,ss.StudID,CONVERT(numeric, ss.StudID) StudID2,ss.name,ss.idno" & vbCrLf
        sql &= " ,ss.phoneD,ss.CellPhone" & vbCrLf
        sql &= " ,j1.BusName,j1.zipName,j1.BusAddr" & vbCrLf
        sql &= " ,j1.BusTel,CONVERT(varchar, j1.MDate, 111) MDate" & vbCrLf
        sql &= " ,j1.SalID,k1.SalName" & vbCrLf
        sql &= " ,j1.NGJobDesc" & vbCrLf
        sql &= " ,j1.NGJobDesc2" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN V_STUDENTINFO ss on ss.ocid =cc.ocid" & vbCrLf
        sql &= " left JOIN V_GETJOBC9 j1 on j1.socid =ss.socid" & vbCrLf
        sql &= " LEFT JOIN KEY_SALARY k1 ON k1.SalID=j1.SalID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and ss.Studstatus IN (2,3)" & vbCrLf
        If vEnterChanl <> "" Then
            sql &= " AND ss.EnterChannel IN (" & vEnterChanl & ")" & vbCrLf
        End If
        sql &= " and (1!=1" & vbCrLf
        sql &= " or (ss.WkAheadOfSch = 'Y' and ss.RTReasonID ='02')" & vbCrLf
        sql &= " or (j1.socid is not null)" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " select cc.ocid" & vbCrLf
        sql &= " ,cc.orgname" & vbCrLf
        sql &= " ,cc.classcname2" & vbCrLf
        sql &= " ,cc.stdate" & vbCrLf
        sql &= " ,cc.ftdate" & vbCrLf
        sql &= " ,oj.lostjob " & vbCrLf '(處理) 'STUD_LOSTJOBWEEK (批次處理：LostJobApp.vbproj)
        sql &= " ,g.Studstatus,g.StudID,g.StudID2,g.name,g.idno" & vbCrLf
        sql &= " ,g.phoneD,g.CellPhone" & vbCrLf
        sql &= " ,g.BusName,g.zipName,g.BusAddr" & vbCrLf
        sql &= " ,g.BusTel,g.MDate" & vbCrLf
        sql &= " ,g.SalID,g.SalName" & vbCrLf
        sql &= " ,g.NGJobDesc" & vbCrLf
        sql &= " ,g.NGJobDesc2" & vbCrLf
        sql &= " from (" & vbCrLf
        sql &= " select * from WJ1" & vbCrLf
        sql &= " UNION select * from WJ9" & vbCrLf
        sql &= " ) g" & vbCrLf
        sql &= " JOIN WC1 cc ON cc.ocid =g.ocid" & vbCrLf
        sql &= " LEFT JOIN STUD_LOSTJOBWEEK oj on oj.socid =g.socid" & vbCrLf
        sql &= " ORDER BY CONVERT(numeric, g.StudID)" & vbCrLf

        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Try
            With oCmd
                .Parameters.Clear()
                dt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            'Throw ex
            Dim strErrmsg As String = ""
            strErrmsg += "/* Function LoadData1(ByVal MyValue As String) As DataTable */" & vbCrLf
            strErrmsg += " vOCID:" & vOCID & vbCrLf
            strErrmsg += " vYears:" & vYears & vbCrLf
            strErrmsg += " vDistID:" & vDistID & vbCrLf
            strErrmsg += " vRID:" & vRID & vbCrLf
            strErrmsg += " vTPlanID:" & vTPlanID & vbCrLf
            strErrmsg += " vEnterChanl:" & vEnterChanl & vbCrLf
            strErrmsg += " vSTDate1:" & vSTDate1 & vbCrLf
            strErrmsg += " vSTDate2:" & vSTDate2 & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf

            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        'sql = "select * from STUD_LOSTJOBWEEK"
        Return dt
    End Function

    'Sub kldfskljdfskljdfsklj()
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " WITH WC1 AS (" & vbCrLf
    '    sql &= " select cc.ocid" & vbCrLf
    '    sql &= " ,cc.orgname" & vbCrLf
    '    sql &= " ,cc.classcname2" & vbCrLf
    '    sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
    '    sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
    '    sql &= " FROM VIEW2 cc" & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " and cc.ocid = 71801" & vbCrLf
    '    sql &= " )" & vbCrLf
    '    sql &= " ,WJ1 AS (" & vbCrLf
    '    sql &= " select cc.ocid,ss.socid" & vbCrLf
    '    sql &= " ,ss.Studstatus  ,ss.StudID  ,CONVERT(numeric, ss.StudID) StudID2  ,ss.name  ,ss.idno" & vbCrLf
    '    sql &= " ,ss.phoneD  ,ss.CellPhone" & vbCrLf
    '    sql &= " ,j1.BusName  ,j1.zipName  ,j1.BusAddr" & vbCrLf
    '    sql &= " ,j1.BusTel  ,CONVERT(varchar, j1.MDate, 111)  MDate" & vbCrLf
    '    sql &= " ,j1.SalID  ,k1.SalName" & vbCrLf
    '    sql &= " ,j1.NGJobDesc" & vbCrLf
    '    sql &= " ,j1.NGJobDesc2" & vbCrLf
    '    sql &= " FROM WC1 cc" & vbCrLf
    '    sql &= " JOIN V_STUDENTINFO ss on ss.ocid =cc.ocid" & vbCrLf
    '    sql &= " JOIN V_GETJOBC1 j1 on j1.socid =ss.socid" & vbCrLf
    '    sql &= " LEFT JOIN KEY_SALARY k1 ON k1.SalID=j1.SalID" & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " and ss.Studstatus NOT IN (2,3)" & vbCrLf
    '    sql &= " )" & vbCrLf
    '    sql &= " ,WJ9 AS (" & vbCrLf
    '    sql &= " select cc.ocid,ss.socid" & vbCrLf
    '    sql &= " ,ss.Studstatus  ,ss.StudID  ,CONVERT(numeric, ss.StudID) StudID2  ,ss.name  ,ss.idno" & vbCrLf
    '    sql &= " ,ss.phoneD  ,ss.CellPhone" & vbCrLf
    '    sql &= " ,j1.BusName  ,j1.zipName  ,j1.BusAddr" & vbCrLf
    '    sql &= " ,j1.BusTel  ,CONVERT(varchar, j1.MDate, 111)  MDate" & vbCrLf
    '    sql &= " ,j1.SalID  ,k1.SalName" & vbCrLf
    '    sql &= " ,j1.NGJobDesc" & vbCrLf
    '    sql &= " ,j1.NGJobDesc2" & vbCrLf
    '    sql &= " FROM WC1 cc" & vbCrLf
    '    sql &= " JOIN V_STUDENTINFO ss on ss.ocid =cc.ocid" & vbCrLf
    '    sql &= " left JOIN V_GETJOBC9 j1 on j1.socid =ss.socid" & vbCrLf
    '    sql &= " LEFT JOIN KEY_SALARY k1 ON k1.SalID=j1.SalID" & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " and ss.Studstatus IN (2,3)" & vbCrLf
    '    sql &= " and (1!=1" & vbCrLf
    '    sql &= " or (ss.WkAheadOfSch = 'Y' and ss.RTReasonID ='02')" & vbCrLf
    '    sql &= " or (j1.socid is not null)" & vbCrLf
    '    sql &= " )" & vbCrLf
    '    sql &= " )" & vbCrLf
    '    sql &= " select cc.ocid" & vbCrLf
    '    sql &= " ,cc.orgname" & vbCrLf
    '    sql &= " ,cc.classcname2" & vbCrLf
    '    sql &= " ,cc.stdate" & vbCrLf
    '    sql &= " ,cc.ftdate" & vbCrLf
    '    sql &= " ,oj.lostjob" & vbCrLf
    '    sql &= " ,g.Studstatus  ,g.StudID  ,g.StudID2  ,g.name  ,g.idno" & vbCrLf
    '    sql &= " ,g.phoneD  ,g.CellPhone" & vbCrLf
    '    sql &= " ,g.BusName ,g.zipName  ,g.BusAddr" & vbCrLf
    '    sql &= " ,g.BusTel  ,g.MDate" & vbCrLf
    '    sql &= " ,g.SalID  ,g.SalName" & vbCrLf
    '    sql &= " ,g.NGJobDesc" & vbCrLf
    '    sql &= " ,g.NGJobDesc2" & vbCrLf
    '    sql &= " from (" & vbCrLf
    '    sql &= " select * from WJ1" & vbCrLf
    '    sql &= " UNION select * from WJ9" & vbCrLf
    '    sql &= " ) g" & vbCrLf
    '    sql &= " JOIN WC1 cc ON cc.ocid =g.ocid" & vbCrLf
    '    sql &= " LEFT JOIN STUD_LOSTJOBWEEK oj on oj.socid =g.socid" & vbCrLf
    '    sql &= " ORDER BY CONVERT(numeric, g.StudID)" & vbCrLf

    'End Sub

    '匯出 Response (結訓學員輔導就業成果名冊) '學員輔導就業成果名冊 ('匯出EXCEL)
    Sub ExpReport1(ByRef dt As DataTable)

        Dim strTitle1 As String = "" '匯出表頭名稱
        'strTitle1 = "結訓學員輔導就業成果名冊"
        strTitle1 = "學員輔導就業成果名冊"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        '套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        '"訓練機構"、"班別名稱"、"訓練起日"、"訓練迄日"、"學號"、"學員姓名"、
        '"身分證號"、"未加保週數"、"電話"、"就業單位名稱"、"就業單位電話"、"就業單位地址"、
        '"到職日期"、"薪資級距"、"就業狀況"欄位。
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td>訓練機構</td>" & vbTab '訓練機構
        ExportStr &= "<td>班別名稱</td>" & vbTab '
        ExportStr &= "<td>訓練起日</td>" & vbTab '
        ExportStr &= "<td>訓練迄日</td>" & vbTab '
        ExportStr &= "<td>學號</td>" & vbTab '
        ExportStr &= "<td>學員姓名</td>" & vbTab '
        ExportStr &= "<td>身分證號</td>" & vbTab '
        ExportStr &= "<td>未加保週數</td>" & vbTab '
        ExportStr &= "<td>電話</td>" & vbTab '
        ExportStr &= "<td>電話2</td>" & vbTab '

        ExportStr &= "<td>就業單位名稱</td>" & vbTab '
        ExportStr &= "<td>就業單位電話</td>" & vbTab '
        ExportStr &= "<td>就業單位縣市</td>" & vbTab '
        ExportStr &= "<td>就業單位地址</td>" & vbTab '

        ExportStr &= "<td>到職日期</td>" & vbTab '
        ExportStr &= "<td>薪資級距</td>" & vbTab '
        ExportStr &= "<td>就業狀況</td>" & vbTab '
        ExportStr &= "<td>切結對象</td>" & vbTab '

        ExportStr += "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Select("1=1", "orgname,classcname2,StudID")
            'For Each dr As DataRow In dt.Rows
            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("classcname2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("stdate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ftdate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("StudID")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("name")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("idno")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("LOSTJOB")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("phoneD")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CellPhone")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("BusName")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("BusTel")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("zipName")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("BusAddr")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("mdate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("SalName")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("NGJobDesc")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("NGJobDesc2")) & "</td>" & vbTab '

            ExportStr += "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        Call TIMS.CloseDbConn(objconn)
        Response.End()

    End Sub

    '匯出EXCEL
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Dim MyValue As String = ""
        '取得查詢參數。outValue@MyValue
        If Not GetSearchValue(MyValue, 2) Then
            Common.MessageBox(Me.Page, "請選擇班別!")
            Exit Sub
        End If

        Dim dt As DataTable
        dt = LoadData1(MyValue)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        '檢核 就業關聯性 是否全部填寫 false:有部份未填寫 true:全部已填寫
        Dim flag2 As Boolean = TIMS.CHK_JOBRELATE_OCID(OCIDValue.Value, 1, objconn)
        hid_CHKJOBRELATE_NG.Value = ""
        If Not flag2 Then hid_CHKJOBRELATE_NG.Value = TIMS.cst_YES '就業關聯性(有部份未填寫)

        '假設該按鈕可新增，但確有資料未填寫 就業關聯性 CHK_JOBRELATE_OCID
        If hid_CHKJOBRELATE_NG.Value = TIMS.cst_YES Then
            Common.MessageBox(Me, cst_ttipMsg1)
            Exit Sub
        End If

        Call ExpReport1(dt) '匯出EXCEL
    End Sub

End Class