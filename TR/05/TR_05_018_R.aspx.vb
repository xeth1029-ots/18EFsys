Partial Class TR_05_018_R
    Inherits AuthBasePage

    'TR_05_018_1

    'select * from v_Key_BusinessKeys where years ='2013' --產業別 與 關鍵字對照表。
    'select * from KEY_BUSINESSKEYS where DEPID IN ('06','07','08') --'關鍵字。
    'select * from Key_Depot where years ='2013' --2013年度產業。
    'select * from KEY_BUSINESS where DEPID IN ('06','07','08') --各產業 產業別。
    ' select * from v_Depot06 ORDER BY KID --2013-六大新興產業
    ' select * from v_Depot10 ORDER BY KID --2013-十大重點服務業
    ' select * from v_Depot04 ORDER BY KID --2013-四大新興智慧型產業
    Const cst_DepID06 As String = "07" ' select * from v_Depot06 ORDER BY KID --六大新興產業
    Const cst_DepID10 As String = "08" ' select * from v_Depot10 ORDER BY KID --十大重點服務業
    Const cst_DepID04 As String = "06" ' select * from v_Depot04 ORDER BY KID --四大新興智慧型產業

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
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
            '產業別鍵詞
            KID_6.Attributes("onclick") = "SelectAll('KID_6','KID_6_hid');" '六大新興產業關鍵字查詢項目如下
            KID_10.Attributes("onclick") = "SelectAll('KID_10','KID_10_hid');" '十大重點服務業關鍵字查詢項目如下
            KID_4.Attributes("onclick") = "SelectAll('KID_4','KID_4_hid');" '四大智慧型產業關鍵字查詢項目如下
            '選擇全部預算來源
            BudgetList.Attributes("onclick") = "SelectAll('BudgetList','hidBudgetList');"

            '列印檢查
            Me.Print.Attributes("onclick") = "javascript:return CheckPrint();"
            '匯出名細檢查
            Export1.Attributes("onclick") = "javascript:return CheckPrint();"
        End If

    End Sub

    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, sm.UserInfo.Years) '預設值

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))

        '計畫
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        '新年度2013
        Get_KeyBusiness(KID_6, "07") '六大新興產業關鍵字查詢項目如下
        Get_KeyBusiness(KID_10, "08") '十大重點服務業關鍵字查詢項目如下
        Get_KeyBusiness(KID_4, "06") '四大智慧型產業關鍵字查詢項目如下

        '預算來源
        BudgetList = TIMS.Get_Budget(BudgetList, 3)
        BudgetList.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    '產業別鍵詞
    Function Get_KeyBusiness(ByVal obj As ListControl, ByVal DepID As String) As ListControl
        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql += " select a.KNAME ,a.KID ,a.SeqNo ,a.DepID,a.Status" & vbCrLf
        sql += " from Key_Business a" & vbCrLf
        sql += " join Key_Depot b on b.Depid=a.DepID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " AND a.Status is NULL " & vbCrLf
        sql += " AND a.DepID=@DepID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable '= Nothing
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("DepID", SqlDbType.VarChar).Value = DepID
            dt.Load(.ExecuteReader())
        End With
        With obj
            .Items.Clear()
            .DataSource = dt
            .DataTextField = "KNAME"
            .DataValueField = "KID" '"SeqNo"
            .DataBind()
            If TypeOf obj Is DropDownList Then
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End If
            If TypeOf obj Is CheckBoxList Then
                .Items.Insert(0, New ListItem("全部", ""))
            End If
        End With
        Return obj
    End Function

    '匯出 Response
    Sub ExpReport1(ByRef dt As DataTable, ByVal sKeyName As String)
        Const cst_title1 As String = "四大、六大、十大產業課程統計表"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(cst_title1, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        ''mso-number-format:"0" 
        'Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""

        '建立抬頭
        ExportStr = ""
        ExportStr += "<tr>" & vbCrLf
        ExportStr &= "<td>序號</td>" & vbTab
        ExportStr &= "<td>轄區</td>" & vbTab
        ExportStr &= "<td>訓練計畫</td>" & vbTab
        ExportStr &= "<td>訓練機構名稱</td>" & vbTab

        ExportStr &= "<td>班別名稱</td>" & vbTab
        ExportStr &= "<td>訓練職類</td>" & vbTab
        ExportStr &= "<td>訓練性質</td>" & vbTab
        ExportStr &= "<td>訓練時段</td>" & vbTab
        ExportStr &= "<td>開訓日期</td>" & vbTab
        ExportStr &= "<td>結訓日期</td>" & vbTab

        ExportStr &= "<td>招生人數</td>" & vbTab
        ExportStr &= "<td>時數</td>" & vbTab

        ExportStr &= "<td>開訓人數</td>" & vbTab
        ExportStr &= "<td>結訓人數</td>" & vbTab
        ExportStr &= "<td>就業人數</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))


        Dim cnt1 As Integer = 0 '招生人數
        Dim cnt2 As Integer = 0 '時數
        Dim cnt3 As Integer = 0 '開訓人數
        Dim cnt4 As Integer = 0 '結訓人數
        Dim cnt5 As Integer = 0 '就業人數

        '建立資料面
        Dim i As Integer = 0

        i = 0
        For Each dr As DataRow In dt.Rows
            i += 1

            cnt1 += CInt(Val(dr("TNum")))
            cnt2 += CInt(Val(dr("THours")))
            cnt3 += CInt(Val(dr("openNum")))
            cnt4 += CInt(Val(dr("closeNum")))
            cnt5 += CInt(Val(dr("jobNum")))

            ExportStr = ""
            ExportStr += "<tr>" & vbCrLf
            ExportStr &= "<td>" & Convert.ToString(i) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("distname")) & "</td>" & vbTab '轄區
            ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab '訓練計畫
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>" & vbTab '訓練機構名稱

            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab '班別名稱
            ExportStr &= "<td>" & Convert.ToString(dr("trainName")) & "</td>" & vbTab '訓練職類
            ExportStr &= "<td>" & Convert.ToString(dr("PropertyID")) & "</td>" & vbTab '訓練性質
            ExportStr &= "<td>" & Convert.ToString(dr("hourRanName")) & "</td>" & vbTab '訓練時段
            ExportStr &= "<td>" & Convert.ToString(dr("STDate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("FTDate")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(Val(dr("TNum"))) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(Val(dr("THours"))) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(Val(dr("openNum"))) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(Val(dr("closeNum"))) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(Val(dr("jobNum"))) & "</td>" & vbTab
            ExportStr += "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        '建立尾1
        ExportStr = ""
        ExportStr += "<tr>" & vbCrLf
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(dt.Rows.Count) & "班</td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab

        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab
        ExportStr &= "<td></td>" & vbTab

        ExportStr &= "<td>" & Convert.ToString(cnt1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(cnt2) & "</td>" & vbTab

        ExportStr &= "<td>" & Convert.ToString(cnt3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(cnt4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(cnt5) & "</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立尾2
        If sKeyName <> "" Then
            ExportStr = ""
            ExportStr += "<tr>" & vbCrLf
            ExportStr &= "<td>查詢關鍵字：</td>" & vbTab
            ExportStr &= "<td colspan=""9"">" & Convert.ToString(sKeyName) & "</td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab

            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab
            'ExportStr &= "<td></td>" & vbTab

            ExportStr &= "<td></td>" & vbTab
            ExportStr &= "<td></td>" & vbTab

            ExportStr &= "<td></td>" & vbTab
            ExportStr &= "<td></td>" & vbTab
            ExportStr &= "<td></td>" & vbTab
            ExportStr += "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        End If


        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        Response.End()

    End Sub

    '取出關鍵字
    Function GetKeyName(ByVal DepID As String, ByVal KID As String) As String
        Dim rst As String = ""
        'Sql = "" & vbCrLf
        'Sql += " select 'x' from v_Key_BusinessKeys x " & vbCrLf
        'Sql += " where 1=1" & vbCrLf
        'Sql += " and CHARINDEX(x.KeyName,cc.classcname )<>0" & vbCrLf
        If KID <> "" Then
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " SELECT k.DepID,k.KID" & vbCrLf
            sql += " ,d.DNAME, d.Years" & vbCrLf
            sql += " ,b.KName" & vbCrLf
            sql += " ,k.KeyName " & vbCrLf
            sql += " FROM KEY_BUSINESS b" & vbCrLf
            sql += " join key_depot d on d.DepID =b.DepID" & vbCrLf
            sql += " JOIN KEY_BUSINESSKEYS k on k.DepID=b.DepID and k.KID=b.KID" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and k.DepID ='" & DepID & "'" & vbCrLf
            sql += " and k.KID IN (" & KID & ")" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
            rst = ""
            For Each dr As DataRow In dt.Rows
                If rst <> "" Then rst += ","
                rst += Convert.ToString(dr("KeyName"))
            Next
        End If
        Return rst
    End Function

    '組合SQL語法
    Function ChgSqlWhere(ByVal KeyNames As String) As String
        Dim rst As String = ""
        Dim cst_dtNickName As String = "x"
        If KeyNames <> "" Then
            rst = ""
            If KeyNames.IndexOf(",") > -1 Then
                Dim aryStr() As String = Split(KeyNames, ",")
                For i As Integer = 0 To aryStr.Length - 1
                    rst += " OR " & cst_dtNickName & ".classcname like N'%" & aryStr(i) & "%'" & vbCrLf '換行
                Next
            Else
                rst += " OR " & cst_dtNickName & ".classcname like N'%" & KeyNames & "%'" & vbCrLf '換行
            End If
        End If
        Return rst
    End Function

    '匯出 SQL
    Private Sub Export1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export1.Click
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        '報表要用的預算來源參數
        Dim BudgetID As String = ""
        BudgetID = TIMS.GetCheckBoxListRptVal(BudgetList, 1)
        '報表要用的新興產業
        Dim sKID_6 As String = ""
        sKID_6 = TIMS.GetCheckBoxListRptVal(KID_6, 1)
        Dim sKID_10 As String = ""
        sKID_10 = TIMS.GetCheckBoxListRptVal(KID_10, 1)
        Dim sKID_4 As String = ""
        sKID_4 = TIMS.GetCheckBoxListRptVal(KID_4, 1)

        Dim KeyNames As String = ""
        KeyNames = ""
        If sKID_6 <> "" Then
            If KeyNames <> "" Then KeyNames += ","
            KeyNames += GetKeyName(cst_DepID06, sKID_6.Replace("\'", "'")) '取出關鍵字
        End If
        If sKID_10 <> "" Then
            If KeyNames <> "" Then KeyNames += ","
            KeyNames += GetKeyName(cst_DepID10, sKID_10.Replace("\'", "'")) '取出關鍵字
        End If
        If sKID_4 <> "" Then
            If KeyNames <> "" Then KeyNames += ","
            KeyNames += GetKeyName(cst_DepID04, sKID_4.Replace("\'", "'")) '取出關鍵字
        End If
        Dim sqlKeyNames As String = ""
        sqlKeyNames = ChgSqlWhere(KeyNames)

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&Years=" & Syear.SelectedValue '年度
        MyValue += "&STDate1=" & Me.STDate1.Text
        MyValue += "&STDate2=" & Me.STDate2.Text
        MyValue += "&FTDate1=" & Me.FTDate1.Text
        MyValue += "&FTDate2=" & Me.FTDate2.Text

        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & TPlanID1
        MyValue += "&BudgetID=" & BudgetID '預算來源

        '參考sd_15_012
        'Dim sFileName As String = ""
        'sFileName = "TR_05_018_1"
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", sFileName, MyValue)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select  ip.distid ,ip.planid "
        sql += " ,ip.distname " & vbCrLf '轄區	
        sql += " ,ip.planname" & vbCrLf '訓練計畫	" & vbCrLf
        sql += " ,oo.orgname " & vbCrLf '訓練機構名稱	" & vbCrLf
        sql += " ,cc.classcname " & vbCrLf '班別名稱	" & vbCrLf
        sql += " ,vt.trainName  " & vbCrLf '訓練職類	" & vbCrLf
        sql += " ,case when dbo.NVL(ip.PropertyID,0)=0 then '職前' " & vbCrLf
        sql += " 	when dbo.NVL(ip.PropertyID,0)=1 then '在職' end PropertyID" & vbCrLf '訓練性質	" & vbCrLf
        sql += " ,k1.hourRanName  hourRanName" & vbCrLf '訓練時段	" & vbCrLf
        sql += " ,CONVERT(varchar, cc.stdate, 111) STDate" & vbCrLf ' '開訓日期'" & vbCrLf
        sql += " ,CONVERT(varchar, cc.ftdate, 111) FTDate" & vbCrLf ' '結訓日期'" & vbCrLf
        sql += " ,cc.TNum " & vbCrLf '招生人數
        sql += " ,cc.THours " & vbCrLf '時數	
        sql += " ,dbo.NVL(gcs.openNum,0) openNum" & vbCrLf ' 開訓人數" & vbCrLf
        sql += " ,dbo.NVL(gcs.closeNum,0) closeNum" & vbCrLf '結訓人數" & vbCrLf
        sql += " ,dbo.NVL(gcs.jobNum,0) jobNum" & vbCrLf '就業人數" & vbCrLf
        'Sql += " " & vbCrLf
        sql += " from class_classinfo cc  " & vbCrLf
        sql += " join plan_planinfo pp   on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno " & vbCrLf
        sql += " join view_plan ip   on ip.planid =cc.planid" & vbCrLf
        sql += " join view_TrainType vt   on vt.TMID=cc.TMID" & vbCrLf
        sql += " join org_orginfo oo   on oo.comidno =cc.comidno" & vbCrLf
        sql += " left join Key_HourRan k1   on k1.HRID=cc.TPeriod" & vbCrLf
        sql += " left join (" & vbCrLf
        sql += " 	select cs.ocid " & vbCrLf
        sql += " 	,sum(case when cs.socid is not null then 1 end) openNum" & vbCrLf
        sql += " 	,sum(case when (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql += "    then 1 end) closeNum" & vbCrLf
        sql += " 	,sum(case when (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql += "   and sg3.IsGetJob='1' then 1 end) jobNum" & vbCrLf
        sql += " 	from Class_StudentsOfClass cs   " & vbCrLf
        sql += " 	join class_classinfo cc   on cc.ocid =cs.ocid" & vbCrLf
        sql += " 	join stud_studentinfo ss   on ss.SID =cs.SID" & vbCrLf
        sql += " 	join Stud_SubData ss2   on ss2.sid=ss.sid" & vbCrLf
        sql += " 	left join Stud_GetJobState3 sg3   on sg3.CPoint=1 and sg3.socid =cs.socid " & vbCrLf
        sql += " 	WHERE 1=1 " & vbCrLf
        sql += "    and cc.IsSuccess='Y' " & vbCrLf
        sql += "    and cc.NotOpen='N'" & vbCrLf
        'If Me.STDate1.Text <> "" Then
        '    Sql += " and cc.STDate >='" & Me.STDate1.Text & "'" & vbCrLf
        'End If
        'If Me.STDate2.Text <> "" Then
        '    Sql += " and cc.STDate <='" & Me.STDate2.Text & "'" & vbCrLf
        'End If
        'If Me.FTDate1.Text <> "" Then
        '    Sql += " and cc.FTDate >='" & Me.FTDate1.Text & "'" & vbCrLf
        'End If
        'If Me.FTDate2.Text <> "" Then
        '    Sql += " and cc.FTDate <='" & Me.FTDate2.Text & "'" & vbCrLf
        'End If
        If BudgetID <> "" Then
            sql += " and cs.BudgetID IN (" & BudgetID.Replace("\'", "'") & ")" & vbCrLf
        End If
        sql += " 	group by cs.ocid " & vbCrLf
        sql += " ) gcs on gcs.ocid =cc.ocid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.IsSuccess='Y' " & vbCrLf
        sql += " and cc.NotOpen='N'" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql += " and ip.Years='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql += " and cc.STDate >=convert(datetime, '" & Me.STDate1.Text & "', 111)" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql += " and cc.STDate <=convert(datetime, '" & Me.STDate2.Text & "', 111)" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql += " and cc.FTDate >=convert(datetime, '" & Me.FTDate1.Text & "', 111)" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql += " and cc.FTDate <=convert(datetime, '" & Me.FTDate2.Text & "', 111)" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql += " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql += " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If

        If sqlKeyNames <> "" Then
            'dt: x
            sql += " AND EXISTS (" & vbCrLf
            sql += "  select 'x' from class_classinfo x WHERE 1=1" & vbCrLf
            sql += "  AND x.ocid =cc.ocid " & vbCrLf
            sql += "  AND (1!=1" & vbCrLf
            sql += sqlKeyNames
            sql += " )" & vbCrLf
            sql += " )" & vbCrLf
        End If
        'Sql += " ORDER BY ip.distid , ip.planid ,cc.classcname" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        dt.DefaultView.Sort = "distid ,planid ,classcname"
        dt = TIMS.dv2dt(dt.DefaultView)

        Call ExpReport1(dt, KeyNames)

    End Sub

    '列印
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        '報表要用的預算來源參數
        Dim BudgetID As String = ""
        BudgetID = TIMS.GetCheckBoxListRptVal(BudgetList, 1)
        '報表要用的新興產業
        Dim sKID6 As String = ""
        sKID6 = TIMS.GetCheckBoxListRptVal2(KID_6, 1, cst_DepID06)
        Dim sKID10 As String = ""
        sKID10 = TIMS.GetCheckBoxListRptVal2(KID_10, 1, cst_DepID10)
        Dim sKID4 As String = ""
        sKID4 = TIMS.GetCheckBoxListRptVal2(KID_4, 1, cst_DepID04)
        Dim sDKID As String = "" 'DKID
        If sKID6 <> "" Then
            If sDKID <> "" Then sDKID += ","
            sDKID += sKID6
        End If
        If sKID10 <> "" Then
            If sDKID <> "" Then sDKID += ","
            sDKID += sKID10
        End If
        If sKID4 <> "" Then
            If sDKID <> "" Then sDKID += ","
            sDKID += sKID4
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&Years=" & Syear.SelectedValue '年度
        MyValue += "&STDate1=" & Me.STDate1.Text
        MyValue += "&STDate2=" & Me.STDate2.Text
        MyValue += "&FTDate1=" & Me.FTDate1.Text
        MyValue += "&FTDate2=" & Me.FTDate2.Text

        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & TPlanID1
        MyValue += "&BudgetID=" & BudgetID '預算來源
        MyValue += "&DKID=" & sDKID 'DKID 

        'If TIMS.sUtl_ChkTest() Then
        '    MyValue = "&Years=2012&STDate1=2012/07/01&STDate2=&FTDate1=2012/07/01&FTDate2=&DistID=\'000\',\'001\',\'002\',\'003\',\'004\',\'005\',\'006\'&TPlanID=\'01\',\'02\',\'03\',\'04\',\'05\',\'06\',\'07\',\'08\',\'09\',\'10\',\'11\',\'12\',\'13\',\'14\',\'15\',\'16\',\'17\',\'18\',\'19\',\'20\',\'21\',\'22\',\'23\',\'24\',\'25\',\'26\',\'27\',\'28\',\'29\',\'30\',\'31\',\'33\',\'34\',\'35\',\'36\',\'37\',\'38\',\'39\',\'40\',\'41\',\'42\',\'43\',\'44\',\'45\',\'46\',\'47\',\'48\',\'49\',\'50\',\'51\',\'52\',\'53\',\'54\',\'55\',\'56\',\'57\'&BudgetID=\'01\',\'02\',\'03\',\'04\',\'97\',\'98\',\'99\'&DKID=\'07,01\',\'07,02\',\'07,03\',\'07,04\',\'07,05\',\'07,06\'"
        'End If

        Dim sFileName As String = ""
        sFileName = "TR_05_018_1"
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", sFileName, MyValue)
    End Sub
End Class
