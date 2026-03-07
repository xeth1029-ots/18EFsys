Partial Class SYS_03_013
    Inherits AuthBasePage

    Dim str_SortAry As String
    Dim SortAry() As String
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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            create1()
            '預設年度
            Common.SetListItem(yearlist, sm.UserInfo.Years)
        End If
        Button1.Attributes("onclick") = "javascript:return search();"
        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"
        ''選擇全部匯出欄位
        'Sort.Attributes("onclick") = "SelectAll('Sort','SortHidden');"
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        If Get_value(Errmsg) Then
            search_value()
        Else
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
    End Sub

    Sub create1()
        Dim dt As DataTable
        Dim sqlstr As String
        yearlist = TIMS.GetSyear(yearlist, 0, 0, False)
        sqlstr = "SELECT Name,DistID FROM ID_District order by DistID"
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Me.DistrictList.DataSource = dt
        Me.DistrictList.DataTextField = "Name"
        Me.DistrictList.DataValueField = "DistID"
        Me.DistrictList.DataBind()
        Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        sqlstr = "select PlanName,TPlanID from Key_Plan order by TPlanID"
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Me.PlanList.DataSource = dt
        Me.PlanList.DataTextField = "PlanName"
        Me.PlanList.DataValueField = "TPlanID"
        Me.PlanList.DataBind()
        Me.PlanList.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    Function Get_value(ByRef Errmsg As String) As Boolean
        Errmsg = ""
        Get_value = True
        Dim int_ok_flag As Integer = 0

        STDate1.Text = STDate1.Text.Trim
        STDate2.Text = STDate2.Text.Trim
        FTDate1.Text = FTDate1.Text.Trim
        FTDate2.Text = FTDate2.Text.Trim

        Dim TPlanID1 As String = ""
        Dim DistID1 As String = ""
        int_ok_flag = 0
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                int_ok_flag += 1
                If DistID1 = "" Then
                    DistID1 = Convert.ToString("'" & Me.DistrictList.Items(i).Value & "'")
                Else
                    DistID1 += "," & Convert.ToString("'" & Me.DistrictList.Items(i).Value & "'")
                End If
            End If
        Next
        If int_ok_flag < 1 Then
            Errmsg += "請選擇轄區 至少1項" & vbCrLf
            Get_value = False
        End If

        int_ok_flag = 0
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected Then
                int_ok_flag += 1
                If TPlanID1 = "" Then
                    TPlanID1 = Convert.ToString("'" & Me.PlanList.Items(i).Value & "'")
                Else
                    TPlanID1 += "," & Convert.ToString("'" & Me.PlanList.Items(i).Value & "'")
                End If
            End If
        Next
        If int_ok_flag < 1 Then
            Errmsg += "請選擇訓練計畫 至少1項" & vbCrLf
            Get_value = False
        End If

        '訓練計畫限定數取消 by AMU 20091006
        'If int_ok_flag > 5 Then
        '    Errmsg += "訓練計畫限定5項" & vbCrLf
        '    Get_value = False
        'End If

        Me.ViewState("DistID1") = DistID1
        Me.ViewState("TPlanID1") = TPlanID1
    End Function


    Sub search_value()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim ExportStr As String = ""

        Dim date_str As String = "" '檔案名稱使用。
        date_str = Replace(Common.FormatDate(Now), "/", "")

        sql = "" & vbCrLf
        sql += " select" & vbCrLf
        sql += " idt.Name as 轄區" & vbCrLf
        sql += " , kp.PlanName as 計畫名稱" & vbCrLf
        sql += " , oo.OrgName as 訓練機構" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班級名稱" & vbCrLf
        sql += " ,CONVERT(varchar, cc.STDate, 111) as 開訓日" & vbCrLf
        sql += " ,CONVERT(varchar, cc.FTDate, 111) as 結訓日" & vbCrLf
        sql += " ,ss.Name as 姓名" & vbCrLf
        sql += " ,sd.Email as 電子郵件" & vbCrLf
        sql += " ,''''+sd.PhoneD as 日間電話" & vbCrLf
        sql += " ,''''+sd.PhoneN as 夜間電話" & vbCrLf
        sql += " ,''''+sd.CellPhone as 手機" & vbCrLf
        sql += " from id_plan ip" & vbCrLf
        sql += " join key_plan kp on kp.TPlanID =ip.TPlanID" & vbCrLf
        sql += " join ID_District idt on idt.DistID =ip.DistID" & vbCrLf
        sql += " Join Plan_PlanINFO pp on pp.PlanID=ip.PlanID" & vbCrLf
        sql += " join Class_ClassInfo cc on  cc.PlanID=pp.PlanID AND cc.ComIDNO=pp.ComIDNO AND cc.Seqno=pp.Seqno" & vbCrLf
        sql += " JOIN Org_OrgInfo oo on oo.ComIDNO=pp.ComIDNO" & vbCrLf
        sql += " JOIN Class_StudentsOfClass cs on cs.OCID =cc.OCID" & vbCrLf
        sql += " join Stud_StudentInfo ss on ss.SID=cs.SID" & vbCrLf
        sql += " join Stud_SubData sd on sd.SID=cs.SID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf

        '排除離退訓學員
        sql += "   and cs.StudStatus NOT IN (2,3) " & vbCrLf
        '排除不開班 '轉入成功
        sql += "   and cc.NotOpen='N' and cc.IsSuccess='Y' " & vbCrLf

        If Me.ViewState("DistID1") <> "" Then
            sql += "  AND idt.DistID in (" & Me.ViewState("DistID1") & ") " & vbCrLf
        Else
            sql += "  AND idt.DistID='' " & vbCrLf '不勾選
        End If
        If Me.ViewState("TPlanID1") <> "" Then
            sql += "  AND kp.TPlanID in (" & Me.ViewState("TPlanID1") & ") " & vbCrLf
        Else
            sql += "  AND kp.TPlanID='' " & vbCrLf '不勾選
        End If
        If yearlist.SelectedValue <> "" Then
            sql += "  and ip.Years='" & yearlist.SelectedValue & "'" & vbCrLf
        End If
        If STDate1.Text <> "" Then
            sql += "   and cc.STDate >= " & TIMS.To_date(Stdate1.Text) & vbCrLf
        End If
        If Stdate2.Text <> "" Then
            sql += "   and cc.STDate <= " & TIMS.To_date(Stdate2.Text) & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            sql += "   and cc.FTDate >= " & TIMS.To_date(Ftdate1.Text) & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += "   and cc.FTDate <= " & TIMS.To_date(Ftdate2.Text) & vbCrLf
        End If

        sql += "   ORDER BY 1,2,3" & vbCrLf
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Me.ViewState("MESSAGE") = ""
            Me.ViewState("MESSAGE") += "資料庫效能異常，請重新查詢" & vbCrLf
            Me.ViewState("MESSAGE") += "若重複出現請縮小查詢範圍，再次查詢" & vbCrLf
            Common.MessageBox(Me, Me.ViewState("MESSAGE"))
            Exit Sub
        End Try

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "此條件下查無資料，請重新設定條件")
            Exit Sub
        End If

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("學員資料", System.Text.Encoding.UTF8) & date_str & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentType = "Application/vnd.ms-excel"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        ExportStr = ""
        ExportStr += "轄區" & vbTab
        ExportStr += "訓練機構" & vbTab
        ExportStr += "班級名稱" & vbTab
        ExportStr += "開訓日" & vbTab
        ExportStr += "結訓日" & vbTab
        ExportStr += "姓名" & vbTab
        ExportStr += "電子郵件" & vbTab
        ExportStr += "日間電話" & vbTab
        ExportStr += "夜間電話" & vbTab
        ExportStr += "手機" & vbTab
        ExportStr += vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        For Each dr As DataRow In dt.Rows
            ExportStr = ""
            ExportStr += dr("轄區") & vbTab
            ExportStr += dr("訓練機構") & vbTab
            ExportStr += dr("班級名稱") & vbTab
            ExportStr += dr("開訓日") & vbTab
            ExportStr += dr("結訓日") & vbTab
            ExportStr += dr("姓名") & vbTab
            ExportStr += dr("電子郵件") & vbTab
            ExportStr += dr("日間電話") & vbTab
            ExportStr += dr("夜間電話") & vbTab
            ExportStr += dr("手機") & vbTab
            ExportStr += vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Response.End()

    End Sub

End Class
