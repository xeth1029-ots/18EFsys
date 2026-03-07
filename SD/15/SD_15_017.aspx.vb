Public Class SD_15_017
    Inherits AuthBasePage

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            yearlist1 = TIMS.GetSyear(yearlist1)
            Common.SetListItem(yearlist1, sm.UserInfo.Years)
            yearlist2 = TIMS.GetSyear(yearlist2)
            Common.SetListItem(yearlist2, sm.UserInfo.Years)

            Distid = TIMS.Get_DistID(Distid)
            Distid.Items.Insert(0, New ListItem("全部", 0))
            Distid.Attributes("onclick") = "SelectAll('Distid','DistHidden');"

            '2013年(Old)
            GovClassName1 = TIMS.Get_GovClass(GovClassName1, 1, objConn) '訓練業別
            GovClassName1.Attributes("onclick") = "SelectAll('GovClassName1','HidGovClass1');"
            '2015年(New)
            GovClassName2 = TIMS.Get_GovClass(GovClassName2, 2, objConn) '訓練業別2
            GovClassName2.Attributes("onclick") = "SelectAll('GovClassName2','HidGovClass2');"

            cblDepot12 = TIMS.Get_KeyBusiness(cblDepot12, "12", objConn) '課程分類
            cblDepot12.Attributes("onclick") = "SelectAll('cblDepot12','HidcblDepot12');"

            Distid.Enabled = True
            If sm.UserInfo.DistID <> "000" Then '若登入者非署(局)署，鎖定轄區
                Common.SetListItem(Distid, sm.UserInfo.DistID)
                Distid.Enabled = False
            End If
        End If

    End Sub

    '匯出
    Protected Sub BtnExp_Click(sender As Object, e As EventArgs) Handles BtnExp.Click
        Dim okFlag As Boolean = False
        'Dim conn As SqlConnection
        okFlag = False '結束狀況有誤
        Try
            Call TIMS.OpenDbConn(objConn)
            'Dim da As New SqlDataAdapter
            'da.SelectCommand = New SqlCommand
            'da.SelectCommand.Connection = objConn
            'da.SelectCommand.CommandTimeout = 100
            Call ExpRpt() '匯出SUB'SQL
            okFlag = True '結束狀況無誤

            Call TIMS.CloseDbConn(objConn)

        Catch ex As Exception
            'If conn.State = ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            Exit Sub
        End Try

        '結束狀況無誤
        If okFlag Then TIMS.Utl_RespWriteEnd(Me, objConn, "") ' Call TIMS.CloseDbConn(objConn) ' Response.End()
    End Sub

    '匯出SUB (SQL)
    Private Sub ExpRpt()
        Dim sYearlist1 As String = TIMS.ClearSQM(yearlist1.SelectedValue)
        Dim sYearlist2 As String = TIMS.ClearSQM(yearlist2.SelectedValue)
        If sYearlist1 = "" Then
            sYearlist1 = TIMS.ClearSQM(sm.UserInfo.Years)
        End If
        If sYearlist2 = "" Then
            sYearlist2 = TIMS.ClearSQM(sm.UserInfo.Years)
        End If
        '置換錯誤問題
        If sYearlist1 > sYearlist2 Then
            Dim TMPVal As String = sYearlist1
            sYearlist1 = sYearlist2
            sYearlist2 = TMPVal
            Common.SetListItem(yearlist1, sYearlist1)
            Common.SetListItem(yearlist2, sYearlist2)
        End If

        '轄區
        Dim sDistID As String = ""
        If sm.UserInfo.DistID = "000" Then
            For i As Integer = 0 To Distid.Items.Count - 1
                If Distid.Items.Item(i).Selected = True AndAlso Distid.Items.Item(i).Value <> "" Then
                    If Distid.Items.Item(i).Text <> "全部" Then
                        If sDistID <> "" Then sDistID += ","
                        sDistID += "'" & TIMS.ClearSQM(Distid.Items.Item(i).Value) & "'"
                    End If
                End If
            Next
        End If
        If sDistID = "" AndAlso sm.UserInfo.DistID <> "000" Then
            sDistID = "'" & TIMS.ClearSQM(sm.UserInfo.DistID) & "'"
        End If

        '訓練業別1
        Dim sGovClass1 As String = ""
        For i As Integer = 0 To GovClassName1.Items.Count - 1
            If GovClassName1.Items.Item(i).Selected = True AndAlso GovClassName1.Items.Item(i).Value <> "" Then
                If GovClassName1.Items.Item(i).Text <> "全部" Then
                    If sGovClass1 <> "" Then sGovClass1 &= ","
                    sGovClass1 += "'" & TIMS.ClearSQM(GovClassName1.Items.Item(i).Value) & "'"
                End If
            End If
        Next

        '訓練業別2
        Dim sGovClass2 As String = ""
        For i As Integer = 0 To GovClassName2.Items.Count - 1
            If GovClassName2.Items.Item(i).Selected = True AndAlso GovClassName2.Items.Item(i).Value <> "" Then
                If GovClassName2.Items.Item(i).Text <> "全部" Then
                    If sGovClass2 <> "" Then sGovClass2 &= ","
                    sGovClass2 += "'" & TIMS.ClearSQM(GovClassName2.Items.Item(i).Value) & "'"
                End If
            End If
        Next

        '訓練課程分類
        Dim sCblDepot12 As String = ""
        For i As Integer = 0 To cblDepot12.Items.Count - 1
            If cblDepot12.Items.Item(i).Selected = True AndAlso cblDepot12.Items.Item(i).Value <> "" Then
                If cblDepot12.Items.Item(i).Text <> "全部" Then
                    If sCblDepot12 <> "" Then sCblDepot12 &= ","
                    sCblDepot12 += "'" & TIMS.ClearSQM(cblDepot12.Items.Item(i).Value) & "'"
                End If
            End If
        Next
        '是否核定
        Dim sIsSuccess As String = TIMS.ClearSQM(rblIsSuccess.SelectedValue)
        '品名
        txtCName.Text = TIMS.ClearSQM(txtCName.Text)

        '序號'年度'分署'訓練單位'課程名稱'訓練業別'課程分類'品名'規格'單位'單價'用途說明
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ip.Years 年度  " & vbCrLf
        sql &= " ,ip.distname 分署" & vbCrLf
        sql &= " ,oo.orgname 訓練單位" & vbCrLf
        sql &= " ,dbo.NVL(cc.classcname,pp.classname) 課程名稱" & vbCrLf
        sql &= " ,tt.jobname 訓練業別" & vbCrLf
        sql &= " ,vd12.KNAME 課程分類" & vbCrLf
        sql &= " ,a.CNAME 品名  " & vbCrLf
        sql &= " ,a.Standard 規格  " & vbCrLf
        sql &= " ,a.Unit 單位" & vbCrLf
        sql &= " ,a.Price 單價" & vbCrLf
        sql &= " ,a.PurPose 用途說明" & vbCrLf
        sql &= " FROM PLAN_PERSONCOST a" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.planid=a.planid and pp.comidno=a.comidno and pp.seqno=a.seqno" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=pp.planid " & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =pp.comidno" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO cc on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= " LEFT JOIN VIEW_TRAINTYPE tt on tt.tmid =pp.tmid " & vbCrLf
        sql &= " LEFT JOIN VIEW_GOVCLASSCAST ig on pp.GCID = ig.GCID" & vbCrLf
        sql &= " LEFT JOIN V_GOVCLASSCAST2 ig2 on pp.GCID2 = ig2.GCID2 " & vbCrLf
        sql &= " LEFT JOIN VIEW_DEPOT12 vd12 on vd12.GCID2 = ig2.GCID2 " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and ip.TPLANID ='28'" & vbCrLf
        sql &= " AND PP.ISAPPRPAPER='Y'" & vbCrLf
        Select Case sIsSuccess
            Case "Y"
                sql &= " AND cc.ISSUCCESS='Y'" & vbCrLf
                sql &= " AND cc.NotOpen='N'" & vbCrLf
            Case "N"
                sql &= " AND cc.OCID IS NULL" & vbCrLf
        End Select
        If sYearlist1 <> "" Then
            sql &= " and ip.years >='" & sYearlist1 & "'" & vbCrLf
        End If
        If sYearlist2 <> "" Then
            sql &= " and ip.years <='" & sYearlist2 & "'" & vbCrLf
        End If
        If sDistID <> "" Then
            sql &= " and ip.distid in (" & sDistID & ") " & vbCrLf
        End If
        If sGovClass1 <> "" Then
            sql &= " and (ig.GOVCLASS+','+ig.GCODE1) IN (" & sGovClass1 & ")" & vbCrLf
        End If
        If sGovClass2 <> "" Then
            sql &= " and ig2.GCODE1 in (" & sGovClass2 & ") " & vbCrLf
        End If
        If sCblDepot12 <> "" Then
            sql &= " and vd12.KID in (" & sCblDepot12 & ") " & vbCrLf
        End If
        If txtCName.Text <> "" Then
            Dim sCompareMode1 As String = TIMS.ClearSQM(rblCompareMode1.SelectedValue)
            Select Case sCompareMode1
                Case "1" '1:模糊比對 2:完整比對
                    sql &= " and a.CNAME LIKE '%" & txtCName.Text & "%' " & vbCrLf
                Case Else
                    sql &= " and a.CNAME = '" & txtCName.Text & "' " & vbCrLf
            End Select
        End If
        'sql += " ORDER BY ip.Years, ip.DistName ,a.CName" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objConn)


        sql = "" & vbCrLf
        sql &= " SELECT ip.Years 年度  " & vbCrLf
        sql &= " ,ip.distname 分署" & vbCrLf
        sql &= " ,oo.orgname 訓練單位" & vbCrLf
        sql &= " ,dbo.NVL(cc.classcname,pp.classname) 課程名稱" & vbCrLf
        sql &= " ,tt.jobname 訓練業別" & vbCrLf
        sql &= " ,vd12.KNAME 課程分類" & vbCrLf
        sql &= " ,a.CNAME 品名  " & vbCrLf
        sql &= " ,a.Standard 規格  " & vbCrLf
        sql &= " ,a.Unit 單位" & vbCrLf
        sql &= " ,a.Price 單價" & vbCrLf
        sql &= " ,a.PurPose 用途說明" & vbCrLf
        sql &= " FROM PLAN_COMMONCOST a" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.planid=a.planid and pp.comidno=a.comidno and pp.seqno=a.seqno" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=pp.planid " & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =pp.comidno" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO cc on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= " LEFT JOIN VIEW_TRAINTYPE tt on tt.tmid =pp.tmid " & vbCrLf
        sql &= " LEFT JOIN VIEW_GOVCLASSCAST ig on pp.GCID = ig.GCID" & vbCrLf
        sql &= " LEFT JOIN V_GOVCLASSCAST2 ig2 on pp.GCID2 = ig2.GCID2 " & vbCrLf
        sql &= " LEFT JOIN VIEW_DEPOT12 vd12 on vd12.GCID2 = ig2.GCID2 " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and ip.TPLANID ='28'" & vbCrLf
        sql &= " AND PP.ISAPPRPAPER='Y'" & vbCrLf
        Select Case sIsSuccess '是否核定
            Case "Y"
                sql &= " AND cc.ISSUCCESS='Y'" & vbCrLf
                sql &= " AND cc.NotOpen='N'" & vbCrLf
            Case "N"
                sql &= " AND cc.OCID IS NULL" & vbCrLf
        End Select
        If sYearlist1 <> "" Then
            sql &= " and ip.years >='" & sYearlist1 & "'" & vbCrLf
        End If
        If sYearlist2 <> "" Then
            sql &= " and ip.years <='" & sYearlist2 & "'" & vbCrLf
        End If
        If sDistID <> "" Then
            sql &= " and ip.distid in (" & sDistID & ") " & vbCrLf
        End If
        'If sGovClassName <> "" Then
        '    sql &= " and ig2.GCODE1 in (" & sGovClassName & ") " & vbCrLf
        'End If
        If sGovClass1 <> "" Then
            sql &= " and (ig.GOVCLASS+','+ig.GCODE1) IN (" & sGovClass1 & ")" & vbCrLf
        End If
        If sGovClass2 <> "" Then
            sql &= " and ig2.GCODE1 in (" & sGovClass2 & ") " & vbCrLf
        End If
        If sCblDepot12 <> "" Then
            sql &= " and vd12.KID in (" & sCblDepot12 & ") " & vbCrLf
        End If
        '品名
        'txtCName.Text = TIMS.ClearSQM(txtCName.Text)
        If txtCName.Text <> "" Then
            Dim sCompareMode1 As String = TIMS.ClearSQM(rblCompareMode1.SelectedValue)
            Select Case sCompareMode1
                Case "1" '1:模糊比對 2:完整比對
                    sql &= " and a.CNAME LIKE '%" & txtCName.Text & "%' " & vbCrLf
                Case Else
                    sql &= " and a.CNAME = '" & txtCName.Text & "' " & vbCrLf
            End Select
        End If
        'sql += " ORDER BY ip.Years, ip.DistName ,a.CName" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objConn)

        TIMS.OpenDbConn(objConn)
        Dim dt As New DataTable
        With sCmd
            .Connection = objConn
            .CommandTimeout = 100
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        With sCmd2
            .Connection = objConn
            .CommandTimeout = 100
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        dt.DefaultView.Sort = "年度,分署,品名"
        dt = TIMS.dv2dt(dt.DefaultView)

        Const cst_StrTitle1 As String = "項目金額查詢"
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(cst_StrTitle1, System.Text.Encoding.UTF8) & ".xls")
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
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")

        Dim ExportStr As String = ""
        ExportStr = ""

        ExportStr &= "<tr>"
        ExportStr &= "<td>序號</td>"
        For iii As Integer = 0 To dt.Columns.Count - 1
            ExportStr &= "<td>" & Convert.ToString(dt.Columns(iii).ColumnName) & "</td>"
        Next
        ExportStr &= "</tr>"

        Dim iMinPrice As Double = -1 '初始值小於1
        Dim iMaxPrice As Double = 0
        Dim iAllPrice As Double = 0
        Dim iRow As Integer = 0
        For Each dr As DataRow In dt.Rows
            iAllPrice += Val(dr("單價"))
            If iMinPrice = -1 Then
                '初始值時 直接塞一個數字
                iMinPrice = Val(dr("單價"))
                iMaxPrice = Val(dr("單價"))
            Else
                If iMinPrice > Val(dr("單價")) Then iMinPrice = Val(dr("單價"))
                If iMaxPrice < Val(dr("單價")) Then iMaxPrice = Val(dr("單價"))
            End If

            iRow += 1
            ExportStr &= "<tr>"
            ExportStr &= "<td class=""noDecFormat"">" & CStr(iRow) & "</td>"
            For iii As Integer = 0 To dt.Columns.Count - 1
                Select Case dt.Columns(iii).ColumnName
                    Case "單價"
                        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr(dt.Columns(iii).ColumnName)) & "</td>"
                    Case "年度"
                        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr(dt.Columns(iii).ColumnName)) & "</td>"
                    Case Else
                        ExportStr &= "<td>" & Convert.ToString(dr(dt.Columns(iii).ColumnName)) & "</td>"
                End Select
            Next
            ExportStr &= "</tr>"
        Next

        ExportStr &= "<tr>"
        Dim tmpStr1 As String = ""
        tmpStr1 = ""
        tmpStr1 &= "&nbsp;&nbsp;　　　　　　　最高價：" & TIMS.ROUND(iMaxPrice, 2)
        tmpStr1 &= "&nbsp;&nbsp;　　　　　　　最低價：" & TIMS.ROUND(iMinPrice, 2)
        tmpStr1 &= "&nbsp;&nbsp;　　　　　　　平均價：" & TIMS.ROUND(iAllPrice / iRow, 2)
        tmpStr1 &= "&nbsp;&nbsp;　　　　　　　"
        Dim iColSpan As Integer = dt.Columns.Count
        ExportStr &= "<td class=""noDecFormat"">　</td>"
        ExportStr &= "<td colSpan =""" & iColSpan & """>" & tmpStr1 & "</td>"
        ExportStr &= "</tr>"

        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
    End Sub


End Class