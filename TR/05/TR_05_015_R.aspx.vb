Partial Class TR_05_015_R
    Inherits AuthBasePage

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
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False

            CreateItem()

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

            ''選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

            btnSearch.Attributes("onclick") = "return search();"
            btnExport.Attributes("onclick") = "return search();"
        End If

    End Sub

    '關鍵字詞建立
    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)
        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Remove(DistID.Items.FindByValue(""))
        DistID.Items.Insert(0, New ListItem("全部", ""))
        '計畫別
        'TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", "TPlanID not in ('28','15','36','54')")
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Me.Syear.SelectedValue = "" Then
            Errmsg += "請選擇結訓年度" & vbCrLf
        End If

        Dim j As Integer = 0
        Dim CBLobj As CheckBoxList
        j = 0
        CBLobj = DistID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇轄區" & vbCrLf
        'j = 0
        'CBLobj = CTID
        'For i As Integer = 1 To CBLobj.Items.Count - 1
        '    Dim objitem As ListItem = CBLobj.Items(i)
        '    If objitem.Selected = True Then
        '        j += 1
        '        Exit For
        '    End If
        'Next
        'If j = 0 Then Errmsg += "請選擇縣市" & vbCrLf
        j = 0
        CBLobj = TPlanID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇訓練計畫" & vbCrLf
        'j = 0
        'CBLobj = cbl_JOBMDATE_MM
        'For i As Integer = 0 To CBLobj.Items.Count - 1
        '    Dim objitem As ListItem = CBLobj.Items(i)
        '    If objitem.Selected = True Then
        '        j += 1
        '        Exit For
        '    End If
        'Next
        'If j = 0 Then Errmsg += "請選擇就業區間" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Function Get_IdentityDt() As DataTable
        Dim sql As String = ""

        sql = "" & vbCrLf
        sql &= " select kd.IdentityID" & vbCrLf
        sql &= " ,kd.Name kdName " ' --'參訓身分別'" & vbCrLf
        sql &= " ,0 as openNum" & vbCrLf
        sql &= " ,0 as closeNum" & vbCrLf
        'Sql &= " ,0 as jobNum" & vbCrLf
        sql &= " FROM KEY_IDENTITY kd" & vbCrLf
        'Sql &= " WHERE 1=1 AND kd.IdentityID IN (" & TIMS.Cst_Identity06_2019_11 & ") " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql &= "   AND kd.IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")" & vbCrLf
        Else
            sql &= "   AND kd.IDENTITYID IN (" & TIMS.Cst_Identity28_2019_11 & ")" & vbCrLf
        End If
        sql &= " ORDER BY  kd.IdentityID" & vbCrLf
        Dim dtIDENT As DataTable
        dtIDENT = DbAccess.GetDataTable(sql, objconn)
        Return dtIDENT
    End Function

    '統計 SQL 查詢
    Sub Search1()
        Dim parms As Hashtable = New Hashtable()
        Dim tmpStr As String = ""
        Dim itemDist As String = "" '轄區
        'Dim itemCTID As String = "" '縣市
        Dim itemTPlanID As String = "" '計畫
        'Dim itemJOBMDATE As String = "" '就業區間

        Dim CBLobj As CheckBoxList
        tmpStr = ""
        CBLobj = DistID
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "'" & objitem.Value & "'"
            End If
        Next
        If tmpStr <> "" Then itemDist = tmpStr

        'tmpStr = ""
        'CBLobj = CTID
        'For Each objitem As ListItem In CBLobj.Items
        '    If objitem.Selected AndAlso objitem.Value <> "" Then
        '        If tmpStr <> "" Then tmpStr += ","
        '        tmpStr += "'" & objitem.Value & "'"
        '    End If
        'Next
        'If tmpStr <> "" Then itemCTID = tmpStr

        tmpStr = ""
        CBLobj = TPlanID
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "'" & objitem.Value & "'"
            End If
        Next
        If tmpStr <> "" Then itemTPlanID = tmpStr

        '取得身份別
        Dim dtIDENT As DataTable = Get_IdentityDt()

        Dim sql As String = ""

        parms.Clear()
        sql = "" & vbCrLf
        sql &= " select 1 openNum" & vbCrLf
        sql &= " , case when cs.StudStatus not in (2,3) and cc.ftdate<=getdate() then 1 else 0 end  closeNum" & vbCrLf
        'Sql &= " , case when (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()) " & vbCrLf
        'Sql &= " 	and j3.socid is not null then 1 else 0 end  jobNum" & vbCrLf
        For i As Integer = 0 To dtIDENT.Rows.Count - 1
            Dim vIdd As String = Convert.ToString(dtIDENT.Rows(i)("IdentityID"))
            sql &= " , case when cs.Identityid like '%" & vIdd & "%' then 1 else 0 end as  Idd" & vIdd & vbCrLf
        Next

        sql &= "  FROM VIEW_PLAN ip" & vbCrLf
        sql &= "  JOIN CLASS_CLASSINFO cc on cc.planid =ip.planid" & vbCrLf
        sql &= "  JOIN PLAN_PLANINFO pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= "  join AUTH_RELSHIP aa on aa.RID=cc.RID" & vbCrLf
        sql &= "  JOIN VIEW_TRAINTYPE tt on tt.tmid=cc.tmid" & vbCrLf
        sql &= "  JOIN ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= "  JOIN VIEW_ZIPNAME vz on vz.ZipCode=cc.TaddressZip" & vbCrLf
        sql &= "  JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid =cc.ocid " & vbCrLf
        'Sql &= " 	--and cs.StudStatus not in (2,3) and cs.closedate<=getdate() and cc.ftdate<=getdate() " & vbCrLf
        sql &= "  JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        'Sql &= "  left join Stud_GetJobState3 j3 ON cs.SOCID=j3.SOCID and j3.CPoint=1 and j3.IsGetJob=1" & vbCrLf
        'Sql &= "  JOIN KEY_IDENTITY kd on kd.IdentityID =cs.mIdentityid and kd.IdentityID IN (" & TIMS.Cst_Identity06_2019_11 & ") " & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        sql &= "  and ip.TPlanID NOT IN ('28','15','36','54') " & vbCrLf
        'Sql &= "  and cc.TPropertyID='0'--0職前 " & vbCrLf
        sql &= "  and cc.TPropertyID='1'--1在職 " & vbCrLf
        sql &= "  and cc.NotOpen='N' -- '排除不開班" & vbCrLf
        sql &= "  and cc.IsSuccess='Y'-- '轉入成功資料" & vbCrLf
        sql &= "  and cc.Evta_NoShow is null " & vbCrLf
        '' Sql &= " -- and ip.years ='2012'" & vbCrLf
        'Sql &= "  and ip.distid IN ('001')" & vbCrLf
        'Sql &= "  and ip.TPlanID IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','16','17','18','19','20','21','22','23','24','25','26','27','29','30','31','33','34','35','37','38','39','40','41','42','43','44','45','46','47','48','49','50','51','52','53','55','56','57')" & vbCrLf
        '年度  
        If Me.Syear.SelectedValue <> "" Then
            sql &= " and ip.years=@years " & vbCrLf
            parms.Add("years", Me.Syear.SelectedValue)
        End If
        '轄區   
        If itemDist <> "" Then
            sql &= " and ip.distid IN (" & itemDist & ")" & vbCrLf
        End If
        '訓練計畫 
        If itemTPlanID <> "" Then
            sql &= " and ip.TPlanID IN (" & itemTPlanID & ")" & vbCrLf
        End If
        '開訓期間
        If Me.STDate1.Text <> "" Then
            'Sql &= " and cc.STDate >=convert(datetime, '" & Me.STDate1.Text & "', 111)" & vbCrLf
            sql &= " and cc.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", Me.STDate1.Text)
        End If
        If Me.STDate2.Text <> "" Then
            'Sql &= " and cc.STDate <=convert(datetime, '" & Me.STDate2.Text & "', 111)" & vbCrLf
            sql &= " and cc.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", Me.STDate2.Text)
        End If

        '結訓期間 
        If Me.FTDate1.Text <> "" Then
            'Sql &= " and cc.FTDate >=convert(datetime, '" & Me.FTDate1.Text & "', 111)" & vbCrLf
            sql &= " and cc.FTDate >= @FTDate1 " & vbCrLf
            parms.Add("FTDate1", Me.FTDate1.Text)
        End If
        If Me.FTDate2.Text <> "" Then
            'Sql &= " and cc.FTDate <=convert(datetime, '" & Me.FTDate2.Text & "', 111)" & vbCrLf
            sql &= " and cc.FTDate <= @FTDate2 " & vbCrLf
            parms.Add("FTDate2", Me.FTDate2.Text)
        End If

        Dim dtTmp As DataTable
        Try
            dtTmp = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sql))
            'Common.RespWrite(Me, ex.ToString)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try

        '混合資料,配合sql語法之欄位，請小心更改
        dtIDENT = Subx1(dtIDENT, dtTmp)

        'Table4.Style("display") = "inline"
        'Print.Visible = False
        'btnExport1.Visible = False
        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dtIDENT.Rows.Count > 0 Then
            'dt.DefaultView
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True

            PageControler1.PageDataTable = dtIDENT
            'PageControler1.Sort = "DistID "
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    '混合資料,配合sql語法之欄位，請小心更改
    Function Subx1(ByVal dt As DataTable, ByVal dtTmp As DataTable) As DataTable
        Dim Rst As New DataTable
        Const cst_0 As String = "參訓身分別"
        Const cst_1 As String = "開訓人數"
        Const cst_2 As String = "結訓人數"
        'Const cst_3 As String = "就業人數"
        Rst.Columns.Add(New DataColumn(cst_0))
        Rst.Columns.Add(New DataColumn(cst_1))
        Rst.Columns.Add(New DataColumn(cst_2))
        'Rst.Columns.Add(New DataColumn(cst_3))

        For Each dr As DataRow In dt.Rows
            Dim dr1 As DataRow
            dr1 = Rst.NewRow
            dr1(cst_0) = dr("kdName")
            For i As Integer = 1 To 2
                Dim filter1 As String = ""
                Dim colName As String = ""
                Select Case i
                    Case 1
                        filter1 = "openNum"
                        colName = cst_1
                    Case 2
                        filter1 = "closeNum"
                        colName = cst_2
                        'Case 3
                        '    filter1 = "jobNum"
                        '    colName = cst_3
                End Select
                If dtTmp.Select("Idd" & dr("IdentityID") & "=1 and " & filter1 & " =1").Length > 0 Then
                    dr(filter1) = CInt(dr(filter1)) + dtTmp.Select("Idd" & dr("IdentityID") & "=1 and " & filter1 & " =1").Length
                End If
                dr1(colName) = dr(filter1)
            Next
            Rst.Rows.Add(dr1)
        Next
        'dt.AcceptChanges()
        'Return dt
        Rst.AcceptChanges()
        Return Rst
    End Function

    '匯出 明細資料
    Sub sExport1()

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        If DataGrid1.Visible = False OrElse msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        'Const Cst_xlsFileName As String = "參訓身分別.xls"
        'Dim sFileName As String = ""
        ''勞保勾稽查詢
        'sFileName = HttpUtility.UrlEncode(Cst_xlsFileName, System.Text.Encoding.UTF8)

        Response.Clear()
        Dim sFileName1 As String = "參訓身分別"
        Dim strSTYLE As String = ""

        ''套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        'strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        DataGrid1.AllowPaging = True '開啟分頁
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '查詢 明細資料
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        '查詢
        Call Search1()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Call sExport1()
    End Sub

End Class
