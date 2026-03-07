Partial Class CP_04_002_add
    Inherits AuthBasePage

    Dim blnPrintFlag As Boolean = False
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
        blnPrintFlag = False
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Me.ViewState("itemstr") = Session("itemstr")
            Me.ViewState("itemplan") = Session("itemplan")
            Me.ViewState("itemcity") = Session("itemcity")
            Me.ViewState("SSTDate") = Session("SSTDate")
            Me.ViewState("ESTDate") = Session("ESTDate")
            Me.ViewState("ConRID") = Session("ConRID")
            Session("itemstr") = Nothing
            Session("itemplan") = Nothing
            Session("itemcity") = Nothing
            Session("SSTDate") = Nothing
            Session("ESTDate") = Nothing
            Session("ConRID") = Nothing

            Dim sErrMsg As String = ""
            sErrMsg = ""
            If Not create(sErrMsg) Then
                Common.MessageBox(Me, sErrMsg)
            End If

        End If

        '回上一頁
        Me.Button2.Attributes.Add("onclick", "location.href='CP_04_002.aspx';return false;")
    End Sub

    Function create(ByRef sErrMsg As String) As Boolean
        Dim Rst As Boolean = True
        sErrMsg = ""

        Dim dt As DataTable
        Dim dr As DataRow
        Dim sqlstr As String
        Dim yearlist As String = Request("yearlist")
        Dim itemstr As String = Me.ViewState("itemstr")
        Dim itemplan As String = Me.ViewState("itemplan")
        Dim itemcity As String = Me.ViewState("itemcity")
        Dim SSTDate As String = Me.ViewState("SSTDate")
        Dim ESTDate As String = Me.ViewState("ESTDate")
        Dim STNum As Integer = 0
        Dim SumTotalCost As Long = 0 '整數
        Dim SumTotalCost2 As Long = 0 '整數

        sqlstr = " SELECT a.*" & vbCrLf
        sqlstr += " ,b.Name AS DistName,b.DistID" & vbCrLf
        sqlstr += " ,c.PlanName,d.OrgName,g.TrainName" & vbCrLf
        sqlstr += " ,dbo.NVL(tb.totalcost,0) TCost" & vbCrLf
        sqlstr += " ,dbo.NVL(h.TNum,0) TrainNum" & vbCrLf
        sqlstr += " from Plan_PlanInfo   a " & vbCrLf
        sqlstr += " JOIN Org_OrgInfo     d ON d.ComIDNO=a.ComIDNO " & vbCrLf
        sqlstr += " JOIN ID_Plan         e ON e.PlanID=a.PlanID " & vbCrLf
        sqlstr += " JOIN ID_District     b ON b.DistID=e.DistID " & vbCrLf
        sqlstr += " JOIN Key_Plan        c ON c.TPlanID=a.TPlanID " & vbCrLf
        sqlstr += " JOIN Key_TrainType   g ON g.TMID=a.TMID " & vbCrLf
        sqlstr += " left outer join ( " & vbCrLf
        sqlstr += "  select planid,comidno,seqno" & vbCrLf
        sqlstr += "  ,case when costmode in (1,2) then sum(oprice*itemage*itemcost)" & vbCrLf
        sqlstr += "  when costmode in (3,4)then sum(oprice*itemage) end totalcost" & vbCrLf
        sqlstr += "  from Plan_CostItem " & vbCrLf
        sqlstr += "  group by planid,comidno,seqno,costmode " & vbCrLf
        sqlstr += " ) tb on a.planid=tb.planid and a.comidno=tb.comidno and a.seqno=tb.seqno " & vbCrLf

        sqlstr += " LEFT JOIN Class_ClassInfo  h ON a.PlanID=h.PlanID and a.ComIDNO=h.ComIDNO and a.SeqNo=h.SeqNo " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = h.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=a.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=a.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        sqlstr += " where a.IsApprPaper='Y' " & vbCrLf

        '選擇年度
        If yearlist <> 0 Then
            sqlstr += " and a.PlanYear= '" & yearlist & "' " & vbCrLf
        End If

        '選擇轄區
        If itemstr <> "" Then
            sqlstr += " and b.DistID IN (" & itemstr & ") " & vbCrLf
        End If

        '選擇縣市
        If itemcity <> "" Then
            sqlstr += " and (1!=1" & vbCrLf
            sqlstr += "    OR iz.CTID IN (" & itemcity & ") " & vbCrLf
            sqlstr += "    OR iz1.CTID IN (" & itemcity & ") " & vbCrLf
            sqlstr += "    OR iz2.CTID IN (" & itemcity & ") " & vbCrLf
            sqlstr += " )" & vbCrLf
        End If

        '選擇訓練計畫
        If itemplan <> "" Then
            sqlstr += "and c.TPlanID IN (" & itemplan & ") " & vbCrLf
        End If

        '開訓日期起
        If SSTDate <> "" Then
            sqlstr += "and a.STDate >= " & TIMS.To_date(SSTDate) & vbCrLf
        End If

        '開訓日期迄
        If ESTDate <> "" Then
            sqlstr += "and a.STDate <= " & TIMS.To_date(ESTDate) & vbCrLf
        End If

        If Me.ViewState("ConRID") <> "" Then
            Dim Relship As String = ""
            Dim RelshipStr As String = "" '多筆含逗號
            For i As Integer = 0 To Split(Me.ViewState("ConRID"), ",").Length - 1
                Relship = DbAccess.ExecuteScalar("SELECT Relship FROM Auth_Relship WHERE RID ='" & Split(Me.ViewState("ConRID"), ",")(i) & "'", objconn)
                RelshipStr &= String.Concat(If(RelshipStr <> "", ",", ""), Relship)
            Next

            If Split(RelshipStr, ",").Length > 1 Then
                '多筆 Split(RelshipStr, ",")(i) 
                sqlstr += " and (1!=1"
                For i As Integer = 0 To Split(RelshipStr, ",").Length - 1
                    sqlstr += " or a.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Split(RelshipStr, ",")(i) & "%')" & vbCrLf
                Next
                sqlstr += " )"
            Else
                '單1筆 (RelshipStr)
                sqlstr += " and a.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & RelshipStr & "%')"
            End If
        End If

        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Dim RecordCount As Integer = dt.Rows.Count ''TIMS.Get_SQLRecordCount(sqlstr, objconn)
        Me.CountLabel.Text = RecordCount

        Me.NoData.Text = "<font color=red>查無資料</font>"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If RecordCount > 0 Then
            If ViewState("sort") = "" Then
                ViewState("sort") = "DistID"
            End If

            Me.NoData.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            'PageControler1.SqlDataCreate(sqlstr, "DistID,PlanID,STDate desc")
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "DistID,PlanID,STDate desc"
            PageControler1.ControlerLoad()
        End If

        ''筆數
        'Me.CountLabel.Text = dt.Rows.Count.ToString

        '訓練總人數
        For Each dr In dt.Rows
            If Not dr("TNum").ToString = "" Then
                STNum += dr("TNum")
            End If
        Next
        Me.STNum.Text = STNum.ToString

        SumTotalCost = 0
        SumTotalCost2 = 0
        '訓練總經費
        Try
            For Each dr In dt.Rows
                If Convert.ToString(dr("TCost")) <> "" AndAlso IsNumeric(Convert.ToString(dr("TCost"))) Then
                    SumTotalCost += Val(Convert.ToString(dr("TCost")))
                    SumTotalCost2 += AdmFee(dr("planid"), dr("comIDNO"), dr("seqNo"))
                End If
            Next
        Catch ex As Exception
            sErrMsg += ex.ToString()
            Rst = False
        End Try

        Me.SumTotalCost.Text = Val(SumTotalCost + SumTotalCost2)
        Me.SumTotalCost.ToolTip = ""
        Me.SumTotalCost.ToolTip += "原總訓練費用：" & SumTotalCost.ToString & vbCrLf
        Me.SumTotalCost.ToolTip += "總行政管理費：" & SumTotalCost2.ToString


        '顯示所選年度
        Me.Year.Visible = True
        Me.YearLabel.Visible = True
        Me.YearLabel.Text = yearlist

        '顯示所選轄區
        Call area()

        Return Rst
    End Function

    Sub area()
        Dim dt As DataTable
        Dim dr As DataRow
        Dim itemstr As String = Me.ViewState("itemstr")
        Dim sqlstr As String = "SELECT NAME,DISTID FROM ID_DISTRICT"
        '選擇轄區
        If itemstr <> "" Then sqlstr += " where DistID IN (" & itemstr & ")"

        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Dim sTMP As String = ""
        For Each dr In dt.Rows
            sTMP &= String.Concat(If(sTMP <> "", ",", ""), dr("Name"))
        Next
        Me.DistrictLabel.Text = sTMP 's.Substring(0, s.Length - 1)
        Me.District.Visible = True
        Me.DistrictLabel.Visible = True
    End Sub

    Private Sub DataGrid1_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '審核狀態 和 班名連結
        Dim yearlist As String = Request("yearlist")
        Select Case e.Item.ItemType
            Case ListItemType.Header

                If Not blnPrintFlag Then
                    '排序功能
                    If Me.ViewState("sort") <> "" Then
                        'Dim mylabel As String
                        Dim mysort As New System.Web.UI.WebControls.Image
                        Dim i As Integer = -1
                        Select Case Me.ViewState("sort")
                            Case "DistID", "DistID DESC"
                                'mylabel = "ComName"
                                i = 1
                                If Me.ViewState("sort") = "DistID" Then
                                    mysort.ImageUrl = "../../images/SortUp.gif"
                                Else
                                    mysort.ImageUrl = "../../images/SortDown.gif"
                                End If
                            Case "PlanID", "PlanID DESC"
                                'mylabel = "ComName"
                                i = 3
                                If Me.ViewState("sort") = "PlanID" Then
                                    mysort.ImageUrl = "../../images/SortUp.gif"
                                Else
                                    mysort.ImageUrl = "../../images/SortDown.gif"
                                End If
                            Case "OrgName", "OrgName DESC"
                                'mylabel = "OrgName"
                                i = 4
                                If Me.ViewState("sort") = "OrgName" Then
                                    mysort.ImageUrl = "../../images/SortUp.gif"
                                Else
                                    mysort.ImageUrl = "../../images/SortDown.gif"
                                End If
                            Case "STDate", "STDate DESC"
                                'mylabel = "ComName"
                                i = 7
                                If Me.ViewState("sort") = "STDate" Then
                                    mysort.ImageUrl = "../../images/SortUp.gif"
                                Else
                                    mysort.ImageUrl = "../../images/SortDown.gif"
                                End If
                        End Select
                        If i <> -1 Then
                            e.Item.Cells(i).Controls.Add(mysort)
                        End If
                    End If

                End If
            Case ListItemType.Item, ListItemType.AlternatingItem

                Dim drv As DataRowView = e.Item.DataItem
                Dim mybtn As LinkButton

                If Not blnPrintFlag Then
                    mybtn = e.Item.Cells(5).Controls(0)
                    mybtn.Attributes("onclick") = "window.open('CP_04_002_01.aspx?SeqNO=" & Me.DataGrid1.DataKeys(e.Item.ItemIndex) & "&PlanName=" & Server.UrlEncode(e.Item.Cells(3).Text) & "&PlanID=" & e.Item.Cells(12).Text & "&ComIDNO=" & e.Item.Cells(13).Text & "&Year=" & yearlist & "&TrainName=" & Server.UrlEncode(e.Item.Cells(14).Text) & "'); return false;"
                Else
                    e.Item.Cells(5).CssClass = "font"
                    e.Item.Cells(5).Text = drv("ClassName").ToString
                End If

                '序號
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                '行政管理費
                e.Item.Cells(11).Text = drv("Tcost") + AdmFee(drv("planid"), drv("comIDNO"), drv("seqNo"))
                e.Item.Cells(11).ToolTip = ""
                e.Item.Cells(11).ToolTip += "原訓練費用：" & drv("Tcost").ToString & vbCrLf
                e.Item.Cells(11).ToolTip += "行政管理費：" & AdmFee(drv("planid"), drv("comIDNO"), drv("seqNo")).ToString


                Select Case drv("AppliedResult").ToString
                    Case "Y"
                        e.Item.Cells(6).Text = "審核通過"
                    Case "N"
                        e.Item.Cells(6).Text = "審核未通過"
                    Case "M"
                        e.Item.Cells(6).Text = "請修正資料"
                    Case Else
                        e.Item.Cells(6).Text = "審核中"
                End Select
        End Select

    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Not blnPrintFlag Then
            If Me.ViewState("sort") <> e.SortExpression Then
                Me.ViewState("sort") = e.SortExpression
            Else
                Me.ViewState("sort") = e.SortExpression & " DESC"
            End If
            PageControler1.ChangeSort(Me.ViewState("sort"))
        End If
    End Sub

    '行政管理費 add by nick 060522
    Function AdmFee(ByVal PlanID As Int32, ByVal ComIDNO As String, ByVal seqNO As Int32) As Integer
        Dim ACsum As Integer = 0
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr, s As String
        'Dim sum As Double
        'Dim ACsum, totalsum, totalus As Integer

        'sqlstr = "SELECT *,(c.OPrice*c.Itemage*c.ItemCost) as AllSmallSum "
        'sqlstr += "from Plan_PlanInfo a "
        'sqlstr += "JOIN Plan_CostItem c ON c.PlanID=a.PlanID and c.ComIDNO=a.ComIDNO and c.SeqNO=a.SeqNO "
        'sqlstr += "where a.PlanID=" & PlanID
        'sqlstr += " and a.ComIDNO='" & ComIDNO & "' "
        'sqlstr += " and a.SeqNO=" & seqNO

        Dim sum As Double = 0
        Dim hPMS As New Hashtable
        hPMS.Add("PlanID", Val(PlanID))
        hPMS.Add("ComIDNO", ComIDNO)
        hPMS.Add("SeqNO", Val(seqNO))
        Dim sqlstr As String = ""
        sqlstr += " SELECT c.AdmFlag" & vbCrLf
        sqlstr += " ,d.CostName" & vbCrLf
        sqlstr += " ,a.AdmPercent" & vbCrLf
        sqlstr += " ,a.TNum" & vbCrLf
        sqlstr += " ,c.CostID" & vbCrLf
        sqlstr += " ,c.ItemOther" & vbCrLf
        sqlstr += " ,c.OPrice" & vbCrLf
        sqlstr += " ,c.Itemage" & vbCrLf
        sqlstr += " ,(c.OPrice*c.Itemage*c.ItemCost) AllSmallSum " & vbCrLf
        sqlstr += " from Plan_PlanInfo a "
        sqlstr += " JOIN Plan_CostItem c ON c.PlanID=a.PlanID and c.ComIDNO=a.ComIDNO and c.SeqNO=a.SeqNO "
        sqlstr += " JOIN Key_CostItem  d ON d.CostID=c.CostID "
        sqlstr += " where a.PlanID=@PlanID and a.ComIDNO=@ComIDNO and a.SeqNO=@SeqNO"
        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn, hPMS)

        If dt.Rows.Count > 0 Then
            '行政管理費
            For Each dr As DataRow In dt.Select("AdmFlag='Y'")
                'If Not dr("ItemCost").ToString = "" Then
                sum += CInt(dr("OPrice")) * CInt(dr("Itemage"))
                'End If
            Next

            If dt.Rows(0)("AdmPercent").ToString = "" Then
                Me.ViewState("AdmPercent") = 0
            Else
                Me.ViewState("AdmPercent") = dt.Rows(0)("AdmPercent")
            End If

            If Not sum = Nothing Then
                ACsum = Math.Round(sum * CDbl(Me.ViewState("AdmPercent").ToString) / 100)
            Else
                ACsum = 0
            End If

        End If
        Return ACsum
    End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出Excel
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        blnPrintFlag = True '列印動作

        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim sErrMsg As String = ""
        sErrMsg = ""
        If Not create(sErrMsg) Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        'Me.EnableEventValidation = False
        'Me.AutoEventWireup = True

        Dim sFileName As String = "計畫資料.xls"
        sFileName = HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8)
        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.Visible = False
    End Sub


End Class
