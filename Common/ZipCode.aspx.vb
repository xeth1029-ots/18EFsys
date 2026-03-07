Partial Class ZipCode1
    Inherits AuthBasePage

    'ID_ZIP2,ID_ZIP6,ID_ZIP,ID_CITY 
    'Dim sTmp As String="" '暫存字

#Region "匯入資料方式"

    '台灣的郵遞區號過去採行3+2制度，2020年3月3日起為改善投遞效率，將郵遞區號改為3+3碼。
    '前3碼為臺灣368個鄉鎮市區專用碼（新竹市與嘉義市其下雖有分區，但郵遞區號不分碼）
    '，加上南海諸島東沙、南沙（以上兩區皆隸屬高雄市旗津區管轄）及釣魚臺列嶼（臺灣將該島劃為宜蘭縣頭城鎮的行政區）
    '，後3碼為大量用戶專用碼或是投遞責任區碼。普通寫前3碼即可，例如新竹市為300，若完整書寫郵遞區號可加快郵件遞送速度。
    '大部分地區的房屋門牌有註明所在地鄉、鎮、市、區名及其郵遞區號，例外如臺北市全市門牌無標示區名和郵遞區號。
    '中華郵政建議在國內直式標準信封背面加印郵遞區號一覽表，以便查閱。台灣除臺北市及其他省轄市地址有時不寫區名外
    '（例如臺北市中正區中山南路簡為臺北市中山南路），郵政單位仍建議各市地址應寫區名及郵遞區號，以減少查閱及分揀的麻煩。

    'GOOGLE SEARCH: 3+2碼郵遞區號Excel檔
    '3+2碼郵遞區號Excel檔 101/02(自解壓縮檔) 
    'http://www.post.gov.tw/post/internet/Download/default.jsp?ID=22

    'drop table ID_Zip2
    'GO
    ' CREATE TABLE [dbo].[ID_Zip2] ( 
    '	[ZipLID] [int]  IDENTITY(1,1)  NOT NULL ,
    '	[ZipCode] [int]  NOT NULL ,
    '	[ZipCode2] [varchar] (2)  NOT NULL ,
    '	[road] [nvarchar] (32)  NOT NULL ,
    '	[note] [nvarchar] (32)  NOT NULL ,
    '	 CONSTRAINT [PK_ID_Zip2]  PRIMARY KEY CLUSTERED 
    '	 ( 
    '	 [ZipLID] 	 )  ON [PRIMARY] 
    '	) ON [PRIMARY] 
    'GO
    '--select * from ID_Zip2
    'INSERT INTO ID_Zip2(
    ' ZipCode
    ' ,ZipCode2
    ' ,road
    ' ,note
    ') 
    'select left([Zip Code],3) ZipCode,right([Zip Code],2) ZipCode2
    ',ltrim(rtrim(replace(road,' ','')))  road
    ',ltrim(rtrim(replace(scope,' ',''))) note
    'from Zip32_10102

    'ZIP_CODE NVARCHAR2 255 Y    
    'CITY NVARCHAR2 255 Y    
    'AREA NVARCHAR2 255 Y    
    'ROAD NVARCHAR2 255 Y    
    'SCOPE NVARCHAR2 255 Y 
    'ZIP32_10406

    '/*
    'select left([ZIP_CODE],3) ZipCode,right([ZIP_CODE],2) ZipCode2
    ',ltrim(rtrim(replace(road,' ','')))  road
    ',ltrim(rtrim(replace(scope,' ',''))) note
    'from ZIP32_10406
    '*/
    'drop table ID_Zip2
    'go
    'CREATE TABLE [dbo].[ID_Zip2] (
    '	[ZipLID] [int]  IDENTITY(1,1)  NOT NULL ,
    '	[ZipCode] [int]  NOT NULL ,
    '	[ZipCode2] [varchar] (2)  NOT NULL ,
    '	[road] [nvarchar] (32)  NOT NULL ,
    '	[note] [nvarchar] (32)  NOT NULL ,
    '	 CONSTRAINT [PK_ID_Zip2]  PRIMARY KEY CLUSTERED
    '	 (
    '	 [ZipLID] 	 )  ON [PRIMARY]
    '	) ON [PRIMARY]
    'go
    'INSERT INTO ID_Zip2(
    'ZipCode
    ',ZipCode2
    ',road
    ',note
    ')
    'select left([ZIP_CODE],3) ZipCode,right([ZIP_CODE],2) ZipCode2
    ',ltrim(rtrim(replace(road,' ','')))  road
    ',ltrim(rtrim(replace(scope,' ',''))) note
    'from ZIP32_10406
    'go
    '--select * from ID_Zip2
    'select count(1) cnt  from ID_Zip2

    '-- truncate table ID_Zip2
    'select 
    ''insert into ID_Zip2(ZIPLID ,ZipCode,ZipCode2,road,note )'
    '+' select '+convert(varchar,ZIPLID)+','+convert(varchar,ZipCode)
    '+',N'''+ZipCode2+''',N'''+road+''',N'''+note+'''  ;/' xstr
    'from ID_Zip2

    ' CREATE TABLE ID_ZIP3 ( 
    '	 ZIPLID  NUMBER (10,0)  NOT NULL ,
    '	 ZIPCODE  NUMBER (10,0)  NOT NULL ,
    '	 ZIPCODE2  VARCHAR2 (2 char)  NOT NULL ,
    '	 ROAD  NVARCHAR2 (32)  NOT NULL ,
    '	 NOTE  NVARCHAR2 (32)  NOT NULL ,
    '	 CONSTRAINT PK_ID_ZIP3 PRIMARY KEY 
    '	 ( 
    '	 ZIPLID 	 )  ENABLE 
    '	)

    'select 
    ''insert into ID_Zip3(ZIPLID ,ZipCode,ZipCode2,road,note )'
    '+' select '+convert(varchar,ZIPLID)+','+convert(varchar,ZipCode)
    '+',N'''+ZipCode2+''',N'''+road+''',N'''+note+'''  ; ' xstr
    'from ID_Zip2

#End Region

    Const cst_city As String = "city"
    Const cst_zip As String = "zip"
    'Const cst_ziptype As Integer=6 '6碼
    'If TIMS.sUtl_ChkTest AndAlso hidSN.Value="" Then hidSN.Value="0" '測試環境

    ''' <summary> 代入DropDownList資料 </summary>
    ''' <param name="strType"></param>
    ''' <param name="objDDL"></param>
    ''' <param name="sTpValue"></param>
    ''' <param name="oConn"></param>
    Public Sub ddlList(ByVal strType As String, ByVal objDDL As DropDownList, ByVal sTpValue As String, ByRef oConn As SqlConnection)
        Dim sql As String = " SELECT ZIPCODE ID, ZIPNAME NAME FROM ID_ZIP WHERE 1<>1" ' order by 1,2
        Select Case strType
            Case cst_city
                sql = " SELECT CTID ID, CTNAME NAME FROM ID_CITY ORDER BY 1"
            Case cst_zip
                Dim flag_Can_Use_sTpValue As Boolean = (sTpValue <> "" AndAlso TIMS.IsNumeric2(sTpValue))
                'If sTpValue <> "" AndAlso TIMS.IsNumeric2(sTpValue) Then flag_Can_Use_sTpValue=True
                If flag_Can_Use_sTpValue Then
                    sql = $" SELECT ZIPCODE ID, ZIPNAME NAME FROM ID_ZIP WHERE CTID={sTpValue} ORDER BY 1,2 "
                End If
        End Select
        'Call TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        Using sCmd As New SqlCommand(sql, oConn)
            With sCmd
                .Parameters.Clear()
                dt.Load(.ExecuteReader())
            End With
            objDDL.Items.Clear()
            If TIMS.dtNODATA(dt) Then Return
        End Using

        With objDDL
            .DataSource = dt 'ds.Tables("Data")
            .DataTextField = "name"
            .DataValueField = "id"
            .DataBind()
            .Items.Insert(0, New ListItem("請選擇", "")) 'Clear()
        End With
    End Sub

    ''' <summary> 查詢 3+2 郵遞區號 </summary>
    ''' <returns></returns>
    Function Search1_Dt() As DataTable
        'Dim sTmp As String="" '暫存1,'Dim sTmp2 As String="" '暫存2
        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " SELECT c.ctid ,a.zipcode ,a.zipcode2, concat(a.zipcode ,a.zipcode2) zipcode6" & vbCrLf
        sql &= " ,CASE WHEN a.ZIPCODE_N IS NOT NULL THEN a.city ELSE c.ctname END ctname" & vbCrLf
        sql &= " ,CASE WHEN a.ZIPCODE_N IS NOT NULL THEN a.area ELSE b.zipname END zipname" & vbCrLf
        sql &= " ,a.road ,a.note ,a.ZIPCODE_N ,a.ZIPLID,a.AREA" & vbCrLf
        sql &= " FROM ID_ZIP2 a" & vbCrLf
        sql &= " JOIN ID_ZIP b ON b.zipcode=a.zipcode" & vbCrLf
        sql &= " JOIN ID_CITY c ON c.ctid=b.ctid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        Select Case hidSN.Value
            Case "0", "1" '0:舊版ZIP取得 1:
            Case Else
                sql &= " AND 1<>1 " & vbCrLf '異常不顯示資料(應有0或1)
                Return dt
        End Select

        Dim myParam As New Hashtable
        txtZip.Text = TIMS.ClearSQM(txtZip.Text)
        If txtZip.Text <> "" Then
            Dim v_txtZip As String = txtZip.Text
            If v_txtZip.Length = 3 Then
                sql &= " AND CONVERT(VARCHAR,a.zipcode)=@TXTZIP " & vbCrLf
                myParam.Add("TXTZIP", v_txtZip)
            Else
                sql &= " AND concat(a.zipcode,a.zipcode2) like '%'+@TXTZIP+'%'" & vbCrLf
                myParam.Add("TXTZIP", v_txtZip)
            End If
        End If

        Dim v_ddlCity As String = TIMS.GetListValue(ddlCity)
        If v_ddlCity <> "" Then
            sql &= " AND c.CTID=@CTID " & vbCrLf
            myParam.Add("CTID", v_ddlCity)
        End If

        Dim v_ddlZip As String = TIMS.GetListValue(ddlZip)
        Dim v_ddlZip_txt As String = TIMS.GetListText(ddlZip)
        If v_ddlZip <> "" AndAlso TIMS.IsNumeric2(v_ddlZip) Then '確認是郵遞區號格式
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR CONVERT(varchar,a.ZIPCODE)=@ZIPCODE " & vbCrLf
            sql &= " OR CONVERT(varchar,a.AREA)=@AREA " & vbCrLf
            sql &= " )" & vbCrLf
            myParam.Add("ZIPCODE", v_ddlZip)
            myParam.Add("AREA", v_ddlZip_txt)
        End If

        txtRoad.Text = TIMS.ClearSQM(txtRoad.Text)
        If txtRoad.Text <> "" Then
            sql &= " AND a.ROAD LIKE '%'+@Road+'%' " & vbCrLf
            myParam.Add("Road", txtRoad.Text)
        End If
        sql &= " ORDER BY 1,2,3,4,5,6,7 " & vbCrLf
        'Call TIMS.OpenDbConn(objconn)
        dt = DbAccess.GetDataTable(sql, objconn, myParam)
        Return dt
    End Function

    ''' <summary> 查詢 3+3(6W) 郵遞區號 </summary>
    ''' <returns></returns>
    Function Search1_Dt2() As DataTable
        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " SELECT c.ctid ,a.zipcode ,SUBSTRING(a.zipcode6,4,3) zipcode2,a.zipcode6" & vbCrLf
        sql &= " ,CASE WHEN a.ZIPCODE_N IS NOT NULL THEN a.city ELSE c.ctname END ctname" & vbCrLf
        sql &= " ,CASE WHEN a.ZIPCODE_N IS NOT NULL THEN a.area ELSE b.zipname END zipname" & vbCrLf
        sql &= " ,a.road,a.note,a.ZIPCODE_N,a.ZIPLID,a.AREA" & vbCrLf
        sql &= " FROM ID_ZIP6 a" & vbCrLf
        sql &= " JOIN ID_ZIP b ON b.zipcode=a.zipcode" & vbCrLf
        sql &= " JOIN ID_CITY c ON c.ctid=b.ctid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        Select Case hidSN.Value
            Case "0", "1" '0:舊版ZIP取得 1:
            Case Else
                sql &= " AND 1<>1 " & vbCrLf '異常不顯示資料(應有0或1)
                Return dt
        End Select

        Dim myParam As New Hashtable
        txtZip.Text = TIMS.ClearSQM(txtZip.Text)
        If txtZip.Text <> "" Then
            Dim v_txtZip As String = txtZip.Text
            sql &= " AND a.zipcode6 like '%'+@TXTZIP+'%'" & vbCrLf
            myParam.Add("TXTZIP", v_txtZip)
        End If
        Dim v_ddlCity As String = TIMS.GetListValue(ddlCity)
        If v_ddlCity <> "" Then
            sql &= " AND c.CTID=@CTID " & vbCrLf
            myParam.Add("CTID", v_ddlCity)
        End If
        Dim v_ddlZip As String = TIMS.GetListValue(ddlZip)
        Dim v_ddlZip_txt As String = TIMS.GetListText(ddlZip)
        If v_ddlZip <> "" AndAlso TIMS.IsNumeric2(v_ddlZip) Then '確認是郵遞區號格式
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR CONVERT(varchar,a.ZIPCODE)=@ZIPCODE " & vbCrLf
            sql &= " )" & vbCrLf
            myParam.Add("ZIPCODE", v_ddlZip)
        End If

        txtRoad.Text = TIMS.ClearSQM(txtRoad.Text)
        If txtRoad.Text <> "" Then
            sql &= " AND a.road LIKE '%'+@Road+'%' " & vbCrLf
            myParam.Add("Road", txtRoad.Text)
        End If
        sql &= " ORDER BY 1,2,3,4,5,6,7 " & vbCrLf
        'Call TIMS.OpenDbConn(objconn)
        dt = DbAccess.GetDataTable(sql, objconn, myParam)
        Return dt
    End Function

    '查詢
    Sub SHOW_DG1(ByVal iPageIndex As Integer)
        '3: 3+3郵遞區號 2: 3+2郵遞區號
        Dim v_rblPOSTTYPE1 As String = TIMS.GetListValue(rblPOSTTYPE1)

        'Dim dt As DataTable=Search1_Dt() 'Session("ZipCodeDt")
        'If(flag_work2022xZIP6, Search1_Dt2(), Search1_Dt())
        Dim dt As DataTable = If(v_rblPOSTTYPE1 = "2", Search1_Dt(), Search1_Dt2())
        DataGrid1.Visible = False
        labMsg.Text = "查無資料!"

        If dt.Rows.Count = 0 Then Return

        'If dt.Rows.Count > 0 Then  End If
        DataGrid1.Visible = True
        labMsg.Text = ""
        DataGrid1.CurrentPageIndex = 0
        DataGrid1.DataSource = dt 'ds.Tables("data")
        DataGrid1.DataBind()
        If iPageIndex > 0 Then
            '無效的 CurrentPageIndex 值。必須是 >= 0 且 < PageCount。
            If Not (iPageIndex >= 0 AndAlso iPageIndex < DataGrid1.PageCount) Then iPageIndex = 0 '(超過系統頁數範圍預設為0)
            DataGrid1.CurrentPageIndex = iPageIndex
            DataGrid1.DataBind()
        End If
        Call TIMS.Set_row_color(DataGrid1)
    End Sub

    'Dim flag_work2022xZIP6 As Boolean=False
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        'test YY work2022xZIP6 全系統+網站【郵遞區號】調整可支援3+3碼功能
        'flag_work2022xZIP6=TIMS.sUtl_ChkTest("work2022xZIP6")

        If Not IsPostBack Then Call Create1()
    End Sub

    Sub Create1()
        'Session("ZipCodeDt")=Nothing ZipCode.aspx
        hidSN.Value = TIMS.ClearSQM(Request("sn"))
        'If TIMS.sUtl_ChkTest AndAlso hidSN.Value="" Then hidSN.Value="0" '測試環境
        Dim flagNG1 As Boolean = False
        Select Case hidSN.Value
            Case "0"
                'sn=0&field=ctid,zipcode,zipcode2,zipcode6,ctname,zipname,road
                If Split(Server.UrlDecode(Request("field")), ",").Length > 5 Then
                    hidCtID.Value = Split(Server.UrlDecode(Request("field")), ",")(0)
                    hidZipCode.Value = Split(Server.UrlDecode(Request("field")), ",")(1)
                    hidZipCode2.Value = Split(Server.UrlDecode(Request("field")), ",")(2)
                    'ZipCode6,flag_work2022xZIP6
                    hidZipCode6.Value = Split(Server.UrlDecode(Request("field")), ",")(3)
                    hidCtName.Value = Split(Server.UrlDecode(Request("field")), ",")(4)
                    hidZIPCODE_N.Value = Split(Server.UrlDecode(Request("field")), ",")(5) 'hidZipName.Value=Split(Server.UrlDecode(Request("field")), ",")(5)
                    hidRoad.Value = Split(Server.UrlDecode(Request("field")), ",")(6)
                Else
                    flagNG1 = True
                End If

            Case "1"
                'sn=1&field=ctid,zipcode,zipcode2,zipcode6,ctname,cityname,zipname,zipcode_n
                If Split(Server.UrlDecode(Request("field")), ",").Length > 7 Then
                    hidCtID.Value = Split(Server.UrlDecode(Request("field")), ",")(0)
                    hidZipCode.Value = Split(Server.UrlDecode(Request("field")), ",")(1)
                    hidZipCode2.Value = Split(Server.UrlDecode(Request("field")), ",")(2)
                    'ZipCode6,flag_work2022xZIP6
                    hidZipCode6.Value = Split(Server.UrlDecode(Request("field")), ",")(3)
                    hidCtName.Value = Split(Server.UrlDecode(Request("field")), ",")(4)
                    hidCityName.Value = Split(Server.UrlDecode(Request("field")), ",")(5)
                    hidZipName.Value = Split(Server.UrlDecode(Request("field")), ",")(6)
                    hidZIPCODE_N.Value = Split(Server.UrlDecode(Request("field")), ",")(7)
                Else
                    flagNG1 = True
                End If

            Case Else
                flagNG1 = True
        End Select
        If flagNG1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg9)
            Exit Sub
        End If

        'hidSN.Value=TIMS.ClearSQM(hidSN.Value)
        hidCtID.Value = TIMS.ClearSQM(hidCtID.Value)
        hidZipCode.Value = TIMS.ClearSQM(hidZipCode.Value)
        hidZipCode2.Value = TIMS.ClearSQM(hidZipCode2.Value)
        'ZipCode6,flag_work2022xZIP6
        hidZipCode6.Value = TIMS.ClearSQM(hidZipCode6.Value)
        hidCtName.Value = TIMS.ClearSQM(hidCtName.Value)
        hidCityName.Value = TIMS.ClearSQM(hidCityName.Value)
        hidZipName.Value = TIMS.ClearSQM(hidZipName.Value)
        hidZIPCODE_N.Value = TIMS.ClearSQM(hidZIPCODE_N.Value)
        hidRoad.Value = TIMS.ClearSQM(hidRoad.Value)
        'hidno.Value=TIMS.ClearSQM(hidno.Value)
        Call ddlList(cst_city, ddlCity, "", objconn)
        btnCancel.Attributes.Add("onclick", "window.close();")
        btnClear1.Attributes.Add("onclick", "getVal('','','','','','','','');")
    End Sub

    ''' <summary>
    ''' 查詢鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSch.Click
        txtZip.Text = TIMS.ClearSQM(txtZip.Text)
        Dim v_ddlCity As String = TIMS.ClearSQM(ddlCity)
        txtRoad.Text = TIMS.ClearSQM(txtRoad.Text)

        '3: 3+3郵遞區號 2: 3+2郵遞區號
        Dim v_rblPOSTTYPE1 As String = TIMS.GetListValue(rblPOSTTYPE1)
        Select Case v_rblPOSTTYPE1
            Case "2", "3"
            Case Else
                Common.MessageBox(Me, "請選擇郵遞區號種類!!")
                Exit Sub
        End Select

        Dim X1 As Boolean = If(txtZip.Text = "", True, False)
        Dim X2 As Boolean = If(v_ddlCity = "", True, False)
        Dim X3 As Boolean = If(txtRoad.Text = "", True, False)
        If X1 AndAlso X2 AndAlso X3 Then
            Common.MessageBox(Me, "請選擇或輸入任1條件!!")
            Exit Sub
        End If

        Call SHOW_DG1(0)
        'Call TIMS.set_row_color(DataGrid1)
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        'Dim iValue As Integer=e.NewPageIndex,'Dim dt As DataTable=Search1_Dt() 'Session("ZipCodeDt")
        Call SHOW_DG1(TIMS.GetValue1(e.NewPageIndex))
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim rdoSelect As RadioButton = e.Item.FindControl("rdoSelect")
                Dim labZipCode As Label = e.Item.FindControl("labZipCode")
                Dim labZipArea As Label = e.Item.FindControl("labZipArea")
                Dim labRoad As Label = e.Item.FindControl("labRoad")
                Dim labNote As Label = e.Item.FindControl("labNote")

                'Dim s_ZipCode As String=If(flag_work2022xZIP6, Convert.ToString(drv("zipcode6")), Convert.ToString(drv("zipcode")) & Convert.ToString(drv("zipcode2")))
                labZipCode.Text = Convert.ToString(drv("zipcode6"))
                Dim s_ZipArea As String = If(Convert.ToString(drv("ctname")) = Convert.ToString(drv("zipname")), Convert.ToString(drv("zipname")), String.Concat(drv("ctname"), drv("zipname")))
                labZipArea.Text = s_ZipArea 'Convert.ToString(drv("ctname")) & Convert.ToString(drv("zipname"))

                labRoad.Text = Convert.ToString(drv("road"))
                labNote.Text = Convert.ToString(drv("note"))
                Dim JSvalue1x As String = ""
                TIMS.AddSQMValue(JSvalue1x, $"{drv("ctid")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("zipcode")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("zipcode2")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("zipcode6")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("ctname")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("zipname")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("ZIPCODE_N")}")
                TIMS.AddSQMValue(JSvalue1x, $"{drv("road")}")
                rdoSelect.Attributes.Add("onclick", String.Format("getVal({0});", JSvalue1x))
        End Select
    End Sub

    Private Sub ddlCity_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCity.SelectedIndexChanged
        Dim v_ddlCity As String = TIMS.GetListValue(ddlCity)
        If v_ddlCity <> "" Then
            ddlZip.Enabled = True
            Call ddlList(cst_zip, ddlZip, v_ddlCity, objconn)
            Return 'Exit Sub
        End If

        ddlZip.Enabled = False
        ddlZip.Items.Clear()
        ddlZip.Items.Add(New ListItem("請選擇縣市", "")) 'Clear()
        'ddlZip.SelectedValue=""
    End Sub
End Class