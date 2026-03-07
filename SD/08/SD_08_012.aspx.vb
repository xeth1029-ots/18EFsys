Partial Class SD_08_012
    Inherits AuthBasePage

    'Dim OCID As String = ""
    'Dim orgname As String = ""
    Dim objconn As SqlConnection ' = DbAccess.GetConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cache.SetExpires(DateTime.Now())
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        '在這裡放置使用者程式碼以初始化網頁
        Dim sql As String = ""
        'Dim timssql As String
        'Dim condsql As String
        Dim condsqltims As String = ""
        Dim condsqlnontims As String = ""

        Dim dt As New DataTable
        Dim i As Int32
        Dim j As Int32
        '第幾筆資料
        Dim k As Int32 = 0
        Dim count As Int32
        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl

        Dim OCID As String = Server.UrlDecode(Request("OCID")) '課程名稱
        Dim orgname As String = Server.UrlDecode(Request("orgname")) '申請職業訓練機構
        If OCID = "" Then Exit Sub
        Dim drC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        sql = ""
        sql &= " select b.ClassName,b.IdentityID,b.Name, b.Birthday,b.IDNO, b.TSDate, b.TEDate"
        sql &= " ,b.ApplyMonth, b.ApplyMoney"
        sql &= " from Sub_SubSidyApply a "
        sql &= " left join Sub_SubSidyApply_all b on a.subid=b.tsubid"
        sql &= " where socid in (select socid from class_studentsofclass where ocid in (" & OCID & ")) "
        sql &= " order by b.idno,b.TSDate desc"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then

            Dim ttlMoney As Integer = 0 '請領金額加總

            If (dt.Rows.Count Mod 10) = 0 Then
                count = (dt.Rows.Count / 10) - 1
            Else
                count = dt.Rows.Count / 10
            End If
            For i = 0 To count

                '表頭
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("colspan", 3)
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("colspan", 5)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("colspan", 3)
                nc.Attributes.Add("align", "right")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("colspan", 3)
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                'stella modify 2007-8-9 北區人員提出
                'nc.InnerHtml = orgname & (Convert.ToInt16(sm.UserInfo.Years) - 1911).ToString & "年度申請"
                nc.InnerHtml = orgname

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("colspan", 5)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                'stella modify 2007-8-9 北區人員提出
                'nc.InnerHtml = "受訓學員訓練生活津貼補助印領清冊"
                nc.InnerHtml = (Convert.ToInt16(sm.UserInfo.Years) - 1911).ToString & "年度申請受訓學員訓練生活津貼補助印領清冊"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("colspan", 3)
                nc.Attributes.Add("align", "right")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "列印日期: " & Format(Today, "yyyy/MM/dd")


                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "2")
                div_print.Controls.Add(nt)
                '欄位名稱
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "編號"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "14%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "參訓班別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "11%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "申請類別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "8%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "姓名"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "8%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "出生日期"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "12%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "身分證編號或<br>身障手冊編號"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "9%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "受訓起<br>迄日期"

                'stella add 2006/12/20
                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "6%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "月數"
                'stella add 2006/12/20

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "7%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "請領總<br>金額"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "9%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "申請人<br>蓋章"

                'stella modify 2006/12/20
                'nc = New HtmlTableCell
                'nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "6%")
                ''nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                'nc.Attributes.Add("align", "center")
                'nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                'nc.InnerHtml = "扶養親<br>屬人數"
                'stella add 2006/12/20

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "13%")
                'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                nc.InnerHtml = "備註"

                For j = 0 To 9
                    If k + 1 > dt.Rows.Count Then
                        GoTo [CONTINUE]
                    End If

                    '資料內容
                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    '編號
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "5%")
                    nc.Attributes.Add("height", "50px")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = k + 1
                    '參訓班別
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "14%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = dt.Rows(k).Item(0)
                    '申請類別
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "11%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = getcname(dt.Rows(k).Item(1), "id")
                    '姓名
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "8%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = dt.Rows(k).Item(2)
                    '出生日期
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "8%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    Dim v_ROC_DATE1 As String = ""
                    If Convert.ToString(dt.Rows(k).Item(3)) <> "" Then
                        v_ROC_DATE1 = TIMS.Cdate17(CDate(Format(dt.Rows(k).Item(3), "yyyy/MM/dd")))
                    End If
                    nc.InnerHtml = v_ROC_DATE1 'CommonUtil.W2ROC(Format(dt.Rows(k).Item(3), "yyyy/MM/dd"))
                    'nc.InnerHtml = Format(DateAdd("yyyy", -1911, Format(dt.Rows(k).Item(3), "yyyy/MM/dd")), "yy/MM/dd")
                    '身分證編號或<br>身障手冊編號
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "10%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = dt.Rows(k).Item(4)
                    '受訓起<br>迄日期
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "9%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = Format(DateAdd("yyyy", -1911, Format(dt.Rows(k).Item(5), "yyyy/MM/dd")), "yy/MM/dd") & "至<br>" & Format(DateAdd("yyyy", -1911, Format(dt.Rows(k).Item(6), "yyyy/MM/dd")), "yy/MM/dd")
                    '月數
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "5%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = dt.Rows(k).Item(7)
                    '請領總<br>金額
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "7%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    nc.InnerHtml = FormatNumber(dt.Rows(k).Item(8), 0)

                    ttlMoney += Convert.ToInt64(dt.Rows(k).Item(8))

                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "9%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    'nc.InnerHtml = "申請人<br>蓋章"

                    'stella modify 2006/12/20
                    'nc = New HtmlTableCell
                    'nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "6%")
                    ''nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    'nc.Attributes.Add("align", "center")
                    'nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    'nc.InnerHtml = "0"
                    'stella modify 2006/12/20

                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("width", "13%")
                    'nc.Attributes.Add("colspan", (colcnt + 1).ToString())
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", "font-size:10pt;font-family:DFKai-SB")
                    'nc.InnerHtml = "備註"
                    k += 1

                    If k = dt.Rows.Count Then
                        ''最後一列顯示總人數
                        'nr = New HtmlTableRow
                        'nt.Controls.Add(nr)

                        'nc = New HtmlTableCell
                        'nr.Controls.Add(nc)
                        'nc.Attributes.Add("colspan", 13)
                        'nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        ''nc.InnerHtml = "上列人數共計<u>&nbsp;&nbsp;" & dt.Rows.Count & "&nbsp;&nbsp;</u>人&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        ''nc.InnerHtml += "合計新臺幣<u>&nbsp;&nbsp;&nbsp;&nbsp;" & FormatNumber(ttlMoney, 0) & "&nbsp;&nbsp;&nbsp;&nbsp;</u>元整"

                        'Dim dts_money As DataTable
                        'dts_money = DbAccess.GetDataTable("select dbo.fn_GET_Money('" & ttlMoney & "') as tobig")
                        'nc.InnerHtml = "上列人數共計<u>&nbsp;&nbsp;" & dt.Rows.Count & "&nbsp;&nbsp;</u>人&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        'nc.InnerHtml += "合計新臺幣<u>&nbsp;&nbsp;&nbsp;&nbsp;" & dts_money.Rows(0).Item("tobig") & "&nbsp;&nbsp;&nbsp;&nbsp;</u>"
                        'nc.InnerHtml += "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        'nc.InnerHtml += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        ''nc.InnerHtml += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        'nc.InnerHtml += "(" & FormatNumber(ttlMoney, 0) & ")"
                        nt = New HtmlTable
                        nt.Attributes.Add("style", "width:100%;BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                        nt.Attributes.Add("align", "center")
                        nt.Attributes.Add("border", "0")
                        div_print.Controls.Add(nt)

                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("width", "73%")

                        Dim sql1 As String = ""
                        sql1 = "select dbo.fn_GET_Money('" & ttlMoney & "') tobig "
                        Dim dts_money As DataTable
                        dts_money = DbAccess.GetDataTable(sql1, objconn)
                        nc.InnerHtml = "上列人數共計<u>&nbsp;&nbsp;" & dt.Rows.Count & "&nbsp;&nbsp;</u>人&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        nc.InnerHtml += "合計新臺幣<u>&nbsp;&nbsp;&nbsp;&nbsp;" & dts_money.Rows(0).Item("tobig") & "&nbsp;&nbsp;&nbsp;&nbsp;</u><br><br>"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("width", "27%")
                        nc.InnerHtml = FormatNumber(ttlMoney, 0) & "元<br><br>"

                        nt = New HtmlTable
                        nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                        nt.Attributes.Add("align", "center")
                        nt.Attributes.Add("border", "0")
                        div_print.Controls.Add(nt)
                        '第一列(空行)
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 6)
                        nc.InnerHtml = "&nbsp;"

                        '第二列
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "(承辦單位)"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "20%")
                        nc.InnerHtml = "&nbsp;"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "會計主管："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "60%")
                        nc.Attributes.Add("colspan", 3)
                        nc.InnerHtml = "&nbsp;"

                        '第三列(空行)
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 6)
                        nc.InnerHtml = "&nbsp;"

                        '第四列
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "承辦人員："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "20%")
                        nc.InnerHtml = "&nbsp;"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "單位主管："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "30%")
                        nc.InnerHtml = "&nbsp;"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "機關首長："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "20%")
                        nc.InnerHtml = "&nbsp;"

                        '第五列(空行)
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 6)
                        nc.InnerHtml = "&nbsp;"

                        '第六列
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("colspan", 2)
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "(委訓單位)承辦人員："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "60%")
                        nc.Attributes.Add("colspan", 2)
                        nc.InnerHtml = "&nbsp;"

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "業務主管："

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("width", "20%")
                        nc.InnerHtml = "&nbsp;"

                        '第七列(空行)
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 6)
                        nc.InnerHtml = "&nbsp;"

                        '第四列(說明文字)
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 6)
                        nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                        nc.InnerHtml = "說明：一、本清冊請分別填繕一式三份，並加蓋申請人私章及原承辦單位、主管人員職章（如係政府機關委託辦理訓練者，須加蓋<br>"
                        nc.InnerHtml += " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;原委訓單位業務主管人員及承辦人員職章）<br>"
                        nc.InnerHtml += " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        nc.InnerHtml += " 二、各項資料請詳實填寫，如有塗改時，請加蓋申請人私章。<br>"
                        nc.InnerHtml += " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        nc.InnerHtml += " 三、本表僅就「就業促進津貼實施辦法」規定，請領訓練生活津貼填報（即依基本工資之60%計算），倘依「就業保險法」規<br>"
                        nc.InnerHtml += " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                        nc.InnerHtml += " 定請領津貼部分，應依該作業流程辦理。"

                    End If

                Next
[CONTINUE]:
                '表尾
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", 11)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "第 " & i + 1 & " 頁"

                If k + 1 > dt.Rows.Count Then
                    GoTo out
                End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
            Next
out:
        End If

    End Sub

    Function getcname(ByVal val As String, ByVal type As String) As String
        Dim strsql As String = ""
        Dim name As String = ""
        Select Case type
            Case "id"
                strsql = "select name from Key_Identity where identityid='" & val & "'"
            Case "un"
                strsql = "select orgname from sub_org where orgid = '" & val & "'"
        End Select
        If strsql <> "" Then
            name = DbAccess.ExecuteScalar(strsql, objconn)
        End If
        Return name
    End Function

End Class