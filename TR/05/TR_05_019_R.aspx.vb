Partial Class TR_05_019_R
    Inherits AuthBasePage

    Const cst_在職 As String = "在職"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
            If sm.UserInfo.DistID <> "000" Then '非署(局)
                Common.SetListItem(DistID, sm.UserInfo.DistID)
                Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            End If

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
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

    End Sub

    '匯出 Response(融合式訓練辦理情形)
    Sub ExpReport1(ByRef dt As DataTable)
        Dim strPlanType As String '計畫分類 (1:自辦、2:委外)
        strPlanType = Me.rblPlanType.SelectedItem.Text
        Dim strPropertyID As String = cst_在職 '訓練類別 (0:職前、1:在職)
        'Dim strPropertyID As String '訓練類別 (0:職前、1:在職)
        'strPropertyID = Me.rblPropertyID.SelectedItem.Text
        'Dim sPropertyID As String '訓練類別 (0:職前、1:在職)
        'sPropertyID = Me.rblPropertyID.SelectedValue

        'Dim strTitle1 As String = sFileName1
        Dim sFileName1 As String = "" '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        sFileName1 = Me.rblType1.SelectedItem.Text

        '套CSS值
        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td rowspan=""3"">訓練類別</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">訓練班數<BR />(" & strPlanType & "訓練)</td>" & vbTab
        ExportStr &= "<td colspan=""8"">參訓人數</td>" & vbTab '一般學員(A) '身障學員(B)
        ExportStr &= "<td colspan=""7"">結訓人數</td>" & vbTab
        'sPropertyID 0:職前
        'If sPropertyID = "0" Then ExportStr &= "<td  colspan=""2"">就業情形</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '第2行
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""3"">一般學員(A)</td>" & vbTab '一般學員(A)
        ExportStr &= "<td colspan=""3"">身障學員(B)</td>" & vbTab '身障學員(B)
        ExportStr &= "<td rowspan=""2"">合計<BR />(C=A+B)</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">身障學員<BR />參訓比率<BR />(D=B/C)</td>" & vbTab
        ExportStr &= "<td colspan=""3"">一般學員(E)</td>" & vbTab '一般學員(E)
        ExportStr &= "<td colspan=""3"">身障學員(F)</td>" & vbTab '身障學員(F)
        ExportStr &= "<td rowspan=""2"">合計<BR />(G=E+F)</td>" & vbTab
        'If sPropertyID = "0" Then
        '    ExportStr &= "<td rowspan=""2"">一般學員<BR />就業人數</td>" & vbTab
        '    ExportStr &= "<td rowspan=""2"">身障學員<BR />就業人數</td>" & vbTab
        'End If
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '第3行
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td>男<BR />(A1)</td>" & vbTab '一般學員(A)
        ExportStr &= "<td>女<BR />(A2)</td>" & vbTab '一般學員(A)
        ExportStr &= "<td>合計<BR />(A=A1+A2)</td>" & vbTab

        ExportStr &= "<td>男<BR />(B1)</td>" & vbTab '身障學員(B)
        ExportStr &= "<td>女<BR />(B2)</td>" & vbTab '身障學員(B)
        ExportStr &= "<td>合計<BR />(B=B1+B2)</td>" & vbTab


        ExportStr &= "<td>男<BR />(E1)</td>" & vbTab '一般學員(E)
        ExportStr &= "<td>女<BR />(E2)</td>" & vbTab '一般學員(E)
        ExportStr &= "<td>合計<BR />(E=E1+E2)</td>" & vbTab

        ExportStr &= "<td>男<BR />(F1)</td>" & vbTab '身障學員(F)
        ExportStr &= "<td>女<BR />(F2)</td>" & vbTab '身障學員(F)
        ExportStr &= "<td>合計<BR />(F=F1+F2)</td>" & vbTab

        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows

            Dim d_D As Double
            d_D = 0
            If Convert.ToString(dr("C")) <> "" _
                AndAlso Convert.ToString(dr("C")) <> "0" Then '避免除0錯誤

                d_D = CDbl(Val(dr("B"))) / CDbl(Val(dr("C")))
            End If

            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            ExportStr &= "<td>" & strPropertyID & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("classCnt")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("C")) & "</td>" & vbTab
            ExportStr &= "<td>" & d_D.ToString("0.00") & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("E1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("G")) & "</td>" & vbTab
            'If sPropertyID = "0" Then
            '    ExportStr &= "<td>" & Convert.ToString(dr("workNum1")) & "</td>" & vbTab
            '    ExportStr &= "<td>" & Convert.ToString(dr("workNum2")) & "</td>" & vbTab
            'End If

            ExportStr += "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Common.RespWrite(Me, "</table>")
        'Common.RespWrite(Me, "</body>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出 Response(融合式訓練職類統計)
    Sub ExpReport2(ByRef dt As DataTable)
        Dim strPlanType As String '計畫分類 (1:自辦、2:委外)
        strPlanType = Me.rblPlanType.SelectedItem.Text
        'Dim strPropertyID As String '訓練類別 (0:職前、1:在職)
        'strPropertyID = Me.rblPropertyID.SelectedItem.Text
        Dim sPropertyID As String '訓練類別 (0:職前、1:在職)
        sPropertyID = "1" '1:在職 'Me.rblPropertyID.SelectedValue
        'Dim strTitle1 As String = "" '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        'strTitle1 = Me.rblType1.SelectedItem.Text

        Dim sFileName1 As String = rblType1.SelectedItem.Text

        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""

        '建立抬頭
        ExportStr = ""
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td rowspan=""2"">項次</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">職類名稱<BR />(" & strPlanType & "訓練)</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">訓練類別</td>" & vbTab '一般學員(A) '身障學員(B)
        ExportStr &= "<td rowspan=""2"">訓練期程</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">訓練時數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">參訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">結訓人數</td>" & vbTab
        If sPropertyID = "0" Then ExportStr &= "<td colspan=""2"">就業情形</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td>一般學員</td>" & vbTab
        ExportStr &= "<td>身障學員</td>" & vbTab
        ExportStr &= "<td>障礙別</td>" & vbTab
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "<td>一般學員</td>" & vbTab
        ExportStr &= "<td>身障學員</td>" & vbTab
        ExportStr &= "<td>障礙別</td>" & vbTab
        ExportStr &= "<td>合計</td>" & vbTab
        'If sPropertyID = "0" Then ExportStr &= "<td>就業人數</td>" & vbTab
        If sPropertyID = "0" Then
            ExportStr &= "<td>一般學員<BR />就業人數</td>" & vbTab
            ExportStr &= "<td>身障學員<BR />就業人數</td>" & vbTab
        End If
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        Dim iRow As Integer = 0
        iRow = 0
        For Each dr As DataRow In dt.Rows
            iRow += 1

            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            ExportStr &= "<td>" & CStr(iRow) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("PropertyID")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("STDate")) & "~" & Convert.ToString(dr("FTDate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("THours")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B")) & "</td>" & vbTab
            '障礙別
            ExportStr &= "<td>" & Convert.ToString(dr("B2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("E")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F")) & "</td>" & vbTab
            '障礙別
            ExportStr &= "<td>" & Convert.ToString(dr("F2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("G")) & "</td>" & vbTab
            If sPropertyID = "0" Then
                ExportStr &= "<td>" & Convert.ToString(dr("workNum1")) & "</td>" & vbTab
                ExportStr &= "<td>" & Convert.ToString(dr("workNum2")) & "</td>" & vbTab
            End If

            ExportStr += "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出 Response((專班)辦理情形)
    Sub ExpReport3(ByRef dt As DataTable)
        'Dim strPlanType As String '計畫分類 (1:自辦、2:委外)
        'strPlanType = Me.rblPlanType.SelectedItem.Text
        'Dim strPropertyID As String '訓練類別 (0:職前、1:在職)
        'strPropertyID = Me.rblPropertyID.SelectedItem.Text
        Dim sPropertyID As String '訓練類別 (0:職前、1:在職)
        sPropertyID = "1" '1:在職 'Me.rblPropertyID.SelectedValue
        'Dim strTitle1 As String = "" '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        'strTitle1 = Me.rblType1.SelectedItem.Text

        Dim sFileName1 As String = rblType1.SelectedItem.Text
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""
        '建立抬頭
        ExportStr = ""
        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td rowspan=""2"">訓練類別</td>" & vbTab
        ExportStr &= "<td colspan=""5"">開班情形</td>" & vbTab
        ExportStr &= "<td colspan=""2"">結訓情形</td>" & vbTab
        If sPropertyID = "0" Then ExportStr &= "<td>就業情形</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td>訓練班別</td>" & vbTab
        ExportStr &= "<td>訓練單位</td>" & vbTab
        ExportStr &= "<td>訓練期程</td>" & vbTab
        ExportStr &= "<td>訓練時數</td>" & vbTab
        ExportStr &= "<td>參訓人數</td>" & vbTab
        ExportStr &= "<td>已結訓班數</td>" & vbTab
        ExportStr &= "<td>結訓人數</td>" & vbTab
        If sPropertyID = "0" Then ExportStr &= "<td>就業人數</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        Dim iRow As Integer = 0
        iRow = 0
        For Each dr As DataRow In dt.Rows
            iRow += 1

            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            If iRow = 1 Then
                ExportStr &= "<td rowspan=""" & CStr(dt.Rows.Count) & """>" & Convert.ToString(dr("PropertyID")) & "</td>" & vbTab
            End If
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("OrgName")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("THours")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("STDate")) & " ~ " & Convert.ToString(dr("FTDate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("trainNum")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("isClose")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("closeNum")) & "</td>" & vbTab
            If sPropertyID = "0" Then ExportStr &= "<td>" & Convert.ToString(dr("workNum")) & "</td>" & vbTab
            ExportStr += "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    Function Get_WC1SQL() As String
        'sType1 '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        Dim v_Syear As String = TIMS.GetListValue(Syear)

        'Dim rst As String = ""
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ip.DISTID ,ip.PLANID" & vbCrLf
        sql &= " ,ip.DISTNAME" & vbCrLf
        sql &= " ,ip.PLANNAME" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,vt.TRAINNAME" & vbCrLf
        sql &= " ,k1.HOURRANNAME" & vbCrLf
        sql &= " ,cc.STDATE" & vbCrLf
        sql &= " ,cc.FTDATE" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        sql &= " ,cc.THOURS" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE vt on vt.TMID=cc.TMID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_HOURRAN k1 on k1.HRID=cc.TPeriod" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and cc.IsSuccess='Y'" & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf

        If v_Syear <> "" Then
            sql &= " and ip.Years='" & v_Syear & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= " and cc.STDate >=" & TIMS.To_date(Me.STDate1.Text) & vbCrLf '" & Me.STDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " and cc.STDate <=" & TIMS.To_date(Me.STDate2.Text) & vbCrLf '('" & Me.STDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >=" & TIMS.To_date(Me.FTDate1.Text) & vbCrLf '('" & Me.FTDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <=" & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '('" & Me.FTDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If

        Return sql
    End Function

    'SQL:1(融合式訓練辦理情形)
    Function LoadData1() As DataTable
        Dim rst As DataTable
        'sType1 '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)

        Dim sql As String = ""
        sql = "" & vbCrLf
        'sql &= " select  ip.distid ,ip.planid "
        'sql &= " ,ip.distname " & vbCrLf '轄區	
        'sql &= " ,ip.planname" & vbCrLf '訓練計畫	" & vbCrLf
        'sql &= " ,oo.orgname " & vbCrLf '訓練機構名稱	" & vbCrLf
        'sql &= " ,cc.classcname " & vbCrLf '班別名稱	" & vbCrLf
        'sql &= " ,vt.trainName  " & vbCrLf '訓練職類	" & vbCrLf
        'sql &= " ,case when ISNULL(ip.PropertyID,0)=0 then '職前' " & vbCrLf
        'sql &= " 	when ISNULL(ip.PropertyID,0)=1 then '在職' end PropertyID" & vbCrLf '訓練性質	" & vbCrLf
        'sql &= " ,k1.hourRanName  hourRanName" & vbCrLf '訓練時段	" & vbCrLf
        'sql &= " ,convert(varchar,cc.stdate,111) STDate" & vbCrLf ' '開訓日期'" & vbCrLf
        'sql &= " ,convert(varchar,cc.ftdate,111) FTDate" & vbCrLf ' '結訓日期'" & vbCrLf
        'sql &= " ,cc.TNum " & vbCrLf '招生人數
        'sql &= " ,cc.THours " & vbCrLf '時數	
        'sql &= " ,ISNULL(gcs.openNum,0) openNum" & vbCrLf ' 開訓人數" & vbCrLf
        'sql &= " ,ISNULL(gcs.closeNum,0) closeNum" & vbCrLf '結訓人數" & vbCrLf
        'sql &= " ,ISNULL(gcs.jobNum,0) jobNum" & vbCrLf '就業人數" & vbCrLf
        ''sql &= " " & vbCrLf

        sql &= " select ISNULL(COUNT(cc.ocid),0) classCnt" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.A1,0)) A1" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.A2,0)) A2" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.A,0)) A" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.B1,0)) B1" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.B2,0)) B2" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.B,0)) B" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.C,0)) C" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.E1,0)) E1" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.E2,0)) E2" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.E,0)) E" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.F1,0)) F1" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.F2,0)) F2" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.F,0)) F" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.G,0)) G" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.workNum1,0)) workNum1" & vbCrLf
        sql &= " ,SUM(ISNULL(gcs.workNum2,0)) workNum2" & vbCrLf

        sql &= " from dbo.class_classinfo cc" & vbCrLf
        sql &= " join dbo.plan_planinfo pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno " & vbCrLf
        sql &= " join dbo.view_plan ip on ip.planid =cc.planid" & vbCrLf
        sql &= " join dbo.view_TrainType vt on vt.TMID=cc.TMID" & vbCrLf
        sql &= " join dbo.org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " left join dbo.Key_HourRan k1 on k1.HRID=cc.TPeriod" & vbCrLf
        sql &= " left join (" & vbCrLf
        sql &= " 	select cs.ocid " & vbCrLf
        '一般身分者
        'sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('01') and ss.Sex='M' then 1 end) A1" & vbCrLf
        'sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('01') and ss.Sex='F' then 1 end) A2" & vbCrLf
        'sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('01') and ss.Sex IN ('M','F') then 1 end) A" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID NOT in ('06') and ss.Sex='M' then 1 end) A1" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID NOT in ('06') and ss.Sex='F' then 1 end) A2" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F') then 1 end) A" & vbCrLf
        '身心障礙者
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('06') and ss.Sex='M' then 1 end) B1" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('06') and ss.Sex='F' then 1 end) B2" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('06') and ss.Sex IN ('M','F') then 1 end) B" & vbCrLf
        '合計
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and ss.Sex IN ('M','F') then 1 end) C" & vbCrLf
        '一般身分者
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01') and ss.Sex='M' then 1 end) E1" & vbCrLf
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01') and ss.Sex='F' then 1 end) E2" & vbCrLf
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01') and ss.Sex IN ('M','F') then 1 end) E" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex='M' then 1 end) E1" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex='F' then 1 end) E2" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F') then 1 end) E" & vbCrLf
        '身心障礙者
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex='M' then 1 end) F1" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex='F' then 1 end) F2" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex IN ('M','F') then 1 end) F" & vbCrLf
        ''合計
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01','06') and ss.Sex IN ('M','F') then 1 end) G" & vbCrLf
        ''就業
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01','06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum" & vbCrLf
        '合計
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID IS NOT NULL and ss.Sex IN ('M','F') then 1 end) G" & vbCrLf
        '一般就業
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum1" & vbCrLf
        '身障就業
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum2" & vbCrLf

        sql &= " 	from dbo.Class_StudentsOfClass cs " & vbCrLf
        sql &= " 	join dbo.class_classinfo cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " 	join dbo.stud_studentinfo ss on ss.SID =cs.SID" & vbCrLf
        sql &= " 	join dbo.Stud_SubData ss2 on ss2.sid=ss.sid" & vbCrLf
        sql &= " 	left join dbo.Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= "    and cc.IsSuccess='Y' " & vbCrLf
        sql &= "    and cc.NotOpen='N'" & vbCrLf

        sql &= " 	group by cs.ocid " & vbCrLf
        sql &= " ) gcs on gcs.ocid =cc.ocid " & vbCrLf

        sql &= " where 1=1" & vbCrLf
        sql &= " and cc.IsSuccess='Y' " & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql &= " and ip.Years='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= " and cc.STDate >=" & TIMS.To_date(Me.STDate1.Text) & vbCrLf '" & Me.STDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " and cc.STDate <=" & TIMS.To_date(Me.STDate2.Text) & vbCrLf '('" & Me.STDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >=" & TIMS.To_date(Me.FTDate1.Text) & vbCrLf '('" & Me.FTDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <=" & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '('" & Me.FTDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If

        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    'SQL:2(融合式訓練職類統計)
    Function LoadData2() As DataTable
        Dim rst As DataTable
        'sType1 '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)

        Dim sql As String = ""
        sql = "" & vbCrLf

        sql &= " select  ip.distid ,ip.planid " & vbCrLf
        sql &= " ,ip.distname " & vbCrLf '轄區	
        sql &= " ,ip.planname" & vbCrLf '訓練計畫	" & vbCrLf
        sql &= " ,oo.orgname " & vbCrLf '訓練機構名稱	" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf '班別代號	" & vbCrLf
        sql &= " ,cc.classcname " & vbCrLf '班別名稱	" & vbCrLf
        sql &= " ,vt.trainName  " & vbCrLf '訓練職類	" & vbCrLf
        'sql &= " ,case when ISNULL(ip.PropertyID,0)=0 then '職前' " & vbCrLf
        'sql &= " 	when ISNULL(ip.PropertyID,0)=1 then '在職' end PropertyID" & vbCrLf '訓練性質	" & vbCrLf
        'sql &= " ,'" & Me.rblPropertyID.SelectedItem.Text & "' PropertyID" & vbCrLf '訓練性質	" & vbCrLf
        sql &= " ,'" & cst_在職 & "' PropertyID" & vbCrLf '訓練性質	" & vbCrLf

        sql &= " ,k1.hourRanName  hourRanName" & vbCrLf '訓練時段	" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) STDate" & vbCrLf ' '開訓日期'" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) FTDate" & vbCrLf ' '結訓日期'" & vbCrLf
        sql &= " ,cc.TNum " & vbCrLf '招生人數
        sql &= " ,cc.THours " & vbCrLf '訓練時數	

        sql &= " ,ISNULL(gcs.A,0) A" & vbCrLf
        sql &= " ,ISNULL(gcs.B,0) B" & vbCrLf
        sql &= " ,convert(varchar(3000),null) B2" & vbCrLf
        sql &= " ,ISNULL(gcs.C,0) C" & vbCrLf
        sql &= " ,ISNULL(gcs.E,0) E" & vbCrLf
        sql &= " ,ISNULL(gcs.F,0) F" & vbCrLf
        sql &= " ,convert(varchar(3000),null) F2" & vbCrLf
        sql &= " ,ISNULL(gcs.G,0) G" & vbCrLf
        'sql &= " ,ISNULL(gcs.workNum,0) workNum" & vbCrLf
        sql &= " ,ISNULL(gcs.workNum1,0) workNum1" & vbCrLf
        sql &= " ,ISNULL(gcs.workNum2,0) workNum2" & vbCrLf

        sql &= " from class_classinfo cc " & vbCrLf
        sql &= " join plan_planinfo pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno " & vbCrLf
        sql &= " join view_plan ip on ip.planid =cc.planid" & vbCrLf
        sql &= " join view_TrainType vt on vt.TMID=cc.TMID" & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " left join Key_HourRan k1 on k1.HRID=cc.TPeriod" & vbCrLf
        sql &= " left join (" & vbCrLf
        sql &= " 	select cs.ocid " & vbCrLf
        '一般身分者
        'sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('01') and ss.Sex IN ('M','F') then 1 end) A" & vbCrLf
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F') then 1 end) A" & vbCrLf
        '身心障礙者
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID in ('06') and ss.Sex IN ('M','F') then 1 end) B" & vbCrLf
        '合計
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.MIdentityID IS NOT NULL and ss.Sex IN ('M','F') then 1 end) C" & vbCrLf

        '一般身分者
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01') and ss.Sex IN ('M','F') then 1 end) E" & vbCrLf
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F') then 1 end) E" & vbCrLf
        '身心障礙者
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex IN ('M','F') then 1 end) F" & vbCrLf
        ''合計
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01','06') and ss.Sex IN ('M','F') then 1 end) G" & vbCrLf
        ''就業
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID in ('01','06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum" & vbCrLf

        '合計
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID IS NOT NULL and ss.Sex IN ('M','F') then 1 end) G" & vbCrLf
        ''就業
        'sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        'sql &= "      and cs.MIdentityID IS NOT NULL and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum" & vbCrLf
        '一般就業
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID NOT in ('06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum1" & vbCrLf
        '身心障礙者就業
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      and cs.MIdentityID in ('06') and ss.Sex IN ('M','F')  and sg3.IsGetJob='1' then 1 end) workNum2" & vbCrLf

        sql &= " 	from Class_StudentsOfClass cs " & vbCrLf
        sql &= " 	join class_classinfo cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " 	join stud_studentinfo ss on ss.SID =cs.SID" & vbCrLf
        sql &= " 	join Stud_SubData ss2 on ss2.sid=ss.sid" & vbCrLf
        sql &= " 	left join Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= "    and cc.IsSuccess='Y' " & vbCrLf
        sql &= "    and cc.NotOpen='N'" & vbCrLf

        sql &= " 	group by cs.ocid " & vbCrLf
        sql &= " ) gcs on gcs.ocid =cc.ocid " & vbCrLf

        sql &= " where 1=1" & vbCrLf
        sql &= " and cc.IsSuccess='Y' " & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql &= " and ip.Years='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= " and cc.STDate >=" & TIMS.To_date(Me.STDate1.Text) & vbCrLf '" & Me.STDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " and cc.STDate <=" & TIMS.To_date(Me.STDate2.Text) & vbCrLf '('" & Me.STDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >=" & TIMS.To_date(Me.FTDate1.Text) & vbCrLf '('" & Me.FTDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <=" & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '('" & Me.FTDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        sql &= " ORDER BY ip.distid , ip.planid ,cc.classcname" & vbCrLf

        rst = DbAccess.GetDataTable(sql, objconn)


        Dim dt2 As DataTable
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,cs.StudStatus" & vbCrLf
        sql &= " ,ss2.HandTypeID HandTypeID " & vbCrLf
        sql &= " ,ss2.HandTypeID2 " & vbCrLf '多選
        sql &= " ,r.Name HandTypeName " & vbCrLf
        '開訓B: 1:1代障礙別。 2:2代障礙別。
        sql &= " ,case when cs.socid is not null and ss.Sex IN ('M','F') and ss2.HandTypeID is not null then 1 " & vbCrLf
        sql &= "   when cs.socid is not null and ss.Sex IN ('M','F') and ss2.HandTypeID2 is not null then 2 end B " & vbCrLf
        '結訓F: 1:1代障礙別。 2:2代障礙別。
        sql &= " ,case when (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate())) and ss.Sex IN ('M','F') and ss2.HandTypeID is not null then 1 " & vbCrLf
        sql &= "   when (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate())) and ss.Sex IN ('M','F') and ss2.HandTypeID2 is not null then 2 end F " & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.SID =cs.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA ss2 on ss2.sid=ss.sid" & vbCrLf
        sql &= " LEFT JOIN KEY_HANDICATTYPE r ON ss2.HandTypeID=r.HandTypeID" & vbCrLf
        'sql &= " left JOIN Key_HandicatType2 r2 ON ss2.HandTypeID2=r2.HandTypeID2" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and cs.MIdentityID in ('06')" & vbCrLf
        sql &= " and cc.IsSuccess='Y'" & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql &= " and ip.Years='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= " and cc.STDate >=" & TIMS.To_date(Me.STDate1.Text) & vbCrLf '" & Me.STDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " and cc.STDate <=" & TIMS.To_date(Me.STDate2.Text) & vbCrLf '('" & Me.STDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >=" & TIMS.To_date(Me.FTDate1.Text) & vbCrLf '('" & Me.FTDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <=" & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '('" & Me.FTDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        dt2 = DbAccess.GetDataTable(sql, objconn)

        For i As Integer = 0 To rst.Rows.Count - 1
            Dim dr As DataRow = rst.Rows(i)
            Dim dt3 As DataTable
            Dim ValueB2 As String = ""
            Dim ValueF2 As String = ""
            Dim ValueB2b As String = ""
            Dim ValueF2b As String = ""

            Dim filter1 As String = "OCID=" & dr("OCID") & " AND B=1"
            Dim filter2 As String = "OCID=" & dr("OCID") & " AND F=1"
            Dim filter1b As String = "OCID=" & dr("OCID") & " AND B=2"
            Dim filter2b As String = "OCID=" & dr("OCID") & " AND F=2"
            If dt2.Select(filter1).Length > 0 Then
                For j As Integer = 0 To dt2.Select(filter1).Length - 1
                    Dim dr2 As DataRow = dt2.Select(filter1)(j)
                    If ValueB2.IndexOf(dr2("HandTypeName")) = -1 Then
                        If ValueB2 <> "" Then ValueB2 &= ","
                        ValueB2 &= Convert.ToString(dr2("HandTypeName"))
                    End If
                Next
            End If
            If dt2.Select(filter1b).Length > 0 Then
                For j As Integer = 0 To dt2.Select(filter1b).Length - 1
                    Dim dr2 As DataRow = dt2.Select(filter1b)(j)
                    If ValueB2b.IndexOf(dr2("HandTypeID2")) = -1 Then
                        If ValueB2b <> "" Then ValueB2 &= ","
                        ValueB2b &= "'" & Convert.ToString(dr2("HandTypeID2")) & "'"
                    End If
                Next
                sql = "SELECT NAME FROM KEY_HANDICATTYPE2 WHERE HANDTYPEID2 IN (" & ValueB2b & ")"
                dt3 = DbAccess.GetDataTable(sql, objconn)
                For j As Integer = 0 To dt3.Rows.Count - 1
                    Dim dr3 As DataRow = dt3.Rows(j)
                    If ValueB2.IndexOf(dr3("Name")) = -1 Then
                        If ValueB2 <> "" Then ValueB2 &= ","
                        ValueB2 &= Convert.ToString(dr3("Name"))
                    End If
                Next
            End If

            If dt2.Select(filter2).Length > 0 Then
                For j As Integer = 0 To dt2.Select(filter2).Length - 1
                    Dim dr2 As DataRow = dt2.Select(filter2)(j)
                    If ValueF2.IndexOf(dr2("HandTypeName")) = -1 Then
                        If ValueF2 <> "" Then ValueB2 &= ","
                        ValueF2 &= Convert.ToString(dr2("HandTypeName"))
                    End If
                Next
            End If
            If dt2.Select(filter2b).Length > 0 Then
                For j As Integer = 0 To dt2.Select(filter2b).Length - 1
                    Dim dr2 As DataRow = dt2.Select(filter2b)(j)
                    If ValueF2b.IndexOf(dr2("HandTypeID2")) = -1 Then
                        If ValueF2b <> "" Then ValueB2 &= ","
                        ValueF2b &= "'" & Convert.ToString(dr2("HandTypeID2")) & "'"
                    End If
                Next
                sql = "SELECT NAME FROM KEY_HANDICATTYPE2 WHERE HANDTYPEID2 IN (" & ValueF2b & ")"
                dt3 = DbAccess.GetDataTable(sql, objconn)
                For j As Integer = 0 To dt3.Rows.Count - 1
                    Dim dr3 As DataRow = dt3.Rows(j)
                    If ValueF2.IndexOf(dr3("Name")) = -1 Then
                        If ValueF2 <> "" Then ValueB2 &= ","
                        ValueF2 &= Convert.ToString(dr3("Name"))
                    End If
                Next
            End If

            dr("B2") = ValueB2
            dr("F2") = ValueF2
        Next
        Return rst
    End Function

    'SQL:3((專班)辦理情形)
    Function LoadData3() As DataTable
        Dim rst As DataTable
        'sType1 '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        Dim strProperty As String = ""
        strProperty = "在職訓練"
        'Select Case Me.rblPropertyID.SelectedValue '訓練類別 (0:職前、1:在職)
        '    Case "0"
        '        strProperty = "養成訓練"
        '    Case "1"
        '        strProperty = "在職訓練"
        'End Select

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select  ip.distid ,ip.planid " & vbCrLf
        sql &= " ,ip.distname " & vbCrLf '轄區	
        sql &= " ,ip.planname" & vbCrLf '訓練計畫	" & vbCrLf
        sql &= " ,oo.orgname " & vbCrLf '訓練機構名稱	" & vbCrLf
        sql &= " ,cc.classcname " & vbCrLf '班別名稱	" & vbCrLf
        sql &= " ,vt.trainName  " & vbCrLf '訓練職類	" & vbCrLf
        'sql &= " ,case when ISNULL(ip.PropertyID,0)=0 then '職前' " & vbCrLf
        'sql &= " 	when ISNULL(ip.PropertyID,0)=1 then '在職' end PropertyID" & vbCrLf '訓練性質	" & vbCrLf
        sql &= " ,'" & strProperty & "' PropertyID" & vbCrLf '訓練類別	" & vbCrLf

        sql &= " ,k1.hourRanName  hourRanName" & vbCrLf '訓練時段	" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) STDate" & vbCrLf ' '開訓日期'" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) FTDate" & vbCrLf ' '結訓日期'" & vbCrLf
        sql &= " ,cc.TNum " & vbCrLf '招生人數
        sql &= " ,cc.THours " & vbCrLf '訓練時數	
        sql &= " ,case when ISNULL(cc.isClosed,'N')='Y' THEN '1' ELSE '0' END  isClose" & vbCrLf '訓練時數	

        sql &= " ,ISNULL(gcs.openNum,0) openNum" & vbCrLf
        sql &= " ,ISNULL(gcs.trainNum,0) trainNum" & vbCrLf
        sql &= " ,ISNULL(gcs.closeNum,0) closeNum" & vbCrLf
        sql &= " ,ISNULL(gcs.workNum,0) workNum" & vbCrLf

        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno " & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE vt on vt.TMID=cc.TMID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " LEFT JOIN KEY_HOURRAN k1 on k1.HRID=cc.TPeriod" & vbCrLf
        sql &= " left join (" & vbCrLf
        sql &= " 	select cs.ocid " & vbCrLf
        '開訓人數
        sql &= " 	,SUM(CASE WHEN cs.socid is not null then 1 end) openNum" & vbCrLf
        '參訓人數
        sql &= " 	,SUM(CASE WHEN cs.socid is not null and cs.StudStatus not in (2,3) then 1 end) trainNum" & vbCrLf
        '結訓人數
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "      then 1 end) closeNum" & vbCrLf
        '就業人數
        sql &= " 	,SUM(CASE WHEN (cs.StudStatus in (5) or (cs.StudStatus not in (2,3) and cc.ftdate<=getdate()))" & vbCrLf
        sql &= "     and sg3.IsGetJob='1' then 1 end) workNum" & vbCrLf

        sql &= " 	from Class_StudentsOfClass cs" & vbCrLf
        sql &= " 	join class_classinfo cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " 	join stud_studentinfo ss on ss.SID =cs.SID" & vbCrLf
        sql &= " 	join Stud_SubData ss2 on ss2.sid=ss.sid" & vbCrLf
        sql &= " 	left join Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= "    and cc.IsSuccess='Y' " & vbCrLf
        sql &= "    and cc.NotOpen='N'" & vbCrLf

        sql &= " 	group by cs.ocid " & vbCrLf
        sql &= " ) gcs on gcs.ocid =cc.ocid " & vbCrLf

        sql &= " where 1=1" & vbCrLf
        sql &= " and cc.IsSuccess='Y' " & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql &= " and ip.Years='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= " and cc.STDate >=" & TIMS.To_date(Me.STDate1.Text) & vbCrLf '" & Me.STDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " and cc.STDate <=" & TIMS.To_date(Me.STDate2.Text) & vbCrLf '('" & Me.STDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >=" & TIMS.To_date(Me.FTDate1.Text) & vbCrLf '('" & Me.FTDate1.Text & "','yyyy/mm/dd')" & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <=" & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '('" & Me.FTDate2.Text & "','yyyy/mm/dd')" & vbCrLf
        End If

        If DistID1 <> "" Then
            sql &= " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        sql &= " ORDER BY ip.distid , ip.planid ,cc.classcname" & vbCrLf

        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    '匯出EXCEL 
    Private Sub Export1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export1.Click
        Dim strType1 As String = ""  '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        strType1 = Me.rblType1.SelectedValue

        Dim dt As DataTable
        Select Case strType1
            Case "1"
                dt = LoadData1()
                Call ExpReport1(dt)
            Case "2"
                dt = LoadData2()
                Call ExpReport2(dt)
            Case "3"
                dt = LoadData3()
                Call ExpReport3(dt)
        End Select

    End Sub

End Class
