Partial Class SD_05_008_E
    Inherits AuthBasePage

    'select * from Stud_ResultStudData where rownum <=10
    'Select Q12 From Stud_ResultTwelveData Where rownum <=10
    'Select IdentityID From Stud_ResultIdentData Where rownum <=10
    Const cst_search As String = "search"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Me.IsPostBack Then
            Button1.Attributes("onclick") = "javascript:return search()"

            TPlan = TIMS.Get_TPlan(TPlan, , , , , objconn)
            Call TIMS.SetvDistID(Me.DistID, Mode.SelectedValue, objconn)
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean

        Dim v_Mode As String = TIMS.GetListValue(Mode)
        Select Case v_Mode
            Case "1", "2"
            Case Else
                Errmsg += "署/非署屬 為必選，請重新選擇!!" & vbCrLf
        End Select

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        If start_date.Text = "" Then
            Errmsg += "起始結訓日期 為必填，請重新填寫!!" & vbCrLf
        End If
        If end_date.Text = "" Then
            Errmsg += "迄止結訓日期 為必填，請重新填寫!!" & vbCrLf
        End If
        If (Errmsg <> "") Then Return False

        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        Errmsg = ""
        If Not TIMS.IsDate1(start_date.Text) Then
            Errmsg += "起始結訓日期，日期格式有誤，請重新填寫!!" & vbCrLf
        End If
        If Not TIMS.IsDate1(end_date.Text) Then
            Errmsg += "迄止結訓日期，日期格式有誤，請重新填寫!!" & vbCrLf
        End If
        If (Errmsg <> "") Then Return False

        Return True
    End Function

    '匯出
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim v_Mode As String = TIMS.GetListValue(Mode)
        'Dim SqlStr As String = ""
        '取學員資料
        'select * from ID_StatistDist where Type = 1 order by DistID
        'select * from ID_StatistDist where Type = 0 order by statID

        Dim sSql As String = " SELECT "
        '1:署(局)屬 2:/非署(局)屬
        sSql &= If(v_Mode = "1", " format(cc.FTDATE,'yyyy/MM/dd') ResultDate", " format(S1.ResultDate,'yyyy/MM/dd') ResultDate") & vbCrLf
        '(轉換如下：)
        sSql &= " ,CASE S1.UnitCode WHEN '001' THEN '001'" & vbCrLf
        sSql &= " WHEN '002' THEN '002'" & vbCrLf
        sSql &= " WHEN '003' THEN '003'" & vbCrLf
        sSql &= " WHEN '004' THEN '004'" & vbCrLf
        sSql &= " WHEN '005' THEN '008'" & vbCrLf
        sSql &= " WHEN '006' THEN '009'" & vbCrLf
        sSql &= " WHEN '007' THEN '005'" & vbCrLf
        sSql &= " WHEN '008' THEN '006'" & vbCrLf
        sSql &= " WHEN '009' THEN '007'" & vbCrLf
        sSql &= " WHEN '010' THEN '010'" & vbCrLf
        sSql &= " WHEN '011' THEN '011' ELSE S1.UnitCode END UnitCode" & vbCrLf
        sSql &= " ,S1.TPlanID,S1.Trainice,S1.TrainCommend1" & vbCrLf
        sSql &= " ,S1.SchoolTime,S1.ResultCount,S1.TrainingTHour" & vbCrLf
        sSql &= " ,S2.DLID,S2.SubNo,S2.StudentID" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(S2.StdPID) STDPID_MK,S2.StdPID" & vbCrLf
        sSql &= " ,S2.Sex" & vbCrLf
        '1:署(局)屬 2:/非署(局)屬
        Dim vDATAS1 As String = If(v_Mode = "1", "cc.FTDATE", "S1.ResultDate")
        sSql &= String.Format(" ,datepart(year,{0})-S2.BirthYear Birth", vDATAS1) & vbCrLf
        sSql &= " ,S2.DegreeID" & vbCrLf
        sSql &= " ,S2.MilitaryID" & vbCrLf
        sSql &= " ,S2.Q7,S2.Q8,S2.Q9,S2.Q9Y,S2.Q10" & vbCrLf
        sSql &= " ,S2.Q11,S2.Q11N" & vbCrLf
        sSql &= " ,S2.Q12A,S2.Q12B" & vbCrLf
        sSql &= " ,S2.Q12V1,S2.Q12V2,S2.Q12V3,S2.Q12V4,S2.Q12V5" & vbCrLf
        sSql &= " ,vt.TMKEY TMID" & vbCrLf
        sSql &= " FROM STUD_DATALID S1" & vbCrLf
        sSql &= " JOIN STUD_RESULTSTUDDATA S2 ON S1.DLID = S2.DLID" & vbCrLf
        If v_Mode = "1" Then '署(局)屬
            sSql &= " JOIN CLASS_CLASSINFO cc ON S1.OCID = cc.OCID" & vbCrLf
            sSql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.SOCID = s2.SOCID AND cs.StudStatus NOT IN (2,3)" & vbCrLf  '排除離退 BY AMU 201508 
        End If
        sSql &= " JOIN VIEW_TRAINTYPE vt on vt.TMID=S1.TMID" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf

        If v_Mode = "1" Then '署(局)屬
            sSql &= " AND cc.FTDate >= " & TIMS.To_date(Me.start_date.Text) & vbCrLf
            sSql &= " AND cc.FTDate <= " & TIMS.To_date(Me.end_date.Text) & vbCrLf
        Else '非署(局)屬
            sSql &= " AND S1.ResultDate >= " & TIMS.To_date(Me.start_date.Text) & vbCrLf
            sSql &= " AND S1.ResultDate <= " & TIMS.To_date(Me.end_date.Text) & vbCrLf
        End If

        Dim Dist As String = ""
        For i As Integer = 0 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                Dist &= String.Concat(If(Dist <> "", ",", ""), "'", Me.DistID.Items(i).Value, "'")
            End If
        Next
        'newDist = Mid(Dist, 1, Dist.Length - 1)
        If Dist <> "" Then sSql &= " AND S1.UnitCode IN (" & Dist & ") " & vbCrLf

        Dim TPlanID As String = ""
        For i As Integer = 0 To Me.TPlan.Items.Count - 1
            If Me.TPlan.Items(i).Selected Then
                TPlanID &= String.Concat(If(TPlanID <> "", ",", ""), "'", Me.TPlan.Items(i).Value, "'")
            End If
        Next
        'newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        If TPlanID <> "" Then sSql &= " AND S1.TPlanID IN (" & TPlanID & ") " & vbCrLf


        sSql &= " ORDER BY S1.TPlanID ,S1.DLID ,S1.OCID " & vbCrLf
        sSql &= "  ,S2.SOCID ,S2.StudentID ,S2.SubNO " & vbCrLf
        Dim FName As String = ""
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sSql, objconn)

        Call ExportX1(dt)
    End Sub

    Sub ExportX1(ByRef dt As DataTable)
        Dim sFileName1 As String = String.Concat("結訓學員資料", TIMS.GetDateNo2())
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("big5")
        ''Dim Y, M, D As String
        'Dim YY As String = ""
        'Dim MM As String = ""
        'Dim DD As String = ""
        'Dim UnitCode As String = ""

        'YY = DateTime.Today.Year()
        'MM = DateTime.Today.Month()
        'DD = DateTime.Today.Day()

        'FName = YY & MM & DD & ".csv"

        ''提示使用者是否要儲存檔案
        'Response.AddHeader("content-disposition", "attachment; filename=" & FName)

        ''直接在ie內頁顯示內容
        ''Response.AddHeader("content-disposition", "inline; filename=" & FName)

        ''文件內容指定為excel
        ''Response.ContentType = "application/octet-stream; charset=big5"
        'Response.ContentType = "application/vnd.ms-excel;charset=Big5"

        Dim str_OutPutStr As String = ""
        str_OutPutStr = ""
        str_OutPutStr &= "結訓日期" 'ResultDate (strDate)
        str_OutPutStr &= ",訓練機構" 'UnitCode
        str_OutPutStr &= ",訓練計劃" 'TPlanID
        str_OutPutStr &= ",訓練職類" 'TMID
        str_OutPutStr &= ",訓練性質" 'Trainice
        str_OutPutStr &= ",委託訓練" 'TrainCommend1
        str_OutPutStr &= ",上課時段" 'SchoolTime
        str_OutPutStr &= ",結訓人數" 'ResultCount
        str_OutPutStr &= ",訓練總時數" 'TrainingTHour
        str_OutPutStr &= ",學號" 'StudentID
        str_OutPutStr &= ",身分證號碼" 'StdPID/STDPID_MK
        str_OutPutStr &= ",性別" 'Sex
        str_OutPutStr &= ",年齡" 'Birth
        str_OutPutStr &= ",學歷" 'DegreeID
        str_OutPutStr &= ",兵役" 'MilitaryID
        str_OutPutStr &= ",學員身分" 'IdentityID
        str_OutPutStr &= ",參訓動機" 'Q7
        str_OutPutStr &= ",結訓後動向" 'Q8
        str_OutPutStr &= ",職訓前一個月有否工作" 'Q9
        str_OutPutStr &= ",職訓前一個月有否尋找工作" 'Q10
        str_OutPutStr &= ",您覺得參加本次訓練後對日後尋找工作幫助的程度" 'Q11.您覺得參加本次訓練後，對日後尋找工作幫助的程度
        'str_OutPutStr &= ",參加本次訓練後是否找到工作"
        str_OutPutStr &= ",參加本次訓練後是否覺得滿意" 'Q12A
        str_OutPutStr &= ",參加本次訓練覺得不滿意需改進為何" 'Q12B
        str_OutPutStr &= ",參訓職類符合就業市場需求" 'Q12V1
        str_OutPutStr &= ",教學課程安排" 'Q12V2
        str_OutPutStr &= ",訓練師專業及熱忱" 'Q12V3
        str_OutPutStr &= ",訓練設備符合產業需求" 'Q12V4
        str_OutPutStr &= ",訓練時數" 'Q12V5
        'str_OutPutStr &= vbCrLf

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">") 'strHTML &= "<table>"
        strHTML &= TIMS.Get_TABLETR(str_OutPutStr)


        '要輸出的ht內容
        'strUtf8 = System.Text.Encoding.Unicode.GetBytes(BarcodeBody)
        'strBig5 = System.Text.Encoding.Convert(System.Text.Encoding.Unicode, System.Text.Encoding.Default, strUtf8)
        'Response.BinaryWrite(System.Text.Encoding.Convert(System.Text.Encoding.Unicode, System.Text.Encoding.Default, System.Text.Encoding.Unicode.GetBytes(str_OutPutStr)))

        For Each dr As DataRow In dt.Rows
            Dim Q9 As String = If(Convert.ToString(dr("Q9")) = "Y", Convert.ToString(dr("Q9Y")), Convert.ToString(dr("Q9")))
            Dim Q11N As String = If(Convert.ToString(dr("Q11")) = "Y", Convert.ToString(dr("Q11")), Convert.ToString(dr("Q11N")))
            Dim Q12 As String = ""
            Dim strDate As String = ""
            Dim dt1 As DataTable
            Dim dt2 As DataTable

            Dim Sql As String = ""
            Dim StrSql As String = ""
            ''參訓身分
            Dim IdentityID As String = ""
            IdentityID = ""

            StrSql = "" & vbCrLf
            StrSql &= " SELECT cs.IdentityID " & vbCrLf
            StrSql &= " FROM Stud_ResultStudData sr " & vbCrLf
            StrSql &= " JOIN Class_StudentsOfClass cs ON sr.SOCID = cs.SOCID " & vbCrLf
            StrSql &= " WHERE 1=1 " & vbCrLf
            StrSql &= " AND sr.DLID = '" & dr("DLID") & "' " & vbCrLf
            StrSql &= " AND sr.SubNo = '" & dr("SubNo") & "' " & vbCrLf
            'StrSql &= " ORDER BY sr.DLID,sr.SubNo" & vbCrLf
            Dim drIden As DataRow = DbAccess.GetOneRow(StrSql, objconn)
            If drIden IsNot Nothing Then
                Dim arIden As String() = drIden("IdentityID").ToString.Split(",")
                For J As Integer = 0 To arIden.Length - 1
                    IdentityID += arIden(J).ToString
                Next
            Else
                '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
                StrSql = " SELECT IdentityID FROM Stud_ResultIdentData WHERE DLID = '" & dr("DLID") & "' AND SubNo = '" & dr("SubNo") & "' "
                dt1 = DbAccess.GetDataTable(StrSql, objconn)
                IdentityID = ""
                For J As Integer = 0 To dt1.Rows.Count - 1
                    IdentityID += dt1.Rows(J)(0)
                Next
            End If

            Sql = " SELECT Q12 FROM STUD_RESULTTWELVEDATA WHERE DLID = '" & dr("DLID") & "' AND SubNo = '" & dr("SubNo") & "' ORDER BY Q12 "
            dt2 = DbAccess.GetDataTable(Sql, objconn)
            'Dim x As Integer
            Q12 = ""
            For x As Integer = 0 To dt2.Rows.Count - 1
                Q12 &= dt2.Rows(x)(0).ToString
            Next

            Dim strLine As System.Text.StringBuilder = New System.Text.StringBuilder
            strDate = Convert.ToString(dr("ResultDate"))
            strLine.Append(strDate).Append(",")
            strLine.Append(Chr(128) & dr("UnitCode").ToString()).Append(",")
            strLine.Append(Chr(128) & dr("TPlanID").ToString()).Append(",")
            strLine.Append(dr("TMID")).Append(",")
            strLine.Append(dr("Trainice")).Append(",")
            strLine.Append(dr("TrainCommend1")).Append(",")
            strLine.Append(dr("SchoolTime")).Append(",")
            strLine.Append(dr("ResultCount")).Append(",")
            strLine.Append(dr("TrainingTHour")).Append(",")
            strLine.Append(dr("StudentID")).Append(",")
            strLine.Append(dr("STDPID_MK")).Append(",")
            'strLine.Append(dr("StdPID")).Append(",")
            strLine.Append(dr("Sex")).Append(",")
            strLine.Append(dr("Birth")).Append(",")
            strLine.Append(dr("DegreeID")).Append(",")
            strLine.Append(dr("MilitaryID")).Append(",")
            strLine.Append(Chr(128) & IdentityID.ToString()).Append(",") '€
            strLine.Append(dr("Q7")).Append(",")
            strLine.Append(dr("Q8")).Append(",")
            strLine.Append(Q9).Append(",")
            strLine.Append(dr("Q10")).Append(",")
            'Q11(Q11N)
            strLine.Append(Q11N).Append(",")
            'Q12A
            strLine.Append(If(Convert.ToString(dr("Q12A")) <> "", dr("Q12A"), " ")).Append(",")
            'Q12 2013前原資料
            'Q12B 2014年資料。
            If Convert.ToString(dr("Q12B")) <> "" Then
                Q12 = Convert.ToString(dr("Q12B"))
                If Q12.IndexOf(",") > -1 Then Q12 = Replace(dr("Q12B"), ",", "")
                strLine.Append(Chr(128) & Q12.ToString()).Append(",") '€
            Else
                strLine.Append(Chr(128) & Q12.ToString()).Append(",") '€
            End If

            strLine.Append(dr("Q12V1")).Append(",")
            strLine.Append(dr("Q12V2")).Append(",")
            strLine.Append(dr("Q12V3")).Append(",")
            strLine.Append(dr("Q12V4")).Append(",")
            strLine.Append(dr("Q12V5"))
            'strLine.Append(vbCrLf)
            strHTML &= TIMS.Get_TABLETR(strLine.ToString())
            'Response.BinaryWrite(System.Text.Encoding.Convert(System.Text.Encoding.Unicode, System.Text.Encoding.Default, System.Text.Encoding.Unicode.GetBytes(strLine.ToString())))
        Next
        '結訓學員資料匯出
        'Response.End()  '結束程式執行
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") '  Response.End()
    End Sub

    Private Sub Mode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Mode.SelectedIndexChanged
        Call TIMS.SetvDistID(Me.DistID, Mode.SelectedValue, objconn)
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Session(cst_search) IsNot Nothing Then Session(cst_search) = Nothing '(清除)
        TIMS.Utl_Redirect1(Me, "SD_05_008.aspx?ID=" & Request("ID"))
    End Sub
End Class