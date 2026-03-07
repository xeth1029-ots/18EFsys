Partial Class CM_03_010_R
    Inherits AuthBasePage

    'Dim conn As SqlConnection = DbAccess.GetConnection
    'Dim sql As String = ""

#Region "Sub"
    '建置報表
    Private Sub crtTable()
        Dim dt As DataTable = getData()
        Dim dr As DataRow = Nothing
        Dim intCnt As Integer = 0

        Dim intM_1 As Integer = 0 '開訓人數(男)
        Dim intF_1 As Integer = 0 '開訓人數(女)
        Dim intM_2 As Integer = 0 '結訓人數(男)
        Dim intF_2 As Integer = 0 '結訓人數(女)

        Dim intBudget01_1 As Integer = 0 '開訓公務人數
        Dim intBudget02_1 As Integer = 0 '開訓就安人數
        Dim intBudget03_1 As Integer = 0 '開訓就保人數
        Dim intBudget98_1 As Integer = 0 '開訓特別預算人數
        Dim intBudget99_1 As Integer = 0 '開訓不補助人數
        Dim intBudgetOdr_1 As Integer = 0 '開訓其他人數

        Dim intBudget01_2 As Integer = 0 '結訓公務人數
        Dim intBudget02_2 As Integer = 0 '結訓就安人數
        Dim intBudget03_2 As Integer = 0 '結訓就保人數
        Dim intBudget98_2 As Integer = 0 '結訓特別預算人數
        Dim intBudget99_2 As Integer = 0 '結訓不補助人數
        Dim intBudgetOdr_2 As Integer = 0 '結訓其他人數

        Dim intIdentity01_1 As Integer = 0 '開訓一般身分者人數
        Dim intIdentity02_1 As Integer = 0 '開訓就業保險被保險人非自願失業者
        Dim intIdentity03_1 As Integer = 0 '開訓負擔家計婦女
        Dim intIdentity04_1 As Integer = 0 '開訓中高齡者
        Dim intIdentity05_1 As Integer = 0 '開訓原住民
        Dim intIdentity06_1 As Integer = 0 '開訓身心障礙者
        Dim intIdentity07_1 As Integer = 0 '開訓生活扶助戶
        Dim intIdentity08_1 As Integer = 0 '開訓急難救助戶
        Dim intIdentity09_1 As Integer = 0 '開訓家庭暴力受害人
        Dim intIdentity10_1 As Integer = 0 '開訓更生受保護人
        Dim intIdentity11_1 As Integer = 0 '開訓農漁民
        Dim intIdentity12_1 As Integer = 0 '開訓屆退官兵(須單位將級以上長官薦送函)
        Dim intIdentity13_1 As Integer = 0 '開訓外籍配偶
        Dim intIdentity14_1 As Integer = 0 '開訓大陸配偶
        Dim intIdentity15_1 As Integer = 0 '開訓遊民
        Dim intIdentity16_1 As Integer = 0 '開訓公營事業民營化員工
        Dim intIdentity17_1 As Integer = 0 '開訓參加職業工會失業者
        Dim intIdentity18_1 As Integer = 0 '開訓921受災戶
        Dim intIdentity19_1 As Integer = 0 '開訓性侵害被害人
        Dim intIdentity20_1 As Integer = 0 '開訓就業保險被保險人自願失業者
        Dim intIdentity21_1 As Integer = 0 '開訓臨時工作津貼人員
        Dim intIdentity22_1 As Integer = 0 '開訓多元就業開發方案人員
        Dim intIdentity23_1 As Integer = 0 '開訓申請失業給付經失業認定者(學習卷專用)
        Dim intIdentity24_1 As Integer = 0 '開訓非失業認定之就業保險失業者(學習卷專用)
        Dim intIdentity25_1 As Integer = 0 '開訓非就業保險失業者(學習卷專用)
        Dim intIdentity26_1 As Integer = 0 '開訓犯罪被害人及其親屬
        Dim intIdentity27_1 As Integer = 0 '開訓長期失業者
        Dim intIdentity28_1 As Integer = 0 '開訓獨力負擔家計者
        Dim intIdentity29_1 As Integer = 0 '開訓天然災害受災民眾
        Dim intIdentity30_1 As Integer = 0 '開訓因應貿易自由化協助勞工
        Dim intIdentity31_1 As Integer = 0 '開訓單一中華民國國籍之無戶籍國民
        Dim intIdentity32_1 As Integer = 0 '開訓取得居留身分之泰國、緬甸、印度或尼泊爾地區無國籍人民
        Dim intIdentityOdr_1 As Integer = 0 '開訓其他

        Dim intIdentity01_2 As Integer = 0 '結訓一般身分者人數
        Dim intIdentity02_2 As Integer = 0 '結訓就業保險被保險人非自願失業者
        Dim intIdentity03_2 As Integer = 0 '結訓負擔家計婦女
        Dim intIdentity04_2 As Integer = 0 '結訓中高齡者
        Dim intIdentity05_2 As Integer = 0 '結訓原住民
        Dim intIdentity06_2 As Integer = 0 '結訓身心障礙者
        Dim intIdentity07_2 As Integer = 0 '結訓生活扶助戶
        Dim intIdentity08_2 As Integer = 0 '結訓急難救助戶
        Dim intIdentity09_2 As Integer = 0 '結訓家庭暴力受害人
        Dim intIdentity10_2 As Integer = 0 '結訓更生受保護人
        Dim intIdentity11_2 As Integer = 0 '結訓農漁民
        Dim intIdentity12_2 As Integer = 0 '結訓屆退官兵(須單位將級以上長官薦送函)
        Dim intIdentity13_2 As Integer = 0 '結訓外籍配偶
        Dim intIdentity14_2 As Integer = 0 '結訓大陸配偶
        Dim intIdentity15_2 As Integer = 0 '結訓遊民
        Dim intIdentity16_2 As Integer = 0 '結訓公營事業民營化員工
        Dim intIdentity17_2 As Integer = 0 '結訓參加職業工會失業者
        Dim intIdentity18_2 As Integer = 0 '結訓921受災戶
        Dim intIdentity19_2 As Integer = 0 '結訓性侵害被害人
        Dim intIdentity20_2 As Integer = 0 '結訓就業保險被保險人自願失業者
        Dim intIdentity21_2 As Integer = 0 '結訓臨時工作津貼人員
        Dim intIdentity22_2 As Integer = 0 '結訓多元就業開發方案人員
        Dim intIdentity23_2 As Integer = 0 '結訓申請失業給付經失業認定者(學習卷專用)
        Dim intIdentity24_2 As Integer = 0 '結訓非失業認定之就業保險失業者(學習卷專用)
        Dim intIdentity25_2 As Integer = 0 '結訓非就業保險失業者(學習卷專用)
        Dim intIdentity26_2 As Integer = 0 '結訓犯罪被害人及其親屬
        Dim intIdentity27_2 As Integer = 0 '結訓長期失業者
        Dim intIdentity28_2 As Integer = 0 '結訓獨力負擔家計者
        Dim intIdentity29_2 As Integer = 0 '結訓天然災害受災民眾
        Dim intIdentity30_2 As Integer = 0 '結訓因應貿易自由化協助勞工
        Dim intIdentity31_2 As Integer = 0 '結訓單一中華民國國籍之無戶籍國民
        Dim intIdentity32_2 As Integer = 0 '結訓取得居留身分之泰國、緬甸、印度或尼泊爾地區無國籍人民
        Dim intIdentityOdr_2 As Integer = 0 '結訓其他

        Call getSch()
        labCityName_1.Text = getCityName(objconn, Me.ViewState("city"))
        labCityName_2.Text = getCityName(objconn, Me.ViewState("city"))

        '人數統計
        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)

            Select Case Convert.ToString(dr("studstatus"))
                Case "5" '結訓
                    '男女統計
                    Select Case Convert.ToString(dr("sex"))
                        Case "M"
                            intM_2 += 1
                        Case "F"
                            intF_2 += 1
                    End Select

                    '預算別統計
                    Select Case Convert.ToString(dr("budgetid"))
                        Case "01"
                            intBudget01_2 += 1
                        Case "02"
                            intBudget02_2 += 1
                        Case "03"
                            intBudget03_2 += 1
                        Case "98"
                            intBudget98_2 += 1
                        Case "99"
                            intBudget99_2 += 1
                        Case Else
                            intBudgetOdr_2 += 1
                    End Select

                    '身分別統計
                    Select Case Convert.ToString(dr("midentityid"))
                        Case "01"
                            intIdentity01_2 += 1
                        Case "02"
                            intIdentity02_2 += 1
                        Case "03"
                            intIdentity03_2 += 1
                        Case "04"
                            intIdentity04_2 += 1
                        Case "05"
                            intIdentity05_2 += 1
                        Case "06"
                            intIdentity06_2 += 1
                        Case "07"
                            intIdentity07_2 += 1
                        Case "08"
                            intIdentity08_2 += 1
                        Case "09"
                            intIdentity09_2 += 1
                        Case "10"
                            intIdentity10_2 += 1
                        Case "11"
                            intIdentity11_2 += 1
                        Case "12"
                            intIdentity12_2 += 1
                        Case "13"
                            intIdentity13_2 += 1
                        Case "14"
                            intIdentity14_2 += 1
                        Case "15"
                            intIdentity15_2 += 1
                        Case "16"
                            intIdentity16_2 += 1
                        Case "17"
                            intIdentity17_2 += 1
                        Case "18"
                            intIdentity18_2 += 1
                        Case "19"
                            intIdentity19_2 += 1
                        Case "20"
                            intIdentity20_2 += 1
                        Case "21"
                            intIdentity21_2 += 1
                        Case "22"
                            intIdentity22_2 += 1
                        Case "23"
                            intIdentity23_2 += 1
                        Case "24"
                            intIdentity24_2 += 1
                        Case "25"
                            intIdentity25_2 += 1
                        Case "26"
                            intIdentity26_2 += 1
                        Case "27"
                            intIdentity27_2 += 1
                        Case "28"
                            intIdentity28_2 += 1
                        Case "29"
                            intIdentity29_2 += 1
                        Case "30"
                            intIdentity30_2 += 1
                        Case "31"
                            intIdentity31_2 += 1
                        Case "32"
                            intIdentity32_2 += 1
                        Case Else
                            intIdentityOdr_2 += 1
                    End Select

                Case Else '開訓
                    '男女統計
                    Select Case Convert.ToString(dr("sex"))
                        Case "M"
                            intM_1 += 1
                        Case "F"
                            intF_1 += 1
                    End Select

                    '預算別統計
                    Select Case Convert.ToString(dr("budgetid"))
                        Case "01"
                            intBudget01_1 += 1
                        Case "02"
                            intBudget02_1 += 1
                        Case "03"
                            intBudget03_1 += 1
                        Case "98"
                            intBudget98_1 += 1
                        Case "99"
                            intBudget99_1 += 1
                        Case Else
                            intBudgetOdr_1 += 1
                    End Select

                    '身分別統計
                    Select Case Convert.ToString(dr("midentityid"))
                        Case "01"
                            intIdentity01_1 += 1
                        Case "02"
                            intIdentity02_1 += 1
                        Case "03"
                            intIdentity03_1 += 1
                        Case "04"
                            intIdentity04_1 += 1
                        Case "05"
                            intIdentity05_1 += 1
                        Case "06"
                            intIdentity06_1 += 1
                        Case "07"
                            intIdentity07_1 += 1
                        Case "08"
                            intIdentity08_1 += 1
                        Case "09"
                            intIdentity09_1 += 1
                        Case "10"
                            intIdentity10_1 += 1
                        Case "11"
                            intIdentity11_1 += 1
                        Case "12"
                            intIdentity12_1 += 1
                        Case "13"
                            intIdentity13_1 += 1
                        Case "14"
                            intIdentity14_1 += 1
                        Case "15"
                            intIdentity15_1 += 1
                        Case "16"
                            intIdentity16_1 += 1
                        Case "17"
                            intIdentity17_1 += 1
                        Case "18"
                            intIdentity18_1 += 1
                        Case "19"
                            intIdentity19_1 += 1
                        Case "20"
                            intIdentity20_1 += 1
                        Case "21"
                            intIdentity21_1 += 1
                        Case "22"
                            intIdentity22_1 += 1
                        Case "23"
                            intIdentity23_1 += 1
                        Case "24"
                            intIdentity24_1 += 1
                        Case "25"
                            intIdentity25_1 += 1
                        Case "26"
                            intIdentity26_1 += 1
                        Case "27"
                            intIdentity27_1 += 1
                        Case "28"
                            intIdentity28_1 += 1
                        Case "29"
                            intIdentity29_1 += 1
                        Case "30"
                            intIdentity30_1 += 1
                        Case "31"
                            intIdentity31_1 += 1
                        Case "32"
                            intIdentity32_1 += 1
                        Case Else
                            intIdentityOdr_1 += 1
                    End Select
            End Select

            '代入統計值
            '開訓統計資料
            labM_1.Text = intM_1.ToString
            labF_1.Text = intF_1.ToString
            labTotal_1.Text = (intM_1 + intF_1).ToString

            labBudget01_1.Text = intBudget01_1.ToString
            labBudget02_1.Text = intBudget02_1.ToString
            labBudget03_1.Text = intBudget03_1.ToString
            labBudget98_1.Text = intBudget98_1.ToString
            labBudget99_1.Text = intBudget99_1.ToString
            labBudgetOdr_1.Text = intBudgetOdr_1.ToString
            labBudgetTotal_1.Text = (intBudget01_1 + intBudget02_1 + intBudget03_1 + intBudget98_1 + intBudget99_1 + intBudgetOdr_1).ToString

            labId01_1.Text = intIdentity01_1.ToString
            labId02_1.Text = intIdentity02_1.ToString
            labId03_1.Text = intIdentity03_1.ToString
            labId04_1.Text = intIdentity04_1.ToString
            labId05_1.Text = intIdentity05_1.ToString
            labId06_1.Text = intIdentity06_1.ToString
            labId07_1.Text = intIdentity07_1.ToString
            labId08_1.Text = intIdentity08_1.ToString
            labId09_1.Text = intIdentity09_1.ToString
            labId10_1.Text = intIdentity10_1.ToString
            labId11_1.Text = intIdentity11_1.ToString
            labId12_1.Text = intIdentity12_1.ToString
            labId13_1.Text = intIdentity13_1.ToString
            labId14_1.Text = intIdentity14_1.ToString
            labId15_1.Text = intIdentity15_1.ToString
            labId16_1.Text = intIdentity16_1.ToString
            labId17_1.Text = intIdentity17_1.ToString
            labId18_1.Text = intIdentity18_1.ToString
            labId19_1.Text = intIdentity19_1.ToString
            labId20_1.Text = intIdentity20_1.ToString
            labId21_1.Text = intIdentity21_1.ToString
            labId22_1.Text = intIdentity22_1.ToString
            labId23_1.Text = intIdentity23_1.ToString
            labId24_1.Text = intIdentity24_1.ToString
            labId25_1.Text = intIdentity25_1.ToString
            labId26_1.Text = intIdentity26_1.ToString
            labId27_1.Text = intIdentity27_1.ToString
            labId28_1.Text = intIdentity28_1.ToString
            labId29_1.Text = intIdentity29_1.ToString
            labId30_1.Text = intIdentity30_1.ToString
            labId31_1.Text = intIdentity31_1.ToString
            labId32_1.Text = intIdentity32_1.ToString
            labIdOdr_1.Text = intIdentityOdr_1.ToString


            '結訓統計資料
            labM_2.Text = intM_2.ToString
            labF_2.Text = intF_2.ToString
            labTotal_2.Text = (intM_2 + intF_2).ToString

            labBudget01_2.Text = intBudget01_2.ToString
            labBudget02_2.Text = intBudget02_2.ToString
            labBudget03_2.Text = intBudget03_2.ToString
            labBudget98_2.Text = intBudget98_2.ToString
            labBudget99_2.Text = intBudget99_2.ToString
            labBudgetOdr_2.Text = intBudgetOdr_2.ToString
            labBudgetTotal_2.Text = (intBudget01_2 + intBudget02_2 + intBudget03_2 + intBudget98_2 + intBudget99_2 + intBudgetOdr_2).ToString

            labId01_2.Text = intIdentity01_2.ToString
            labId02_2.Text = intIdentity02_2.ToString
            labId03_2.Text = intIdentity03_2.ToString
            labId04_2.Text = intIdentity04_2.ToString
            labId05_2.Text = intIdentity05_2.ToString
            labId06_2.Text = intIdentity06_2.ToString
            labId07_2.Text = intIdentity07_2.ToString
            labId08_2.Text = intIdentity08_2.ToString
            labId09_2.Text = intIdentity09_2.ToString
            labId10_2.Text = intIdentity10_2.ToString
            labId11_2.Text = intIdentity11_2.ToString
            labId12_2.Text = intIdentity12_2.ToString
            labId13_2.Text = intIdentity13_2.ToString
            labId14_2.Text = intIdentity14_2.ToString
            labId15_2.Text = intIdentity15_2.ToString
            labId16_2.Text = intIdentity16_2.ToString
            labId17_2.Text = intIdentity17_2.ToString
            labId18_2.Text = intIdentity18_2.ToString
            labId19_2.Text = intIdentity19_2.ToString
            labId20_2.Text = intIdentity20_2.ToString
            labId21_2.Text = intIdentity21_2.ToString
            labId22_2.Text = intIdentity22_2.ToString
            labId23_2.Text = intIdentity23_2.ToString
            labId24_2.Text = intIdentity24_2.ToString
            labId25_2.Text = intIdentity25_2.ToString
            labId26_2.Text = intIdentity26_2.ToString
            labId27_2.Text = intIdentity27_2.ToString
            labId28_2.Text = intIdentity28_2.ToString
            labId29_2.Text = intIdentity29_2.ToString
            labId30_2.Text = intIdentity30_2.ToString
            labId31_2.Text = intIdentity31_2.ToString
            labId32_2.Text = intIdentity32_2.ToString
            labIdOdr_2.Text = intIdentityOdr_2.ToString
        Next

        '美化排版
        '開訓統計資料
        intCnt = 6 - labM_1.Text.Length
        For i As Integer = 1 To intCnt
            labM_1.Text = "&nbsp;&nbsp;" & labM_1.Text
        Next

        intCnt = 6 - labF_1.Text.Length
        For i As Integer = 1 To intCnt
            labF_1.Text = "&nbsp;&nbsp;" & labF_1.Text
        Next

        intCnt = 6 - labTotal_1.Text.Length
        For i As Integer = 1 To intCnt
            labTotal_1.Text = "&nbsp;&nbsp;" & labTotal_1.Text
        Next

        intCnt = 6 - labBudget01_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudget01_1.Text = "&nbsp;&nbsp;" & labBudget01_1.Text
        Next

        intCnt = 6 - labBudget02_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudget02_1.Text = "&nbsp;&nbsp;" & labBudget02_1.Text
        Next

        intCnt = 6 - labBudget03_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudget03_1.Text = "&nbsp;&nbsp;" & labBudget03_1.Text
        Next

        intCnt = 6 - labBudget98_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudget98_1.Text = "&nbsp;&nbsp;" & labBudget98_1.Text
        Next

        intCnt = 6 - labBudget99_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudget99_1.Text = "&nbsp;&nbsp;" & labBudget99_1.Text
        Next

        intCnt = 6 - labBudgetOdr_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudgetOdr_1.Text = "&nbsp;&nbsp;" & labBudgetOdr_1.Text
        Next

        intCnt = 6 - labBudgetTotal_1.Text.Length
        For i As Integer = 1 To intCnt
            labBudgetTotal_1.Text = "&nbsp;&nbsp;" & labBudgetTotal_1.Text
        Next

        '結訓統計資料
        intCnt = 6 - labM_2.Text.Length
        For i As Integer = 1 To intCnt
            labM_2.Text = "&nbsp;&nbsp;" & labM_2.Text
        Next

        intCnt = 6 - labF_2.Text.Length
        For i As Integer = 1 To intCnt
            labF_2.Text = "&nbsp;&nbsp;" & labF_2.Text
        Next

        intCnt = 6 - labTotal_2.Text.Length
        For i As Integer = 1 To intCnt
            labTotal_2.Text = "&nbsp;&nbsp;" & labTotal_2.Text
        Next

        intCnt = 6 - labBudget01_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudget01_2.Text = "&nbsp;&nbsp;" & labBudget01_2.Text
        Next

        intCnt = 6 - labBudget02_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudget02_2.Text = "&nbsp;&nbsp;" & labBudget02_2.Text
        Next

        intCnt = 6 - labBudget03_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudget03_2.Text = "&nbsp;&nbsp;" & labBudget03_2.Text
        Next

        intCnt = 6 - labBudget98_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudget98_2.Text = "&nbsp;&nbsp;" & labBudget98_2.Text
        Next

        intCnt = 6 - labBudget99_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudget99_2.Text = "&nbsp;&nbsp;" & labBudget99_2.Text
        Next

        intCnt = 6 - labBudgetOdr_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudgetOdr_2.Text = "&nbsp;&nbsp;" & labBudgetOdr_2.Text
        Next

        intCnt = 6 - labBudgetTotal_2.Text.Length
        For i As Integer = 1 To intCnt
            labBudgetTotal_2.Text = "&nbsp;&nbsp;" & labBudgetTotal_2.Text
        Next
    End Sub

    '代入查詢條件
    Private Sub getSch()
        Dim strRtn As String = ""

        If Me.ViewState("year") <> "" Then
            labYear_1.Text = Me.ViewState("year")
            labYear_2.Text = Me.ViewState("year")
        Else
            trYear_1.Visible = False
            trYear_2.Visible = False
        End If

        If Me.ViewState("stsdate") <> "" Or Me.ViewState("stedate") <> "" Then
            labSDate_1.Text = If(Me.ViewState("stsdate") <> "", Me.ViewState("stsdate"), "－") & "～" & If(Me.ViewState("stedate") <> "", Me.ViewState("stedate"), "－")
            labSDate_2.Text = If(Me.ViewState("stsdate") <> "", Me.ViewState("stsdate"), "－") & "～" & If(Me.ViewState("stedate") <> "", Me.ViewState("stedate"), "－")
        Else
            trSDate_1.Visible = False
            trSDate_2.Visible = False
        End If

        If Me.ViewState("etsdate") <> "" Or Me.ViewState("etedate") <> "" Then
            labEDate_1.Text = If(Me.ViewState("etsdate") <> "", Me.ViewState("etsdate"), "－") & "～" & If(Me.ViewState("etedate") <> "", Me.ViewState("etedate"), "－")
            labEDate_2.Text = If(Me.ViewState("etsdate") <> "", Me.ViewState("etsdate"), "－") & "～" & If(Me.ViewState("etedate") <> "", Me.ViewState("etedate"), "－")
        Else
            trEDate_1.Visible = False
            trEDate_2.Visible = False
        End If
    End Sub
#End Region

#Region "Function"
    '取得縣市名稱
    Public Shared Function getCityName(ByRef oConn As SqlConnection, ByVal city1 As String) As String
        Dim strRtn As String = ""
        Dim sql As String = ""
        sql = ""
        sql &= " select ctname from id_city where CTID=@CTID"
        sql &= " order by ctid"
        Dim dt As New DataTable
        'Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, oConn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("CTID", SqlDbType.VarChar).Value = city1
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            strRtn = dt.Rows(0)("ctname")
        End If
        Return strRtn
    End Function

    '取得班級學員資料
    Function getData() As DataTable
        Dim tmpDT As New DataTable
        tmpDT.Columns.Add(New DataColumn("sex"))
        tmpDT.Columns.Add(New DataColumn("budgetid"))
        tmpDT.Columns.Add(New DataColumn("midentityid"))
        tmpDT.Columns.Add(New DataColumn("studstatus"))

        '查ocid
        Dim sql As String = ""
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select a.ocid" & vbCrLf
        sql &= " from class_classinfo a " & vbCrLf
        sql &= " join id_plan b on b.planid=a.planid " & vbCrLf
        sql &= " join id_zip c on c.zipcode=a.taddresszip " & vbCrLf
        sql &= " join id_city d on d.ctid=c.ctid " & vbCrLf
        sql &= " where d.ctid='" & Me.ViewState("city") & "'" & vbCrLf
        '排除產投
        sql &= " and b.tplanid NOT IN (" & TIMS.Cst_TPlanID28_2 & ")" & vbCrLf

        If Me.ViewState("year") <> "" Then
            sql &= " and b.years='" & Me.ViewState("year") & "' " & vbCrLf
        End If

        If Me.ViewState("stsdate") <> "" Then
            sql &= " and a.stdate>=convert(datetime, '" & Me.ViewState("stsdate") & "', 111) " & vbCrLf
        End If

        If Me.ViewState("stedate") <> "" Then
            sql &= " and a.stdate<=convert(datetime, '" & Me.ViewState("stedate") & "', 111) " & vbCrLf
        End If

        If Me.ViewState("etsdate") <> "" Then
            sql &= " and a.ftdate>=convert(datetime, '" & Me.ViewState("etsdate") & "', 111) " & vbCrLf
        End If

        If Me.ViewState("etedate") <> "" Then
            sql &= " and a.ftdate<=convert(datetime, '" & Me.ViewState("etedate") & "', 111) " & vbCrLf
        End If
        sql &= " )" & vbCrLf
        sql &= " select ss.sex,cs.sid,cs.budgetid,cs.midentityid,cs.studstatus " & vbCrLf
        sql &= " from class_studentsofclass cs" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid=cs.sid" & vbCrLf
        sql &= " where cs.OCID in ( SELECT OCID FROM WC1)" & vbCrLf
        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        '查學員資料
        For Each dr As DataRow In dt.Rows
            Dim tmpDR As DataRow = tmpDT.NewRow
            tmpDT.Rows.Add(tmpDR)
            tmpDR("budgetid") = dr("budgetid")
            tmpDR("midentityid") = dr("midentityid")
            tmpDR("studstatus") = dr("studstatus")
            tmpDR("sex") = dr("sex") '判斷學員性別
        Next

        Return tmpDT
    End Function
#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Me.ViewState("year") = TIMS.ClearSQM(Request("year"))
            Me.ViewState("stsdate") = TIMS.ClearSQM(Request("stsdate"))
            Me.ViewState("stedate") = TIMS.ClearSQM(Request("stedate"))
            Me.ViewState("etsdate") = TIMS.ClearSQM(Request("etsdate"))
            Me.ViewState("etedate") = TIMS.ClearSQM(Request("etedate"))
            Me.ViewState("city") = TIMS.ClearSQM(Request("city"))

            crtTable()
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub bt_excel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_excel.Click
        Dim fileName As String = "職業訓練各類身分別人數統計.xls"

        If Request.Browser.Browser = "IE" Then
            fileName = Server.UrlPathEncode(fileName)
        End If

        Dim strContentDisposition As String = [String].Format("{0}; filename=""{1}""", "attachment", fileName)

        Response.AddHeader("Content-Disposition", strContentDisposition)
        Response.ContentType = "application/vnd.ms-excel"

        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)

        div1.RenderControl(htw)
        Common.RespWrite(Me, sw.ToString().Replace("<div>", "").Replace("</div>", ""))
        Response.End()
    End Sub
End Class
