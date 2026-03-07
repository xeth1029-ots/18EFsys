Partial Class TR_04_010_R
    Inherits AuthBasePage

    Const Cst_CommandTimeout As Integer = 30 '1000
    Dim TitleTable As DataTable
    Dim SearchStr As String
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

        Dim MyCell As TableCell = Nothing
        Dim MyRow As TableRow = Nothing

        Dim dr1 As DataRow = Nothing
        Dim sql As String = ""

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If Session("TitleTable") Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If Session("SearchStr") Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        TitleTable = Session("TitleTable")
        SearchStr = Session("SearchStr")
        Dim dr As DataRow = Nothing
        dr = TitleTable.Rows(0)
        STDate.Text = dr("STDate1") & "~" & dr("STDate2")
        FTDate.Text = dr("FTDate1") & "~" & dr("FTDate2")

        If dr("DistID").ToString = "" Then
            DistID.Text = "無選擇"
            TD2_1.Visible = False
            TD2_2.Visible = False
        Else
            sql = "SELECT Name FROM ID_District WHERE DistID='" & dr("DistID").ToString & "'"
            dr1 = DbAccess.GetOneRow(sql, objconn)
            If dr1 Is Nothing Then
                DistID.Text = "無選擇"
                TD2_1.Visible = False
                TD2_2.Visible = False
            Else
                DistID.Text = dr1("Name")
            End If
        End If

        If dr("TPlanID").ToString = "" Then
            TPlanID.Text = "無選擇"
            TD2_3.Visible = False
            TD2_4.Visible = False
        Else
            sql = "SELECT PlanName FROM Key_Plan WHERE TPlanID='" & dr("TPlanID").ToString & "'"
            dr1 = DbAccess.GetOneRow(sql, objconn)
            If dr1 Is Nothing Then
                TPlanID.Text = "無選擇"
                TD2_3.Visible = False
                TD2_4.Visible = False
            Else
                TPlanID.Text = dr1("PlanName")
            End If
        End If

        If TD2_1.Visible = False And TD2_3.Visible = False Then
            TR2.Visible = False
        End If

        If dr("RIDValue").ToString = "" Then
            RIDValue.Text = "無選擇"
            TD3_1.Visible = False
            TD3_2.Visible = False
        Else
            sql = "SELECT b.OrgName  "
            sql += "FROM Auth_Relship a "
            sql += "JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
            sql += "WHERE a.RID='" & dr("RIDValue").ToString & "' "
            dr1 = DbAccess.GetOneRow(sql, objconn)
            If dr1 Is Nothing Then
                RIDValue.Text = "無選擇"
                TD3_1.Visible = False
                TD3_2.Visible = False
            Else
                RIDValue.Text = dr1("OrgName")
            End If
        End If

        Dim ClassCName As String
        If dr("OCIDValue").ToString = "" Then
            ClassCName = "無選擇"
            TD3_3.Visible = False
            TD3_4.Visible = False
        Else
            sql = "SELECT ClassCName,CyclType FROM Class_ClassInfo WHERE OCID='" & dr("OCIDValue").ToString & "'"
            dr1 = DbAccess.GetOneRow(sql, objconn)
            If dr1 Is Nothing Then
                ClassCName = "無選擇"
                TD3_3.Visible = False
                TD3_4.Visible = False
            Else
                ClassCName = dr1("ClassCName").ToString
                If IsNumeric(dr1("CyclType")) Then
                    If Int(dr1("CyclType")) <> 0 Then
                        ClassCName += "第" & Int(dr1("CyclType")) & "期"
                    End If
                End If
            End If
            OCIDValue.Text = ClassCName
        End If

        If TD3_1.Visible = False And TD3_3.Visible = False Then
            TR3.Visible = False
        End If

        Dim sInputText As String = ""
        sInputText = ""
        sInputText &= "&BorderWidth=1"
        sInputText &= "&TPlanID=" & Convert.ToString(dr("TPlanID"))
        sInputText &= "&Range1=" & Convert.ToString(dr("Range1"))
        sInputText &= "&Range2=" & Convert.ToString(dr("Range2"))
        sInputText &= "&Range3=" & Convert.ToString(dr("Range3"))
        sInputText &= "&Range4=" & Convert.ToString(dr("Range4"))

        Select Case Request("Mode")
            Case "1"
                TR_04_010.CreateRow(ShowDataTable, MyRow)

                TR_04_010.CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , "TR_04002_TR", 1)
                'CreateCell(MyRow, MyCell, "訓前已加保者", , , "TR_04002_TR")
                'MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, "無加退保紀錄", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range1").ToString & "週(含)以下", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range2").ToString & "週至" & TitleTable.Rows(0)("Range3").ToString & "週", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range4").ToString & "週(含)以上", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, "合計", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)

                For j As Integer = 1 To 3
                    Call TR_04_010.CreateData(3, j, SearchStr, False, Me, objconn, sInputText)
                Next
            Case "2"
                TR_04_010.CreateRow(ShowDataTable, MyRow)
                TR_04_010.CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , "TR_04002_TR", 1)
                'CreateCell(MyRow, MyCell, "訓前已加保者", , , "TR_04002_TR")
                'MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, "無加退保紀錄", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range1").ToString & "週(含)以下", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range2").ToString & "週至" & TitleTable.Rows(0)("Range3").ToString & "週", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, TitleTable.Rows(0)("Range4").ToString & "週(含)以上", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)
                TR_04_010.CreateCell(MyRow, MyCell, "合計", , , "TR_04002_TR", 1)
                MyCell.Width = Unit.Pixel(90)

                Call TR_04_010.CreateData(1, 1, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(2, 1, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(1, 2, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(2, 2, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(4, 2, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(1, 3, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(2, 3, SearchStr, False, Me, objconn, sInputText)
                Call TR_04_010.CreateData(4, 3, SearchStr, False, Me, objconn, sInputText)
            Case "3"
        End Select
    End Sub


End Class
