Partial Class CP_04_002_01
    Inherits AuthBasePage

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

        If Not Page.IsPostBack Then
        End If

    End Sub

    Sub create1()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        Dim PlanID As String = Request("PlanID")
        Dim ComIDNO As String = Request("ComIDNO")
        Dim SeqNO As String = Request("SeqNO")
        '訓練年度
        Me.train_year.Text = Request("Year")
        '訓練計畫
        Me.train_plan.Text = Request("PlanName")
        '訓練職類
        Me.train_name.Text = Request("TrainName")

        PlanID = TIMS.ClearSQM(PlanID)
        ComIDNO = TIMS.ClearSQM(ComIDNO)
        SeqNO = TIMS.ClearSQM(SeqNO)
        train_year.Text = TIMS.ClearSQM(train_year.Text)
        train_plan.Text = TIMS.ClearSQM(train_plan.Text)
        train_name.Text = TIMS.ClearSQM(train_name.Text)
        If PlanID = "" Then Exit Sub
        If ComIDNO = "" Then Exit Sub
        If SeqNO = "" Then Exit Sub

        Dim sqlstr As String = ""
        sqlstr = " SELECT a.*,b.Name "
        sqlstr += " from Plan_PlanInfo a "
        sqlstr += " JOIN Key_Degree b ON b.DegreeID=a.CapDegree "
        sqlstr += " where a.PlanID=" & PlanID
        sqlstr += " and a.ComIDNO='" & ComIDNO & "'"
        sqlstr += " and a.SeqNO=" & SeqNO
        Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
        If dr Is Nothing Then Exit Sub

        '表格全不允許修改
        Me.PlanCause.ReadOnly = True
        Me.PurScience.ReadOnly = True
        Me.PurTech.ReadOnly = True
        Me.PurMoral.ReadOnly = True

        If Not dr Is Nothing Then
            '目標
            Me.PlanCause.Text = dr("PlanCause").ToString
            Me.PurScience.Text = dr("PurScience").ToString
            Me.PurTech.Text = dr("PurTech").ToString
            Me.PurMoral.Text = dr("PurMoral").ToString

            '受訓資格
            Me.CapDegree.Text = dr("Name").ToString
            Me.CapAge1.Text = dr("CapAge1").ToString
            Me.CapAge2.Text = dr("CapAge2").ToString

            If dr("CapSex").ToString = "0" Then
                Me.CapSex.Text = "不分"
            ElseIf dr("CapSex").ToString = "M" Then
                Me.CapSex.Text = "男"
            ElseIf dr("CapSex").ToString = "F" Then
                Me.CapSex.Text = "女"
            End If

            'Me.CapMilitary.Text = "不分"
            If dr("CapMilitary").ToString = "00" Then
                Me.CapMilitary.Text = "不分"
            ElseIf dr("CapMilitary").ToString = "04" Then
                Me.CapMilitary.Text = "在役"
            ElseIf dr("CapMilitary").ToString = "0103" Then
                Me.CapMilitary.Text = "役畢(含免役)"
            ElseIf dr("CapMilitary").ToString = "02" Then
                Me.CapMilitary.Text = "未役"
            End If

            Me.CapOther1.Text = dr("CapOther1").ToString
            Me.CapOther2.Text = dr("CapOther2").ToString
            Me.CapOther3.Text = dr("CapOther3").ToString

            '訓練方式
            Me.TMScience.Text = dr("TMScience").ToString
            Me.TMTech.Text = dr("TMTech").ToString

            ''課程編配
            '一般學科
            If dr("GenSciHours").ToString = "" Then
                Me.GenSciHours.Text = "0"
            Else
                Me.GenSciHours.Text = dr("GenSciHours").ToString
            End If

            '專業學科
            If dr("ProSciHours").ToString = "" Then
                Me.ProSciHours.Text = "0"
            Else
                Me.ProSciHours.Text = dr("ProSciHours").ToString
            End If

            '學科
            If dr("GenSciHours").ToString <> "" Or dr("ProSciHours").ToString <> "" Then
                Dim a As Int32 = dr("GenSciHours") + dr("ProSciHours")
                Me.SciHours.Text = a.ToString
            End If

            '術科
            If dr("ProTechHours").ToString = "" Then
                Me.ProTechHours.Text = "0"
            Else
                Me.ProTechHours.Text = dr("ProTechHours").ToString
            End If

            '其他時數
            If dr("OtherHours").ToString = "" Then
                Me.OtherHours.Text = "0"
            Else
                Me.OtherHours.Text = dr("OtherHours").ToString
            End If

            '總計
            If dr("TotalHours").ToString = "" Then
                Me.TotalHours.Text = "0"
            Else
                Me.TotalHours.Text = dr("TotalHours").ToString
            End If

            '班別資料
            Me.ClassName.Text = dr("ClassName").ToString
            Me.TNum.Text = dr("TNum").ToString
            Me.THours.Text = dr("THours").ToString
            Me.STDate.Text = FormatDateTime(dr("STDate"), 2)
            Me.FDDate.Text = FormatDateTime(dr("FDDate"), 2)
            Me.CyclType.Text = dr("CyclType").ToString

            '訓練費用
            Me.ViewState("AdmPercent") = dr("AdmPercent")
            Me.Panel1.Visible = False
            Me.Panel2.Visible = False
            Me.Panel3.Visible = False
            Me.Panel4.Visible = False
            AdmFee()

            '經費來源
            If dr("DefGovCost").ToString = "" Then
                Me.MainCost.Text = "0"
            Else
                Me.MainCost.Text = dr("DefGovCost").ToString
            End If

            If dr("DefUnitCost").ToString = "" Then
                Me.CenterCost.Text = "0"
            Else
                Me.CenterCost.Text = dr("DefUnitCost").ToString
            End If

            If dr("DefStdCost").ToString = "" Then
                Me.UnitCost.Text = "0"
            Else
                Me.UnitCost.Text = dr("DefStdCost").ToString
            End If

            '備註
            Me.Note.Value = dr("Note").ToString

        End If
    End Sub

    Sub AdmFee()
        Dim dt As DataTable
        Dim dr As DataRow

        Dim PlanID As Int32 = Request("PlanID")
        Dim ComIDNO As String = Request("ComIDNO")
        Dim SeqNO As Int32 = Request("SeqNO")
        Dim sum, AllSmallSum As Double
        Dim ACsum, totalsum, totalus As Integer

        Dim s As String = ""
        Dim sqlstr As String = ""
        sqlstr += " select c.*" & vbCrLf
        sqlstr += " ,d.CostName" & vbCrLf
        sqlstr += " ,a.AdmPercent" & vbCrLf
        sqlstr += " ,a.TNum" & vbCrLf
        sqlstr += " ,(c.OPrice*c.Itemage*c.ItemCost) AllSmallSum" & vbCrLf
        sqlstr += " from Plan_PlanInfo a "
        sqlstr += " JOIN Plan_CostItem c ON c.PlanID=a.PlanID and c.ComIDNO=a.ComIDNO and c.SeqNO=a.SeqNO "
        sqlstr += " JOIN Key_CostItem d ON d.CostID=c.CostID "
        sqlstr += " where a.PlanID=" & PlanID
        sqlstr += " and a.ComIDNO='" & ComIDNO & "' "
        sqlstr += " and a.SeqNO=" & SeqNO
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        For Each dr In dt.Rows
            If Not dr("AllSmallSum").ToString = "" Then
                AllSmallSum += dr("AllSmallSum")
            End If
        Next

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Select Case dr("CostMode")
                Case 1
                    Me.Panel1.Visible = True
                    DataGrid1.DataSource = dt
                    DataGrid1.DataBind()

                    '行政管理費
                    For Each dr In dt.Select("AdmFlag='Y'")
                        sum += dr("AllSmallSum")
                        If dr("CostID").ToString = "99" Then
                            s += dr("CostName").ToString & "-" & dr("ItemOther").ToString & "+"
                        Else
                            s += dr("CostName").ToString & "+"
                        End If
                    Next

                    If Me.ViewState("AdmPercent").ToString = "" Then
                        Me.ViewState("AdmPercent") = 0
                    End If

                    If Not s = "" Then
                        ACsum = Math.Round(sum * CDbl(Me.ViewState("AdmPercent").ToString) / 100)
                        AdmCost.Text = "(" & s.Substring(0, s.Length - 1) & ")*" & Me.ViewState("AdmPercent").ToString & "% = " & ACsum.ToString
                        totalsum = Math.Round(ACsum + AllSmallSum)
                        Me.TotalCost1.Text = totalsum.ToString
                    Else
                        AdmCost.Text = "0"
                        Me.TotalCost1.Text = AllSmallSum.ToString
                    End If
                Case 2
                    '每人每時單價計價
                    Me.Panel2.Visible = True
                    DataGrid2.DataSource = dt
                    DataGrid2.DataBind()

                    '加入行政管理費 bynick
                    '行政管理費
                    For Each dr In dt.Select("AdmFlag='Y'")
                        sum += dr("AllSmallSum")
                        If dr("CostID").ToString = "99" Then
                            s += dr("CostName").ToString & "-" & dr("ItemOther").ToString & "+"
                        Else
                            s += dr("CostName").ToString & "+"
                        End If
                    Next
                    For Each dr In dt.Rows
                        totalus += dr("AllSmallSum")
                    Next
                    If Me.ViewState("AdmPercent").ToString = "" Then
                        Me.ViewState("AdmPercent") = 0
                    End If

                    If Not s = "" Then
                        ACsum = Math.Round(sum * CDbl(Me.ViewState("AdmPercent").ToString) / 100)
                        Me.AdmCost2.Text = "(" & s.Substring(0, s.Length - 1) & ")*" & Me.ViewState("AdmPercent").ToString & "% = " & ACsum.ToString
                        totalsum = Math.Round(ACsum + AllSmallSum)
                        Me.TotalCost2.Text = totalus + ACsum
                    Else
                        Me.AdmCost2.Text = "0"
                        Me.TotalCost2.Text = totalus
                    End If
                    '------end----

                    ' Me.TotalCost2.Text = AllSmallSum.ToString
                Case 3
                    '每人輔助單價計費
                    Me.Panel3.Visible = True
                    AdmFee3()
                Case 4
                    Me.Panel4.Visible = True
                    AdmFee4()
            End Select
        Else
            Me.TrainCostStatus.Text = "查無資料"
        End If

    End Sub

    Sub AdmFee3()
        Dim dt As DataTable
        Dim dr As DataRow

        Dim PlanID As Int32 = Request("PlanID")
        Dim ComIDNO As String = Request("ComIDNO")
        Dim SeqNO As Int32 = Request("SeqNO")
        Dim sum, AllSmallSum As Double
        Dim ACsum, totalsum, totalus As Integer

        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr += " select c.*" & vbCrLf
        sqlstr += " ,d.CostName" & vbCrLf
        sqlstr += " ,a.AdmPercent" & vbCrLf
        sqlstr += " ,a.TNum" & vbCrLf
        sqlstr += " ,(c.OPrice*c.Itemage) AllSmallSum" & vbCrLf
        sqlstr += " from Plan_PlanInfo a "
        sqlstr += " JOIN Plan_CostItem c ON c.PlanID=a.PlanID and c.ComIDNO=a.ComIDNO and c.SeqNO=a.SeqNO "
        sqlstr += " JOIN Key_CostItem d ON d.CostID=c.CostID "
        sqlstr += " where a.PlanID=" & PlanID
        sqlstr += " and a.ComIDNO='" & ComIDNO & "' "
        sqlstr += " and a.SeqNO=" & SeqNO
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        DataGrid3.DataSource = dt
        DataGrid3.DataBind()

        '加入行政管理費 by nick 00519
        '行政管理費
        Dim s As String = ""
        For Each dr In dt.Select("AdmFlag='Y'")
            sum += dr("AllSmallSum")
            If dr("CostID").ToString = "99" Then
                s += dr("CostName").ToString & "-" & dr("ItemOther").ToString & "+"
            Else
                s += dr("CostName").ToString & "+"
            End If
        Next
        For Each dr In dt.Rows
            totalus += dr("AllSmallSum")
        Next
        If Me.ViewState("AdmPercent").ToString = "" Then
            Me.ViewState("AdmPercent") = 0
        End If

        If Not s = "" Then
            ACsum = Math.Round(sum * CDbl(Me.ViewState("AdmPercent").ToString) / 100)
            Me.AdmCost3.Text = "(" & s.Substring(0, s.Length - 1) & ")*" & Me.ViewState("AdmPercent").ToString & "% = " & ACsum.ToString
            totalsum = Math.Round(ACsum + AllSmallSum)
            Me.TotalCost3.Text = totalus + ACsum
        Else
            Me.AdmCost3.Text = "0"
            Me.TotalCost3.Text = totalus
        End If
        '--------end---------
    End Sub

    Sub AdmFee4()
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim PlanID As Int32 = Request("PlanID")
        Dim ComIDNO As String = Request("ComIDNO")
        Dim SeqNO As Int32 = Request("SeqNO")
        Dim sum As Double = 0
        Dim AllSmallSum As Double = 0
        Dim PeopleNum As Integer = 0
        Dim ACsum As Integer = 0
        Dim totalsum As Integer = 0
        Dim totalus As Integer = 0

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " select c.*" & vbCrLf
        sqlstr += " ,d.CostName" & vbCrLf
        sqlstr += " ,a.AdmPercent" & vbCrLf
        sqlstr += " ,a.TNum" & vbCrLf
        sqlstr += " ,(c.OPrice*c.Itemage) AllSmallSum" & vbCrLf
        sqlstr += " from Plan_PlanInfo a "
        sqlstr += " JOIN Plan_CostItem c ON c.PlanID=a.PlanID and c.ComIDNO=a.ComIDNO and c.SeqNO=a.SeqNO "
        sqlstr += " JOIN Key_CostItem  d ON d.CostID=c.CostID "
        sqlstr += " where a.PlanID=" & PlanID
        sqlstr += " and a.ComIDNO='" & ComIDNO & "' "
        sqlstr += " and a.SeqNO=" & SeqNO
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        DataGrid4.DataSource = dt
        DataGrid4.DataBind()

        '加入行政管理費 by nick 00519
        '行政管理費
        Dim s As String = ""
        For Each dr In dt.Select("AdmFlag='Y'")
            sum += dr("AllSmallSum")
            If dr("CostID").ToString = "99" Then
                s += dr("CostName").ToString & "-" & dr("ItemOther").ToString & "+"
            Else
                s += dr("CostName").ToString & "+"
            End If
        Next

        For Each dr In dt.Rows
            totalus += dr("AllSmallSum")
        Next
        If Me.ViewState("AdmPercent").ToString = "" Then
            Me.ViewState("AdmPercent") = 0
        End If

        If Not s = "" Then
            ACsum = Math.Round(sum * CDbl(Me.ViewState("AdmPercent").ToString) / 100)
            Me.AdmCost4.Text = "(" & s.Substring(0, s.Length - 1) & ")*" & Me.ViewState("AdmPercent").ToString & "% = " & ACsum.ToString
            totalsum = Math.Round(ACsum + AllSmallSum)
            Me.TotalCost4.Text = totalus + ACsum
            Me.PerCost.Text = Math.Round((totalus + ACsum) / CInt(dr("TNum")))
        Else
            Me.AdmCost4.Text = "0"
            Me.TotalCost4.Text = totalus
            Me.PerCost.Text = Math.Round(totalus / CInt(dr("TNum")))
        End If
        '--------end---------

    End Sub

End Class
