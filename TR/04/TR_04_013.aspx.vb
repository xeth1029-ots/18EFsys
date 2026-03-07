Partial Class TR_04_013
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (眻諉婓 AuthBasePage ?燴, 祥蚚?脤 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '1.就業率 = (就業人數含提前就業人數) / (結訓人數含提前就業人數)
            Dim sMsg1 As String = ""
            sMsg1 &= " 就業人數=勞保勾稽就業人數+人工判定就業人數" & vbCrLf
            'sMsg1 &= " 2.就業率1=(就業人數+提前就業人數-公法救助人數)/(結訓人數+提前就業人數-不就業人數-公法救助人數)" & vbCrLf
            'sMsg1 &= " 3.就業率2=(就業人數+提前就業人數-公法救助人數)/(結訓人數+提前就業人數-在職者-公法救助人數)" & vbCrLf
            labMsg1.Text = sMsg1

            Create1()
        End If
    End Sub

    Sub Create1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.CTID" & vbCrLf
        sql &= " ,a.CTName" & vbCrLf
        sql &= " ,dbo.NVL(b.FinCount,0) FinCount" & vbCrLf
        sql &= " ,dbo.NVL(b.InJobCount,0) InJobCount" & vbCrLf
        sql &= " ,ROUND(dbo.NVL(c.TotalCost,0),2) TotalCost" & vbCrLf
        sql &= " FROM ID_City a" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "  SELECT a.CTID" & vbCrLf
        sql &= "  ,Count(1) FinCount" & vbCrLf
        sql &= "  ,Count(case when a.JobState='1' then 1 end) InJobCount" & vbCrLf
        sql &= "  FROM Stud_GetJobStateByCity a" & vbCrLf
        sql &= "  JOIN CLASS_CLASSINFO cc on cc.OCID=a.OCID " & vbCrLf
        sql &= "  JOIN Plan_PlanInfo pp ON pp.PlanID=cc.PlanID and pp.ComIDNO=cc.ComIDNO and pp.SeqNo=cc.SeqNo" & vbCrLf
        sql &= "  JOIN view_ZipName iz ON cc.TaddressZip=iz.Zipcode" & vbCrLf
        sql &= "  join ID_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        sql &= "  and cc.IsSuccess='Y'" & vbCrLf
        sql &= "  and cc.NotOpen='N'" & vbCrLf
        sql &= "  and ip.TPlanID IN (" & TIMS.Cst_TPlanID_PreUseLimited17c & ")" & vbCrLf
        sql &= "  and a.CPoint=@CPoint" & vbCrLf
        sql &= "  Group By a.CTID" & vbCrLf
        sql &= " ) b ON b.CTID=a.CTID" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "  SELECT iz.CTID" & vbCrLf
        sql &= "  ,SUM(dbo.NVL(d3.TotalCost,0)+(dbo.NVL(d3.AdmCost,0) * dbo.NVL(pp.AdmPercent,0) /100)) TotalCost" & vbCrLf
        sql &= "  FROM Class_ClassInfo cc" & vbCrLf
        sql &= "  JOIN Plan_PlanInfo pp ON pp.PlanID=cc.PlanID and pp.ComIDNO=cc.ComIDNO and pp.SeqNo=cc.SeqNo" & vbCrLf
        sql &= "  JOIN view_ZipName iz ON cc.TaddressZip=iz.Zipcode" & vbCrLf
        sql &= "  join ID_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= "  LEFT JOIN (" & vbCrLf
        sql &= "   SELECT PlanID,ComIDNO,SeqNo" & vbCrLf
        sql &= "   ,SUM(dbo.NVL(OPrice,1)*dbo.NVL(Itemage,1)*dbo.NVL(ItemCost,1)) TotalCost" & vbCrLf
        sql &= "   ,SUM(case when AdmFlag='Y' then dbo.NVL(OPrice,1)*dbo.NVL(Itemage,1)*dbo.NVL(ItemCost,1) end) AdmCost" & vbCrLf
        sql &= "   FROM Plan_CostItem" & vbCrLf
        sql &= "   Group By PlanID,ComIDNO,SeqNo) d3 ON d3.PlanID=cc.PlanID and d3.ComIDNO=cc.ComIDNO and d3.SeqNo=cc.SeqNo" & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        sql &= "  and cc.IsSuccess='Y'" & vbCrLf
        sql &= "  and cc.NotOpen='N'" & vbCrLf
        sql &= "  and ip.TPlanID IN (" & TIMS.Cst_TPlanID_PreUseLimited17c & ")" & vbCrLf
        sql &= "  GROUP BY iz.CTID" & vbCrLf
        sql &= " ) c ON c.CTID=a.CTID" & vbCrLf
        sql &= " Order By a.CTID" & vbCrLf

        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("CPoint", SqlDbType.VarChar).Value = Convert.ToString(CPoint.SelectedIndex)
            dt.Load(.ExecuteReader())
        End With
        'Call CloseDbConn(conn)
        'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))

        'Dim sql As String = ""
        'Dim dt As DataTable = Nothing
        'sql = "SELECT a.CTName,dbo.NVL(b.FinCount,0) FinCount,dbo.NVL(c.InJobCount,0) InJobCount,dbo.NVL(TotalCost,0) TotalCost " & vbCrLf
        'sql += "FROM ID_City a " & vbCrLf
        'sql += "LEFT JOIN (SELECT CTID,Count(CTID) FinCount FROM Stud_GetJobStateByCity WHERE CPoint='" & CPoint.SelectedIndex & "' Group By CTID) b ON b.CTID=a.CTID " & vbCrLf
        'sql += "LEFT JOIN (SELECT CTID,Count(CTID) InJobCount FROM Stud_GetJobStateByCity WHERE JobState='1' and CPoint='" & CPoint.SelectedIndex & "' Group By CTID) c ON c.CTID=a.CTID " & vbCrLf
        'sql += "LEFT JOIN ( " & vbCrLf
        'sql += "SELECT d5.CTID,SUM(dbo.NVL(d3.TotalCost,0)+(dbo.NVL(d4.AdmCost,0) * dbo.NVL(d2.AdmPercent,0) /100)) TotalCost FROM " & vbCrLf
        'sql += "(SELECT * FROM Class_ClassInfo WHERE IsSuccess='Y' and NotOpen='N' and OCID IN (SELECT OCID FROM Stud_GetJobStateByCity WHERE CPoint='" & CPoint.SelectedIndex & "')) d1 " & vbCrLf
        'sql += "JOIN Plan_PlanInfo d2 ON d1.PlanID=d2.PlanID and d1.ComIDNO=d2.ComIDNO and d1.SeqNo=d2.SeqNo " & vbCrLf
        'sql += "LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,SUM(dbo.NVL(OPrice,1)*dbo.NVL(Itemage,1)*dbo.NVL(ItemCost,1)) TotalCost FROM Plan_CostItem Group By PlanID,ComIDNO,SeqNo) d3 ON d3.PlanID=d2.PlanID and d3.ComIDNO=d2.ComIDNO and d3.SeqNo=d2.SeqNo " & vbCrLf
        'sql += "LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,SUM(dbo.NVL(OPrice,1)*dbo.NVL(Itemage,1)*dbo.NVL(ItemCost,1)) AdmCost FROM Plan_CostItem WHERE AdmFlag='Y' Group By PlanID,ComIDNO,SeqNo) d4 ON d4.PlanID=d2.PlanID and d4.ComIDNO=d2.ComIDNO and d4.SeqNo=d2.SeqNo " & vbCrLf
        'sql += "JOIN view_ZipName d5 ON d1.TaddressZip=d5.Zipcode Group by d5.CTID" & vbCrLf
        'sql += ") d ON d.CTID=a.CTID " & vbCrLf
        'sql += "Order By a.CTID" & vbCrLf
        'dt = DbAccess.GetDataTable(sql, objconn)

        With DataGrid1
            .DataSource = dt
            .DataBind()
        End With
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "TR_TD3"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If Not e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "TR_TD4"

                If drv("FinCount") = 0 Then
                    e.Item.Cells(3).Text = "0%"
                Else
                    e.Item.Cells(3).Text = Math.Round(drv("InJobCount") / drv("FinCount") * 100, 2) & "%"
                End If

            Case ListItemType.Footer
                For Each item As DataGridItem In DataGrid1.Items
                    e.Item.Cells(1).Text += Int(item.Cells(1).Text)
                    e.Item.Cells(2).Text += Int(item.Cells(2).Text)
                    e.Item.Cells(4).Text += CDbl(item.Cells(4).Text)
                Next
                e.Item.Cells(4).Text = Format(CDbl(e.Item.Cells(4).Text), "#,##0.00")
                If Int(e.Item.Cells(2).Text) = 0 Then
                    e.Item.Cells(3).Text = "0%"
                Else
                    e.Item.Cells(3).Text = Math.Round(e.Item.Cells(1).Text / e.Item.Cells(2).Text * 100, 2) & "%"
                End If
        End Select
    End Sub

    Private Sub CPoint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CPoint.SelectedIndexChanged
        Call Create1()
    End Sub
End Class
