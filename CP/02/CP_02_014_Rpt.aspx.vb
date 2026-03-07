'Imports System.Data
'Imports Oracle.DataAccess.Client

Partial Class CP_02_014_Rpt
    Inherits AuthBasePage

    'Dim arrItem(113, 113) As String
    'Dim arrSpanRow(114) As String
    'Dim sourceDT As DataTable = Nothing
    'Dim PName As String = ""
    'Dim plankind As String = ""

    Public Class myrow
        Public name As String
        'Public layer As Integer
        Public group As List(Of myrow)
        Public count As List(Of Integer)
    End Class

    'Dim SYMD As String = "2012/01/01"
    'Dim EYMD As String = "2012/02/01"
    Dim gSYMD As String = "2012/01/01"
    Dim gEYMD As String = "2012/02/01"

    Dim tDt_main As New DataTable

    Dim col As List(Of String) = New List(Of String)
    Dim col_total As List(Of Integer) = New List(Of Integer)
    Dim row As List(Of myrow) = New List(Of myrow)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Dim start_date As String = "2012/01/01"
        Dim end_date As String = "2012/02/01"
        Dim TPlan As String = "37"
        Dim OC_Type As String = "1"
        Dim SYMD As String = "2012/01/01"
        Dim EYMD As String = "2012/02/01"

        start_date = Convert.ToString(Request("start_date")) '"2012/01/01"
        end_date = Convert.ToString(Request("end_date")) '"2012/02/01"
        TPlan = Convert.ToString(Request("TPlan")) '"37"
        OC_Type = Convert.ToString(Request("OC_Type")) '"1"
        SYMD = Convert.ToString(Request("SYMD"))
        EYMD = Convert.ToString(Request("EYMD"))
        gSYMD = Convert.ToString(Request("SYMD"))
        gEYMD = Convert.ToString(Request("EYMD"))

        Dim tDt As New DataTable

        tDt_main = db_main(start_date, end_date, TPlan, OC_Type, SYMD, EYMD)

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim c As Integer
        Dim str As String
        Dim b As Boolean
        Dim parent As List(Of myrow)
        Dim layer As Integer
        Dim r As myrow
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim cv As Integer

        For i = 0 To tDt_main.Rows.Count - 1
            str = tDt_main.Rows(i).Item("Field1").ToString()
            b = False
            For j = 0 To col.Count - 1
                If str = col.Item(j) Then
                    str = tDt_main.Rows(i).Item("sex_desc").ToString()
                    cv = tDt_main.Rows(i).Item("COUNTValue").ToString()
                    If str = "男" Then
                        col_total.Item(j * 2) += cv
                    ElseIf str = "女" Then
                        col_total.Item(j * 2 + 1) += cv
                    End If
                    b = True
                    Exit For
                End If
            Next

            If Not b Then
                col.Add(New String(str))
                c = col_total.Count
                col_total.Add(New Integer())
                col_total.Add(New Integer())

                str = tDt_main.Rows(i).Item("sex_desc").ToString()
                cv = tDt_main.Rows(i).Item("COUNTValue").ToString()
                If str = "男" Then
                    col_total.Item(c) += cv
                ElseIf str = "女" Then
                    col_total.Item(c + 1) += cv
                End If
            End If
        Next


        For i = 0 To tDt_main.Rows.Count - 1
            str = tDt_main.Rows(i).Item("Field1").ToString()
            str2 = tDt_main.Rows(i).Item("Field2").ToString()
            str3 = tDt_main.Rows(i).Item("Unit_desc").ToString()
            str4 = tDt_main.Rows(i).Item("Train1_desc").ToString()
            b = False
            layer = 0
            parent = row
            For k = 0 To col.Count - 1
                If str = col.Item(k) Then
                    Exit For
                End If
            Next
            str = tDt_main.Rows(i).Item("sex_desc").ToString()
            cv = tDt_main.Rows(i).Item("COUNTValue").ToString()
            For j = 0 To parent.Count - 1
                If str2 = parent.Item(j).name Then
                    If str = "男" Then
                        parent.Item(j).count(k * 2) += cv
                    ElseIf str = "女" Then
                        parent.Item(j).count(k * 2 + 1) += cv
                    End If
                    b = True
                    layer = 1
                    parent = parent.Item(j).group
                    Exit For
                End If
            Next

            If b And str3 <> "" Then
                b = False
                For j = 0 To parent.Count - 1
                    If str3 = parent.Item(j).name Then
                        If str = "男" Then
                            parent.Item(j).count(k * 2) += cv
                        ElseIf str = "女" Then
                            parent.Item(j).count(k * 2 + 1) += cv
                        End If
                        b = True
                        layer = 2
                        parent = parent.Item(j).group
                        Exit For
                    End If
                Next
            End If

            If b And str4 <> "" Then
                b = False
                For j = 0 To parent.Count - 1
                    If str4 = parent.Item(j).name Then
                        If str = "男" Then
                            parent.Item(j).count(k * 2) += cv
                        ElseIf str = "女" Then
                            parent.Item(j).count(k * 2 + 1) += cv
                        End If
                        b = True
                        layer = 3
                        parent = parent.Item(j).group
                        Exit For
                    End If
                Next
            End If

            If Not b Then
                If layer = 0 Then
                    r = New myrow
                    r.name = str2
                    r.group = New List(Of myrow)
                    r.count = New List(Of Integer)
                    For j = 0 To col.Count - 1
                        r.count.Add(New Integer())
                        r.count.Add(New Integer())
                    Next
                    parent.Add(r)
                    If str = "男" Then
                        r.count.Item(k * 2) += cv
                    ElseIf str = "女" Then
                        r.count.Item(k * 2 + 1) += cv
                    End If
                    layer = 1
                    parent = r.group
                End If

                If layer = 1 Then
                    If str3 <> "" Then
                        r = New myrow
                        r.name = str3
                        r.group = New List(Of myrow)
                        r.count = New List(Of Integer)
                        For j = 0 To col.Count - 1
                            r.count.Add(New Integer())
                            r.count.Add(New Integer())
                        Next
                        parent.Add(r)
                        If str = "男" Then
                            r.count.Item(k * 2) += cv
                        ElseIf str = "女" Then
                            r.count.Item(k * 2 + 1) += cv
                        End If
                        layer = 2
                        parent = r.group
                    End If
                End If

                If layer = 2 Then
                    If str4 <> "" Then
                        r = New myrow
                        r.name = str4
                        r.group = New List(Of myrow)
                        r.count = New List(Of Integer)
                        For j = 0 To col.Count - 1
                            r.count.Add(New Integer())
                            r.count.Add(New Integer())
                        Next
                        parent.Add(r)
                        If str = "男" Then
                            r.count.Item(k * 2) += cv
                        ElseIf str = "女" Then
                            r.count.Item(k * 2 + 1) += cv
                        End If
                    End If
                End If

            End If
        Next

        PrintDiv(tDt_main, "AAAA", "150", "150", 40, "7", "V")

    End Sub


    Private Function db_main(ByVal start_date As String, ByVal end_date As String, ByVal TPlan As String, ByVal OC_Type As String, ByVal SYMD As String, ByVal EYMD As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " Select Unit_type,Unit_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc  "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc  "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID)='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID)='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc   "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc   "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "  union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID)='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += "   "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.UnitCode=S1.Unit_type"
        sql += "  "
        sql += "  UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "    left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        'sql += "  (Select '001' as Unit_type,'職訓局泰山職訓中心' as Unit_desc  "
        sql += "  (Select '001' as Unit_type,'泰山職業訓練場' as Unit_desc  "
        'sql += "   union Select '002' as Unit_type,'職訓局中區職訓中心' as Unit_desc  "
        sql += "   union Select '002' as Unit_type,'勞動部勞動力發展署中彰投分署' as Unit_desc  "
        'sql += "   union Select '003' as Unit_type,'職訓局北區職訓中心' as Unit_desc  "
        sql += "   union Select '003' as Unit_type,'勞動部勞動力發展署北基宜花金馬分署' as Unit_desc  "
        'sql += "   union Select '004' as Unit_type,'職訓局南區職訓中心' as Unit_desc  "
        sql += "   union Select '004' as Unit_type,'勞動部勞動力發展署高屏澎東分署' as Unit_desc  "
        'sql += "   union Select '005' as Unit_type,'職訓局桃園職訓中心' as Unit_desc  "
        sql += "   union Select '005' as Unit_type,'勞動部勞動力發展署桃竹苗分署' as Unit_desc  "
        'sql += "   union Select '006' as Unit_type,'職訓局台南職訓中心' as Unit_desc    "
        sql += "   union Select '006' as Unit_type,'勞動部勞動力發展署雲嘉南分署' as Unit_desc    "
        sql += "   union Select '007' as Unit_type,'退輔會訓練中心' as Unit_desc  "
        sql += "   union Select '008' as Unit_type,'青輔會青年職訓中心' as Unit_desc  "
        sql += "   union Select '009' as Unit_type,'農委會漁業署遠洋漁業開發中心' as Unit_desc  "
        sql += "   union Select '010' as Unit_type,'台北市職訓中心' as Unit_desc  "
        sql += "   union Select '011' as Unit_type,'高雄市訓練就業中心' as Unit_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Unit_type "
        sql += " ) A_table "
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Trainice_type,Trainice_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )   S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )   S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )   S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )   S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc  "
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc  "
        sql += "   )  T1)  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc  "
        sql += "    union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )   S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " )B_table"
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Q7_type,Q7_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += ")  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3  From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc  "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc  "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc  "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc  "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc  "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " )C_table"
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Q7_type,Q7_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,3 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc  "
        sql += "   union Select '2' as s_type,'女' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " ) D_table"
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Train_type,Train_desc,Train1_type,Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc )  S2"
        sql += "  left join"
        sql += " (Select "
        sql += " Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null) "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'N' as Train_type,'已找到工作' as Train_desc )  Z1"
        sql += "    Union all"
        sql += "    Select Z1.*,Z2.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "     (SELECT 'Y' as Train_type,'未找到工作' as Train_desc )  Z1 cross join"
        sql += "     (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc  "
        sql += "      union SELECT '2' as  Train1_type,'有幫助' as Train1_desc  "
        sql += "      union SELECT '3' as  Train1_type,'普通' as Train1_desc  "
        sql += "      union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc  "
        sql += "      union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc  "
        sql += "     )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += " end as Q11"
        sql += " ,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11=S1.Train_type and (CONVERT(varchar, S3.Q11N)=S1.Train1_type or S3.Q11N is null)"
        sql += " ) E_table"
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Train_type,Train_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += ")  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Train_type,S1.Train_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'訓練後或訓練期間有否找到工作' as Field2,4 as Field3 From "
        sql += "   (SELECT 'Z' as Train_type,'未填答' as Train_desc )  Z1"
        sql += "  )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  )  S2"
        sql += "  left join"
        sql += " (Select B.Q11,B.Q11N,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,C.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.Q11 is null "
        sql += " ) F_table"
        sql += " Where 1=1"
        sql += " "
        sql += " UNION ALL"
        sql += " Select Q12_type,Q12_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,ISNULL(COUNTValue,0) as COUNTValue,IdentityID,Field1,Field2,Field3 From ("
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'02' as IdentityID, '就業保險被保險人失業者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='02'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'03' as IdentityID, '負擔家計婦女' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "  left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='03'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'04' as IdentityID, '中高齡(45歲)' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='04'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += ")  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'05' as IdentityID, '原住民' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='05'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'06' as IdentityID, '身心障礙者' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='06'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'07' as IdentityID, '生活扶助戶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "  left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='07'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += "  "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'09' as IdentityID, '家庭暴力受害人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "  left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='09'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type "
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'10' as IdentityID, '更生保護人' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='10'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'12' as IdentityID, '屆退官兵' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='12'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'13' as IdentityID, '外籍配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='13'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL"
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,ISNULL(S3.COUNTValue,0) as COUNTValue,'14' as IdentityID, '大陸配偶' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,5 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc  "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc  "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc  "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc  "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc  "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc  "
        sql += "   union Select '1' as s_type,'男' as sex_desc  "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left join Stud_ResultIdentData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "   left  join Key_Identity ky on D.IdentityID = ky.IdentityID"
        sql += "  where ISNULL(ky.MergeID,D.IdentityID) ='14'"
        If start_date <> "" Then
            sql += "  And ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  And A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " ) G_table"
        sql += " Where 1=1"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function



    Private Sub PrintDiv(ByVal dt As DataTable, ByVal selRpt As String, ByVal Field1_width As String, ByVal Field2_width As String, ByVal RCount As Integer, ByVal font_size As String, ByVal portrait As String)
        'dt:要顯示的資料,selRpt:,Field1_width:標題題目的寬度,Field2_width:標題題目的寬度,RCount:每頁筆數,font_size:內容字型大小,portrait:直式/橫式

        Dim tmpDT As New DataTable
        'Dim tmpDR As DataRow
        'Dim tmpObj As Object
        Dim sql As String = ""
        Dim PageCount As Int32 = 0  'Pages
        Dim ReportCount As Integer = RCount '每頁筆數
        Dim ColCount As Integer = 0
        Dim intTmp As Integer = 0
        Dim rsCursor As Integer = 0   '報表內容列印的NO
        Dim intPageRecord As Integer = RCount '每頁列印幾筆

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:" + font_size + "pt;font-family:DFKai-SB"
        Dim int_width As Integer
        Dim strWatermarkImg As String
        Dim strWatermarkDiv As String
        Dim intWatermarkTop As Integer

        tmpDT = dt
        ColCount = dt.Columns.Count

        intTmp = tmpDT.Rows.Count
        PageCount = 1
        'If (intTmp Mod ReportCount) = 0 Then
        '    PageCount = (intTmp / ReportCount) - 1
        'Else
        '    PageCount = intTmp / ReportCount
        'End If

        '表格寬度的設定
        'If portrait = "H" Then
        '    int_width = Int((550 - Field1_width - Field2_width) / 19)
        '    strWatermarkImg = "TIMS_1.jpg"
        'Else
        int_width = 100 'Int((820 - Field1_width - Field2_width) / 19)
        strWatermarkImg = "TIMS_2.jpg"
        'End If

        If dt.Rows.Count > 0 Then

            'For i As Integer = 0 To PageCount
            '加背景圖的div
            If portrait = "H" Then
                intWatermarkTop = 0 * 800
            Else
                intWatermarkTop = 0 * 550
            End If
            strWatermarkDiv = "<div style='position:absolute;z-index:-1; margin:0;padding:0;left:0px;top: " + intWatermarkTop.ToString + "px;'><img src='../../images/rptpic/temple/" + strWatermarkImg + "' /></div>"
            nl = New HtmlGenericControl
            div_print.Controls.Add(nl)
            nl.InnerHtml = strWatermarkDiv

            '表頭
            nt = New HtmlTable
            nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", "0")
            div_print.Controls.Add(nt)

            'nr = New HtmlTableRow
            'nt.Controls.Add(nr)

            'nc = New HtmlTableCell
            'nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "100%")
            'nc.Attributes.Add("colspan", "2")
            'nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
            'nc.InnerHtml = ""

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "100%")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
            nc.InnerHtml = "公立職訓機構特定對象結訓人數"


            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "95%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", "font-size:9pt;font-family:DFKai-SB")
            nc.InnerHtml = gSYMD + "~" + gEYMD

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "5%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", "font-size:9pt;font-family:DFKai-SB")
            nc.InnerHtml = "單位：人"


            'Column Header
            nt = New HtmlTable
            nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", "2")
            div_print.Controls.Add(nt)

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "10%")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "訓練機構別"

            For j As Integer = 0 To col.Count - 1
                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2")
                nc.InnerHtml = col.Item(j)
            Next

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            For j As Integer = 0 To col.Count - 1
                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "男"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "女"
            Next

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "10%")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.InnerHtml = "總　　　　計"

            For j As Integer = 0 To col_total.Count - 1
                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = col_total.Item(j)
            Next

            For j As Integer = 0 To row.Count - 1
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(j).name

                For k As Integer = 0 To row.Item(j).count.Count - 1
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "10%")
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", strStyle)
                    nc.InnerHtml = row.Item(j).count.Item(k)
                Next

                For k As Integer = 0 To row.Item(j).group.Count - 1
                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "10%")
                    nc.Attributes.Add("align", "left")
                    nc.Attributes.Add("style", strStyle)
                    nc.InnerHtml = "&nbsp;" + row.Item(j).group.Item(k).name

                    For l As Integer = 0 To row.Item(j).group.Item(k).count.Count - 1
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = row.Item(j).group.Item(k).count.Item(l)
                    Next

                    For l As Integer = 0 To row.Item(j).group.Item(k).group.Count - 1
                        nr = New HtmlTableRow
                        nt.Controls.Add(nr)

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = "&nbsp;&nbsp;" + row.Item(j).group.Item(k).group.Item(l).name

                        For m As Integer = 0 To row.Item(j).group.Item(k).group.Item(l).count.Count - 1
                            nc = New HtmlTableCell
                            nr.Controls.Add(nc)
                            'nc.Attributes.Add("width", "10%")
                            nc.Attributes.Add("align", "center")
                            nc.Attributes.Add("style", strStyle)
                            nc.InnerHtml = row.Item(j).group.Item(k).group.Item(l).count.Item(m)
                        Next
                    Next

                Next

            Next

            nt = New HtmlTable
            nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", "0")
            div_print.Controls.Add(nt)

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "90%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
            nc.InnerHtml = "列印日期：" + Now().ToShortDateString()

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "10%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
            nc.InnerHtml = "頁數：" + (0 + 1).ToString + " / " + PageCount.ToString



[CONTINUE]:
            '表尾
            'If rsCursor + 1 > tmpDT.Rows.Count Then
            '    GoTo out
            'End If
            '換頁列印
            nl = New HtmlGenericControl
            div_print.Controls.Add(nl)
            nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
            'Next
out:
        End If

    End Sub

    Private Sub exc_Print(ByVal portrait As String)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return
        'Dim strScript As String = ""
        'strScript = "<script language=""javascript"">" + vbCrLf
        ''strScript = "function print() {"
        'strScript += "if (!factory.object) {"
        ''strScript += "return"
        'strScript += "} else {"
        'strScript += "factory.printing.header = """";"
        'strScript += "factory.printing.footer = """";"
        'strScript += "factory.printing.leftMargin = 5; "
        'strScript += "factory.printing.topMargin = 10; "
        'strScript += "factory.printing.rightMargin = 5; "
        'strScript += "factory.printing.bottomMargin = 10; "
        'strScript += "factory.printing.portrait = " + portrait + ";"
        'strScript += "factory.printing.Print(true);"
        'strScript += "window.close();"
        'strScript += "}"
        ''strScript += "}"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub


End Class
