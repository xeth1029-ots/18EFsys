Partial Class CP_02_001_R_Rpt
    Inherits AuthBasePage

    'Dim arrItem(113, 113) As String
    'Dim arrSpanRow(114) As String
    'Dim sourceDT As DataTable = Nothing
    'Dim PName As String = ""
    'Dim plankind As String = ""

    Public Class myrow
        Public name As String
        Public type As String
        Public total As Integer
        Public group As List(Of myrow)
        Public count As List(Of Integer)
        Public layer As Integer
    End Class

    Public Class mycol
        Public name As String
        Public group As List(Of String)
    End Class

    'Dim SYMD As String = "2012/01/01"
    'Dim EYMD As String = "2012/02/01"
    'Dim X As String = "性別,教育程度,學員身分"
    'Dim Y As String = "訓練機構,訓練職類,結訓後動向"

    Dim gSYMD As String = "2012/01/01"
    Dim gEYMD As String = "2012/02/01"

    Dim tDt_main As New DataTable
    Dim tDt_q1 As New DataTable

    Dim col As List(Of mycol) = New List(Of mycol)
    'Dim col_total As List(Of Integer) = New List(Of Integer)
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
        Dim X As String = "性別,教育程度,學員身分"
        Dim Y As String = "訓練機構,訓練職類,結訓後動向"

        start_date = Convert.ToString(Request("start_date")) '"2012/01/01"
        end_date = Convert.ToString(Request("end_date")) '"2012/02/01"
        TPlan = Convert.ToString(Request("TPlan")) '"37"
        OC_Type = Convert.ToString(Request("OC_Type")) '"1"
        X = Convert.ToString(Request("P1")) '"性別,教育程度,學員身分"
        Y = Convert.ToString(Request("P2")) '"訓練機構,訓練職類,結訓後動向"
        SYMD = Convert.ToString(Request("SYMD"))
        EYMD = Convert.ToString(Request("EYMD"))

        gSYMD = Convert.ToString(Request("SYMD"))
        gEYMD = Convert.ToString(Request("EYMD"))

        Dim tDt As New DataTable

        'STDate = Request("STDate")
        'STDate2 = Request("STDate2")
        'DistID = Request("DistID")
        'DistName = Request("DistName")
        'title = Request("title")

        tDt_main = db_main(start_date, end_date, TPlan, OC_Type, X, Y)
        tDt_q1 = db_q1(start_date, end_date, TPlan, OC_Type, X, Y)

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim l As Integer
        'Dim c As Integer
        Dim str As String
        Dim b As Boolean
        Dim parent As List(Of myrow)
        Dim layer As Integer
        Dim r As myrow
        Dim c As mycol
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim str6 As String
        Dim str7 As String
        Dim cv As Integer

        For i = 0 To tDt_main.Rows.Count - 1
            str = tDt_main.Rows(i).Item("Field1").ToString()
            str5 = tDt_main.Rows(i).Item("sex_desc").ToString()
            b = False
            For j = 0 To col.Count - 1
                If str = col.Item(j).name Then
                    b = True
                    Exit For
                End If
            Next

            If Not b Then
                c = New mycol
                c.name = str
                c.group = New List(Of String)
                c.group.Add(str5)
                col.Add(c)
            Else
                b = False
                For k = 0 To col.Item(j).group.Count - 1
                    If str5 = col.Item(j).group.Item(k) Then
                        b = True
                        Exit For
                    End If
                Next
                If Not b Then
                    col.Item(j).group.Add(str5)
                End If
            End If
        Next


        For i = 0 To tDt_main.Rows.Count - 1
            str = tDt_main.Rows(i).Item("Field1").ToString()
            str2 = tDt_main.Rows(i).Item("Field2").ToString()
            str3 = tDt_main.Rows(i).Item("Stat_desc").ToString()
            str4 = tDt_main.Rows(i).Item("Train1_desc").ToString()
            str5 = tDt_main.Rows(i).Item("sex_desc").ToString()
            str6 = tDt_main.Rows(i).Item("Stat_type").ToString()
            str7 = tDt_main.Rows(i).Item("Train1_type").ToString()
            b = False
            layer = 0
            parent = row
            k = 0
            For j = 0 To col.Count - 1
                If str = col.Item(j).name Then
                    For l = 0 To col.Item(j).group.Count - 1
                        If str5 = col.Item(j).group.Item(l) Then
                            Exit For
                        Else
                            k += 1
                        End If
                    Next
                    Exit For
                Else
                    k += col.Item(j).group.Count
                End If
            Next

            cv = tDt_main.Rows(i).Item("COUNTValue").ToString()
            For j = 0 To parent.Count - 1
                If str2 = parent.Item(j).name Then
                    parent.Item(j).total += cv
                    parent.Item(j).count(k) += cv
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
                        parent.Item(j).total += cv
                        parent.Item(j).count(k) += cv
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
                        parent.Item(j).total += cv
                        parent.Item(j).count.Item(k) += cv
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
                        For l = 0 To col.Item(j).group.Count - 1
                            r.count.Add(New Integer())
                        Next
                    Next
                    parent.Add(r)
                    r.total += cv
                    r.count.Item(k) += cv
                    layer = 1
                    parent = r.group
                End If

                If layer = 1 Then
                    If str3 <> "" Then
                        r = New myrow
                        r.name = str3
                        r.type = str6
                        r.group = New List(Of myrow)
                        r.count = New List(Of Integer)
                        For j = 0 To col.Count - 1
                            For l = 0 To col.Item(j).group.Count - 1
                                r.count.Add(New Integer())
                            Next
                        Next
                        parent.Add(r)
                        r.total += cv
                        r.count.Item(k) += cv
                        layer = 2
                        parent = r.group
                    End If
                End If

                If layer = 2 Then
                    If str4 <> "" Then
                        r = New myrow
                        r.name = str4
                        r.type = str7
                        r.group = New List(Of myrow)
                        r.count = New List(Of Integer)
                        For j = 0 To col.Count - 1
                            For l = 0 To col.Item(j).group.Count - 1
                                r.count.Add(New Integer())
                            Next
                        Next
                        parent.Add(r)
                        r.total += cv
                        r.count.Item(k) += cv
                    End If
                End If

            End If
        Next

        For j = 0 To row.Count - 1
            row.Item(j).total = 0
            For k = 0 To row.Item(j).group.Count - 1
                row.Item(j).group.Item(k).total = 0
                For l = 0 To row.Item(j).group.Item(k).group.Count - 1
                    row.Item(j).group.Item(k).group.Item(l).total = 0
                Next
            Next
        Next
        For i = 0 To tDt_q1.Rows.Count - 1
            str = tDt_q1.Rows(i).Item("Field2").ToString
            str2 = tDt_q1.Rows(i).Item("Stat_type").ToString
            str3 = tDt_q1.Rows(i).Item("Train1_type").ToString
            cv = tDt_q1.Rows(i).Item("T_Sex").ToString
            For j = 0 To row.Count - 1
                If str = row.Item(j).name Then
                    row.Item(j).total += cv
                    For k = 0 To row.Item(j).group.Count - 1
                        If str2 = row.Item(j).group.Item(k).type Then
                            row.Item(j).group.Item(k).total += cv
                            If str3 <> "" Then
                                For l = 0 To row.Item(j).group.Item(k).group.Count - 1
                                    If str3 = row.Item(j).group.Item(k).group.Item(l).type Then
                                        row.Item(j).group.Item(k).group.Item(l).total += cv
                                        Exit For
                                    End If
                                Next
                            End If
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        Next

        PrintDiv(tDt_main, "AAAA", "150", "150", 40, "5", "V")

    End Sub

    'SQL
    Private Function db_main(ByVal start_date As String, ByVal end_date As String, ByVal TPlan As String, ByVal OC_Type As String, ByVal X As String, ByVal Y As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " Select Stat_type,CONVERT(varchar, Stat_desc) as Stat_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Stat_type,S1.Stat_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From "
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc " & vbCrLf
        sql += "   union Select '2' as s_type,'女' as sex_desc " & vbCrLf
        sql += "  )  S2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += " (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '14' as h_age,'未滿15歲' as age_desc " & vbCrLf
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc " & vbCrLf
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc " & vbCrLf
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc " & vbCrLf
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc " & vbCrLf
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc " & vbCrLf
        sql += "  )  T1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age  and T2.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1 "
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += "  (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " ( Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  ( select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc " & vbCrLf
        sql += "   )  T1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,A.Trainice,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.UnitCode=S1.Stat_type"
        sql += "  "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID And I2.UnitCode=S1.Stat_type "
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc " & vbCrLf
        sql += " )  Q91 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc " & vbCrLf
        sql += " )  Q101 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 And Q102.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q11_type,'沒有' as Q11_desc " & vbCrLf
        sql += " )  Q111 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 And Q112.UnitCode=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,U2.Unit_type,CONVERT(varchar, U2.Unit_desc),dbo.NVL(U3.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'訓練機構' as Field2,1 as Field3 From"
        sql += "  (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += "  )  U1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U2 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U3 on U2.Unit_type=U3.UnitCode And U3.UnitCode=S1.Stat_type"
        sql += " ) A_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' + Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Trainice_type,Trainice_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Trainice_type,S1.Trainice_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "   (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "    union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc " & vbCrLf
        sql += "   union Select '1' as s_type,'男' as sex_desc " & vbCrLf
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " ) S3 on S2.s_type=S3.sex and S3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "  )  T1 )  S1 cross join"
        sql += " ( SELECT '14' as h_age,'未滿15歲' as age_desc " & vbCrLf
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc " & vbCrLf
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc " & vbCrLf
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc " & vbCrLf
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc " & vbCrLf
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc " & vbCrLf
        sql += "  )  T1 left join"
        sql += " (Select A.Trainice,A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age  and T2.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "  )  T1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name   as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.Trainice,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.Trainice,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID  and M2.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc " & vbCrLf
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.Trainice,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND  B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc " & vbCrLf
        sql += " )  Q91 left join"
        sql += " (Select A.Trainice,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc " & vbCrLf
        sql += " )  Q101 left join"
        sql += " (Select A.Trainice,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc " & vbCrLf
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc " & vbCrLf
        sql += " )  Q111 left join"
        sql += " (Select A.Trainice,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Trainice=S1.Trainice_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Trainice_type,S1.Trainice_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select T1.*,'訓練性質' as Field2,2 as Field3 From "
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc " & vbCrLf
        sql += "   )  T1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.Trainice,A.UnitCode,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Trainice=S1.Trainice_type"
        sql += " ) B_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' + Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Stat_type,Stat_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Stat_type,S1.Stat_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "   (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1"
        sql += " )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc " & vbCrLf
        sql += "   union Select '1' as s_type,'男' as sex_desc " & vbCrLf
        sql += "  )  S2"
        sql += "  left join"
        sql += " (Select "
        sql += "  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on  S2.s_type=S3.sex and S3.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "   (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += "  )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc " & vbCrLf
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc " & vbCrLf
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc " & vbCrLf
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc " & vbCrLf
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc " & vbCrLf
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc " & vbCrLf
        sql += "  )  T1 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,"
        sql += " case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as  h_age,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    union SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += "  )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    union SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf

        sql += "  )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   M2 on M1.Military_type=M2.MilitaryID and M2.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf

        sql += " )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc " & vbCrLf
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc " & vbCrLf
        sql += "   )  T2 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.Trainice,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   T3 on T2.Trainice_type=T3.Trainice and T3.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += "  )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   I2 on I1.Identity_type=I2.IdentityID and I2.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc " & vbCrLf
        sql += " )  Q91 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q92 on Q91.Q9_type=Q92.Q9 and Q92.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc " & vbCrLf
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc " & vbCrLf
        sql += " )  Q101 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q102 on Q101.Q10_type=Q102.Q10 and Q102.TrainCommend1=S1.Stat_type"
        sql += " /*Q111*/ "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc " & vbCrLf
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc " & vbCrLf
        sql += " )  Q111 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q112 on Q111.Q11_type=Q112.Q11 and Q112.TrainCommend1=S1.Stat_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Stat_type,S1.Stat_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as Field3,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'' as Train1_type,'' as Train1_desc,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT 'Y' as Stat_type,'否' as Stat_desc " & vbCrLf
        sql += "    UNION SELECT 'N' as Stat_type,'是' as Stat_desc " & vbCrLf
        sql += "    )  Z1" & vbCrLf
        sql += " )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select  Case when A.TrainCommend1='N' then 'Y'"
        sql += "   when A.TrainCommend1='Y' then 'N'"
        sql += "  end as TrainCommend1,A.UnitCode,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   U2 on U1.Unit_type=U2.UnitCode and U2.TrainCommend1=S1.Stat_type"
        sql += " ) C_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' + Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select 'N' as Stat_type,'是' as Stat_desc,Train1_type,Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "  )  Z1 "
        sql += " )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2"
        sql += "  left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and (S3.TrainCommend2=S1.Train1_type or S3.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age AND (T2.TrainCommend2=S1.Train1_type or T2.TrainCommend2 is null)"
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "  (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "  )  Z1 )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID AND (D2.TrainCommend2=S1.Train1_type or D2.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   M2 on M1.Military_type=M2.MilitaryID AND (M2.TrainCommend2=S1.Train1_type or M2.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.TrainCommend2,A.Trainice,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   T3 on T2.Trainice_type=T3.Trainice AND (T3.TrainCommend2=S1.Train1_type or T3.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "      (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   I2 on I1.Identity_type=I2.IdentityID AND (I2.TrainCommend2=S1.Train1_type or I2.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q92 on Q91.Q9_type=Q92.Q9 AND (Q92.TrainCommend2=S1.Train1_type or Q92.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.TrainCommend2,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q102 on Q101.Q10_type=Q102.Q10 AND (Q102.TrainCommend2=S1.Train1_type or Q102.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select  A.TrainCommend2,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   Q112 on Q111.Q11_type=Q112.Q11 AND (Q112.TrainCommend2=S1.Train1_type or Q112.TrainCommend2 is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as Field3,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,'委託訓練' as Field2,3 as Field3 From "
        sql += "     (SELECT '1' as Train1_type,'學校' as Train1_desc "
        sql += "      union SELECT '2' as  Train1_type,'民間企業或法人團體' as Train1_desc "
        sql += "      union SELECT '3' as  Train1_type,'公營企業' as Train1_desc "
        sql += "      union SELECT '4' as  Train1_type,'其他' as Train1_desc "
        sql += "     )  Z1 )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.TrainCommend2,A.UnitCode,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )"
        sql += "   U2 on U1.Unit_type=U2.UnitCode AND (U2.TrainCommend2=S1.Train1_type or U2.TrainCommend2 is null)"
        sql += " ) D_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%'+ Field1 +'%'"
        sql += " AND '" + Y + "' Like '%'+ Field2 +'%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Bus_type,CONVERT(varchar, Bus_desc),Unit_type,Unit_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where K1.TMID=K2.Parent"
        sql += "   and K2.TMID=K3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as  Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.Trainice,A.ResultDate,A.TMID,1 as COUNTValue,B.Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Trainice=S1.Unit_type and S3.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where K1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as  Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.Trainice,A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,A.TMID,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Trainice=S1.Unit_type and T2.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID "
        sql += "  from key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as  Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Trainice=S1.Unit_type and D2.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*, Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   from key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Trainice=S1.Unit_type and M2.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Trainice_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Trainice=S1.Unit_type and T3.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where K1.TMID=K2.Parent"
        sql += "   and K2.TMID=K3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as  Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.Trainice,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,A.TMID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND  B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Trainice=S1.Unit_type and I2.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   from key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Trainice=S1.Unit_type and Q92.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Trainice=S1.Unit_type and Q102.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,Q11.Q11_type,Q11.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "   union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q11 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q11.Q11_type=Q112.Q11 and Q112.Trainice=S1.Unit_type and Q112.TMID=S1.TMID"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Bus_type,S1.Bus_desc,S1.Unit_type,S1.Unit_desc,S2.Stat_type,CONVERT(varchar, S2.Stat_desc),dbo.NVL(S3.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Z1.*,Z2.*,'訓練職類' as Field2,4 as Field3 From "
        sql += "  (Select k1.BusID as Bus_type,k1.BusName as Bus_desc,k3.TMID"
        sql += "   From key_traintype k1,key_traintype k2,key_traintype k3 where k1.TMID=k2.Parent"
        sql += "   and k2.TMID=k3.Parent and K1.BusID <>'G')  Z1 cross join"
        sql += "  (SELECT '1' as Unit_type,'職前' as Unit_desc "
        sql += "   union SELECT '2' as Unit_type,'進修' as Unit_desc "
        sql += "  )  Z2 )  S1 cross join"
        sql += " (select StatID as Stat_type,StatName as Stat_desc from ID_StatistDist"
        sql += " )  S2 left join"
        sql += " (Select A.Trainice,A.TMID,A.ResultDate,A.UnitCode,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S3.UnitCode=S2.Stat_type and S3.Trainice=S1.Unit_type and S3.TMID=S1.TMID"
        sql += " ) E_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Time_type,Time_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Time_type,S1.Time_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  (Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "   )  S2 left join"
        sql += "  (Select A.SchoolTime,A.ResultDate,1 as COUNTValue,B.Sex"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex  and S3.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " ( Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += "  ( SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " ( Select A.SchoolTime,A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age  and T2.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,B.DegreeID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,B.MilitaryID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,A.Trainice,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.SchoolTime,A.ResultDate,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.SchoolTime=S1.Time_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Time_type,S1.Time_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'上課時段' as Field2,5 as Field3 From"
        sql += "  ( Select '1' as Time_type,'日間' as Time_desc "
        sql += "   union Select '2' as Time_type,'夜間' as Time_desc "
        sql += "   )  U1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.SchoolTime,A.UnitCode,A.ResultDate,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.SchoolTime=S1.Time_type"
        sql += " ) F_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Identity_type,CONVERT(varchar, Identity_desc),'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Identity_type,S1.Identity_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += "  (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "   Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.ResultDate,A.Trainice,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,I2.Identity_type,CONVERT(varchar, I2.Identity_desc),dbo.NVL(I3.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2, 6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I2 left join"
        sql += " (Select A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,B.Q8"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND  B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I3 on I2.Identity_type=I3.IdentityID and I3.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q11,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.IdentityID=S1.Identity_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Identity_type,S1.Identity_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select I1.*,'學員身分' as Field2,6 as Field3 From"
        sql += "  (SELECT IdentityID as Identity_type,Name as Identity_desc FROM Key_Identity"
        sql += "  )  I1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.IdentityID=S1.Identity_type"
        sql += " ) G_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q7_type,Q7_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue,B.Q7 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  ( Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  ( Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  ( Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.ResultDate,A.Trainice,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  ( Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  ( Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q7=S1.Q7_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '1' as Q7_type,'想學得一技之長以利就業' as Q7_desc "
        sql += "   union Select '2' as Q7_type,'為學習第二專長以利轉業' as Q7_desc "
        sql += "   union Select '3' as Q7_type,'為進一步學得技能以利升遷' as Q7_desc "
        sql += "   union Select '4' as Q7_type,'為充實實務經驗以利升學' as Q7_desc "
        sql += "   union Select '5' as Q7_type,'為參加技能檢定' as Q7_desc "
        sql += "   union Select '6' as Q7_type,'其他' as Q7_desc "
        sql += "   )  Q1)   S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q7=S1.Q7_type"
        sql += " ) H_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q7_type,Q7_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q7_type,S1.Q7_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc )  Q1 "
        sql += "   )  S1 cross join"
        sql += "  (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue,B.Q7 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.ResultDate,A.Trainice,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,B.Q7"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " ( Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q7,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q7 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q7_type,S1.Q7_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓動機' as Field2,7 as Field3 From"
        sql += "  (Select '7' as Q7_type,'未填答' as Q7_desc ) Q1"
        sql += "    )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q7,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q7 is null"
        sql += " ) H1_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q8_type,Q8_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q8_type,S1.Q8_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q8"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += "  (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue,B.Q8 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.ResultDate,A.Trainice,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,B.Q8"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q8=S1.Q8_type "
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc  "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q8=S1.Q8_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '1' as Q8_type,'尋找工作' as Q8_desc "
        sql += "   union Select '2' as Q8_type,'繼續進階課程或升學' as Q8_desc "
        sql += "   union Select '3' as Q8_type,'等待兵役' as Q8_desc "
        sql += "   union Select '4' as Q8_type,'留在原場(廠)服務' as Q8_desc "
        sql += "   union Select '5' as Q8_type,'其他' as Q8_desc "
        sql += "  )  U1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc   from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q8=S1.Q8_type"
        sql += " ) N_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q8_type,Q8_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q8_type,S1.Q8_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (Select '1' as s_type,'男' as sex_desc "
        sql += "   union Select '2' as s_type,'女' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q8"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += " )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   end as h_age,1 as COUNTValue,B.Q8 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " ( select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,T1.Trainice_type,T1.Trainice_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T1 left join"
        sql += " (Select A.ResultDate,A.Trainice,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.Trainice_type=T2.Trainice and T2.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,B.Q8"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q8,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q8 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q8_type,S1.Q8_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select U1.*,'結訓後動向' as Field2,8 as Field3 From"
        sql += "  (Select '6' as Q8_type,'未填答' as Q8_desc  "
        sql += "  )  U1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc   from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,B.Q8,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q8 is null"
        sql += " ) N1_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q9_type,Q9_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += "  Select S1.Q9_type,S1.Q9_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_Type, '性別' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as Q9_type,'沒有工作' as Q9_desc )  Q1 "
        sql += "   )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,"
        sql += "  Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " "
        sql += " Select S1.Q9_type,S1.Q9_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9 is not null"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'' as Train1_type,'' as Train1_desc,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "   union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q9=S1.Q9_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'N' as Q9_type,'有工作' as Q9_desc "
        sql += "    union SELECT 'Y' as  Q9_type,'沒有工作' as Q9_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,Case when B.Q9='N' then 'Y'"
        sql += "  when B.Q9='Y' then 'N'"
        sql += "  end as Q9"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q9=S1.Q9_type"
        sql += " )O_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select 'N' as Q9_type,'有' as Q9_desc,Train1_type,Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += "  Select S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_Type, '性別' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and (S3.Q9Y=S1.Train1_type or S3.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,B.Q9Y"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and (T2.Q9Y=S1.Train1_type or T2.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and (D2.Q9Y=S1.Train1_type or D2.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and (M2.Q9Y=S1.Train1_type or M2.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "   (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice And (T3.Q9Y=S1.Train1_type or T3.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "   (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q9Y,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and (I2.Q9Y=S1.Train1_type or I2.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += "  (Select Q2.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q2)  S1 cross join"
        sql += " (SELECT 'N' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'Y' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and (Q92.Q9Y=S1.Train1_type or Q92.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q9Y,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and (Q102.Q9Y=S1.Train1_type or Q102.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and (Q112.Q9Y=S1.Train1_type or Q112.Q9Y is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From"
        sql += "    (Select '1' as Train1_type,'從事1小時以上有報酬的工作' as Train1_desc "
        sql += "     union SELECT '2' as  Train1_type,'從事每週15小時以上無酬家屬工作' as Train1_desc "
        sql += "     union SELECT '3' as  Train1_type,'有工作而未做領有報酬' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and (U2.Q9Y=S1.Train1_type or U2.Q9Y is null)"
        sql += " ) P_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q9_type,Q9_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += "  Select S1.Q9_type,S1.Q9_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_Type, '性別' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += "  (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "   )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q9,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "   when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,B.Q9,B.Q9Y From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q9,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military )  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q9,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Q9,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q9Y,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q9 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q9_type,S1.Q9_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += "  (Select Q1.*,'參訓前一個月工作情形' as Field2,9 as Field3 From "
        sql += "   (SELECT 'Z' as Q9_type,'未填答' as Q9_desc  )  Q1"
        sql += "  )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Q9,B.Q9Y"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q9 is null"
        sql += " ) Q_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q10_type,Q10_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q10_type,S1.Q10_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,"
        sql += " Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9  and Q92.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q10=S1.Q10_type "
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q10=S1.Q10_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'N' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'Y' as  Q10_type,'沒有' as Q10_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc   from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,Case when B.Q10='N' then 'Y'"
        sql += "  when B.Q10='Y' then 'N'"
        sql += "  end as Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q10=S1.Q10_type"
        sql += " ) R_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q10_type,Q10_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q10_type,S1.Q10_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,B.Q10 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q10,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "   left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q10,B.Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q10 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q10_type,S1.Q10_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'參訓前一個月有否尋找工作' as Field2,91 as Field3 From "
        sql += "   (SELECT 'Z' as Q10_type,'未填答' as Q10_desc  "
        sql += "    )  Q1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc   from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Q10"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q10 is null"
        sql += " ) S_Table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q11_type,Q11_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q11_type,S1.Q11_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += " )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,"
        sql += " Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc "
        sql += "   )  Q1 )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += "  )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += " )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += "   )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "     union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "    union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N' end as Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "     union SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "     UNION SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += "  )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q11=S1.Q11_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'N' as Q11_type,'已找到工作' as Q11_desc "
        sql += "     UNION SELECT 'Y' as  Q11_type,'未找到工作' as Q11_desc )  Q1"
        sql += "  )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,Case when B.Q11='N' then 'Y'"
        sql += "  when B.Q11='Y' then 'N'"
        sql += "  end as Q11"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q11=S1.Q11_type"
        sql += " ) T_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select 'Y' AS Q11_type,'未找到工作' as Q11_desc,Train1_type,Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Train1_type,S1.Train1_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and (S3.Q11N=S1.Train1_type or S3.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,B.Q11N"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and (T2.Q11N=S1.Train1_type or T2.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and (D2.Q11N=S1.Train1_type or D2.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and (M2.Q11N=S1.Train1_type or M2.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "  )  Q1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and (T3.Q11N=S1.Train1_type or T3.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q11N,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and (I2.Q11N=S1.Train1_type or I2.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'N'as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and (Q92.Q11N=S1.Train1_type or Q92.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and (Q102.Q11N=S1.Train1_type or Q102.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and (Q112.Q11N=S1.Train1_type or Q112.Q11N is null)"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Train1_type,S1.Train1_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From"
        sql += "  (SELECT '1' as Train1_type,'非常有幫助' as Train1_desc "
        sql += "   union SELECT '2' as  Train1_type,'有幫助' as Train1_desc "
        sql += "   union SELECT '3' as  Train1_type,'普通' as Train1_desc "
        sql += "   union SELECT '4' as  Train1_type,'沒有幫助' as Train1_desc "
        sql += "   union SELECT '5' as  Train1_type,'完全沒有幫助' as Train1_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode And (U2.Q11N=S1.Train1_type or U2.Q11N is null)"
        sql += " ) U_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Q11_type,Q11_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q11_type,S1.Q11_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,B.Q11,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "    (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,B.Q11,B.Q11N From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,B.Q11,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q11,B.Q11N,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q11,B.Q11N,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q11 is null"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q11_type,S1.Q11_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'訓練後或訓練期間有否找到工作' as Field2,92 as Field3 From "
        sql += "   (SELECT 'Z' as Q11_type,'未填答' as Q11_desc  )  Q1"
        sql += " )  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,B.Q11,B.Q11N"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q11 is null"
        sql += " ) V_table"
        sql += " Where 1=1"
        sql += " AND '" + X + "' Like '%' +  Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"

        sql += " UNION ALL" & vbCrLf
        sql += " Select Q12_type,Q12_desc,'' as Train1_type,'' as Train1_desc,s_type,sex_desc,dbo.NVL(COUNTValue,0) as COUNTValue,A_type,Field1,Field2,Field3 From ("
        sql += " Select S1.Q12_type,S1.Q12_desc,S2.s_type,S2.sex_desc,dbo.NVL(S3.COUNTValue,0) as COUNTValue,1 as A_type, '性別' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (Select '2' as s_type,'女' as sex_desc "
        sql += "   union Select '1' as s_type,'男' as sex_desc "
        sql += "  )  S2 left join"
        sql += " (Select A.ResultDate,1 as COUNTValue,B.Sex,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  S3 on S2.s_type=S3.sex and S3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,T1.h_age,T1.age_desc,dbo.NVL(T2.COUNTValue,0) as COUNTValue,2 as A_type, '年齡' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT '14' as h_age,'未滿15歲' as age_desc "
        sql += "   union SELECT '24' as h_age,'15歲-24歲' as age_desc "
        sql += "   union SELECT '34' as h_age,'25歲-34歲' as age_desc "
        sql += "   union SELECT '44' as h_age,'35歲-44歲' as age_desc "
        sql += "   union SELECT '54' as h_age,'45歲-54歲' as age_desc "
        sql += "   union SELECT '99' as h_age,'55歲以上' as age_desc "
        sql += "  )  T1 left join"
        sql += " (Select A.ResultDate,case"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=0"
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=14 then '14'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=15 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=24 then '24'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=25 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=34 then '34'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=35 "
        sql += "  AND  (datepart(year,getdate())-datepart(year,B.BirthYear))<=44 then '44'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=45 "
        sql += "  AND (datepart(year,getdate())-datepart(year,B.BirthYear))<=54 then '54'"
        sql += "  when (datepart(year,getdate())-datepart(year,B.BirthYear))>=55 then '99'"
        sql += "  end as h_age,1 as COUNTValue,C.Q12 From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T2 on T1.h_age=T2.h_age and T2.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,D1.Degree_type,CONVERT(varchar, D1.Degree_desc),dbo.NVL(D2.COUNTValue,0) as COUNTValue,3 as A_type,'教育程度' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select DegreeID as Degree_type, Name as Degree_desc  from Key_Degree)  D1 left join"
        sql += " (Select A.ResultDate,B.DegreeID,C.Q12,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  D2 on D1.Degree_type=D2.DegreeID and D2.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,M1.Military_type,CONVERT(varchar, M1.Military_desc),dbo.NVL(M2.COUNTValue,0) as COUNTValue,4 as A_type,'兵役' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (select  MilitaryID as Military_type,Name as Military_desc from Key_Military)  M1 left join"
        sql += " (Select A.ResultDate,B.MilitaryID,C.Q12,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  M2 on M1.Military_type=M2.MilitaryID and M2.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,T2.Trainice_type,T2.Trainice_desc,dbo.NVL(T3.COUNTValue,0) as COUNTValue,5 as A_type,'訓練性質' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += "  (SELECT '1' as Trainice_type,'職前' as Trainice_desc "
        sql += "   union SELECT '2' as  Trainice_type,'進修' as Military_desc "
        sql += "   )  T2 left join"
        sql += " (Select A.Trainice,A.ResultDate,1 as COUNTValue,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  T3 on T2.Trainice_type=T3.Trainice and T3.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,I1.Identity_type,CONVERT(varchar, I1.Identity_desc),dbo.NVL(I2.COUNTValue,0) as COUNTValue,6 as A_type,'學員身分' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT IdentityID as Identity_type,Name as Identity_desc  FROM Key_Identity"
        sql += " )  I1 left join"
        sql += " (Select A.ResultDate,B.Q11,dbo.NVL(ky.MergeID,C.IdentityID) as IdentityID,1 as COUNTValue,D.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Left Join Stud_ResultTwelveData D on B.DLID=D.DLID AND B.SubNo=D.SubNo"
        sql += "  left  join Key_Identity ky on C.IdentityID = ky.IdentityID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  I2 on I1.Identity_type=I2.IdentityID and I2.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,Q91.Q9_type,Q91.Q9_desc,dbo.NVL(Q92.COUNTValue,0) as COUNTValue,7 as A_type,'參訓前一個月工作情形' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q9_type,'有' as Q9_desc "
        sql += "   union SELECT 'N' as  Q9_type,'沒有' as Q9_desc "
        sql += " )  Q91 left join"
        sql += " (Select A.ResultDate,B.Q9,C.Q12,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q92 on Q91.Q9_type=Q92.Q9 and Q92.Q12=S1.Q12_type "
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,Q101.Q10_type,Q101.Q10_desc,dbo.NVL(Q102.COUNTValue,0) as COUNTValue,8 as A_type,'參訓前一個月有否尋找工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q10_type,'有' as Q10_desc "
        sql += "   union SELECT 'N' as  Q10_type,'沒有' as Q10_desc "
        sql += " )  Q101 left join"
        sql += " (Select A.ResultDate,B.Q10,C.Q12,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q102 on Q101.Q10_type=Q102.Q10 and Q102.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,Q111.Q11_type,Q111.Q11_desc,dbo.NVL(Q112.COUNTValue,0) as COUNTValue,9 as A_type,'您參加本次訓練後是否找到工作' as Field1,Field2,Field3 From"
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc "
        sql += "   )  Q1)  S1 cross join"
        sql += " (SELECT 'Y' as Q11_type,'有' as Q11_desc "
        sql += "  union SELECT 'N' as  Q11_type,'沒有' as Q11_desc "
        sql += " )  Q111 left join"
        sql += " (Select A.ResultDate,B.Q11,C.Q12,1 as COUNTValue"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  Q112 on Q111.Q11_type=Q112.Q11 and Q112.Q12=S1.Q12_type"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select S1.Q12_type,S1.Q12_desc,U1.Unit_type,CONVERT(varchar, U1.Unit_desc),dbo.NVL(U2.COUNTValue,0) as COUNTValue,10 as A_type,'訓練機構' as Field1,Field2,Field3 From"
        'ID_StatistDist
        sql += " (Select Q1.*,'本次訓練後覺得不滿意需要改進的為何' as Field2,93 as Field3 From "
        sql += "   (SELECT '1' as Q12_type,'參訓職類不符就業市場需求' as Q12_desc "
        sql += "    union SELECT '2' as  Q12_type,'教學課程安排不當' as Q12_desc "
        sql += "    union SELECT '3' as  Q12_type,'訓練師專業及熱忱不足' as Q12_desc "
        sql += "    union SELECT '4' as  Q12_type,'訓練設備不符產業需求' as Q12_desc "
        sql += "    union SELECT '5' as  Q12_type,'其他' as Q12_desc from ID_StatistDist"
        sql += "   )  Q1)  S1 cross join"
        sql += " (select StatID as Unit_type,StatName as Unit_desc   from ID_StatistDist"
        sql += " )  U1 left join"
        sql += " (Select A.UnitCode,A.ResultDate,1 as COUNTValue,C.Q12"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += "  And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += "  AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += "  and I.Type=" + OC_Type
        End If
        sql += " )  U2 on U1.Unit_type=U2.UnitCode and U2.Q12=S1.Q12_type"
        sql += " ) W_table Where 1=1"
        sql += " AND '" + X + "' Like '%' + Field1 + '%'"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    'SQL
    Private Function db_q1(ByVal start_date As String, ByVal end_date As String, ByVal TPlan As String, ByVal OC_Type As String, ByVal X As String, ByVal Y As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex from"
        sql += " (Select  '訓練機構' as Field2,CONVERT(varchar, A.UnitCode) as Stat_type,A.ResultDate,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.UnitCode,A.ResultDate"
        sql += " ) A_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex from"
        sql += " (Select '訓練性質' as Field2,convert(varchar, A.Trainice ) as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += " Group by A.Trainice"
        sql += " ) B_table "
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2,"
        sql += "  Case when A.TrainCommend1='Y' then 'N'"
        sql += "  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1"
        sql += " ) C_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'1' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2,"
        sql += "  Case when A.TrainCommend1='Y' then 'N'"
        sql += "  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y' And A.TrainCommend2=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1,A.TrainCommend2"
        sql += " ) C1_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'2' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2,"
        sql += "  Case when A.TrainCommend1='Y' then 'N' end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y' And A.TrainCommend2=2"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1,A.TrainCommend2"
        sql += " ) C2_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'3' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2,"
        sql += "  Case when A.TrainCommend1='Y' then 'N' end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y' And A.TrainCommend2=3"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1,A.TrainCommend2"
        sql += " ) C3_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'4' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2,"
        sql += "  Case when A.TrainCommend1='Y' then 'N' end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.TrainCommend1='Y' And A.TrainCommend2=4"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1,A.TrainCommend2"
        sql += " ) C4_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex from"
        sql += " (Select '委託訓練' as Field2, Case when A.TrainCommend1='N' then 'Y'"
        sql += "  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where  A.TrainCommend1='N'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by A.TrainCommend1"
        sql += " ) D_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Bus_type,'1' as Train1_type,T_Sex From "
        sql += "  (Select '訓練職類' as Field2,k1.BusID as Bus_type,Count(B.Sex) as T_Sex"
        sql += "  From key_traintype K1"
        sql += "  Join key_traintype K2 on K1.TMID=K2.Parent"
        sql += "  Join key_traintype K3 on K2.TMID=K3.Parent"
        sql += "  Join Stud_DataLid A on K3.TMID=A.TMID"
        sql += "  Join ID_StatistDist I on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.Trainice=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by K1.BusID"
        sql += " ) E1_table"
        sql += " WHERE 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Bus_type,'2' as Train1_type,T_Sex From "
        sql += "  (Select '訓練職類' as Field2,k1.BusID as Bus_type,Count(B.Sex) as T_Sex"
        sql += "  From key_traintype K1"
        sql += "  Join key_traintype K2 on K1.TMID=K2.Parent"
        sql += "  Join key_traintype K3 on K2.TMID=K3.Parent"
        sql += "  Join Stud_DataLid A on K3.TMID=A.TMID"
        sql += "  Join ID_StatistDist I on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where A.Trainice=2"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by K1.BusID"
        sql += " ) E2_table"
        sql += " WHERE 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From"
        sql += " (Select '上課時段' as Field2, CONVERT(varchar, A.SchoolTime) as Stat_type,Count(B.Sex) as T_Sex"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "   Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "   Where 1=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group by A.SchoolTime"
        sql += " ) F_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '學員身分' as Field2,CONVERT(varchar, C.IdentityID) as Stat_type,Count(B.Sex) as T_Sex"
        sql += "   FROM ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  left join Stud_ResultIdentData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where 1=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By C.IdentityID"
        sql += " ) G_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓動機' as Field2,CONVERT(varchar, B.Q7) as Stat_type,Count(B.Sex) as T_Sex"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1 "
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += " Group By B.Q7"
        sql += " ) H_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓動機' as Field2,'7' as Stat_type,count(B.Sex) as T_Sex"
        sql += "   From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where 1=1 AND B.Q7 is null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += " Group By B.Q7"
        sql += " ) H1_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From"
        sql += " (Select '結訓後動向' as Field2,CONVERT(varchar, B.Q8) as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q8 is not null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q8"
        sql += " ) N_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From"
        sql += " (Select '結訓後動向' as Field2,'6' as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q8 is null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q8"
        sql += " ) N1_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,"
        sql += "  Case when B.Q9='Y' then 'N'  end as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) O_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'1' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,"
        sql += "  Case when B.Q9='Y' then 'N'  end as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'  And B.Q9Y=1"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) O1_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'2' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,"
        sql += "  Case when B.Q9='Y' then 'N'  end as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'  And B.Q9Y=2"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) O2_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'3' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,"
        sql += "  Case when B.Q9='Y' then 'N'  end as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='Y'  And B.Q9Y=3"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) O3_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,"
        sql += "  Case when B.Q9='N' then 'Y'  end as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9='N'"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) P_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月工作情形' as Field2,'Z' as Stat_type,count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q9 is null"
        If start_date <> "" Then
            sql += "  And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q9"
        sql += " ) Q_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月有否尋找工作' as Field2,"
        sql += " Case  when B.Q10='Y' then 'N' end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q10='Y'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q10"
        sql += " ) R1_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月有否尋找工作' as Field2,"
        sql += " Case when B.Q10='N' then 'Y' end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q10='N'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q10"
        sql += " ) R2_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '參訓前一個月有否尋找工作' as Field2,'Z' as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode  "
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q10 is null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  Group By B.Q10"
        sql += " ) S_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='Y' then 'N'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='Y'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) T_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N'"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'1' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N' and B.Q11N=1"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U1_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'2' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N' and B.Q11N=2"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U2_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'3' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N' and B.Q11N=3"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U3_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'4' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N' and B.Q11N=4"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U4_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'5' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,"
        sql += " Case when B.Q11='N' then 'Y'  end as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11='N' and B.Q11N=5"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) U5_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From "
        sql += " (Select '訓練後或訓練期間有否找到工作' as Field2,'Z' as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Where B.Q11 is null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += "  group by B.Q11"
        sql += " ) V_table"
        sql += " Where 1=1"
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"
        sql += " "
        sql += " UNION ALL" & vbCrLf
        sql += " Select Field2,Stat_type,'' as Train1_type,T_Sex From"
        sql += " (Select '本次訓練後覺得不滿意需要改進的為何' as Field2,CONVERT(varchar, C.Q12) as Stat_type,Count(B.Sex) as T_Sex"
        sql += "  From ID_StatistDist I Join Stud_DataLid A on I.StatID=A.UnitCode"
        sql += "  Join Stud_ResultStudData B on A.DLID=B.DLID"
        sql += "  Left Join Stud_ResultTwelveData C on B.DLID=C.DLID AND B.SubNo=C.SubNo"
        sql += "  Where C.Q12 is not null"
        If start_date <> "" Then
            sql += " And A.ResultDate>=to_date('" + start_date + "','YYYY/MM/DD')"
        End If
        If end_date <> "" Then
            sql += " And A.ResultDate<=to_date('" + end_date + "','YYYY/MM/DD')"
        End If
        If TPlan <> "" Then
            sql += " AND A.TPlanID in (" + TPlan + ")"
        End If
        If OC_Type <> "" Then
            sql += " and I.Type=" + OC_Type
        End If
        sql += " group by C.Q12"
        sql += " ) W_table"
        sql += " Where 1=1 "
        sql += " AND '" + Y + "' Like '%' + Field2 + '%'"

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

        Dim r As myrow
        Dim list As List(Of Object)
        list = New List(Of Object)

        Dim indexs As List(Of Integer)
        indexs = New List(Of Integer)
        Dim le As Integer
        Dim len As Integer
        Dim olen As Integer

        tmpDT = dt
        ColCount = dt.Columns.Count

        intTmp = tmpDT.Rows.Count
        PageCount = 1

        len = 0
        indexs.Add(list.Count)
        For i As Integer = 0 To row.Count - 1
            r = row.Item(i)
            r.layer = 0
            list.Add(r)

            le = System.Text.Encoding.Default.GetBytes(r.name).Length
            olen = len
            If (le / 4) * 4 < le Then
                len += (le / 4) + 1
            Else
                len += (le / 4)
            End If
            If len > 25 Then
                indexs.Add(list.Count - 1)
                len -= olen
            End If

            For j As Integer = 0 To row.Item(i).group.Count - 1
                r = row.Item(i).group.Item(j)
                r.layer = 1
                list.Add(r)

                le = System.Text.Encoding.Default.GetBytes(r.name).Length + 2
                olen = len
                If (le / 4) * 4 < le Then
                    len += (le / 4) + 1
                Else
                    len += (le / 4)
                End If
                If len > 25 Then
                    indexs.Add(list.Count - 1)
                    len -= olen
                End If

                For k As Integer = 0 To row.Item(i).group.Item(j).group.Count - 1
                    r = row.Item(i).group.Item(j).group.Item(k)
                    r.layer = 2
                    list.Add(r)

                    le = System.Text.Encoding.Default.GetBytes(r.name).Length + 4
                    olen = len
                    If (le / 4) * 4 < le Then
                        len += (le / 4) + 1
                    Else
                        len += (le / 4)
                    End If
                    If len > 25 Then
                        indexs.Add(list.Count - 1)
                        len -= olen
                    End If

                Next
            Next
        Next
        If len > 0 Then
            indexs.Add(list.Count)
        End If
        PageCount = indexs.Count - 1
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


            For i As Integer = 0 To PageCount - 1
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

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "公立職訓機構結訓學員概況統計表"


                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "95%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = gSYMD + "~" + gEYMD

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
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
                nc.Attributes.Add("width", "100")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", "2")
                nc.InnerHtml = "訓練機構別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "10")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", "2")
                nc.InnerHtml = "總計"

                For j As Integer = 0 To col.Count - 1
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "10%")
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", strStyle)
                    nc.Attributes.Add("colspan", col.Item(j).group.Count)
                    nc.InnerHtml = col.Item(j).name
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                For j As Integer = 0 To col.Count - 1
                    For k As Integer = 0 To col.Item(j).group.Count - 1
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = Mid(col.Item(j).group.Item(k), 1, 4)
                    Next
                Next

                'nr = New HtmlTableRow
                'nt.Controls.Add(nr)

                'nc = New HtmlTableCell
                'nr.Controls.Add(nc)
                ''nc.Attributes.Add("width", "10%")
                'nc.Attributes.Add("align", "center")
                'nc.Attributes.Add("style", strStyle)
                'nc.InnerHtml = "總　　　　計"

                'For j As Integer = 0 To col_total.Count - 1
                '    nc = New HtmlTableCell
                '    nr.Controls.Add(nc)
                '    'nc.Attributes.Add("width", "10%")
                '    nc.Attributes.Add("align", "center")
                '    nc.Attributes.Add("style", strStyle)
                '    nc.InnerHtml = col_total.Item(j)
                'Next

                For j As Integer = indexs.Item(i) To indexs.Item(i + 1) - 1
                    r = list.Item(j)
                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "10%")
                    nc.Attributes.Add("align", "left")
                    nc.Attributes.Add("style", strStyle)
                    If r.layer = 1 Then
                        nc.InnerHtml = "&nbsp;" + r.name
                    ElseIf r.layer = 2 Then
                        nc.InnerHtml = "&nbsp;&nbsp;" + r.name
                    Else
                        nc.InnerHtml = r.name
                    End If


                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "10%")
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", strStyle)
                    nc.InnerHtml = r.total
                    For k As Integer = 0 To r.count.Count - 1
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "10%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = r.count.Item(k)
                    Next
                Next

                'For j As Integer = 0 To row.Count - 1
                '    nr = New HtmlTableRow
                '    nt.Controls.Add(nr)

                '    nc = New HtmlTableCell
                '    nr.Controls.Add(nc)
                '    'nc.Attributes.Add("width", "10%")
                '    nc.Attributes.Add("align", "left")
                '    nc.Attributes.Add("style", strStyle)
                '    nc.InnerHtml = row.Item(j).name

                '    nc = New HtmlTableCell
                '    nr.Controls.Add(nc)
                '    'nc.Attributes.Add("width", "10%")
                '    nc.Attributes.Add("align", "center")
                '    nc.Attributes.Add("style", strStyle)
                '    nc.InnerHtml = row.Item(j).total
                '    For k As Integer = 0 To row.Item(j).count.Count - 1
                '        nc = New HtmlTableCell
                '        nr.Controls.Add(nc)
                '        'nc.Attributes.Add("width", "10%")
                '        nc.Attributes.Add("align", "center")
                '        nc.Attributes.Add("style", strStyle)
                '        nc.InnerHtml = row.Item(j).count.Item(k)
                '    Next

                '    For k As Integer = 0 To row.Item(j).group.Count - 1
                '        nr = New HtmlTableRow
                '        nt.Controls.Add(nr)

                '        nc = New HtmlTableCell
                '        nr.Controls.Add(nc)
                '        'nc.Attributes.Add("width", "10%")
                '        nc.Attributes.Add("align", "left")
                '        nc.Attributes.Add("style", strStyle)
                '        nc.InnerHtml = "&nbsp;" + row.Item(j).group.Item(k).name

                '        nc = New HtmlTableCell
                '        nr.Controls.Add(nc)
                '        'nc.Attributes.Add("width", "10%")
                '        nc.Attributes.Add("align", "center")
                '        nc.Attributes.Add("style", strStyle)
                '        nc.InnerHtml = row.Item(j).group.Item(k).total
                '        For l As Integer = 0 To row.Item(j).group.Item(k).count.Count - 1
                '            nc = New HtmlTableCell
                '            nr.Controls.Add(nc)
                '            'nc.Attributes.Add("width", "10%")
                '            nc.Attributes.Add("align", "center")
                '            nc.Attributes.Add("style", strStyle)
                '            nc.InnerHtml = row.Item(j).group.Item(k).count.Item(l)
                '        Next

                '        For l As Integer = 0 To row.Item(j).group.Item(k).group.Count - 1
                '            nr = New HtmlTableRow
                '            nt.Controls.Add(nr)

                '            nc = New HtmlTableCell
                '            nr.Controls.Add(nc)
                '            'nc.Attributes.Add("width", "10%")
                '            nc.Attributes.Add("align", "left")
                '            nc.Attributes.Add("style", strStyle)
                '            nc.InnerHtml = "&nbsp;&nbsp;" + row.Item(j).group.Item(k).group.Item(l).name

                '            nc = New HtmlTableCell
                '            nr.Controls.Add(nc)
                '            'nc.Attributes.Add("width", "10%")
                '            nc.Attributes.Add("align", "center")
                '            nc.Attributes.Add("style", strStyle)
                '            nc.InnerHtml = row.Item(j).group.Item(k).group.Item(l).total

                '            For m As Integer = 0 To row.Item(j).group.Item(k).group.Item(l).count.Count - 1
                '                nc = New HtmlTableCell
                '                nr.Controls.Add(nc)
                '                'nc.Attributes.Add("width", "10%")
                '                nc.Attributes.Add("align", "center")
                '                nc.Attributes.Add("style", strStyle)
                '                nc.InnerHtml = row.Item(j).group.Item(k).group.Item(l).count.Item(m)
                '            Next
                '        Next

                '    Next

                'Next

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
                nc.InnerHtml = "頁數：" + (i + 1).ToString + " / " + PageCount.ToString



[CONTINUE]:
                '表尾

                'If rsCursor + 1 > tmpDT.Rows.Count Then
                '    GoTo out
                'End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"

            Next
out:
        End If

        exc_Print("false")

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
