'Imports Turbo
'Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml
'Imports System.Data
'Imports System.Web.HttpServerUtility

Public Class ETRAIN

    Public Shared Function Main(ByRef MyPage As Page) As String
        Dim ErrorMsg As String = ""
        'Dim strXmlFilePath, path_str As String

        '產生要輸出的XML檔案
        Dim date_str As String = Common.FormatNow()

        'Dim MyDir As System.IO.Directory
        Dim DELXML As String = Path.GetFullPath(System.Configuration.ConfigurationSettings.AppSettings("DELXML"))
        Dim MyDirPath1 As String = Path.GetFullPath(System.Configuration.ConfigurationSettings.AppSettings("XMLpath1"))

        Dim pDELXML As String = MyPage.Request.PhysicalApplicationPath & "\SYS\07\DELXML.CMD"

        Try
            If IO.Directory.Exists(MyDirPath1) = False Then
                Try
                    '建立暫存目錄，如果失敗，表示該目錄的位置安全性不足(要設定Everyone)
                    IO.Directory.CreateDirectory(MyDirPath1)
                Catch ex As Exception
                    MyDirPath1 = Path.GetFullPath("\")
                End Try
            End If

            Try
                Shell(pDELXML, AppWinStyle.Hide, False, -1)
            Catch ex As Exception
                Shell(DELXML, AppWinStyle.Hide, False, -1)
            End Try

            Dim strXmlFilePath As String = ""
            '104
            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Auth_Relship" & Common.FormatNow2() & ".xml"
            ETRAIN.Auth_Relship(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Class_ClassInfo" & Common.FormatNow2() & ".xml"
            ETRAIN.Class_ClassInfo(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Class_ClassLevel" & Common.FormatNow2() & ".xml"
            ETRAIN.Class_ClassLevel(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Class_Plan_Info" & Common.FormatNow2() & ".xml"
            ETRAIN.Class_Plan_Info(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\ID_Plan" & Common.FormatNow2() & ".xml"
            ETRAIN.ID_Plan(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Org_OrgInfo" & Common.FormatNow2() & ".xml"
            ETRAIN.Org_OrgInfo(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Org_OrgPlanInfo" & Common.FormatNow2() & ".xml"
            ETRAIN.Org_OrgPlanInfo(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Org_PlanYear" & Common.FormatNow2() & ".xml"
            ETRAIN.Org_PlanYear(strXmlFilePath)

            strXmlFilePath = System.Configuration.ConfigurationSettings.AppSettings("XMLpath1") & "\Plan_PlanInfo" & Common.FormatNow2() & ".xml"
            ETRAIN.Plan_PlanInfo(strXmlFilePath)

        Catch ex As Exception
            ErrorMsg = ex.ToString()

        End Try

        Return ErrorMsg
    End Function

    '業務關係檔
    Public Shared Sub Auth_Relship(ByVal strXmlFilePath As String)
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter
        'Dim Server As System.Web.HttpServerUtility

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Auth_Relship"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案

        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("RSID", objrow("RSID"))

            writer.WriteElementString("PlanID", objrow("PlanID")) '計畫代碼

            writer.WriteElementString("RID", objrow("RID"))

            writer.WriteElementString("OrgID", objrow("OrgID")) '機構代碼

            writer.WriteElementString("Relship", objrow("Relship")) '業務關係

            writer.WriteElementString("OrgLevel", objrow("OrgLevel")) '機構階層

            writer.WriteElementString("DistID", objrow("DistID")) '轄區中心代碼

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Sub

    '班級基本資料
    Public Shared Sub Class_ClassInfo(ByVal strXmlFilePath As String)
        Dim conn As OracleConnection

        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter
        'Dim Server As System.Web.HttpServerUtility

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Class_ClassInfo where ModifyDate>='" & Common.FormatDate(Now()) & "' or FTDate >='" & Common.FormatDate(Now()) & "'"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("OCID", objrow("OCID"))

            writer.WriteElementString("RID", objrow("RID"))

            writer.WriteElementString("CLSID", objrow("CLSID"))

            writer.WriteElementString("PlanID", objrow("PlanID")) '計畫代碼

            writer.WriteElementString("CyclType", objrow("CyclType")) '期別

            If Convert.IsDBNull(objrow("LevelType")) Then '階段
                writer.WriteElementString("LevelType", "")
            Else
                writer.WriteElementString("LevelType", objrow("LevelType"))
            End If

            writer.WriteElementString("ClassCName", objrow("ClassCName")) '班別中文名稱

            If Convert.IsDBNull(objrow("ClassEngName")) Then '班別英文名稱
                writer.WriteElementString("ClassEngName", "")
            Else
                writer.WriteElementString("ClassEngName", objrow("ClassEngName"))
            End If

            writer.WriteElementString("FTDate", objrow("FTDate")) '結訓日期

            writer.WriteElementString("STDate", objrow("STDate")) '開訓日期

            If Convert.IsDBNull(objrow("SEnterDate")) Then '報名起日期
                writer.WriteElementString("SEnterDate", "")
            Else
                writer.WriteElementString("SEnterDate", objrow("SEnterDate"))
            End If

            If Convert.IsDBNull(objrow("FEnterDate")) Then '報名迄日期
                writer.WriteElementString("FEnterDate", "")
            Else
                writer.WriteElementString("FEnterDate", objrow("FEnterDate"))
            End If

            writer.WriteElementString("TNum", objrow("TNum")) '訓練人數

            writer.WriteElementString("THours", objrow("THours")) '訓練時數

            If Convert.IsDBNull(objrow("ComIDNo")) Then '廠商統編
                writer.WriteElementString("ComIDNo", "")
            Else
                writer.WriteElementString("ComIDNo", objrow("ComIDNo"))
            End If

            If Convert.IsDBNull(objrow("TaddressZip")) Then '訓練地點Zip
                writer.WriteElementString("TaddressZip", "")
            Else
                writer.WriteElementString("TaddressZip", objrow("TaddressZip"))
            End If

            If Convert.IsDBNull(objrow("TAddress")) Then '訓練地點
                writer.WriteElementString("TAddress", "")
            Else
                writer.WriteElementString("TAddress", objrow("TAddress"))
            End If

            writer.WriteElementString("TMID", objrow("TMID")) '訓練職類代碼

            If Convert.IsDBNull(objrow("DefGovCost")) Then '政府負擔費用
                writer.WriteElementString("DefGovCost", "")
            Else
                writer.WriteElementString("DefGovCost", objrow("DefGovCost"))
            End If

            If Convert.IsDBNull(objrow("DefStdCost")) Then '學員負擔費用
                writer.WriteElementString("DefStdCost", "")
            Else
                writer.WriteElementString("DefStdCost", objrow("DefStdCost"))
            End If

            If Convert.IsDBNull(objrow("Tperiod")) Then '訓練時段代碼
                writer.WriteElementString("Tperiod", "")
            Else
                writer.WriteElementString("Tperiod", objrow("Tperiod"))
            End If

            writer.WriteElementString("TPropertyID", objrow("TPropertyID")) '訓練性質

            writer.WriteElementString("CapDegree", objrow("CapDegree")) '學歷資格

            writer.WriteElementString("CapAge1", IIf(IsDBNull(objrow("CapAge1")), "", objrow("CapAge1"))) '年齡資格上限

            writer.WriteElementString("CapAge2", IIf(IsDBNull(objrow("CapAge2")), "", objrow("CapAge2"))) '年齡資格下限

            If Convert.IsDBNull(objrow("CapSex")) Then
                writer.WriteElementString("CapSex", "")
            Else
                writer.WriteElementString("CapSex", objrow("CapSex")) '性別資格
            End If
            If Convert.IsDBNull(objrow("CapMilitary")) Then
                'writer.WriteElementString("CapMilitary", "") '未填寫 '兵役資格
                writer.WriteElementString("CapMilitary", "00") '不限 '兵役資格
            Else
                writer.WriteElementString("CapMilitary", objrow("CapMilitary")) '兵役資格
            End If
            If Convert.IsDBNull(objrow("CapOther1")) Then '其他資格一
                writer.WriteElementString("CapOther1", "")
            Else
                writer.WriteElementString("CapOther1", objrow("CapOther1"))
            End If

            If Convert.IsDBNull(objrow("CapOther2")) Then '其他資格二
                writer.WriteElementString("CapOther2", "")
            Else
                writer.WriteElementString("CapOther2", objrow("CapOther2"))
            End If

            If Convert.IsDBNull(objrow("CapOther3")) Then '其他資格三
                writer.WriteElementString("CapOther3", "")
            Else
                writer.WriteElementString("CapOther3", objrow("CapOther3"))
            End If

            If Convert.IsDBNull(objrow("Content")) Then '課程內容
                writer.WriteElementString("Content", "")
            Else
                writer.WriteElementString("Content", objrow("Content"))
            End If

            writer.WriteElementString("Purpose", objrow("Purpose")) '課程目標

            writer.WriteElementString("NotOpen", objrow("NotOpen")) '不開班

            writer.WriteElementString("SeqNo", objrow("SeqNo")) '序號

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Sub

    '開班階段檔
    Public Shared Sub Class_ClassLevel(ByVal strXmlFilePath As String)
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Class_ClassLevel"
        objtable = DbAccess.GetDataTable(sqlstr, conn)


        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("CCLID", objrow("CCLID"))

            writer.WriteElementString("OCID", objrow("OCID")) '班別序號

            If Convert.IsDBNull(objrow("LevelName")) Then '階段名稱
                writer.WriteElementString("LevelName", "")
            Else
                writer.WriteElementString("LevelName", objrow("LevelName"))
            End If

            writer.WriteElementString("LevelSDate", objrow("LevelSDate")) '階段起始日

            writer.WriteElementString("LevelEDate", objrow("LevelEDate")) '階段結束日

            writer.WriteElementString("LevelHour", objrow("LevelHour")) '階段時數

            If Convert.IsDBNull(objrow("Num")) Then '名額
                writer.WriteElementString("Num", "")
            Else
                writer.WriteElementString("Num", objrow("Num"))
            End If

            If Convert.IsDBNull(objrow("LSDate")) Then '階段報名起日
                writer.WriteElementString("LSDate", "")
            Else
                writer.WriteElementString("LSDate", objrow("LSDate"))
            End If

            If Convert.IsDBNull(objrow("LEDate")) Then '階段報名迄日
                writer.WriteElementString("LEDate", "")
            Else
                writer.WriteElementString("LEDate", objrow("LEDate"))
            End If

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Sub

    '計畫課程資料檔
    Public Shared Sub Class_Plan_Info(ByVal strXmlFilePath As String) 'As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Class_Plan_Info"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("CPID", objrow("CPID"))

            writer.WriteElementString("PlanYear", objrow("PlanYear")) '年度

            writer.WriteElementString("OrgName", objrow("OrgName")) '機構名稱

            writer.WriteElementString("ComIDNo", objrow("ComIDNo")) '廠商統編

            writer.WriteElementString("ClassCName", objrow("ClassCName")) '班級名稱

            If Convert.IsDBNull(objrow("Content")) Then '課程內容
                writer.WriteElementString("Content", "")
            Else
                writer.WriteElementString("Content", objrow("Content"))
            End If

            If Convert.IsDBNull(objrow("Purpose")) Then '課程目標
                writer.WriteElementString("Purpose", "")
            Else
                writer.WriteElementString("Purpose", objrow("Purpose"))
            End If

            writer.WriteElementString("TPropertyID", Convert.ToString(objrow("TPropertyID"))) '訓練性質

            writer.WriteElementString("TMID", Convert.ToString(objrow("TMID"))) '訓練職類代碼

            If Convert.IsDBNull(objrow("ExamDate")) Then '甄試日期
                writer.WriteElementString("ExamDate", "")
            Else
                writer.WriteElementString("ExamDate", objrow("ExamDate"))
            End If

            writer.WriteElementString("STDate", objrow("STDate")) '開訓日期

            writer.WriteElementString("FTDate", objrow("FTDate")) '結束日期

            writer.WriteElementString("SEnterDate", objrow("SEnterDate")) '報名起日期

            writer.WriteElementString("FEnterDate", objrow("FEnterDate")) '報名迄日期

            If Convert.IsDBNull(objrow("CheckInDate")) Then '報到日期
                writer.WriteElementString("CheckInDate", "")
            Else
                writer.WriteElementString("CheckInDate", objrow("CheckInDate"))
            End If

            If Convert.IsDBNull(objrow("TaddressZip")) Then '訓練地點Zip
                writer.WriteElementString("TaddressZip", "")
            Else
                writer.WriteElementString("TaddressZip", objrow("TaddressZip"))
            End If

            If Convert.IsDBNull(objrow("TAddress")) Then '訓練地點
                writer.WriteElementString("TAddress", "")
            Else
                writer.WriteElementString("TAddress", objrow("TAddress"))
            End If

            writer.WriteElementString("THours", objrow("THours")) '訓練時數

            writer.WriteElementString("TNum", objrow("TNum")) '訓練人數

            If Convert.IsDBNull(objrow("TDeadline")) Then '訓練期限代碼
                writer.WriteElementString("TDeadline", "")
            Else
                writer.WriteElementString("TDeadline", objrow("TDeadline"))
            End If

            writer.WriteElementString("Tperiod", objrow("Tperiod")) '訓練時段代碼

            If Convert.IsDBNull(objrow("PlanCause")) Then '訓練目標
                writer.WriteElementString("PlanCause", "")
            Else
                writer.WriteElementString("PlanCause", objrow("PlanCause"))
            End If

            writer.WriteElementString("CapDegree", objrow("CapDegree")) '學歷資格

            writer.WriteElementString("CapAge1", IIf(IsDBNull(objrow("CapAge1")), "", objrow("CapAge1"))) '年齡資格上限

            writer.WriteElementString("CapAge2", IIf(IsDBNull(objrow("CapAge2")), "", objrow("CapAge2")))  '年齡資格下限

            If Convert.IsDBNull(objrow("CapSex")) Then '性別資格
                writer.WriteElementString("CapSex", "")
            Else
                writer.WriteElementString("CapSex", objrow("CapSex"))
            End If

            If Convert.IsDBNull(objrow("CapMilitary")) Then '兵役資格
                'writer.WriteElementString("CapMilitary", "") '未填寫 '兵役資格
                writer.WriteElementString("CapMilitary", "00") '不限 '兵役資格
            Else
                writer.WriteElementString("CapMilitary", objrow("CapMilitary"))
            End If

            If Convert.IsDBNull(objrow("TMScience")) Then '學科
                writer.WriteElementString("TMScience", "")
            Else
                writer.WriteElementString("TMScience", objrow("TMScience"))
            End If

            If Convert.IsDBNull(objrow("TMTech")) Then '術科
                writer.WriteElementString("TMTech", "")
            Else
                writer.WriteElementString("TMTech", objrow("TMTech"))
            End If

            If Convert.IsDBNull(objrow("GenSciHours")) Then '課程編配-學科一般時數
                writer.WriteElementString("GenSciHours", "")
            Else
                writer.WriteElementString("GenSciHours", objrow("GenSciHours"))
            End If

            If Convert.IsDBNull(objrow("ProSciHours")) Then '課程編配-學科專業時數
                writer.WriteElementString("ProSciHours", "")
            Else
                writer.WriteElementString("ProSciHours", objrow("ProSciHours"))
            End If

            If Convert.IsDBNull(objrow("ProTechHours")) Then '課程編配-術科時數
                writer.WriteElementString("ProTechHours", "")
            Else
                writer.WriteElementString("ProTechHours", objrow("ProTechHours"))
            End If

            If Convert.IsDBNull(objrow("OtherHours")) Then '課程編配-其他時數
                writer.WriteElementString("OtherHours", "")
            Else
                writer.WriteElementString("OtherHours", objrow("OtherHours"))
            End If

            If Convert.IsDBNull(objrow("TotalHours")) Then '課程編配-總時數
                writer.WriteElementString("TotalHours", "")
            Else
                writer.WriteElementString("TotalHours", objrow("TotalHours"))
            End If

            If Convert.IsDBNull(objrow("DefMainCost")) Then '經費來源-局
                writer.WriteElementString("DefMainCost", "")
            Else
                writer.WriteElementString("DefMainCost", objrow("DefMainCost"))
            End If

            If Convert.IsDBNull(objrow("DefUnitCost")) Then '經費來源-單位
                writer.WriteElementString("DefUnitCost", "")
            Else
                writer.WriteElementString("DefUnitCost", objrow("DefUnitCost"))
            End If

            If Convert.IsDBNull(objrow("DefStdCost")) Then '經費來源-學員
                writer.WriteElementString("DefStdCost", "")
            Else
                writer.WriteElementString("DefStdCost", objrow("DefStdCost"))
            End If

            writer.WriteElementString("AppliedDate", objrow("AppliedDate")) '申請日期

            If Convert.IsDBNull(objrow("CapOther1")) Then '其他資格一
                writer.WriteElementString("CapOther1", "")
            Else
                writer.WriteElementString("CapOther1", objrow("CapOther1"))
            End If

            If Convert.IsDBNull(objrow("CapOther2")) Then '其他資格二
                writer.WriteElementString("CapOther2", "")
            Else
                writer.WriteElementString("CapOther2", objrow("CapOther2"))
            End If

            If Convert.IsDBNull(objrow("CapOther3")) Then '其他資格三
                writer.WriteElementString("CapOther3", "")
            Else
                writer.WriteElementString("CapOther3", objrow("CapOther3"))
            End If

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Sub

    '計畫代碼
    Public Shared Function ID_Plan(ByVal strXmlFilePath As String) As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From ID_Plan"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("PlanID", objrow("PlanID")) '計畫代碼

            writer.WriteElementString("Years", objrow("Years")) '西元年度

            writer.WriteElementString("DistID", objrow("DistID")) '轄區中心代碼

            writer.WriteElementString("TPlanID", objrow("TPlanID")) '訓練計畫代碼

            writer.WriteElementString("Seq", objrow("Seq")) '序號

            If Convert.IsDBNull(objrow("Sponsor")) Then '主辦單位
                writer.WriteElementString("Sponsor", "")
            Else
                writer.WriteElementString("Sponsor", objrow("Sponsor"))
            End If

            If Convert.IsDBNull(objrow("Cosponsor")) Then '協辦單位
                writer.WriteElementString("Cosponsor", "")
            Else
                writer.WriteElementString("Cosponsor", objrow("Cosponsor"))
            End If

            writer.WriteElementString("SDate", objrow("SDate")) '時效起日

            writer.WriteElementString("EDate", objrow("EDate")) '時效迄日

            If Convert.IsDBNull(objrow("PlanKind")) Then '計畫種類
                writer.WriteElementString("PlanKind", "")
            Else
                writer.WriteElementString("PlanKind", objrow("PlanKind"))
            End If

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Function

    '計畫機構
    Public Shared Function Org_OrgInfo(ByVal strXmlFilePath As String) As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Org_OrgInfo"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("OrgID", objrow("OrgID")) '機構代碼

            writer.WriteElementString("OrgKind", objrow("OrgKind")) '機構別

            writer.WriteElementString("OrgName", objrow("OrgName")) '機構名稱

            writer.WriteElementString("ComIDNO", objrow("ComIDNO")) '統編

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Function

    '機構計畫資料
    Public Shared Function Org_OrgPlanInfo(ByVal strXmlFilePath As String) As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        sqlstr = "Select * From Org_OrgPlanInfo"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow As DataRow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("RSID", objrow("RSID"))

            If Convert.IsDBNull(objrow("OrgPName")) Then '分支單位名稱
                writer.WriteElementString("OrgPName", "")
            Else
                writer.WriteElementString("OrgPName", objrow("OrgPName"))
            End If

            writer.WriteElementString("ZipCode", objrow("ZipCode")) '郵遞區號

            writer.WriteElementString("Address", objrow("Address")) '公司地址

            If Convert.IsDBNull(objrow("Phone")) Then '聯絡人電話
                writer.WriteElementString("Phone", "")
            Else
                writer.WriteElementString("Phone", objrow("Phone"))
            End If

            If Convert.IsDBNull(objrow("MasterName")) Then '負責人姓名
                writer.WriteElementString("MasterName", "")
            Else
                writer.WriteElementString("MasterName", objrow("MasterName"))
            End If

            writer.WriteElementString("ContactName", objrow("ContactName")) '聯絡人姓名

            If Convert.IsDBNull(objrow("ContactEmail")) Then '聯絡人email
                writer.WriteElementString("ContactEmail", "")
            Else
                writer.WriteElementString("ContactEmail", objrow("ContactEmail"))
            End If

            If Convert.IsDBNull(objrow("ContactCellPhone")) Then '聯絡人行動電話
                writer.WriteElementString("ContactCellPhone", "")
            Else
                writer.WriteElementString("ContactCellPhone", objrow("ContactCellPhone"))
            End If

            If Convert.IsDBNull(objrow("TrainCap")) Then '訓練容量
                writer.WriteElementString("TrainCap", "")
            Else
                writer.WriteElementString("TrainCap", objrow("TrainCap"))
            End If

            If Convert.IsDBNull(objrow("ProTrainKind")) Then '專長訓練職類
                writer.WriteElementString("ProTrainKind", "")
            Else
                writer.WriteElementString("ProTrainKind", objrow("ProTrainKind"))
            End If

            If Convert.IsDBNull(objrow("ComSumm")) Then '機構簡介
                writer.WriteElementString("ComSumm", "")
            Else
                writer.WriteElementString("ComSumm", objrow("ComSumm"))
            End If

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Function

    '年度訓練機構
    Public Shared Function Org_PlanYear(ByVal strXmlFilePath As String) As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        Dim objrow As DataRow = Nothing
        sqlstr = "Select * From Org_PlanYear"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("PlanYear", objrow("PlanYear")) '計畫年度

            writer.WriteElementString("OrgID", objrow("OrgID")) '機構代碼

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Function

    '計畫主檔
    Public Shared Function Plan_PlanInfo(ByVal strXmlFilePath As String) As String
        Dim conn As OracleConnection
        Dim objtable As DataTable
        'Dim objrow, objkey, objsql, dr As DataRow
        'Dim SqlStr, Sql As String
        'Dim objadapter As OracleDataAdapter
        Dim writer As XmlTextWriter

        conn = DbAccess.GetConnection2()

        '讀取資料
        Dim sqlstr As String = ""
        Dim objrow As DataRow = Nothing
        sqlstr = "Select * From Plan_PlanInfo where ModifyDate>='" & Common.FormatDate(Now()) & "'"
        objtable = DbAccess.GetDataTable(sqlstr, conn)

        '產生要輸出的XML檔案
        writer = New XmlTextWriter(strXmlFilePath, System.Text.Encoding.GetEncoding("UTF-8"))

        '輸出XML宣告部份
        writer.WriteStartDocument()

        '輸出「SGTPInfo」開始元素
        writer.WriteStartElement("SGTPInfo")

        For Each objrow In objtable.Select(Nothing, Nothing, DataViewRowState.CurrentRows)

            '輸出「Course」開始元素
            writer.WriteStartElement("Course")

            writer.WriteElementString("PlanSerial", objrow("PlanSerial")) '流水號

            writer.WriteElementString("PlanID", objrow("PlanID")) '計畫代碼

            writer.WriteElementString("ComIDNO", objrow("ComIDNO")) '統編

            writer.WriteElementString("SeqNO", objrow("SeqNO")) '序號

            If Convert.IsDBNull(objrow("PlanCause")) Then '計畫緣由
                writer.WriteElementString("PlanCause", "")
            Else
                writer.WriteElementString("PlanCause", objrow("PlanCause"))
            End If

            writer.WriteElementString("PurScience", objrow("PurScience")) '目標-學科

            writer.WriteElementString("PurTech", objrow("PurTech")) '目標-術科

            writer.WriteElementString("ModifyDate", objrow("ModifyDate")) '異動時間

            If Convert.IsDBNull(objrow("GetTrain1")) Then '錄訓方式
                writer.WriteElementString("GetTrain1", "")
            Else
                writer.WriteElementString("GetTrain1", objrow("GetTrain1"))
            End If

            If Convert.IsDBNull(objrow("GetTrain2")) Then '自行報名
                writer.WriteElementString("GetTrain2", "")
            Else
                writer.WriteElementString("GetTrain2", objrow("GetTrain2"))
            End If

            If Convert.IsDBNull(objrow("GetTrain3")) Then '甄試方式
                writer.WriteElementString("GetTrain3", "")
            Else
                writer.WriteElementString("GetTrain3", objrow("GetTrain3"))
            End If

            If Convert.IsDBNull(objrow("GetTrain3Other")) Then '甄試方式_說明
                writer.WriteElementString("GetTrain3Other", "")
            Else
                writer.WriteElementString("GetTrain3Other", objrow("GetTrain3Other"))
            End If

            If Convert.IsDBNull(objrow("GetTrain4")) Then '錄訓方式-其他
                writer.WriteElementString("GetTrain4", "")
            Else
                writer.WriteElementString("GetTrain4", objrow("GetTrain4"))
            End If

            If Convert.IsDBNull(objrow("GetTrain4Other")) Then '錄訓方式_其他說明
                writer.WriteElementString("GetTrain4Other", "")
            Else
                writer.WriteElementString("GetTrain4Other", objrow("GetTrain4Other"))
            End If

            '輸出「Course」結尾元素
            writer.WriteEndElement()
        Next

        '輸出「SGTPInfo」結尾元素
        writer.WriteEndElement()

        '輸出文件結尾
        writer.WriteEndDocument()

        '關閉並釋放XmlTextWriter物件
        writer.Close()
    End Function

End Class
