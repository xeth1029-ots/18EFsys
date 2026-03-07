Imports System.IO
Imports System.Management

Public Class WDAP

    Public Shared LOG As ILog = LogManager.GetLogger("WDAP") 'log4net

    Private ReadOnly ObjLock_SYS_HISTORY11 As New Object

    ''' <summary> 防止駭客攻擊(紀錄) 錯誤才記 - SYS_HISTORY1</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="oConn"></param>
    Public Sub SUtl_SaveLoginData1(ByRef MyPage As Page, ByRef oConn As SqlConnection)
        If oConn Is Nothing Then oConn = DbAccess.GetConnection()

        Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress v_IpAddress '
        '(查無ip或太短-無法判斷，放棄-防止駭客攻擊)
        If String.IsNullOrEmpty(v_IpAddress) OrElse v_IpAddress.Length < 4 Then Return
        '(資料庫異常，放棄-防止駭客攻擊)
        If oConn Is Nothing OrElse Not TIMS.OpenDbConn(oConn) Then Return 'Call OpenDbConn(oConn)

        Dim s_UserHostName As String = "NO-UserHostName"
        Try
            s_UserHostName = If(MyPage IsNot Nothing, MyPage.Request.UserHostName, HttpContext.Current.Request.UserHostName)
        Catch ex As Exception
            LOG.Error(ex.Message, ex)
        End Try

        SyncLock ObjLock_SYS_HISTORY11
            '當日查詢
            Dim s_sql As String = " SELECT 'X' FROM SYS_HISTORY1 WITH(NOLOCK) WHERE L_IP=@L_IP AND convert(date,L_CREATETIME)=convert(date,GETDATE())" '當日
            Using dt As New DataTable
                Using sCmd As New SqlCommand(s_sql, oConn)
                    With sCmd
                        .Parameters.Clear()
                        .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress
                        dt.Load(.ExecuteReader())
                    End With
                End Using

                If TIMS.dtHaveDATA(dt) Then
                    '非第1次
                    Dim u_sql As String = " UPDATE SYS_HISTORY1 SET L_COUNT=L_COUNT+1,L_MODIFYTIME=GETDATE() WHERE L_IP=@L_IP" 'PK
                    Using uCmd As New SqlCommand(u_sql, oConn)
                        With uCmd
                            .Parameters.Clear()
                            .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress 'MyPage.Request.UserHostAddress
                            .ExecuteNonQuery()
                        End With
                    End Using
                    Return
                End If
            End Using

            Try
                '刪除非當日的資訊(任何)-當日查無資料-第1次新增
                Dim d_sql As String = " DELETE SYS_HISTORY1 WHERE L_IP=@L_IP"
                Using dCmd As New SqlCommand(d_sql, oConn)
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress 'MyPage.Request.UserHostAddress
                        .ExecuteNonQuery()
                    End With
                End Using
            Catch ex As Exception
                LOG.Warn($"sUtl_SaveLoginData1: {ex.Message}", ex)
            End Try
            Dim iFlag As Boolean = True '新增正常
            Try
                '第1次新增
                Dim i_sql As String = " INSERT INTO SYS_HISTORY1(L_IP ,L_HOSTNAME ,L_CREATETIME ,L_COUNT) VALUES (@L_IP ,@L_HOSTNAME ,GETDATE() ,1)" '當日
                Using iCmd As New SqlCommand(i_sql, oConn)
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress 'MyPage.Request.UserHostAddress
                        .Parameters.Add("L_HOSTNAME", SqlDbType.VarChar).Value = s_UserHostName 'MyPage.Request.UserHostName
                        .ExecuteNonQuery()
                    End With
                End Using

            Catch ex As Exception
                '瞬間大量(新增時有重複)
                LOG.Warn($"sUtl_SaveLoginData1: {ex.Message}", ex)
                iFlag = False '新增異常改用UPDATE
            End Try
            If Not iFlag Then
                '非第1次
                Dim u_sql As String = " UPDATE SYS_HISTORY1 SET L_COUNT=L_COUNT+1,L_MODIFYTIME=GETDATE() WHERE L_IP=@L_IP"
                Using uCmd As New SqlCommand(u_sql, oConn)
                    With uCmd
                        .Parameters.Clear()
                        .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress 'MyPage.Request.UserHostAddress
                        .ExecuteNonQuery()
                    End With
                End Using

            End If
        End SyncLock

    End Sub

    ''' <summary>防止駭客攻擊(清理紀錄),刪除資訊(呼叫者)</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="oConn"></param>
    Public Sub SUTL_SYS_HISTORY1_RST(ByRef MyPage As Page, ByRef oConn As SqlConnection)
        Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress v_IpAddress '
        '(查無ip或太短-無法判斷，放棄-防止駭客攻擊)
        If String.IsNullOrEmpty(v_IpAddress) OrElse v_IpAddress.Length < 4 Then Return
        '(資料庫異常，放棄-防止駭客攻擊)
        If oConn Is Nothing OrElse Not TIMS.OpenDbConn(oConn) Then Return 'Call OpenDbConn(oConn)
        Try
            Dim d_sql As String = " DELETE SYS_HISTORY1 WHERE L_IP=@L_IP"
            Using dCmd As New SqlCommand(d_sql, oConn)
                With dCmd
                    .Parameters.Clear()
                    .Parameters.Add("L_IP", SqlDbType.VarChar).Value = v_IpAddress 'MyPage.Request.UserHostAddress
                    .ExecuteNonQuery()
                End With
            End Using
        Catch ex As Exception
            LOG.Warn($"SUTL_SYS_HISTORY1_RST: {ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>查詢檔案-最後異動日期</summary>
    ''' <returns></returns>
    Public Function FileLastModified(MyPage As Page, iPMS As Integer) As String
        Dim rst As String = ""
        Const dll_path1 As String = "C:\web\SYS.Net\WDAIIP.Net\Bin\WDAIIP.SYS.Net.dll"
        Const dll_web_path2 As String = "~\Bin\WDAIIP.SYS.Net.dll"
        'Dim filePath As String = "\\192.168.3.121\c$\web\SYS.Net\WDAIIP.Net\Bin\WDAIIP.SYS.Net.dll" 25.svn_up_ojtims_sysNet_1_rc
        Try
            Dim fg_test As Boolean = TIMS.sUtl_ChkTest()
            Dim filePath As String = If(fg_test, MyPage.Server.MapPath(dll_web_path2), dll_path1)
            Dim fileInfo As New FileInfo(filePath)
            If fileInfo.Exists Then
                Dim lastModified As DateTime = fileInfo.LastWriteTime
                rst &= If(iPMS = 2, $"檔案：{filePath}{vbCrLf}最後異動日期：{lastModified}", $"異動日期：{lastModified}")
            Else
                rst &= $"檔案不存在：{filePath}"
            End If
        Catch ex As Exception
            rst &= $",#發生錯誤：{ex.Message}"
        End Try
        Return rst
    End Function

    ''' <summary>查詢目前主機的作業系統 (OS) 版本</summary>
    ''' <returns></returns>
    Public Function GetEnvOsVersion() As String
        'GetEnvironmentOsVersion()  獲取 OperatingSystem 實例
        Dim os As OperatingSystem = Environment.OSVersion
        ' 從 OperatingSystem 實例中獲取 Version 物件
        Dim version As Version = os.Version

        Dim sb1 As New StringBuilder
        sb1.AppendLine($"💻 OS 平台: {os.Platform} ,OS Service Pack: {os.ServicePack}")
        sb1.AppendLine($"🔢 OS 版本號 (Major.Minor.Build.Revision): {version}")
        sb1.AppendLine($" - 主要版本號 (Major): {version.Major}")
        sb1.AppendLine($" - 次要版本號 (Minor): {version.Minor}")

        ' 警告: 在 Windows 10/11 上可能返回舊版號 (如 6.2.9200.0)
        ' VB.NET 中的 If 判斷式
        If version.Major >= 10 Then
            ' 使用 vbCrLf 進行換行，或者在新的 .NET 版本中使用 \n
            sb1.AppendLine($"{Environment.NewLine}⚠️ 注意: 由於相容性，此方法在 Win 10/11 上可能顯示舊的 Windows 版本號，{Environment.NewLine}  請改用 RuntimeInformation.OSDescription 獲取準確資訊。")
        End If
        Return sb1.ToString()
    End Function

    ''' <summary>查詢本機系統的記憶體狀態（包括總量、已使用與空閒空間）memory</summary>
    ''' <returns></returns>
    Public Function GetMemoryStatus() As String
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
        Dim sb1 As New StringBuilder
        For Each queryObj As ManagementObject In searcher.Get()
            ' 取得總實體記憶體 (KB 轉為 GB)
            Dim totalMemory As Double = Math.Round(CDbl(queryObj("TotalVisibleMemorySize")) / 1024 / 1024, 2)
            ' 取得目前可用實體記憶體 (KB 轉為 GB)
            Dim freeMemory As Double = Math.Round(CDbl(queryObj("FreePhysicalMemory")) / 1024 / 1024, 2)
            ' 計算已使用記憶體
            Dim usedMemory As Double = totalMemory - freeMemory

            sb1.AppendLine($"總記憶體: {totalMemory} GB")
            sb1.AppendLine($"已使用: {usedMemory} GB")
            sb1.AppendLine($"空閒中: {freeMemory} GB")
        Next
        Return sb1.ToString()
    End Function

End Class
