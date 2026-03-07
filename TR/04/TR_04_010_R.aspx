<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_010_R.aspx.vb" Inherits="WDAIIP.TR_04_010_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印就業率</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        function PrintReport() {
            window.print();
            window.close();
        }
    </script>
</head>
<body onload="PrintReport();">
    <form id="form1" method="post" runat="server">
        <font face="標楷體">
            <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
                <tr>
                    <td align="center">
                        <font face="標楷體" size="6">就業追蹤成果統計表</font><br>
                        &nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center">

                        <table id="Table2" style="border-collapse: collapse" cellspacing="1" cellpadding="1" width="100%" border="1">
                            <tr>
                                <td width="100" id="TD1_1" runat="server">開訓期間：
                                </td>
                                <td id="TD1_2" runat="server">
                                    <asp:Label ID="STDate" runat="server"></asp:Label>
                                </td>
                                <td width="100" id="TD1_3" runat="server">結訓期間：
                                </td>
                                <td id="TD1_4" runat="server">
                                    <asp:Label ID="FTDate" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr id="TR2" runat="server">
                                <td id="TD2_1" runat="server">轄&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 區：
                                </td>
                                <td id="TD2_2" runat="server">
                                    <asp:Label ID="DistID" runat="server"></asp:Label>
                                </td>
                                <td id="TD2_3" runat="server">訓練計畫：
                                </td>
                                <td id="TD2_4" runat="server">
                                    <asp:Label ID="TPlanID" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr id="TR3" runat="server">
                                <td id="TD3_1" runat="server">訓練機構：
                                </td>
                                <td id="TD3_2" runat="server">
                                    <asp:Label ID="RIDValue" runat="server"></asp:Label>
                                </td>
                                <td id="TD3_3" runat="server">班級名稱：
                                </td>
                                <td id="TD3_4" runat="server">
                                    <asp:Label ID="OCIDValue" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>


                    </td>
                </tr>
                <tr>
                    <td align="center">

                        <asp:Table ID="ShowDataTable" runat="server" CellPadding="1" CellSpacing="0">
                        </asp:Table>

                    </td>
                </tr>
                <tr>
                    <td align="left">備註：提前就業人數：學員實際參訓時數達總訓練時數1/2以上，經分署專案核定免負擔退訓賠償費用<br>
                        者。
                    </td>
                </tr>
            </table>
        </font>
    </form>
</body>
</html>
