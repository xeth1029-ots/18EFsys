<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TechID2.aspx.vb" Inherits="WDAIIP.TechID2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇老師</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link href="../css/css.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        //<!--
        function setStyle(object, styleText) {
            if (object.style.setAttribute) {
                object.style.setAttribute("cssText", styleText);
            }
            else {
                object.setAttribute("style", styleText);
            }
        }

        function setDiv(object) {
            var result = object; // 取得div元素
            result.innerHTML = "";
        }

        function create_techtable(itemName) {
            var TeachID = "0"
            var TeachName = "教師姓名"
            var DegreeName = "學歷"
            var MajorName = "專業領域"
            if (form1.TeachID.value != "") {
                TeachID += "," + form1.TeachID.value;
                TeachName += "," + form1.TeachName.value;
                DegreeName += "," + form1.DegreeName.value;
                MajorName += "," + form1.Major.value;
            }
            var stid = TeachID.split(",");
            var stname = TeachName.split(",");
            var sdname = DegreeName.split(",");
            var Maname = MajorName.split(",");
            var element0, element1, element2, element3, element3b, element4, element5;
            var element4a, element4b, element4c, element4d, element4e;
            var newSpana, newSpanb, newSpand, newSpane; //= document.createElement("span");
            element0 = opener.document.createDocumentFragment();
            element1 = opener.document.createElement("table");
            //element1.setAttribute("style",);
            setStyle(element1, "height:6px;width:100%;border-collapse:collapse;background-color:#FFFFFF;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
            //element1.setAttribute("class","font");
            element1.setAttribute("cellspacing", "0");
            element1.setAttribute("rules", "all");
            element1.setAttribute("DESIGNTIMEDRAGDROP", "347");
            element1.setAttribute("border", "1");
            element2 = opener.document.createElement("tbody");
            //alert(stid.length);
            for (i = 0; i < stid.length; i++) {
                //alert(i);
                if (i == 0) {
                    element3 = opener.document.createElement("tr");
                    setStyle(element3, "height:6px;width:100%;border-collapse:collapse;background-color:#CCCCFF;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
                    element4a = opener.document.createElement("td");
                    element4a.setAttribute("id", i);
                    element4b = opener.document.createElement("td");
                    element4b.setAttribute("id", i);
                    element4d = opener.document.createElement("td");
                    element4d.setAttribute("id", i);
                    element4e = opener.document.createElement("td");
                    element4e.setAttribute("id", i);
                    newSpana = opener.document.createElement("span");
                    newSpana.appendChild(opener.document.createTextNode("序號"));
                    element4a.appendChild(newSpana);
                    newSpanb = opener.document.createElement("span");
                    newSpanb.appendChild(opener.document.createTextNode(stname[i]));
                    element4b.appendChild(newSpanb);
                    newSpand = opener.document.createElement("span");
                    newSpand.appendChild(opener.document.createTextNode(sdname[i]));
                    element4d.appendChild(newSpand);
                    newSpane = opener.document.createElement("span");
                    newSpane.appendChild(opener.document.createTextNode(Maname[i]));
                    element4e.appendChild(newSpane);
                    element3.appendChild(element4a);
                    element3.appendChild(element4b);
                    element3.appendChild(element4d);
                    element3.appendChild(element4e);
                    element2.appendChild(element3);
                }
                else {
                    element3b = opener.document.createElement("tr");
                    setStyle(element3b, "height:6px;width:100%;border-collapse:collapse;background-color:#FFFFFF;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
                    element4a = opener.document.createElement("td");
                    element4a.setAttribute("id", i);
                    element4b = opener.document.createElement("td");
                    element4b.setAttribute("id", i);
                    element4d = opener.document.createElement("td");
                    element4d.setAttribute("id", i);
                    element4e = opener.document.createElement("td");
                    element4e.setAttribute("id", i);
                    newSpana = opener.document.createElement("span");
                    newSpana.appendChild(opener.document.createTextNode(i));
                    element4a.appendChild(newSpana);
                    newSpanb = opener.document.createElement("span");
                    newSpanb.appendChild(opener.document.createTextNode(stname[i]));
                    element4b.appendChild(newSpanb);
                    newSpand = opener.document.createElement("span");
                    newSpand.appendChild(opener.document.createTextNode(sdname[i]));
                    element4d.appendChild(newSpand);
                    newSpane = opener.document.createElement("span");
                    newSpane.appendChild(opener.document.createTextNode(Maname[i]));
                    element4e.appendChild(newSpane);
                    element3b.appendChild(element4a);
                    element3b.appendChild(element4b);
                    element3b.appendChild(element4d);
                    element3b.appendChild(element4e);
                    element2.appendChild(element3b);
                }
            }
            element1.appendChild(element2);
            element0.appendChild(element1);
            setDiv(opener.document.getElementById(getParamValue('Table')));
            opener.document.getElementById(getParamValue('Table')).appendChild(element0);
        }

        function godo(element) {
            window.alert(element.getAttribute("id"));
        }

        function ReturnTechID(TechID, TechName, DegID, DegName, Major) {
            document.getElementById('TeachID').value = TechID;
            document.getElementById('TeachName').value = TechName;
            document.getElementById('DegreeID').value = DegID;
            document.getElementById('DegreeName').value = DegName;
            document.getElementById('Major').value = Major;
            opener.document.getElementById('CTName').value = document.getElementById('TeachID').value;
            opener.document.getElementById('TechName').value = document.getElementById('TeachName').value;
            opener.document.getElementById('DegreeID').value = document.getElementById('DegreeID').value;
            opener.document.getElementById('DegreeNAME').value = document.getElementById('DegreeName').value;
            opener.document.getElementById('Major').value = document.getElementById('Major').value;
            //opener.document.getElementById('CTName').value=TechID;
            //opener.document.getElementById('TechName').value=TechName;
            //opener.document.getElementById('DegreeID').value=DegID;
            //opener.document.getElementById('DegreeNAME').value=DegName;
            //opener.document.getElementById('Major').value=Major;
            create_techtable(getParamValue('Table'));
            window.close();
        }

        function ReturnTechID2() {
            //debugger;
            opener.document.getElementById('CTName').value = document.getElementById('TeachID').value;
            opener.document.getElementById('TechName').value = document.getElementById('TeachName').value;
            opener.document.getElementById('DegreeID').value = document.getElementById('DegreeID').value;
            opener.document.getElementById('DegreeNAME').value = document.getElementById('DegreeName').value;
            opener.document.getElementById('Major').value = document.getElementById('Major').value;
            create_techtable(getParamValue('Table'));
            window.close();
        }

        function OpenProMenu(num) {
            document.getElementById('State').value = num;
            if (num == 1) {
                document.getElementById('ProTR1').style.display = 'inline';
                document.getElementById('ProTR2').style.display = 'none';
            }
            else {
                document.getElementById('ProTR1').style.display = 'none';
                document.getElementById('ProTR2').style.display = 'inline';
            }
        }

        function SelectTechID(Flag, TechID, TechName, DegID, DegName, Major) {
            if (Flag) {
                if (document.getElementById('TeachID').value == '') {
                    document.getElementById('TeachID').value = TechID;
                    document.getElementById('TeachName').value = TechName;
                    document.getElementById('DegreeID').value = DegID;
                    document.getElementById('DegreeName').value = DegName;
                    document.getElementById('Major').value = Major;
                }
                else {
                    document.getElementById('TeachID').value += ',' + TechID;
                    document.getElementById('TeachName').value += ',' + TechName;
                    document.getElementById('DegreeID').value += ',' + DegID;
                    document.getElementById('DegreeName').value += ',' + DegName;
                    document.getElementById('Major').value += ',' + Major;
                }
            }
            else {
                if (document.getElementById('TeachID').value.indexOf(',' + TechID + ',') != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(',' + TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(',' + TechName, '')
                    document.getElementById('DegreeID').value = document.getElementById('DegreeID').value.replace(',' + DegID, '')
                    document.getElementById('DegreeName').value = document.getElementById('DegreeName').value.replace(',' + DegName, '')
                    document.getElementById('Major').value = document.getElementById('Major').value.replace(',' + Major, '')
                }
                else if (document.getElementById('TeachID').value.indexOf(',' + TechID) != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(',' + TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(',' + TechName, '')
                    document.getElementById('DegreeID').value = document.getElementById('DegreeID').value.replace(',' + DegID, '')
                    document.getElementById('DegreeName').value = document.getElementById('DegreeName').value.replace(',' + DegName, '')
                    document.getElementById('Major').value = document.getElementById('Major').value.replace(',' + Major, '')
                }
                else if (document.getElementById('TeachID').value.indexOf(TechID + ',') != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(TechID + ',', '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(TechName + ',', '')
                    document.getElementById('DegreeID').value = document.getElementById('DegreeID').value.replace(DegID + ',', '')
                    document.getElementById('DegreeName').value = document.getElementById('DegreeName').value.replace(DegName + ',', '')
                    document.getElementById('Major').value = document.getElementById('Major').value.replace(Major + ',', '')
                }
                else if (document.getElementById('TeachID').value.indexOf(TechID) != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(TechName, '')
                    document.getElementById('DegreeID').value = document.getElementById('DegreeID').value.replace(DegID, '')
                    document.getElementById('DegreeName').value = document.getElementById('DegreeName').value.replace(DegName, '')
                    document.getElementById('Major').value = document.getElementById('Major').value.replace(Major, '')
                }
            }
        }
        //-->
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <font face="新細明體">
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="300" border="0">
                <tr id="ProTR1" runat="server">
                    <td align="right">
                        <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td width="80">講師代碼： </td>
                                <td>
                                    <asp:TextBox ID="TeacherID" runat="server"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td width="80">講師姓名： </td>
                                <td>
                                    <asp:TextBox ID="TeachCName" runat="server"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td width="80">內外聘： </td>
                                <td>
                                    <asp:DropDownList ID="KindEngage1" runat="server">
                                        <asp:ListItem Value="%">全部</asp:ListItem>
                                        <asp:ListItem Value="1">內</asp:ListItem>
                                        <asp:ListItem Value="2">外</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2">
                                    <asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button><input onclick="ReturnTechID('', '', '', '')" type="button" value="清除"></td>
                            </tr>
                        </table>
                        <asp:HyperLink ID="Close" runat="server" ForeColor="Blue">關閉進階搜尋</asp:HyperLink>
                    </td>
                </tr>
                <tr id="ProTR2" runat="server">
                    <td align="right">
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td width="80">內外聘： </td>
                                <td>
                                    <asp:DropDownList ID="KindEngage" runat="server" AutoPostBack="True">
                                        <asp:ListItem Value="%">全部</asp:ListItem>
                                        <asp:ListItem Value="1">內</asp:ListItem>
                                        <asp:ListItem Value="2">外</asp:ListItem>
                                    </asp:DropDownList>
                                    <input onclick="ReturnTechID('', '', '', '', '')" type="button" value="清除" id="Button3" name="Button3" runat="server">
                                </td>
                            </tr>
                        </table>
                        <asp:HyperLink ID="Open" runat="server" ForeColor="Blue">進階搜尋</asp:HyperLink>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%">
                            <Columns>
                                <asp:TemplateColumn>
                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <input id="Radio1" type="radio" value="Radio1" runat="server"><input id="Checkbox1" type="checkbox" runat="server">
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="KindEngage" HeaderText="內外聘"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TeachCName" HeaderText="講師姓名"></asp:BoundColumn>
                                <asp:BoundColumn DataField="DegreeID" HeaderText="學歷代碼"></asp:BoundColumn>
                                <asp:BoundColumn DataField="DegreeName" HeaderText="學歷"></asp:BoundColumn>
                                <asp:BoundColumn Visible="False" DataField="Major" HeaderText="專業領域"></asp:BoundColumn>
                            </Columns>
                        </asp:DataGrid><input id="Button2" onclick="ReturnTechID2();" type="button" value="送出" name="Button2" runat="server">
                    </td>
                </tr>
            </table>
        </font>
        <input id="State" type="hidden" value="0" runat="server">
        <input id="TeachID" type="hidden" name="TeachID" runat="server">
        <input id="TeachName" type="hidden" name="TeachName" runat="server">
        <input id="DegreeID" type="hidden" name="DegreeID" runat="server">
        <input id="DegreeName" type="hidden" name="DegreeName" runat="server">
        <input id="Major" type="hidden" name="Major" runat="server">
        <asp:HiddenField ID="HidCTName" runat="server" />
    </form>
</body>
</html>