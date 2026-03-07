/*列印*/
function PrintPart(controlId, isPortrait, title) //isPortrait =false 代表?打
{
    var sdiv = document.getElementById(controlId);
    var strTemp = sdiv.innerHTML;
    var ifrdoc = window.open("../js/PrintPage.aspx ", "_new", "height=" + screen.availHeight + ",width=" + screen.availWidth + ",top=0,left=-400,toolbar=no,menubar=no,scrollbars=no,resizable=no,location=no,status=no");
    ifrdoc.document.open(); 
    //ifrdoc.document.write("<object id='factory' style='display:none' classid='clsid:1663ed61-23eb-11d2-b92f-008048fdd814' codebase='../ScriptX/smsx.cab#Version=7.0.0.8'></object>  ");
    ifrdoc.document.write("<span align='center'>" + title + "</span>");
    strTemp = strTemp.replace('#9bd1e6', '#000000');
    strTemp = strTemp.replace('#9bd1e6', '#000000');
    strTemp = strTemp.replace('#9bd1e6', '#000000');
    strTemp = strTemp.replace('#9bd1e6', '#000000');

    ifrdoc.document.write(strTemp);
    
    var fc = ifrdoc.document.getElementById('factory');

    fc.printing.portrait = isPortrait; //是否?向打印
    fc.printing.header = ""; //頁首，空白為不印頁首，也就不會佔空間
    fc.printing.footer = ""; //註腳，空白為不印註腳，也就不會佔空間
    fc.printing.leftMargin = 5; //左邊界
    fc.printing.topMargin = 10; //上邊界
    fc.printing.rightMargin = 5; //右邊界
    fc.printing.bottomMargin = 10; //下邊界
	factory.printing.printBackground = true;	//背景圖列印
    fc.printing.Print();
    ifrdoc.window.close();
}