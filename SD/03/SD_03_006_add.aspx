<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_006_add.aspx.vb" Inherits="TIMS.SD_03_006_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>結訓學員資料維護</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script src="../../js/common.js"></script>
	<script>
		function ChangeSubsidy() {
			if (document.form1.SubsidyHidden.value == '1') {
				if (confirm('變更津貼類型將會將「職業訓練生活津貼申請」相關資料刪除，確定要變更?')) {
					document.form1.SubsidyHidden.value = '0';
				}
				else {
					document.form1.SubsidyID.selectedIndex = 3;
					document.form1.SubsidyHidden.value = '1';
				}
			}
		}
		function EnterChannelChange() {
			if (document.form1.EnterChannel.value == '4') {
				document.getElementById('TRNDTR').style.display = 'inline';
			}
			else {
				document.getElementById('TRNDTR').style.display = 'none';
			}
		}
		function TRNDModeChange() {
			for (var i = 0; i < document.form1.TRNDType.length; i++) {
				document.form1.TRNDType[i].checked = false;
			}
			if (document.form1.TRNDMode.selectedIndex != 0) {
				if (document.form1.TRNDMode.value == '1' || document.form1.TRNDMode.value == '3') {
					for (var i = 0; i < document.form1.TRNDType.length; i++) {
						document.form1.TRNDType[i].disabled = false;
					}
				}
				else {
					for (var i = 0; i < document.form1.TRNDType.length; i++) {
						document.form1.TRNDType[i].disabled = true;
					}
				}
			}
		}

		//改變國籍身分
		function ChangePassPort() {
			var cst_pt1 = 0;
			var cst_pt2 = 1;
			if (document.getElementsByName('ChinaOrNot').length > 2) {
				cst_pt1 = 1; //cst_pt
				cst_pt2 = 2;
			}
			var cst_pp1 = 0;
			var cst_pp2 = 1;
			if (document.getElementsByName('PPNO').length > 2) {
				cst_pp1 = 1; //cst_pp
				cst_pp2 = 2;
			}
			var cst_fs1 = 0;
			var cst_fs2 = 1;
			if (document.getElementsByName('ForeSex').length > 2) {
				cst_fs1 = 1; //cst_fs
				cst_fs2 = 2;
			}
			if (getRadioValue(document.form1.PassPortNO) == 1) {
				document.getElementById('ChinaOrNotTable').style.display = 'none';
				document.getElementById('PPNO').style.display = 'none';
				document.getElementsByName('ChinaOrNot')[cst_pt1].checked = false;
				document.getElementsByName('ChinaOrNot')[cst_pt2].checked = false;
				document.getElementById('Nationality').value = '';
				document.getElementsByName('PPNO')[cst_pp1].checked = false;
				document.getElementsByName('PPNO')[cst_pp2].checked = false;
				for (i = 1; i <= 5; i++) {
					document.getElementById('ForeTr' + i).style.display = 'none';
				}
				document.getElementById('ForeName').value = '';
				document.getElementById('ForeTitle').value = '';
				document.getElementsByName('ForeSex')[cst_fs1].checked = false;
				document.getElementsByName('ForeSex')[cst_fs2].checked = false;
				document.getElementById('ForeBirth').value = '';
				document.getElementById('ForeIDNO').value = '';
				document.getElementById('City6').value = '';
				document.getElementById('ForeZip').value = '';
				document.getElementById('ForeAddr').value = '';
			}
			else {
				document.getElementById('ChinaOrNotTable').style.display = 'inline';
				document.getElementById('PPNO').style.display = 'inline';
				for (i = 1; i <= 5; i++) {
					document.getElementById('ForeTr' + i).style.display = 'inline';
				}
			}
		}

		//變更銀行
		function ChangeBank() {
			document.getElementById('PortTR').style.display = 'none';
			document.getElementById('BankTR1').style.display = 'none';
			document.getElementById('BankTR2').style.display = 'none';
			document.getElementById('BankTR3').style.display = 'none';
			document.getElementById('PostNo_1').value = '';
			document.getElementById('PostNo_2').value = '';
			document.getElementById('AcctNo1_1').value = '';
			document.getElementById('AcctNo1_2').value = '';
			document.getElementById('AcctHeadNo').value = '';
			//document.getElementById('AcctExNo').value='';
			document.getElementById('AcctNo2').value = '';

			if (isChecked(document.getElementsByName('AcctMode'))) {
				switch (getRadioValue(document.getElementsByName('AcctMode'))) {
					case '0':
						document.getElementById('PortTR').style.display = 'inline';
						break;
					case '1':
						document.getElementById('BankTR1').style.display = 'inline';
						document.getElementById('BankTR2').style.display = 'inline';
						document.getElementById('BankTR3').style.display = 'inline';
						break;
				}
			}
		}

		function chkdata() {
			var msg = '';
			var Item = '';
			var Page = 0;

			if (document.form1.LevelNo.disabled == false)
				if (document.form1.LevelNo.selectedIndex == 0) { msg += '請選擇報名階段\n'; if (Item == '') Item = 'LevelNo'; Page = 1; }
			if (document.form1.Name.value == '') { msg += '請輸入姓名\n'; if (Item == '') Item = 'Name'; Page = 1; }
			if (document.form1.StudentID.value == '') { msg += '請輸入學號\n'; if (Item == '') Item = 'StudentID'; Page = 1; }
			if (document.form1.StudentID.value != '' && !isUnsignedInt(document.form1.StudentID.value)) { msg += '學號必須為數字\n'; if (Item == '') Item = 'StudentID'; Page = 1; }
			// 如果是產學訓就不擋英文姓名 緊急通知人 失業週數
			if (document.form1.TPlanID.value != '28') {
				if (document.form1.LName.value == '' || document.form1.FName.value == '') {
					msg += '請填寫英文姓名\n'; if (Item == '') Item = 'LName'; Page = 1;
				}
				else {
					if (!isEng(document.form1.LName.value)) { msg += 'LastName必須為英文字\n'; if (Item == '') Item = 'LName'; Page = 1; }
					if (!isEng(document.form1.FName.value)) { msg += 'FirstName必須為英文字\n'; if (Item == '') Item = 'FName'; Page = 1; }
				}

				if (document.form1.EmergencyContact.value == '') { msg += '請輸入緊急通知人\n'; if (Item == '') Item = 'EmergencyContact'; Page = 1; }
				if (document.form1.EmergencyPhone.value == '') { msg += '請輸入緊急通知人電話\n'; if (Item == '') Item = 'EmergencyPhone'; Page = 1; }
				if (document.form1.EmergencyRelation.value == '') { msg += '請輸入緊急通知人關系\n'; if (Item == '') Item = 'EmergencyRelation'; Page = 1; }
				if (document.form1.ZipCode3.value == '') { msg += '請輸入緊急聯絡人通訊地址(區域)\n'; if (Item == '') Item = 'City3'; Page = 1; }
				if (document.form1.EmergencyAddress.value == '') { msg += '請輸入緊急連絡人通訊地址\n'; if (Item == '') Item = 'EmergencyAddress'; Page = 1; }

				if (document.form1.JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; if (Item == '') Item = 'JoblessID'; Page = 1; }
				if (document.form1.RealJobless.value != '' && !isUnsignedInt(document.form1.RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }

			}
			//end

			for (i = 0, j = 0; i < document.form1.PassPortNO.length; i++) {
				if (!document.form1.PassPortNO[i].checked) j++;
			}
			if (!isChecked(document.form1.PassPortNO)) { msg = msg + '請選擇身分別!\n'; if (Item == '') Item = 'PassPortNO'; Page = 1; }
			else {
				if (document.form1.PassPortNO[1].checked) {
					if (!isChecked(document.form1.ChinaOrNot)) { msg = msg + '請選擇是否為大陸人士!\n'; if (Item == '') Item = 'ChinaOrNot'; Page = 1; }
					if (document.getElementById('Nationality').value == '') { msg = msg + '請輸入原屬國籍!\n'; if (Item == '') Item = 'Nationality'; Page = 1; }
					if (!isChecked(document.form1.PPNO)) { msg = msg + '請選擇護照或居留(工作)證號!\n'; if (Item == '') Item = 'PPNO'; Page = 1; }
				}
			}
			if (document.form1.IDNO.value == '') { msg += '請輸入身分證號碼\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
			else if (document.form1.PassPortNO[0].checked == true) {
				if (document.getElementById('RoleID').value != '99' || document.getElementById('Process') == 'edit') {
					var pattern = /^[A-Z][1-2]{1}\d{8}$/;
					if (!pattern.test(document.form1.IDNO.value)) { msg += '身分證號碼錯誤\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
				}
				else {
					if (!checkId(document.form1.IDNO.value)) { msg += '身分證號碼錯誤(如果有此身分證號碼，請聯絡系統管理者)\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
				}
			}

			if (!isChecked(document.form1.Sex)) {
				msg = msg + '請選擇性別!\n'; if (Item == '') Item = 'Sex'; Page = 1;
			}
			else {
				if (document.form1.PassPortNO[0].checked == true) {
					//if (document.form1.IDNO.value!='' && !checkId(document.form1.IDNO.value)) msg+='身分證號碼不正確\n';
					if (document.form1.IDNO.value.charAt(1) == 1 && getRadioValue(document.form1.Sex) == 'F') { msg += '性別與身分證號碼不符合\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
					else if (document.form1.IDNO.value.charAt(1) == 2 && getRadioValue(document.form1.Sex) == 'M') { msg += '性別與身分證號碼不符合\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
				}
			}
			if (document.form1.Birthday.value == '') { msg += '請輸入出生日期\n'; if (Item == '') Item = 'Birthday'; Page = 1; }
			// for (i = 0, j = 0; i < document.form1.MaritalStatus.length; i++) {
			//      if (!document.form1.MaritalStatus[i].checked) j++;
			// }
			//if (j==document.form1.MaritalStatus.length) msg=msg+'請選擇婚姻狀況!\n';
			if (document.form1.Birthday.value != '' && !checkDate(document.form1.Birthday.value)) { msg += '出生日期格式不正確\n'; if (Item == '') Item = 'Birthday'; Page = 1; }
			if (document.form1.EnterChannel.value == '4') {
				if (document.form1.TRNDMode.selectedIndex == 0) {
					msg += '請選擇推介種類\n'; if (Item == '') Item = 'TRNDMode'; Page = 1;
				}
				else {
					//if(document.form1.TRNDMode.value=='1' || document.form1.TRNDMode.value=='3'){
					if (document.form1.TRNDMode.value == '1') {
						if (!isChecked(document.form1.TRNDType)) { msg += '請選擇券別種類\n'; if (Item == '') Item = 'TRNDMode'; Page = 1; }
					}
				}
			}
			if (document.form1.OpenDate.value != '' && !checkDate(document.form1.OpenDate.value)) { msg += '開訓日期格式不正確\n'; if (Item == '') Item = 'OpenDate'; Page = 1; }
			if (document.form1.CloseDate.value != '' && !checkDate(document.form1.CloseDate.value)) { msg += '結訓日期格式不正確\n'; if (Item == '') Item = 'CloseDate'; Page = 1; }
			if (document.form1.EnterDate.value != '' && !checkDate(document.form1.EnterDate.value)) { msg += '報到日期格式不正確\n'; if (Item == '') Item = 'EnterDate'; Page = 1; }
			if (document.form1.DegreeID.selectedIndex == 0) { msg += '請選擇最高學歷\n'; if (Item == '') Item = 'DegreeID'; Page = 1; }
			if (document.form1.School.value == '') { msg += '請輸入學校\n'; if (Item == '') Item = 'School'; Page = 1; }
			if (document.form1.Department.value == '') { msg += '請輸入科系\n'; if (Item == '') Item = 'Department'; Page = 1; }
			if (document.form1.GraduateStatus.selectedIndex == 0) { msg += '請選擇畢業狀況\n'; if (Item == '') Item = 'GraduateStatus'; Page = 1; }
			if (document.form1.MilitaryID.selectedIndex == 0) { msg += '請選擇兵役狀況\n'; if (Item == '') Item = 'MilitaryID'; Page = 1; }
			if (document.form1.MilitaryID.selectedIndex == 4) {
				if (document.form1.ServiceID.value == '') { msg += '請輸入姓軍種\n'; if (Item == '') Item = 'ServiceID'; Page = 1; }
				if (document.form1.MilitaryRank.value == '') { msg += '請輸入階級\n'; if (Item == '') Item = 'MilitaryRank'; Page = 1; }
				if (document.form1.ServiceOrg.value == '') { msg += '請輸入服務單位名稱\n'; if (Item == '') Item = 'ServiceOrg'; Page = 1; }
				if (document.form1.ServicePhone.value == '') { msg += '請輸入服務單位電話\n'; if (Item == '') Item = 'ServicePhone'; Page = 1; }
				if (document.form1.SServiceDate.value == '') { msg += '請輸入起始服役日期\n'; if (Item == '') Item = 'SServiceDate'; Page = 1; }
				if (document.form1.FServiceDate.value == '') { msg += '請輸入終至服役日期\n'; if (Item == '') Item = 'FServiceDate'; Page = 1; }
			}
			if (document.form1.PhoneD.value == '') { msg += '請輸入聯絡電話(日)\n'; if (Item == '') Item = 'PhoneD'; Page = 1; }
			if (document.form1.ZipCode1.value == '') { msg += '請輸入通訊地址(區域)\n'; if (Item == '') Item = 'City1'; Page = 1; }
			if (document.form1.Address.value == '') { msg += '請輸入通訊地址\n'; if (Item == '') Item = 'Address'; Page = 1; }
			if (document.form1.Email.value != '' && !checkEmail(document.form1.Email.value)) { msg += '請輸入正確的E-mail格式\n'; if (Item == '') Item = 'Email'; Page = 1; }
			if (document.form1.SubsidyID.selectedvalue == '') { msg += '請選擇申請津貼類別\n'; if (Item == '') Item = 'SubsidyID'; Page = 1; }
			var Identity = getCheckBoxListValue('IdentityID');
			var j = 0;
			var Identity = getCheckBoxListValue('IdentityID');
			if (document.form1.MIdentityID.selectedIndex == 0) { msg += '請選擇主要參訓身分別\n'; if (Item == '') Item = 'MIdentityID'; Page = 1; }
			else if (Identity.charAt(document.form1.MIdentityID.selectedIndex - 1) != '1') { msg += '主要參訓身分別必須為下列選的身分別之一\n'; if (Item == '') Item = 'Name'; Page = 1; }
			if (parseInt(Identity) == 0) {
				msg += '請選擇參訓身分別\n';
			}
			else {
				for (var i = 0; i < Identity.length; i++) {
					if (Identity.charAt(i) == '1') j++;
				}
				if (j > 3) msg += '參訓身分別最多只能選擇三項\n';
			}
			if (document.form1.MIdentityID.value == '05') {
				if (document.form1.NativeID.selectedIndex == 0) { msg += '請選擇民族別\n'; if (Item == '') Item = 'NativeID'; Page = 1; }
			}
			if (document.form1.HandTypeID.disabled == false) {
				if (document.form1.HandTypeID.selectedIndex == 0) { msg += '請選擇障礙類別\n'; if (Item == '') Item = 'HandTypeID'; Page = 1; }
				if (document.form1.HandLevelID.selectedIndex == 0) { msg += '請選擇障礙等級\n'; if (Item == '') Item = 'HandLevelID'; Page = 1; }
			}
			if (document.form1.RejectTDate1.value != '' && !checkDate(document.form1.RejectTDate1.value)) { msg += '離訓日期格式不正確\n'; if (Item == '') Item = 'RejectTDate1'; Page = 1; }
			if (document.form1.RejectTDate2.value != '' && !checkDate(document.form1.RejectTDate2.value)) { msg += '退訓日期格式不正確\n'; if (Item == '') Item = 'RejectTDate2'; Page = 1; }


			if (document.getElementById('ForeIDNO').value != '' && !checkId(document.getElementById('ForeIDNO').value)) { msg += '國內聯絡人身分證號碼不正確\n'; if (Item == '') Item = 'ForeIDNO'; Page = 1; }
			if (document.form1.SOfficeYM1.value != '' && !checkDate(document.form1.SOfficeYM1.value)) { msg += '受訓前任職1起始月日期格式不正確\n'; if (Item == '') Item = 'SOfficeYM1'; Page = 1; }
			if (document.form1.SOfficeYM2.value != '' && !checkDate(document.form1.SOfficeYM2.value)) { msg += '受訓前任職2起始月日期格式不正確\n'; if (Item == '') Item = 'SOfficeYM2'; Page = 1; }
			if (document.form1.FOfficeYM1.value != '' && !checkDate(document.form1.FOfficeYM1.value)) { msg += '受訓前任職1終至月日期格式不正確\n'; if (Item == '') Item = 'FOfficeYM1'; Page = 1; }
			if (document.form1.SOfficeYM2.value != '' && !checkDate(document.form1.SOfficeYM2.value)) { msg += '受訓前任職2起始月日期格式不正確\n'; if (Item == '') Item = 'SOfficeYM2'; Page = 1; }
			if (document.form1.FOfficeYM2.value != '' && !checkDate(document.form1.FOfficeYM2.value)) { msg += '受訓前任職2終至月日期格式不正確\n'; if (Item == '') Item = 'FOfficeYM2'; Page = 1; }
			if (document.form1.PriorWorkPay.value != '' && !isUnsignedInt(document.form1.PriorWorkPay.value)) { msg += '受訓前薪資必須為數字\n'; if (Item == '') Item = 'PriorWorkPay'; Page = 1; }

			if (document.form1.ShowDetail.selectedIndex == 0) { msg += '請選擇是否提供基本資料供查詢\n'; if (Item == '') Item = 'ShowDetail'; Page = 1; }
			if (document.form1.BudID) {
				if (!isChecked(document.form1.BudID)) { msg += '請選擇預算別\n'; if (Item == '') Item = 'BudID'; Page = 1; }
			}
			if (document.getElementById('PMode') && document.form1.TPlanID.value == '12') {
				if (!isChecked(document.form1.PMode)) msg += '請選擇公費/自費\n'
			}
			if (!isChecked(document.form1.IsAgree)) { msg += '請選擇是否同意將個人資料提供 勞動部勞動力發展署 暨所屬機關運用\n'; if (Item == '') Item = 'IsAgree'; Page = 1; }
			if (document.getElementById('ActNo')) {
				if (document.getElementById('ActNo').value == '') { msg += '請輸入保險證號\n'; if (Item == '') Item = 'ActNo'; Page = 1; }
			}
			if (document.getElementById('TPlanID').value == '15') {
				var JoinUnit = getCheckBoxListValue('RelClass_Unit');
				if (parseInt(JoinUnit) == 0) {
					msg += '請勾選學習單元\n'
				}
				else {
					if (JoinUnit.charAt(0) == '1') {
						if (document.getElementById('Unit1Hour').value == '') { msg += '請輸入第一單元的實際時數\n'; if (Item == '') Item = 'Unit1Hour'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit1Hour').value)) { msg += '第一單元的實際時數必須為數字\n'; if (Item == '') Item = 'Unit1Hour'; Page = 1; }
						//nick
						if (document.getElementById('Unit1Score').value == '') { msg += '請輸入第一單元的實際分數\n'; if (Item == '') Item = 'Unit1Score'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit1Score').value)) { msg += '第一單元的實際分數必須為數字\n'; if (Item == '') Item = 'Unit1Score'; Page = 1; }
					}
					if (JoinUnit.charAt(1) == '1') {
						if (document.getElementById('Unit2Hour').value == '') { msg += '請輸入第二單元的實際時數\n'; if (Item == '') Item = 'Unit2Hour'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit2Hour').value)) { msg += '第二單元的實際時數必須為數字\n'; if (Item == '') Item = 'Unit2Hour'; Page = 1; }
						//nick
						if (document.getElementById('Unit2Score').value == '') { msg += '請輸入第二單元的實際分數\n'; if (Item == '') Item = 'Unit2Score'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit2Score').value)) { msg += '第二單元的實際分數必須為數字\n'; if (Item == '') Item = 'Unit2Score'; Page = 1; }

					}
					if (JoinUnit.charAt(2) == '1') {
						if (document.getElementById('Unit3Hour').value == '') { msg += '請輸入第三單元的實際時數\n'; if (Item == '') Item = 'Unit3Hour'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit3Hour').value)) { msg += '第三單元的實際時數必須為數字\n'; if (Item == '') Item = 'Unit3Hour'; Page = 1; }
						//nick
						if (document.getElementById('Unit3Score').value == '') { msg += '請輸入第三單元的實際分數\n'; if (Item == '') Item = 'Unit3Score'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit3Score').value)) { msg += '第三單元的實際分數必須為數字\n'; if (Item == '') Item = 'Unit3Score'; Page = 1; }

					}
					if (JoinUnit.charAt(3) == '1') {
						if (document.getElementById('Unit4Hour').value == '') { msg += '請輸入第四單元的實際時數\n'; if (Item == '') Item = 'Unit4Hour'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit4Hour').value)) { msg += '第四單元的實際時數必須為數字\n'; if (Item == '') Item = 'Unit4Hour'; Page = 1; }
						//nick
						if (document.getElementById('Unit4Score').value == '') { msg += '請輸入第四單元的實際分數\n'; if (Item == '') Item = 'Unit4Score'; Page = 1; }
						else if (!isUnsignedInt(document.getElementById('Unit4Score').value)) { msg += '第四單元的實際分數必須為數字\n'; if (Item == '') Item = 'Unit4Score'; Page = 1; }

					}
					//以上,加入判斷輸入分數 by nick 060316						

				}
			}

			//企訓專用
			if (document.getElementById('BackTable')) {
				if (!isChecked(document.getElementsByName('AcctMode'))) {
					msg += '請輸入郵政或銀行帳號\n'; if (Item == '') { Item = 'AcctMode'; Page = 2; }
				}
				else {
					if (getRadioValue(document.getElementsByName('AcctMode')) == '0') {
						if (document.getElementById('PostNo_1').value == '' || document.getElementById('PostNo_2').value == '') { msg += '請輸入局號\n'; if (Item == '') { Item = 'PostNo_1'; Page = 2; } }
						if (document.getElementById('AcctNo1_1').value == '' || document.getElementById('AcctNo1_2').value == '') { msg += '請輸入帳號\n'; if (Item == '') { Item = 'AcctNo1_1'; Page = 2; } }
					}
					else if (getRadioValue(document.getElementsByName('AcctMode')) == '1') {
						if (document.getElementById('BankName').value == '') { msg += '請輸入銀行名稱\n'; if (Item == '') { Item = 'BankName'; Page = 2; } }
						//	if(document.getElementById('ExBankName').value=='') {msg+='請輸入分行名稱\n';if(Item=='') {Item='ExBankName';Page=2;}}
						if (document.getElementById('AcctHeadNo').value == '') { msg += '請輸入總代號\n'; if (Item == '') { Item = 'AcctHeadNo'; Page = 2; } }
						//	if(document.getElementById('AcctExNo').value=='') {msg+='請輸入分支代號\n';if(Item=='') {Item='AcctExNo';Page=2;}}
						if (document.getElementById('AcctNo2').value == '') { msg += '請輸入帳號\n'; if (Item == '') { Item = 'AcctNo2'; Page = 2; } }
					}
				}
				if (document.getElementById('FirDate').value != '' && !checkDate(document.getElementById('FirDate').value)) { msg += '第一次投保日期不是正確的日期格式\n'; if (Item == '') { Item = 'FirDate'; Page = 2; } }
				if (document.getElementById('Tel').value == '') { msg += '請輸入服務單位公司電話\n'; if (Item == '') { Item = 'Tel'; Page = 2; } }
				if (document.getElementById('Zip').value == '') { msg += '請輸入服務單位公司地址[地區]\n'; if (Item == '') { Item = 'City5'; Page = 2; } }
				if (document.getElementById('Addr').value == '') { msg += '請輸入服務單位公司地址\n'; if (Item == '') { Item = 'Addr'; Page = 2; } }
				if (document.getElementById('SDate').value != '' && !checkDate(document.getElementById('SDate').value)) { msg += '個人到任目前任職公司起日不是正確的日期格式\n'; if (Item == '') { Item = 'SDate'; Page = 2; } }
				if (document.getElementById('SJDate').value != '' && !checkDate(document.getElementById('SJDate').value)) { msg += '個人到任目前職務起日不是正確的日期格式\n'; if (Item == '') { Item = 'SJDate'; Page = 2; } }
				if (document.getElementById('SPDate').value != '' && !checkDate(document.getElementById('SPDate').value)) { msg += '最近升遷日期不是正確的日期格式\n'; if (Item == '') { Item = 'SPDate'; Page = 2; } }

				if (!isChecked(document.getElementsByName('Q1'))) { msg += '請選擇是否由公司推薦參訓\n'; if (Item == '') { Item = 'Q1'; Page = 2; } }
				if (parseInt(getCheckBoxListValue('Q2')) == 0) { msg += '請選擇參訊動機\n'; if (Item == '') { Page = 2; } }
				if (document.getElementById('Q4').selectedIndex == 0) { msg += '請選擇服務單位行業別\n'; if (Item == '') { Item = 'Q4'; Page = 2; } }
				if (document.getElementById('Q61').value != '' && !isUnsignedInt(document.getElementById('Q61').value)) { msg += '個人工作年資必須為數字\n'; if (Item == '') { Item = 'Q61'; Page = 2; } }
				if (document.getElementById('Q62').value != '' && !isUnsignedInt(document.getElementById('Q62').value)) { msg += '在這家公司的年資必須為數字\n'; if (Item == '') { Item = 'Q62'; Page = 2; } }
				if (document.getElementById('Q63').value != '' && !isUnsignedInt(document.getElementById('Q63').value)) { msg += '在這職位的年資必須為數字\n'; if (Item == '') { Item = 'Q63'; Page = 2; } }
				if (document.getElementById('Q64').value != '' && !isUnsignedInt(document.getElementById('Q64').value)) { msg += '最近升遷離本職幾年必須為數字\n'; if (Item == '') { Item = 'Q64'; Page = 2; } }
			}

			if (msg != '') {
				ChangeMode(Page);
				if (document.getElementById(Item))
					document.getElementById(Item).focus();
				alert(msg);
				return false;
			}
		}

		function sol(nn) {
			var myTR = document.getElementById("SolTR");
			if (nn == '04') {
				myTR.style.display = 'inline';
			}
			else {
				myTR.style.display = 'none';
			}
		}

		function checkNativeID() {
			var myTr1 = document.getElementById("Tr1");
			if (document.form1.MIdentityID.value == '05') {
				myTr1.style.display = 'inline';
			}
			else {
				myTr1.style.display = 'none';
			}
		}

		function hard() {
			if (document.getElementById('TPlanID').value == '28') {
				if (document.form1.IdentityID_4.checked) {
					document.form1.HandTypeID.disabled = false;
					document.form1.HandLevelID.disabled = false;
				}
				else {
					document.form1.HandTypeID.disabled = true;
					document.form1.HandLevelID.disabled = true;
				}
			}
			else {
				if (document.form1.IdentityID_5.checked) {
					document.form1.HandTypeID.disabled = false;
					document.form1.HandLevelID.disabled = false;
				}
				else {
					document.form1.HandTypeID.disabled = true;
					document.form1.HandLevelID.disabled = true;
				}
			}
		}

		function chknum(value) {
			if (value >= 48 && value <= 57) return true;
			else return false;
		}
		function ChangeMode(num) {
			if (document.getElementById('DetailTable') && document.getElementById('BackTable')) {
				if (num == 1) {
					document.getElementById('DetailTable').style.display = 'inline';
					document.getElementById('BackTable').style.display = 'none';
				}
				else {
					document.getElementById('BackTable').style.display = 'inline';
					document.getElementById('DetailTable').style.display = 'none';
				}
			}
		}
	</script>
</head>
<body ms_positioning="FlowLayout">
	<form id="form1" method="post" runat="server">
	<font face="新細明體">
		<table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
			<tr>
				<td>
					<table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
						<tr>
							<td>
								<asp:Label ID="TitleLab1" runat="server"></asp:Label>
								<asp:Label ID="TitleLab2" runat="server">
                                        首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<FONT color="#990000">結訓學員資料維護</FONT>
								</asp:Label>
							</td>
						</tr>
					</table>
					<table class="font" id="MenuTable" style="cursor: hand" height="20" cellspacing="0" cellpadding="0" border="0" runat="server">
						<tr>
							<td onclick="ChangeMode(1);" width="1" background="../../images/BookMark_01.gif">
							</td>
							<td onclick="ChangeMode(1);" align="center" width="100" background="../../images/BookMark_02.gif">
								個人基本資料
							</td>
							<td onclick="ChangeMode(1);" width="11" background="../../images/BookMark_03.gif">
							</td>
							<td onclick="ChangeMode(2);" width="1" background="../../images/BookMark_01.gif">
							</td>
							<td onclick="ChangeMode(2);" align="center" width="100" background="../../images/BookMark_02.gif">
								參訓背景
							</td>
							<td onclick="ChangeMode(2);" width="11" background="../../images/BookMark_03.gif">
							</td>
						</tr>
					</table>
					<table class="table_sch" id="DetailTable" runat="server">
						<tr id="StdTr" runat="server">
							<td class="bluecol">
								學員
							</td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								班別名稱
							</td>
							<td width="200" class="whitecol">
								<asp:Label ID="ClassName" runat="server"></asp:Label>
							</td>
							<td class="bluecol_need">
								報名階段
							</td>
							<td class="whitecol">
								<asp:DropDownList ID="LevelNo" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								中文姓名
							</td>
							<td class="whitecol">
								<asp:TextBox ID="Name" runat="server" Columns="15"></asp:TextBox>
							</td>
							<td class="bluecol_need">
								學 號(兩碼)
							</td>
							<td class="whitecol">
								<asp:TextBox ID="StudentID" runat="server" Columns="3" MaxLength="2"></asp:TextBox><input id="StudentIDValue" style="width: 32px; height: 22px" type="hidden" size="1" name="StudentIDValue" runat="server"><input id="StudentIDstring" style="width: 32px; height: 22px" type="hidden" size="1" name="StudentIDstring" runat="server">
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								英文姓名
								<asp:Label ID="star1" runat="server"></asp:Label>
							</td>
							<td class="whitecol">
								Last Name(姓)
								<asp:TextBox ID="LName" runat="server" Width="100px"></asp:TextBox>
							</td>
							<td class="bluecol">
								First Name(名)
							</td>
							<td class="whitecol">
								<asp:TextBox ID="FName" runat="server" Width="100px"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								身分別
							</td>
							<td style="height: 91px" class="whitecol">
								<asp:RadioButtonList ID="PassPortNO" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal">
									<asp:ListItem Value="1">本國</asp:ListItem>
									<asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
								</asp:RadioButtonList>
								<table class="font" id="ChinaOrNotTable" style="border-collapse: collapse" bordercolor="darkseagreen" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server">
									<tr>
										<td class="whitecol">
											<asp:RadioButtonList ID="ChinaOrNot" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" CellPadding="0" CellSpacing="0">
												<asp:ListItem Value="1">大陸人士</asp:ListItem>
												<asp:ListItem Value="2">非大陸人士</asp:ListItem>
											</asp:RadioButtonList>
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Nationality" runat="server"></asp:TextBox>
										</td>
									</tr>
								</table>
							</td>
							<td class="bluecol_need">
								身分證號碼
							</td>
							<td bgcolor="#ecf7ff" style="height: 91px">
								<table id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
									<tr>
										<td class="whitecol">
											<asp:RadioButtonList ID="PPNO" runat="server" Width="150px" CssClass="font" CellPadding="0" CellSpacing="0">
												<asp:ListItem Value="1">護照號碼</asp:ListItem>
												<asp:ListItem Value="2">居留(工作)證號</asp:ListItem>
											</asp:RadioButtonList>
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="IDNO" runat="server" Columns="15"></asp:TextBox>
											<asp:Button ID="Button4" runat="server" Text="檢查" CssClass="asp_button_S"></asp:Button>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								性 別
							</td>
							<td class="whitecol">
								<asp:RadioButtonList ID="Sex" runat="server" CssClass="font" RepeatDirection="Horizontal">
									<asp:ListItem Value="M">男</asp:ListItem>
									<asp:ListItem Value="F">女</asp:ListItem>
								</asp:RadioButtonList>
							</td>
							<td class="bluecol_need">
								出生日期
							</td>
							<td class="whitecol">
								<asp:TextBox ID="Birthday" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= Birthday.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								婚姻狀況
							</td>
							<td class="whitecol">
								<asp:RadioButtonList ID="MaritalStatus" runat="server" CssClass="font" RepeatDirection="Horizontal">
									<asp:ListItem Value="1">已婚</asp:ListItem>
									<asp:ListItem Value="2">未婚</asp:ListItem>
									<asp:ListItem Value="3">暫不提供</asp:ListItem>
								</asp:RadioButtonList>
							</td>
							<td class="bluecol">
								報名管道
							</td>
							<td class="whitecol">
								<asp:DropDownList ID="EnterChannel" runat="server">
									<asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
									<asp:ListItem Value="1">網路</asp:ListItem>
									<asp:ListItem Value="2">現場</asp:ListItem>
									<asp:ListItem Value="3">通訊</asp:ListItem>
									<asp:ListItem Value="4">推介</asp:ListItem>
								</asp:DropDownList>
							</td>
						</tr>
						<tr id="TRNDTR" runat="server">
							<td class="bluecol">
								推介種類
							</td>
							<td class="whitecol">
								<asp:DropDownList ID="TRNDMode" runat="server">
									<asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
									<asp:ListItem Value="1">職訓券</asp:ListItem>
									<asp:ListItem Value="2">學習券</asp:ListItem>
									<asp:ListItem Value="3">推介券</asp:ListItem>
								</asp:DropDownList>
							</td>
							<td class="bluecol">
								券別
							</td>
							<td class="whitecol">
								<asp:RadioButtonList ID="TRNDType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
									<asp:ListItem Value="1">甲式</asp:ListItem>
									<asp:ListItem Value="2">乙式</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr id="DGTR" runat="server">
							<td class="bluecol">
								學習券身分別
							</td>
							<td colspan="3" style="height: 27px" class="whitecol">
								<asp:Label ID="DGIdentValue" runat="server"></asp:Label>
							</td>
						</tr>
						<tr id="GovTR" runat="server">
							<td class="bluecol">
								推介單個案區分
							</td>
							<td class="whitecol">
								<asp:Label ID="GovObject_Type" runat="server"></asp:Label>
							</td>
							<td class="bluecol">
								推介單身分別
							</td>
							<td class="whitecol">
								<asp:Label ID="GovSpecial_Type" runat="server"></asp:Label>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								開訓日期
							</td>
							<td class="whitecol">
								<asp:TextBox ID="OpenDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= OpenDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
							</td>
							<td class="bluecol">
								結訓日期
							</td>
							<td class="whitecol">
								<asp:TextBox ID="CloseDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= CloseDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								報到日期
							</td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="EnterDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= EnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								最高學歷
							</td>
							<td class="whitecol">
								<asp:DropDownList ID="DegreeID" runat="server">
								</asp:DropDownList>
							</td>
							<td class="bluecol_need">
								學校名稱
							</td>
							<td class="whitecol">
								<asp:TextBox ID="School" runat="server">不詳</asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								科 系
							</td>
							<td class="whitecol">
								<asp:TextBox ID="Department" runat="server">不詳</asp:TextBox>
							</td>
							<td class="bluecol_need">
								畢業狀況
							</td>
							<td class="whitecol">
								<asp:DropDownList ID="GraduateStatus" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								兵役狀況
							</td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="MilitaryID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr id="SolTR" runat="server">
							<td colspan="4" style="height: 148px">
								<font face="新細明體">
									<table class="font" id="SoldierTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
										<tr>
											<td width="100" class="bluecol_need">
												軍種
											</td>
											<td class="whitecol">
												<asp:TextBox ID="ServiceID" runat="server"></asp:TextBox>
											</td>
											<td class="bluecol">
												職務(兵役)
											</td>
											<td class="whitecol">
												<asp:TextBox ID="MilitaryAppointment" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need">
												階級
											</td>
											<td class="whitecol">
												<asp:TextBox ID="MilitaryRank" runat="server"></asp:TextBox>
											</td>
											<td class="bluecol_need">
												服務單位名稱
											</td>
											<td class="whitecol">
												<asp:TextBox ID="ServiceOrg" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol">
												主管階級姓名
											</td>
											<td class="whitecol">
												<asp:TextBox ID="ChiefRankName" runat="server"></asp:TextBox>
											</td>
											<td class="bluecol_need">
												單位電話
											</td>
											<td class="whitecol">
												<asp:TextBox ID="ServicePhone" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need">
												服役日期
											</td>
											<td colspan="3" class="whitecol">
												<asp:TextBox ID="SServiceDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SServiceDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">～
												<asp:TextBox ID="FServiceDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= FServiceDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
											</td>
										</tr>
										<tr>
											<td class="bluecol">
												服役單位地址
											</td>
											<td colspan="3" class="whitecol">
												<asp:TextBox ID="City4" runat="server" Width="130px"></asp:TextBox><input id="ZipCode4" type="hidden" size="1" name="ZipCode4" runat="server">
												<input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City4', 'ZipCode4')" type="button" value="..." class="button_b_Mini">
												<asp:TextBox ID="ServiceAddress" runat="server" Width="250px"></asp:TextBox>
											</td>
										</tr>
									</table>
								</font>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								聯絡電話
							</td>
							<td bgcolor="#ecf7ff">
								<table class="font" id="Table7" cellspacing="1" cellpadding="1" width="100%" border="0">
									<tr>
										<td class="whitecol">
											(日)
										</td>
										<td class="whitecol">
											<asp:TextBox ID="PhoneD" runat="server" Columns="13"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											(夜)
										</td>
										<td class="whitecol">
											<asp:TextBox ID="PhoneN" runat="server" Columns="13"></asp:TextBox>
										</td>
									</tr>
								</table>
							</td>
							<td class="bluecol">
								行動電話
							</td>
							<td class="whitecol">
								<asp:TextBox ID="CellPhone" runat="server" Columns="13"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								通訊地址
							</td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="City1" runat="server" Width="130px"></asp:TextBox><input id="ZipCode1" type="hidden" size="1" name="ZipCode1" runat="server">
								<input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City1', 'ZipCode1')" type="button" value="..." class="button_b_Mini">
								<asp:TextBox ID="Address" runat="server" Width="250px"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								戶籍地址
							</td>
							<td bgcolor="#ecf7ff" colspan="3" class="whitecol">
								<asp:CheckBox ID="CheckBox1" runat="server" CssClass="font" Text="同通訊地址"></asp:CheckBox><br>
								<asp:TextBox ID="City2" runat="server" Width="130px"></asp:TextBox><input id="ZipCode2" type="hidden" size="1" name="ZipCode2" runat="server">
								<input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City2', 'ZipCode2')" type="button" value="..." class="button_b_Mini">
								<asp:TextBox ID="HouseholdAddress" runat="server" Width="250px"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								電子郵件
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<asp:TextBox ID="Email" runat="server"></asp:TextBox>
							</td>
							<td class="bluecol_need">
								津貼類別
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<font face="新細明體">
									<asp:DropDownList ID="SubsidyID" runat="server">
									</asp:DropDownList>
									<input id="SubsidyHidden" type="hidden" size="1" name="Hidden" runat="server"></font>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								主要參訓<br>
								身分別
							</td>
							<td style="height: 34px" bgcolor="#ecf7ff" colspan="3" class="whitecol">
								<asp:DropDownList ID="MIdentityID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr id="Tr1" runat="server">
							<td class="bluecol_need">
								民族別
							</td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="NativeID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								參訓身分別
								<p>
									(可複選，最多三項)
								</p>
							</td>
							<td bgcolor="#ecf7ff" colspan="3" class="whitecol">
								<asp:CheckBoxList ID="IdentityID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
								</asp:CheckBoxList>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								障礙類別
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<asp:DropDownList ID="HandTypeID" runat="server">
								</asp:DropDownList>
							</td>
							<td class="bluecol">
								障礙等級
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<asp:DropDownList ID="HandLevelID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								離訓日期
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<font face="新細明體">
									<asp:TextBox ID="RejectTDate1" runat="server" Width="75px" onfocus="this.blur()"></asp:TextBox></font>
							</td>
							<td class="bluecol">
								退訓日期
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<asp:TextBox ID="RejectTDate2" runat="server" Width="75px" onfocus="this.blur()"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								緊急通知人<br>
								姓名
								<asp:Label ID="star2" runat="server"></asp:Label>
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<font face="新細明體">
									<asp:TextBox ID="EmergencyContact" runat="server"></asp:TextBox></font>
							</td>
							<td class="bluecol_need">
								緊急通知人<br>
								電話
								<asp:Label ID="star3" runat="server"></asp:Label>
							</td>
							<td bgcolor="#ecf7ff" class="whitecol">
								<asp:TextBox ID="EmergencyPhone" runat="server"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								緊急通知人<br>
								關係
								<asp:Label ID="star4" runat="server"></asp:Label>
							</td>
							<td bgcolor="#ecf7ff" colspan="3" class="whitecol">
								<asp:TextBox ID="EmergencyRelation" runat="server"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								緊急通知人<br>
								地址
								<asp:Label ID="star5" runat="server"></asp:Label>
							</td>
							<td bgcolor="#ecf7ff" colspan="3" class="whitecol">
								<asp:TextBox ID="City3" runat="server" Width="130px"></asp:TextBox><input id="ZipCode3" type="hidden" size="1" name="ZipCode3" runat="server">
								<input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City3', 'ZipCode3')" type="button" value="..." class="button_b_Mini">
								<asp:TextBox ID="EmergencyAddress" runat="server" Width="250px"></asp:TextBox>
							</td>
						</tr>
						<tr id="ForeTr1" runat="server">
							<td align="center" colspan="4" class="bluecol">
								國內親屬資料
							</td>
						</tr>
						<tr id="ForeTr2" runat="server">
							<td class="bluecol">
								姓名
							</td>
							<td class="whitecol">
								<asp:TextBox ID="ForeName" runat="server"></asp:TextBox>
							</td>
							<td class="bluecol">
								稱謂
							</td>
							<td class="whitecol">
								<asp:TextBox ID="ForeTitle" runat="server" Columns="15"></asp:TextBox>
							</td>
						</tr>
						<tr id="ForeTr3" runat="server">
							<td class="bluecol">
								性別
							</td>
							<td class="whitecol">
								<asp:RadioButtonList ID="ForeSex" runat="server" CssClass="font" RepeatDirection="Horizontal">
									<asp:ListItem Value="M">男</asp:ListItem>
									<asp:ListItem Value="F">女</asp:ListItem>
								</asp:RadioButtonList>
							</td>
							<td class="bluecol">
								出生日期
							</td>
							<td class="whitecol">
								<asp:TextBox ID="ForeBirth" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= ForeBirth.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
							</td>
						</tr>
						<tr id="ForeTr4" runat="server">
							<td class="bluecol">
								身分證號碼
							</td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="ForeIDNO" runat="server"></asp:TextBox>
							</td>
						</tr>
						<tr id="ForeTr5" runat="server">
							<td class="bluecol">
								戶籍地址
							</td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="City6" runat="server" Width="130px"></asp:TextBox><input id="ForeZip" type="hidden" size="1" runat="server">
								<input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City6', 'ForeZip')" type="button" value="..." class="button_b_Mini">
								<asp:TextBox ID="ForeAddr" runat="server" Width="250px"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td rowspan="2" class="bluecol">
								受訓服務單位
							</td>
							<td class="whitecol">
								<font face="新細明體">1.
									<asp:TextBox ID="PriorWorkOrg1" runat="server"></asp:TextBox></font>
							</td>
							<td rowspan="2" class="bluecol">
								職稱
							</td>
							<td class="whitecol">
								<font face="新細明體">1.
									<asp:TextBox ID="Title1" runat="server"></asp:TextBox></font>
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								<font face="新細明體">2.
									<asp:TextBox ID="PriorWorkOrg2" runat="server"></asp:TextBox></font>
							</td>
							<td class="whitecol">
								<font face="新細明體">2.
									<asp:TextBox ID="Title2" runat="server"></asp:TextBox></font>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								受訓前任職起<br>
								迄年月
							</td>
							<td colspan="3" class="whitecol">
								<table class="font" id="Table6" cellspacing="1" cellpadding="1" border="0">
									<tr>
										<td class="whitecol">
											1.
										</td>
										<td>
											<asp:TextBox ID="SOfficeYM1" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SOfficeYM1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
										</td>
										<td>
											～
										</td>
										<td>
											<asp:TextBox ID="FOfficeYM1" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= FOfficeYM1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											2.
										</td>
										<td>
											<asp:TextBox ID="SOfficeYM2" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SOfficeYM2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
										</td>
										<td>
											～
										</td>
										<td>
											<asp:TextBox ID="FOfficeYM2" runat="server" Width="75px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= FOfficeYM2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								受訓前薪資
							</td>
							<td class="whitecol">
								<asp:TextBox ID="PriorWorkPay" runat="server" Width="100px"></asp:TextBox>
							</td>
							<td class="bluecol_need">
								受訓前失業週數
								<asp:Label ID="star6" runat="server"></asp:Label>
							</td>
							<td class="whitecol">
								<asp:TextBox ID="RealJobless" runat="server" Width="50px"></asp:TextBox><asp:DropDownList ID="JoblessID" runat="server">
								</asp:DropDownList>
								<br>
								<asp:Label ID="lb_msg" runat="server" ForeColor="Red"></asp:Label>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								交通方式
							</td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="Traffic" runat="server">
									<asp:ListItem Value="0">請選擇</asp:ListItem>
									<asp:ListItem Value="1">住宿</asp:ListItem>
									<asp:ListItem Value="2">通勤</asp:ListItem>
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								提供基本資料<br>
								供查詢
							</td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="ShowDetail" runat="server">
									<asp:ListItem Value="0">請選擇</asp:ListItem>
									<asp:ListItem Value="Y">是</asp:ListItem>
									<asp:ListItem Value="N">否</asp:ListItem>
								</asp:DropDownList>
								<font face="新細明體">(姓名、出生年月日、性別、學歷、科系、電話、電子郵件帳號、專長)</font>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								預算別
							</td>
							<td colspan="3" class="whitecol">
								<asp:RadioButtonList ID="BudID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
								</asp:RadioButtonList>
								<asp:Literal ID="BudIDMsg" runat="server"></asp:Literal>
							</td>
						</tr>
						<tr>
							<td class="bluecol">
								公費/自費<br>
								(職訓券必填)
							</td>
							<td colspan="3" class="whitecol">
								<asp:RadioButtonList ID="PMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
									<asp:ListItem Value="1">公費</asp:ListItem>
									<asp:ListItem Value="2">自費</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr>
							<td colspan="4" class="whitecol">
								&nbsp;&nbsp;&nbsp; <font color="red">*</font>本人
								<asp:RadioButtonList ID="IsAgree" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
									<asp:ListItem Value="Y">同意</asp:ListItem>
									<asp:ListItem Value="N">不同意</asp:ListItem>
								</asp:RadioButtonList>
								個人基本資料，供 勞動部勞動力發展署 暨所屬機關運用，以從事職業訓練及就業服務
							</td>
						</tr>
					</table>
					<table class="table_sch" id="Table2">
						<!-- NICK CHANGE 060316-->
						<tr id="LearnTR1" runat="server">
							<td width="40%" class="bluecol">
								學習券課程單元
							</td>
							<td width="30%" class="bluecol">
								實際上課時數
							</td>
							<td width="30%" class="bluecol">
								單元成績(0~100分)
							</td>
						</tr>
						<tr id="LearnTR2" runat="server">
							<td rowspan="4" class="whitecol">
								<asp:CheckBoxList ID="RelClass_Unit" runat="server" CssClass="font" CellSpacing="10" CellPadding="1" Height="5px">
								</asp:CheckBoxList>
							</td>
							<td class="whitecol">
								<asp:TextBox ID="Unit1Hour" runat="server" Columns="5" MaxLength="2"></asp:TextBox>小時(
								<asp:Label ID="Label1" runat="server">Label</asp:Label>H)
							</td>
							<td class="whitecol">
								&nbsp;
								<asp:TextBox ID="Unit1Score" runat="server" Width="50px"></asp:TextBox>分
							</td>
						</tr>
						<tr id="LearnTR3" runat="server">
							<td class="whitecol">
								<asp:TextBox ID="Unit2Hour" runat="server" Columns="5" MaxLength="2"></asp:TextBox>小時(
								<asp:Label ID="Label2" runat="server">Label</asp:Label>H)
							</td>
							<td class="whitecol">
								&nbsp;
								<asp:TextBox ID="Unit2Score" runat="server" Width="50px"></asp:TextBox>分
							</td>
						</tr>
						<tr id="LearnTR4" runat="server">
							<td class="whitecol">
								<asp:TextBox ID="Unit3Hour" runat="server" Columns="5" MaxLength="2"></asp:TextBox>小時(
								<asp:Label ID="Label3" runat="server">Label</asp:Label>H)
							</td>
							<td class="whitecol">
								&nbsp;
								<asp:TextBox ID="Unit3Score" runat="server" Width="50px"></asp:TextBox>分
							</td>
						</tr>
						<tr id="LearnTR5" runat="server">
							<td class="whitecol">
								<asp:TextBox ID="Unit4Hour" runat="server" Columns="5" MaxLength="2"></asp:TextBox>小時(
								<asp:Label ID="Label4" runat="server">Label</asp:Label>H)
							</td>
							<td class="whitecol">
								&nbsp;
								<asp:TextBox ID="Unit4Score" runat="server" Width="50px"></asp:TextBox>分
							</td>
						</tr>
						<tr id="TPlan23TR" runat="server">
							<td class="bluecol_need">
								指定投保單位<br>
								保險證號
							</td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="ActNo" runat="server"></asp:TextBox>
							</td>
						</tr>
						<!-- END--->
					</table>
				</td>
			</tr>
		</table>
		<table class="table_nw" id="BackTable" runat="server" width="740">
			<tr>
				<td class="bluecol" colspan="4">
					服務單位資料
				</td>
			</tr>
			<tr>
				<td class="bluecol_need" width="100">
					郵政/銀行帳號
				</td>
				<td class="whitecol" colspan="3">
					<asp:RadioButtonList ID="AcctMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
						<asp:ListItem Value="0">郵局帳號</asp:ListItem>
						<asp:ListItem Value="1">銀行帳號</asp:ListItem>
					</asp:RadioButtonList>
				</td>
			</tr>
			<tr id="PortTR" runat="server">
				<td class="bluecol_need">
					局號
				</td>
				<td class="whitecol" width="200">
					<asp:TextBox ID="PostNo_1" runat="server" Columns="8"></asp:TextBox>－
					<asp:TextBox ID="PostNo_2" runat="server" Columns="1"></asp:TextBox>
				</td>
				<td class="bluecol_need">
					帳號
				</td>
				<td class="whitecol" width="200">
					<asp:TextBox ID="AcctNo1_1" runat="server" Columns="8"></asp:TextBox>－
					<asp:TextBox ID="AcctNo1_2" runat="server" Columns="1"></asp:TextBox>
				</td>
			</tr>
			<tr id="BankTR1" runat="server">
				<td class="bluecol_need">
					銀行名稱
				</td>
				<td class="whitecol" width="200" colspan="3">
					<asp:TextBox ID="BankName" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr id="BankTR2" runat="server">
				<td class="bluecol_need">
					總代號
				</td>
				<td class="whitecol" width="200" colspan="3">
					<asp:TextBox ID="AcctHeadNo" runat="server" Columns="8"></asp:TextBox>
				</td>
			</tr>
			<tr id="BankTR3" runat="server">
				<td class="bluecol_need">
					帳號
				</td>
				<td class="whitecol" colspan="3">
					<asp:TextBox ID="AcctNo2" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					第一次投保日
				</td>
				<td class="whitecol" colspan="3">
					<asp:TextBox ID="FirDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= FirDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					目前任職<br>
					公司名稱
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Uname" runat="server"></asp:TextBox>
				</td>
				<td class="bluecol">
					統一編號
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Intaxno" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol_need">
					公司電話
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Tel" runat="server"></asp:TextBox>
				</td>
				<td class="bluecol">
					公司傳真
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Fax" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol_need">
					公司地址
				</td>
				<td class="whitecol" colspan="3">
					<asp:TextBox ID="City5" runat="server" Width="130px"></asp:TextBox><input onclick="getZip('../../js/Openwin/zipcode.aspx', 'City5', 'Zip')" type="button" value="..."><input id="Zip" type="hidden" size="1" runat="server">
					<asp:TextBox ID="Addr" runat="server" Width="250px"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					目前任職部門
				</td>
				<td class="whitecol">
					<asp:TextBox ID="ServDept" runat="server"></asp:TextBox>
				</td>
				<td class="bluecol">
					職稱
				</td>
				<td class="whitecol">
					<asp:TextBox ID="JobTitle" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					個人到任<br>
					目前任職<br>
					公司起日
				</td>
				<td class="whitecol">
					<asp:TextBox ID="SDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
				</td>
				<td class="bluecol">
					個人到任<br>
					目前職務<br>
					起日
				</td>
				<td class="whitecol">
					<asp:TextBox ID="SJDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SJDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					最近升遷日期
				</td>
				<td class="whitecol" colspan="3">
					<asp:TextBox ID="SPDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= SPDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
				</td>
			</tr>
			<tr>
				<td class="bluecol" colspan="4">
					參訓背景資料
				</td>
			</tr>
			<tr>
				<td class="bluecol_need">
					是否由公司<br>
					推薦參訓
				</td>
				<td class="whitecol" colspan="3">
					<asp:RadioButtonList ID="Q1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
						<asp:ListItem Value="1">是</asp:ListItem>
						<asp:ListItem Value="0">否</asp:ListItem>
					</asp:RadioButtonList>
				</td>
			</tr>
			<tr>
				<td class="bluecol_need">
					參訓動機
				</td>
				<td class="whitecol" colspan="3">
					<asp:CheckBoxList ID="Q2" runat="server" CssClass="font" RepeatDirection="Horizontal" CellPadding="0" CellSpacing="0" RepeatColumns="2">
						<asp:ListItem Value="1">為補充與原專長相關之技能</asp:ListItem>
						<asp:ListItem Value="2">轉換其他行職業所需技能</asp:ListItem>
						<asp:ListItem Value="3">拓展工作領域及視野</asp:ListItem>
						<asp:ListItem Value="4">其他</asp:ListItem>
					</asp:CheckBoxList>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					訓後動向
				</td>
				<td class="whitecol" colspan="3">
					<asp:RadioButtonList ID="Q3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
						<asp:ListItem Value="1">轉換工作</asp:ListItem>
						<asp:ListItem Value="2">留任</asp:ListItem>
						<asp:ListItem Value="3">其他</asp:ListItem>
					</asp:RadioButtonList>
					<asp:TextBox ID="Q3_Other" runat="server"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol_need">
					服務單位行業別
				</td>
				<td class="whitecol" colspan="3">
					<asp:DropDownList ID="Q4" runat="server">
					</asp:DropDownList>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					服務單位是否<br>
					屬於中小企業
				</td>
				<td class="whitecol" colspan="3">
					<asp:RadioButtonList ID="Q5" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
						<asp:ListItem Value="是">是</asp:ListItem>
						<asp:ListItem Value="否">否</asp:ListItem>
					</asp:RadioButtonList>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					個人工作年資
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Q61" runat="server" Columns="5"></asp:TextBox>
				</td>
				<td class="bluecol">
					在這家公司的年資
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Q62" runat="server" Columns="5"></asp:TextBox>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					在這職位的年資
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Q63" runat="server" Columns="5"></asp:TextBox>
				</td>
				<td class="bluecol">
					最近升遷離本職幾年
				</td>
				<td class="whitecol">
					<asp:TextBox ID="Q64" runat="server" Columns="5"></asp:TextBox>
				</td>
			</tr>
		</table>
		<table id="Table4" cellspacing="1" cellpadding="1" width="740" border="0" style="width: 740px; height: 28px">
			<tr>
				<td align="center">
					<asp:Button ID="Button1" runat="server" Text="儲存回查詢頁面" CssClass="asp_button_M"></asp:Button>
					<asp:Button ID="Button2" runat="server" Text="維護下一位學員" CssClass="asp_button_M"></asp:Button>
					<asp:Button ID="Button3" runat="server" Text="不儲存回上一頁" CssClass="asp_button_M"></asp:Button>
				</td>
			</tr>
		</table>
	</font>
	<input id="RoleID" type="hidden" size="1" runat="server"><input id="Process" type="hidden" size="1" name="Process" runat="server"><input id="TPlanID" type="hidden" size="1" runat="server">
	</form>
</body>
</html>
