<%
currUAModule = "DO005"
currUASubModule = "DO005024"
currUACategory = "DO"
%>
<!--#include file="../include/glob_conn_open.asp"-->
<!--#include file="../include/global_value.asp" -->
<!--#include file="../include/gen_css.asp" -->
<!--#include file="../include/gen_data.asp" -->
<!--#include file="../include/glob_session_check.asp"-->
<!--#include file="../include/glob_defparser.asp"-->
<!--#include file="../include/glob_akses_check.asp"-->
<%
IF currUAAllAccess=0 AND currUAEdit=0 THEN
	Response.Write glob_WriteBlockPage("")
%>
	<!--#include File="../include/glob_conn_close.asp" -->
<%
	Response.End
END IF	
%>
<HTML>
<%
currAct = Trim(Request.QueryString("crAct"))
currCopy = Trim(Request.QueryString("crCopy"))
Dim xLainTypWar(4,1)
xLainTypWar(0,0) = ""
xLainTypWar(0,1) = "--Choose--"
xLainTypWar(1,0) = "n"
xLainTypWar(1,1) = "No Warranty"
xLainTypWar(2,0) = "w"
xLainTypWar(2,1) = "Warranty"
xLainTypWar(3,0) = "l"
xLainTypWar(3,1) = "Life Time"	 
Dim xLain(1,1)
xLain(0,0) = ""
xLain(0,1) = "--Choose--"
Dim xLainStatus(1,1)
xLainStatus(0,0) = ""
xLainStatus(0,1) = "--Not Selected--"
Dim xLainBrand(1,1)
xLainBrand(0,0) = ""
xLainBrand(0,1) = "Not Selected"
xMouseBack = " src=""../image/btn_back_012.gif"" onmouseover=this.src=""../image/btn_back_112.gif""; onmouseout=this.src=""../image/btn_back_012.gif"";"
xMouseSave = " src=""../image/btn_save_012.gif"" onmouseover=this.src=""../image/btn_save_112.gif""; onmouseout=this.src=""../image/btn_save_012.gif"";"
xMouseCancel = " src=""../image/btn_cancel_012.gif"" onmouseover=this.src=""../image/btn_cancel_112.gif""; onmouseout=this.src=""../image/btn_cancel_012.gif"";"
xMouseAdd = " src=""../image/btn_add_012.gif"" onmouseover=this.src=""../image/btn_add_112.gif""; onmouseout=this.src=""../image/btn_add_012.gif"";"
xMouseEdit = " src=""../image/btn_editSO1.gif"" onmouseover=this.src=""../image/btn_editSO2.gif"" onmouseout=this.src=""../image/btn_editSO1.gif"";"
xMouseSaveNext = " src=""../image/btn_savenext_012.gif"" onmouseover=this.src=""../image/btn_savenext_112.gif""; onmouseout=this.src=""../image/btn_savenext_012.gif"";"
xMouseSaveNew= " src=""../image/btn_savenew_012.gif"" onmouseover=this.src=""../image/btn_savenew_112.gif""; onmouseout=this.src=""../image/btn_savenew_012.gif"";"
xMouseRemove = " src=""../image/btn_remove_010.gif"" onmouseover=this.src=""../image/btn_remove_110.gif""; onmouseout=this.src=""../image/btn_remove_010.gif"";"
xMouseFind= " src=""../image/btn_findmaster_012.gif"" onmouseover=this.src=""../image/btn_findmaster_112.gif""; onmouseout=this.src=""../image/btn_findmaster_012.gif"";"
xMouseDelete =  " src=""../image/btn_delete_012.gif"" onmouseover=this.src=""../image/btn_delete_112.gif""; onmouseout=this.src=""../image/btn_delete_012.gif"";"
xMousePreview = " src=""../image/btn_preview_012.gif"" onmouseover=this.src=""../image/btn_preview_112.gif""; onmouseout=this.src=""../image/btn_preview_012.gif"";"
adaPromo = 0
isCheckedAuthorisedWarranty = "checked"	
isCheckedMerchantWarranty = "disabled"
rbWar0000112 = "checked"
rbWar0000212 = "checked"
currVPLPriceID = "CUR01"
rbPPNInclude = "checked"
strDisplayPeriode = "style=""display:none"""
strTitle = "[Profile]>Add"
flagCat = 0
flagBrand = 0
FUNCTION namaCat(byval id)
	Dim n,nama_Cat
	IF id="" THEN
		nama_Cat = "<option value='' selected>Not Selected"
	ELSE
		nama_Cat = "<option value=''>Not Selected"
	END IF

	FOR n=0 TO countCat-1
		IF CatId(n)=id THEN
			nama_Cat = nama_Cat&"<option value='"&CatId(n)&"' selected>"&CatName(n) 
		ELSE
			nama_Cat = nama_Cat&"<option value='"&CatId(n)&"'>"&CatName(n) 
		END IF
	NEXT
	namaCat = nama_Cat
END FUNCTION
FUNCTION namaBrand(byval id)
	Dim n,nama_Brand
	IF id="" THEN
		nama_Brand = "<option value='' selected>Not Selected"
	ELSE
		nama_Brand = "<option value=''>Not Selected"
	END IF

	FOR n=0 TO countBrand-1
		IF BrandId(n)=id THEN
			nama_Brand = nama_Brand&"<option value='"&BrandId(n)&"' selected>"& BrandName(n) 
		ELSE
			nama_Brand = nama_Brand&"<option value='"&BrandId(n)&"'>"&BrandName(n) 
		END IF
	NEXT
	namaBrand = nama_Brand
END FUNCTION
FUNCTION escapeVal(s)
	Dim xxx
	IF Len(s)>0 THEN
		xxx = Server.HTMLEncode(s)
	ELSE
		xxx = s
	END IF
	IF isEmpty(s) OR isNull(s) THEN
		escapeVal = ""
	ELSE
		escapeVal = Replace(Replace(xxx,"""","&dblquote;"),"'","&#39;")
	END IF
END FUNCTION
FUNCTION ceknull(pString)
	str = Trim(pString)
	IF (isnull(str) OR str="" OR isEmpty(str)) THEN
		ceknull = "<font class='wordredregular'>n/a</font>"
	ELSE
		ceknull = Server.HTMLEncode(str)
	END IF
END FUNCTION
FUNCTION gen_date(model,prmFuncFlg,prmFunc,stsFld,frmnm,objnm,dt,x,y,stsPos)
	Dim isKite2,imgStandard,imgHover
	Dim cdtview,cdt,cWajibIsi
	dim s
	IF model="kite" THEN
		isKite2 = True
		model = "model1"
		imgStandard	= "../image/date.gif"
		imgHover = "../image/date.gif"
	ELSE
		imgStandard	= "../Image/btn_clock_000.gif"
		imgHover = "../Image/btn_clock_100.gif"
	END IF
	IF inStr(model,Chr(162)) THEN
		arrModel = Split(model,Chr(162))
		glob_datemodel = arrModel(0)
		param = arrModel(1)
	ELSE
		glob_datemodel=model
	END IF
	IF isnull(x) THEN x = "null"
	IF isnull(y) THEN y = "null"
	IF stsPos THEN 
		stsPos = "true" 
	ELSE
		stsPos = "false"
	END IF
	cWajibIsi = ""
	IF dt<>"" THEN
		cdt = CDate(dt)
		tgl = Day(cdt)
		FOR i=1 TO 2-Len(tgl)
	 		tgl  = "0"&tgl
		NEXT
		cdtview = tgl&" "&Left(Monthname(Month(cdt)),3)&" "&Year(cdt)
	END IF
	IF stsFld=1 THEN cWajibIsi = " class=wordfieldnormalmustdisabled readonly"
	IF stsFld=2 THEN cWajibIsi = " class=wordfieldnormaldisabled readonly"
	IF stsFld=3 THEN cWajibIsi = " class=fielddatanewmustinput readonly"
	IF stsFld=4 THEN cWajibIsi = " class=fielddatanewmustinputgreen readonly"
	s = s&"<table cellpadding=0 cellspacing=0 border=0>"&vbCrLf&_
		"<tr height=25>"&vbCrLf&_
		"<td valign=top style='border:0px'>"&vbCrLf&_
		"<input type=hidden name=""Cr"&objnm&""" value="""&dt&""" >"&vbCrLf&_
		"<input type=text size=15 name=""Cr"&objnm&"View"" value="""&cdtview&""""&cWajibIsi&" "&param&">"&vbCrLf&_
		"</td>"&vbCrLf&_
		"<td valign=top style='border:0px;padding-left:5px' style=""cursor:hand;"">"&vbCrLf&_
		"<a onmouseover=""document['"&objnm&"'].imgRolln=document['"&objnm&"'].src;document['"&objnm&"'].src=document['"&objnm&"'].lowsrc;"" "&_
		"onmouseout=""document['"&objnm&"'].src=document['"&objnm&"'].imgRolln"" onclick=""javascript:getCalendarFor('"&glob_datemodel&"','"&prmFuncFlg&"','"&prmFunc&"','"&now()&"',document."&frmnm&".Cr"&objnm&",null,document."&frmnm&".Cr"&objnm&".value,document."&frmnm&".Cr"&objnm&"View,'d s Y',"&x&","&y&","&stsPos&")"">"&vbCrLf&_
		"<img border=""0"" src="""&imgStandard&""" id="""&objnm&""" name="""&objnm&""" dynamicanimation="""&objnm&""" lowsrc="""&imgHover&""" align=""middle"" >"&vbCrLf&_
		"</a>"&vbCrLf&_
		"</td>"&vbCrLf&_
		"</tr>"&vbCrLf&_
		"</table>"
	gen_date = s
END FUNCTION
FUNCTION inveprodcatalog_menu(id)
	Dim s,jmlmenu
	s = ""
	Dim mnu(8,3)
	mnu(0,0) = "profile"
	mnu(0,1) = "Profile"
	mnu(0,2) = "digoff_inve_prodcatalog_view.asp?crPartID="&currPartID&xLoadNext
	mnu(1,0) = "detail"
	mnu(1,1) = "Detail"
	mnu(1,2) = "digoff_inve_prodcatalogdet_view.asp?crPartID="&currPartID&xLoadNext
	mnu(2,0) = "image"
	mnu(2,1) = "Image"
	mnu(2,2) = "digoff_inve_prodcatalogimg_view.asp?crPartID="&currPartID&xLoadNext
	mnu(3,0) = "overview"
	mnu(3,1) = "Overview"
	mnu(3,2) = "digoff_inve_prodcatalogoverview_view.asp?crPartID="&currPartID&xLoadNext
	mnu(4,0) = "brochure"
	mnu(4,1) = "Brochure"
	mnu(4,2) = "digoff_inve_prodcatalogbrochure_view.asp?crPartID="&currPartID&xLoadNext
	mnu(5,0) = "offer"
	mnu(5,1) = "Offer"
	mnu(5,2) = "digoff_inve_prodcatalogoffer_view.asp?crPartID="&currPartID&xLoadNext
	mnu(6,0) = "vs"
	mnu(6,1) = "Voucher"
	mnu(6,2) = "digoff_inve_vssku_view.asp?crPartID="&currPartID&xLoadNext
	mnu(7,0) = "market"
	mnu(7,1) = "Marketing"
	mnu(7,2) = "digoff_inve_prodcatalogmarket_view.asp?crPartID="&currPartID&xLoadNext
	IF id<>"" THEN
		s = gen_menu(mnu,id)
	END IF
	inveprodcatalog_menu = s
END FUNCTION


FUNCTION gen_menu(mnu,id)
	Dim s,jmlmenu
	Dim sts
	sts = False
	jmlmenu = Ubound(mnu)
	FOR i=0 TO jmlmenu-1
		IF mnu(i,0)=id THEN
			sts = True
			IF i=0 THEN
				s = s&"<td><img src=""../image/tab/slc_tab_actleft_010.gif"">"
			ELSE
				s = s&"<td><img src=""../image/tab/slc_tab_noactrightactleft_010.gif""></td>"
			END IF
			s = s&"<td align=""center"" background=""../image/tab/slc_tab_bgact_010.gif""><font class=wordactivetab><nobr>"&mnu(i,1)&"</nobr></font></td>"
			IF i=jmlmenu-1 THEN
				s = s&"<td><img src=""../image/tab/slc_tab_actright_010.gif""</td>"
			ELSE
				s = s&"<td><img src=""../image/tab/slc_tab_actrightnoactleft_010.gif""</td>"
			END IF
		ELSE
			IF i=0 THEN
				s = s&"<td><img src=""../image/tab/slc_tab_noactleft_010.gif""</td>"
			ELSE
				IF NOT sts THEN
					 s = s&"<td><img src=""../image/tab/slc_tab_noactrightnoactleft_010.gif""></td>"
				END IF
			END IF
			IF (LCase(currAct)="edit" OR LCase(currAct)="mutation" OR LCase(currAct)="add") THEN
				s = s&"<td align=""center"" background=""../image/tab/slc_tab_bgnoact_010.gif""><nobr><font class=wordnonactivetab>"&mnu(i,1)&"</font></nobr></td>"
			ELSE
				s = s&"<td align=""center"" background=""../image/tab/slc_tab_bgnoact_010.gif""><nobr><a class=""Tab"" href="&mnu(i,2)&" title="""&mnu(i,1)&"""><nobr>"&mnu(i,1)&"</a></nobr></td>"
			END IF
			IF i=(jmlmenu-1) THEN
				s=s&"<td><img src=""../image/tab/slc_tab_noactright_010.gif""</td>"
				sts = True
			ELSE
				sts = False
			END IF
		END IF
		s = s&"</td>"&vbCrLf
	NEXT
	gen_menu = s
END FUNCTION
FUNCTION FormatDateView(currDate)
	IF (currDate<>"") THEN
		cdt = CDate(currDate)
		tgl = Day(cdt)
		FOR i=1 TO 2-Len(tgl)
			tgl  = "0"&tgl
		NEXT
		FormatDateView = tgl&" "&Left(MonthName(Month(cdt)),3)&" "&Year(cdt)
	END IF
END FUNCTION
FUNCTION FormatNumberCustom(currValue,num)
	FormatNumberCustom = Replace(FormatNumber(currValue,num),",","")
END FUNCTION
FUNCTION FormatNumberRoundPrice(currValue,currID,num)
	IF currID="CUR01" THEN
		IF (((Int(currValue)+1) Mod 100)>0) THEN
			FormatNumberRoundPrice = FormatNumber(Int(currValue),num)
		ELSE
			FormatNumberRoundPrice = FormatNumber((Int(currValue)+1),num)
		END IF
	ELSE
		FormatNumberRoundPrice = FormatNumber(currValue,num)
	END IF
END FUNCTION
IF (LCase(currAct)="edit" OR currCopy="Yes") THEN
	currPartID = Trim(Request.QueryString("crPartID"))
	Set getRs = Server.CreateObject("ADODB.Recordset")
	Set RSCurrency = Server.CreateObject("ADODB.Recordset")
	Set getRsLain = Server.CreateObject("ADODB.Recordset")
	flagCat = 1
	flagBrand = 1		
	xstring = "SELECT TOP 1 ISNULL(i.vPartID,'') AS vPartID,ISNULL(i.vBrandID,'') AS vBrandID,ISNULL(i.vSeri,'') AS vSeri,ISNULL(i.vDesc,'') AS vDesc,"&_
			  "ISNULL(vStatusID,0) AS vStatusID,vStartPeriod,vEndPeriod,ISNULL(vShipWeight,0) AS vShipWeight,"&_
			  "ISNULL(i.vManufact,'') AS vManufact,ISNULL(i.vActivation,0) AS vActivation,ISNULL(i.vMarketingInfo,'') AS vMarketingInfo,"&_
			  "ISNULL(i.vNote,'') AS vNote,ISNULL(i.vJaminanMurah,0) AS vJaminanMurah,"&_
			  "ISNULL(i.vAuthorisedWarranty,'TRUE') AS vAuthorisedWarranty,ISNULL(i.vExtTxtWarranty,'') AS vExtTxtWarranty,"&_
			  "ISNULL(i.vCreatorNo,'') AS vCreatorNo,i.vCreatorDateTime,ISNULL(i.vCreatorIP,'') AS vCreatorIP,ISNULL(i.vEditorNo,'') AS vEditorNo,"&_
			  "i.vEditorDateTime,ISNULL(i.vEditorIP,'') AS vEditorIP,"&_
			  "(SELECT TOP 1 ISNULL(vName,'') FROM tlu_inveCPBrand WHERE vBrandID=i.vBrandID) AS vBrandName,"&_
			  "(SELECT TOP 1 ISNULL(vCatID,'') FROM trx_inveCPCatRel WHERE vPartID=i.vPartID AND LOWER(vTypeCat)='p') AS vCatPrimaryID,"&_
			  "(SELECT TOP 1 ISNULL(vName,'') FROM trx_inveCPCatRel AS r INNER JOIN tlu_inveCPCategory AS c ON r.vCatID=c.vCatID "&_
			  "AND r.vPartID=i.vPartID AND LOWER(r.vTypeCat)='p') AS vCatPrimaryName,"&_
			  "(SELECT TOP 1 ISNULL(vGoodsDescription,'') FROM trx_inveCPCatRel AS r INNER JOIN tlu_inveCPCategory AS c ON r.vCatID=c.vCatID "&_ 
			  "AND r.vPartID=i.vPartID AND LOWER(r.vTypeCat)='p') AS vGoodDesc,"&_
			  "(SELECT TOP 1 ISNULL(a.vNickName,'') FROM trx_HRDEmployment AS a WHERE a.vNoEmp=i.vCreatorNo) AS creatorName,"&_
			  "(SELECT TOP 1 ISNULL(a.vNickName,'') FROM trx_HRDEmployment AS a WHERE a.vNoEmp=i.vEditorNo) AS editorName,"&_
			  "ISNULL(s.vSVndID,''),(SELECT TOP 1 ISNULL(vName,'') FROM trx_vendorMain WHERE vVndID=s.vSVndID) AS vSVndName,"&_
			  "ISNULL(buy.vCntnPrcCurrID,'CUR01') AS vCntnPrcCurrID,ISNULL(buy.vCntnPrc,0) AS vCntnPrc,"&_
			  "ISNULL(s.vPrcCurrID,'CUR01') AS vPrcCurrID,ISNULL(s.vPrice,0) AS vPrice,ISNULL(s.vSPrcCurrID,'CUR01') AS vSPrcCurrID,"&_
			  "ISNULL(s.vSPrice,0) AS vSPrice,s.vStartSPrice,vEndSPrice,ISNULL(s.vMinPrice,0) AS vMinPrice,"&_
			  "(SELECT TOP 1 ISNULL(vMarginPct,0) FROM trx_invePriceSetting WHERE vPartID=i.vPartID) AS vMarginPct,"&_
			  "(SELECT TOP 1 vStsPPN FROM trx_invePriceSetting WHERE vPartID=i.vPartID) AS vStsPPN,"&_
			  "(SELECT TOP 1 ISNULL(bun.vTrxID,'') FROM trx_inveCPBundle AS bun WHERE bun.vPartID=i.vPartID) AS vTrxID,"&_
			  "(SELECT TOP 1 ISNULL(buy_vendor.vVPLFromID,'') FROM trx_InveBuyingVendor AS buy_vendor "&_ 
			  "WHERE buy_vendor.vPartID=i.vPartID) AS vVPLFromID,"&_
			  "(SELECT TOP 1 ISNULL(vendor_main.vName,'') FROM trx_vendorMain AS vendor_main INNER JOIN trx_InveBuyingVendor AS buy_vendor "&_ 
			  "ON vendor_main.vVndID=buy_vendor.vVPLFromID AND buy_vendor.vPartID=i.vPartID) AS vVPLFromName,"&_
			  "(SELECT TOP 1 ISNULL(vCurrValue,0) FROM trx_Currency WHERE vCurrID='CUR01' ORDER BY vCurrDate DESC) AS vValueIDR,"&_
			  "(SELECT TOP 1 ISNULL(vCurrValue,0) FROM trx_Currency WHERE vCurrID='CUR02' ORDER BY vCurrDate DESC) AS vValueUSD,"&_
			  "(SELECT TOP 1 ISNULL(vCurrValue,0) FROM trx_Currency WHERE vCurrID='CUR03' ORDER BY vCurrDate DESC) AS vValueJPY,"&_
			  "(SELECT CASE vNeedSN WHEN 1 THEN 'Need SN' ELSE 'Not Need SN' END FROM tlu_InveCPCategory AS cat "&_ 
			  "INNER JOIN trx_InveCPCatRel AS catrel ON cat.vCatID=catrel.vCatID AND catrel.vPartID=i.vPartID) AS vNeedSN,"&_
			  "(SELECT ISNULL(vLongWarranty,0) FROM trx_InveWarranty WHERE vPartID=i.vPartID AND vType=1 AND vModelID='00001') AS vLongWarranty00001,"&_
			  "(SELECT ISNULL(vLongWarranty,0) FROM trx_InveWarranty WHERE vPartID=i.vPartID AND vType=1 AND vModelID='00002') AS vLongWarranty00002 "&_
			  "FROM trx_inveComputer AS i INNER JOIN trx_inveSelling AS s ON i.vPartID=s.vPartID "&_ 
			  "INNER JOIN trx_inveBuying AS buy ON buy.vPartID=i.vPartID AND i.vPartID='"&currPartID&"';"
	getRs.Open xstring,conn,3,1,0
	IF NOT getRs.eof THEN
		Dim Regex
		SET Regex = New RegExp
		allow = False
		symbol = "*"
		currPartID = getRs("vPartID")
		currBrandID = getRs("vBrandID")
		currSeri = getRs("vSeri")
		currDesc = getRs("vDesc")
		currStatusID = getRs("vStatusID")
		currStartStatus = getRs("vStartPeriod")
		currEndStatus = getRs("vEndPeriod")
		currShipWeight = getRs("vShipWeight")
		currManufacturer = getRs("vManufact")
		currActivation = getRs("vActivation")
		currMarketinginfo = getRs("vMarketingInfo")
		currNote = getRs("vNote")
		currJaminanMurah = getRs("vJaminanMurah")
		currAuthorisedWarranty = getRs("vAuthorisedWarranty")
		currExtTxtWarranty = getRs("vExtTxtWarranty")
		crCreatorNo = getRs("vCreatorNo")
		crCreatorDateTime = getRs("vCreatorDateTime")
		crCreatorIP = getRs("vCreatorIP")
		crEditorNo = getRs("vEditorNo")
		crEditorDateTime = getRs("vEditorDateTime")
		crEditorIP = getRs("vEditorIP")
		currBrandName = getRs("vBrandName")
		currCatPrimaryID = getRs("vCatPrimaryID")
		currCatPrimaryName = getRs("vCatPrimaryName")
		currGoodDesc = getRs("vGoodDesc")
		crCreatorName = getRs("creatorName")
		crEditorName = getRs("editorName")
		currSVndID = getRs("vSVndID")
		currSVndName = getRs("vSVndName")
		currVPLPriceID = getRs("vCntnPrcCurrID")
		currVPLPrice = getRs("vCntnPrc")
		currWebPriceID = getRs("vPrcCurrID")
		currWebPrice = getRs("vPrice")
		currSPriceID = getRs("vSPrcCurrID")
		currSPrice = getRs("vSPrice")
		currStartSPrice = getRs("vStartSPrice")
		currEndSPrice = getRs("vEndSPrice")
		currMinPrice = getRs("vMinPrice")
		currMarginPct = getRs("vMarginPct")
		currStsPPN = getRs("vStsPPn")
		currBundleID = getRs("vTrxID")
		currVPLFromID = getRs("vVPLFromID")
		currVPLFromName = getRs("vVPLFromName")
		currValueIDR = getRs("vValueIDR")
		currValueUSD = getRs("vValueUSD")
		currValueJPY = getRs("vValueJPY")
		currNeedSN = getRs("vNeedSN")
		currLongWarranty00001 = getRs("vLongWarranty00001")
		currLongWarranty00002 = getRs("vLongWarranty00002")
		getRs.Close
		IF (currPrcCurrID="CUR01") AND ((currPrice Mod 1)<>0) THEN
			allow = True
			currPrice = FormatNumber(currPrice,2)
		END IF
		IF (currJaminanMurah) THEN
			chkJaminanMurah = "checked"
		END IF
		SELECT CASE currActivation
		CASE 1
			selUnpublished = "selected"
		CASE 2 
			selPublished = "selected"
		CASE 3
			selInActive = "selected"
		END SELECT
		IF (currVPLPriceID<>"") THEN
			basicCurrency = currVPLPriceID
			currencyRate = currVPLPriceID
		ELSE
			basicCurrency = "CUR01"
			currencyRate = "CUR01"
		END IF
		IF (basicCurrency="") OR (basicCurrency="CUR01") THEN 
			basicCurrency = "CUR01"
			currencyRate = "CUR01"
			currRateNow = currValueIDR
		ELSEIF (basicCurrency="CUR02") THEN
			currencyRate = basicCurrency
			symbol = "/"
			currRateNow = currValueUSD
		ELSEIF (basicCurrency="CUR03") THEN
			currencyRate = basicCurrency
			symbol = "/"
			currRateNow = currValueJPY
		END IF
		IF (symbol="*") THEN
			IF (currVPLPriceID<>basicCurrency) THEN
				currVPLPriceID = basicCurrency
				currVPLPrice = currVPLPrice*currRateNow
			END IF
			IF (currPrcCurrID<>basicCurrency) THEN
				currPrcCurrID = basicCurrency
				IF (currPrice>0) THEN
					currPrice = currPrice*currRateNow
				END IF
			END IF
		ELSEIF (symbol="/") then
			IF (currVPLPriceID<>basicCurrency) THEN
				currVPLPriceID = basicCurrency
				currVPLPrice = currVPLPrice/currRateNow
			END IF
			IF (currPrcCurrID<>basicCurrency) THEN
				currPrcCurrID = basicCurrency
				IF (currPrice>0) THEN
					currPrice = currPrice/currRateNow
				END IF
			END IF
		END IF
		Regex.IgnoreCase = False
		Regex.Global = True
		Regex.Pattern = "Merchant"
		Matches = Regex.Test(currExtTxtWarranty)
		IF ((currAuthorisedWarranty=False) AND (Matches=True)) THEN
			isCheckedAuthorisedWarranty = "disabled"
			isCheckedMerchantWarranty = "checked"			
		END IF
		IF (currBundleID<>"") THEN
			adaPromo = 1
		END IF
		IF (inStr(1,currDesc,"<br>")) THEN
			currDescOnly = Left(currDesc,inStr(currDesc,"<br>")-1)
		ELSE
			currDescOnly = currDesc
		END IF 	
		IF (currLongWarranty00001=12) THEN
			rbWar0000112 = ""
			rbWar0000112 = "checked"
		ELSEIF (currLongWarranty00001=36) THEN
			rbWar0000112 = ""
			rbWar0000136 = "checked"
		ELSEIF (currLongWarranty00001>0) THEN
			rbWar0000112 = ""
			rbWar00001WW = "checked"
			valWar00001WW = currLongWarranty00001
		ELSE
			rbWar0000112 = ""
			rbWar00001NN = "checked"
		END IF
		IF (currLongWarranty00002=12) THEN
			rbWar0000212 = ""
			rbWar0000212 = "checked"
		ELSEIF (currLongWarranty00002=36) THEN
			rbWar0000212 = ""
			rbWar0000236 = "checked"
		ELSEIF (currLongWarranty00002>0) THEN
			rbWar0000212 = ""
			rbWar00002WW = "checked"
			valWar00002WW = currLongWarranty00002
		ELSE
			rbWar0000212 = ""
			rbWar00002NN = "checked"
		END IF
		SELECT CASE currStatusID
		CASE 1
			selStatusIDNew = "selected"
		CASE 2
			selStatusIDSales = "selected"
		CASE 3
			selStatusIDHot = "selected"
		CASE 4
			selStatusIDLimited = "selected"
		END SELECT
		SELECT CASE currVPLPriceID
		CASE "CUR01"
			rbVPLPriceIDIDR = "checked"
		CASE "CUR02"
			rbVPLPriceIDUSD = "checked"
		CASE "CUR03"
			rbVPLPriceIDJPY = "checked"
		END SELECT
		IF (currVPLPrice>0 AND currVPLPriceID<>"") THEN
			currVPLPriceTax = currVPLPrice*1.1
		END IF
		IF (currStatusID>0) THEN
			strDisplayPeriode = "style=""display:inline"""
		END IF
		IF (currCopy="Yes") THEN 
			strTitle = "[Profile]>Copy"
		ELSE
			strTitle = "[eProfile]>"&Right(currPartID,8)
		END IF
		SELECT CASE currStsPPN
		CASE 0
			rbPPNInclude = ""
			rbPPNExclude = "checked"
		END SELECT
		IF (currWebPrice<0) THEN
			rbCall = "checked"
			disabledOOS = "disabled"
		ELSEIF (currWebPrice=0) THEN
			rbOOS = "checked"
			disabledCall = "disabled"
		END IF
		IF (currWebPrice<=0) THEN
			disabledExclude = "disabled"
			disabledMarginPct = "readonly"
			disabledMarginValue = "readonly"
			disabledTdWebPriceID = "disabled"
			disabledWebPrice = "disabled"
			disabledWebPriceTax = "disabled"	
			disabledTdSPriceID = "disabled"
			disabledSPrice = "disabled"
			disabledSPriceTax = "disabled"
			disabledValidStart = "disabled"
			disabledValidEnd = "disabled"
			disabledSPriceStartCalendar = "disabled"
			disabledSPriceEndCalendar = "disabled"
			disabledSamePeriod = "disabled"
		END IF
		IF currWebPrice>currVPLPrice THEN
			currMarginValue = currWebPrice-currVPLPrice
		END IF
		SELECT CASE currWebPriceID
		CASE "CUR01"
			rbWebPriceIDIDR = "checked"
		CASE "CUR02"
			rbWebPriceIDUSD = "checked"
		CASE "CUR03"
			rbWebPriceIDJPY = "checked"
		END SELECT
		IF (currWebPrice>0 AND currWebPriceID<>"" AND currStsPPN=1) THEN
			currWebPriceTax = currWebPrice*1.1
		ELSE
			currWebPriceTax = currWebPrice
		END IF
		SELECT CASE currSPriceID
		CASE "CUR01"
			rbSPriceIDIDR = "checked"
		CASE "CUR02"
			rbSPriceIDUSD = "checked"
		CASE "CUR03"
			rbSPriceIDJPY = "checked"
		END SELECT
		IF (currSPrice>0 AND currSPrice<>"") THEN
			currSPriceTax = currSPrice*1.1
		END IF
		IF (currCopy="Yes") THEN
			displayHeader = "style=""display:none"""
		END IF
	END IF
END IF
%>
<HEAD>
	<TITLE>
		<%=strTitle%>
    </TITLE>
	<LINK href="../css/gnr_all.css" rel=STYLESHEET type="text/css">
	<LINK href="../css/style.css" rel=STYLESHEET type="text/css">
    <SCRIPT type="text/javascript" src="../js/gjs_pupdate.js"></SCRIPT>
   	<SCRIPT type="text/javascript" src="../js/gjs_global.js"></SCRIPT>
   	<SCRIPT type="text/javascript" src="../js/gjs_formatcurrency.js"></SCRIPT>
   	<SCRIPT type="text/javascript" src="../js/gjs_validate.js"></SCRIPT>	
</HEAD>
<body leftmargin=0 topmargin=0 bgcolor="#25519A" style="text-align: center">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000"><img border="0" src="../image/spacer_tr.gif" width="995" height="1"></td></tr>
	<tr height=10><td></td></tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td>&nbsp;</td></tr>
    <tr>
   		<td>
        	<!--#include file="../include/glob_tab_top.asp"-->
		    <%
			Response.Write(inveprodcatalog_menu("profile")) 
			%>
            <!--#include file="../include/glob_tab_mid.asp"-->
			<br />
            <form name="frm" method="post" action="digoff_inve_prodcatalog_save.asp" onSubmit="javascript:return fjs_chkdt();">
			<table border=0 width="95%" align=center cellpadding=2 cellspacing=2>
				<tr <%=displayHeader%>>
					<td style="border:solid 1 black">
						&nbsp;<font face=verdana,arial size=2><strong><%=currPartID%></strong></font>
					</td>
					<td style="border:solid 1 black" align=right>
						&nbsp;<font face=verdana,arial size=2><strong><%=currBrandName+" "+currSeri%></strong></font>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
                <tr>
					<td colspan=2 align=right class=""wordfield"" >
						<font size=2><strong>Go To?</strong>
                			[&nbsp;<a href="digoff_inve_prodcatalog_list.asp?CrTypePage=2&crBhs=2&crLoad=1">List</a>&nbsp;]
                			&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;
                			[&nbsp;<a href="digoff_inve_prodcatalog.asp?crAct=add&crLoad=1">Add</a>&nbsp;]
                			&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;
                			[&nbsp;<a href="digoff_inve_prodcatalog_upd.asp?crAct=add&crLoad=1">Update</a>&nbsp;]
                			&nbsp;&nbsp;
						</font>
            		</td>
                </tr>
       			<tr><td>&nbsp;</td></tr>
				<tr>
					<td colspan="2" align="center">
                        <img STYLE="cursor:hand;" onClick="fjs_cancel();" <%=xMouseCancel%> border="0" alt="Cancel" id="Cancel" name="Cancel" dynamicanimation="Cancel">
                        <img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='save';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSave%> border="0" alt="Save" id="Save" name="Save" dynamicanimation="Save">
						<img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='savenext';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSaveNext%> border="0" alt="SaveNext" id="SaveNext" name="SaveNext" dynamicanimation="SaveNext">
						<img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='savenew';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSaveNew%> border="0" alt="SaveNew" id="SaveNew" name="SaveNew" dynamicanimation="SaveNew">
                        <br>
						<br>
					</td>
				</tr>
				<tr>
					<td class="headerdata" height="20" align="center" colspan="3">
						Product Catalogue
					</td>
				</tr>		
				<tr>
					<td rowspan="2" width="50%" valign="top"  style="border:1px solid <%=xWarnaLine%>">
						<table width="100%">
                        	<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%">
									<u>A</u>ctivation
								</td>
								<td class="wordfield" width="1%">
									:
								</td>
								<td class="wordfield">
									<select id="crActivation" name="crActivation" class="wordfieldnormalmust" accesskey="a" style="border:1 solid #0A4F9A;border-width:1;">
                                    	<option value="">--Choose--</option>
                                        <option value="2" <%=selPublished%> >Active & publish</option>
                                        <option value="1" <%=selUnpublished%> >Active & Unpublish</option>
                                        <option value="3" <%=selInActive%> >Inactive & Deleted</option>
                                    </select>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									<u>C</u>ategory Primary
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
                                	 <div id="divCategory">
                         				<input id="crCatPrimaryID" name="crCatPrimaryID" type="hidden" size="48" value="<%=currCatPrimaryID%>" class="wordfieldnormalmust" readonly />
                                        <input id="crCatPrimaryNameOri" name="crCatPrimaryNameOri" type="text" size="48" value="<%=currCatPrimaryName%>" onClick="fjs_clkCat('<%=currCatPrimaryID%>','<%=currNeedSN%>')" class="wordfieldnormalmust" readonly />
                                     	<img id="btn_clkCat" name="btn_clkCat" border="0" align="top" src="../image/btnscrollme.gif" onClick="fjs_clkCat('<%=currCatPrimaryID%>','<%=currNeedSN%>')" />	
                                    	<span id="needSNInfo" style="display:inline"><%=currNeedSN%></span>
                                     </div>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									<u>B</u>rand
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
                                	 <div id="divBrand">
                         				<input id="crBrandID" name="crBrandID" type="hidden" size="19" value="<%=currBrandID%>" class="wordfieldnormalmust" readonly />
                                        <input id="crBrandNameOri" name="crBrandNameOri" type="text" size="19" value="<%=currBrandName%>" onClick="fjs_clkBrand('<%=currBrandID%>')" class="wordfieldnormalmust" readonly />
                                    	<img id="btn_clkBrand" name="btn_clkBrand" border="0" align="top" src="../image/btnscrollme.gif" onClick="fjs_clkBrand('<%=currBrandID%>')" />
                                     </div>
								</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top"></td>
								<td class="wordfield" width="1%"  valign="top"></td>
								<td class="wordfield">
                                    <nobr><input type="text" value="" name="crBrandNewCode" id="crBrandNewCode" size="3" maxlength="3" class="wordfield" disabled onChange="chk_brandCode();" onKeyPress="return gjs_string(event);" />
                                    &nbsp;/&nbsp;
                                    <input id="crBrandNewName" name="crBrandNewName" type="text" value=""   size="27" maxlength="27" class="wordfield" disabled onChange="chk_brandName();" onKeyPress="return gjs_string(event);" />
									<nobr><input id="cbNewBrand" name="cbNewBrand" type="checkbox" onClick="fjs_newBrand();">Add New Brand
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									<u>S</u>eri
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="crSeri" name="crSeri" type="text" class="wordfieldnormalmust" value="<%=escapeVal(currSeri)%>" accesskey="s" size="51" maxlength="100" />
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									<u>D</u>escription
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<textarea id="crDescOnly" name="crDescOnly" cols="38" rows="3" class="wordfieldnormalmust" accesskey="d" onChange="javascript:fjs_descAll();"><%=escapeVal(currDescOnly)%></textarea>
									<textarea id="crDesc" name="crDesc" cols="38" rows="3" class="wordfieldnormalmust" style="display:none"><%=escapeVal(currDesc)%></textarea>
                                    <img id="btn_Description" name="btn_Description" <%=xMousePreview%> style="vertical-align:top;cursor:hand;" alt="Description" onClick="fjs_pupDescription();"/>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Ship<u>W</u>eight
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="crShipWeight" name="crShipWeight" type="text" size="3" maxlength="3" class="wordfieldnormalmust" value="<%=currShipWeight%>" accesskey="w" onKeyPress="return gjs_bilangan(event.KeyCode)" style="text-align:right" />&nbsp;Kg
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
                            	<td width="5">&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Warranty
								</td>
								<td class="wordfield" width="1%"valign="top">
									:
								</td>
								<td class="wordfield">					
									<input id="crAuthWarranty" name="crAuthWarranty" type="checkbox" onClick="fjs_Warranty()" <%=isCheckedAuthorisedWarranty%>>Authorised<nobr>
									<input id="crMchWarranty" name="crMchWarranty" type="checkbox" onClick="fjs_Warranty()" <%=isCheckedMerchantWarranty%>>Merchant<nobr>
								</td>
								<td valign="bottom">&nbsp;</td>
                            </tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" valign="top" width="25%">
									Customer Warranty
								</td>
								<td class="wordfield" valign="top" width="1%">
									:
								</td>
								<td class="wordfield">
                                	<table width="100%" border="0" cellpadding="0" cellspacing="0">
                                    	<tr>
											<td class="wordfield" valign="top" style="border-left:1 solid silver;border-top:1 solid silver;border-right:1 solid silver;">
												&nbsp;&nbsp;Hardware:
											</td>
										</tr>
                                        <tr>
											<td class="wordfield" style="border-left:1 solid silver;border-bottom:1 solid silver;border-right:1 solid silver;">
												<table> 	
													<tr>
														<td class="wordfield">
															<input id="crTypeCustWar00001" name="crTypeCustWar00001" type="radio" value="N" <%=rbWar00001NN%> />No Warranty
														</td>		
														<td class="wordfield">
															<input id="crTypeCustWar00001" name="crTypeCustWar00001" type="radio" value="12" <%=rbWar0000112%> />12 Month
														</td>		
														<td class="wordfield">
															<input id="crTypeCustWar00001" name="crTypeCustWar00001" type="radio" value="36" <%=rbWar0000136%> />36 Month
														</td>
													</tr>	
													<tr>		
														<td class="wordfield">
															<input id="crTypeCustWar00001" name="crTypeCustWar00001" type="radio" value="L" <%=rbWar00001LL%> />Life Time
														</td>		
														<td class="wordfield" colspan="2" >
															<input id="crTypeCustWar00001" name="crTypeCustWar00001" type="radio" value="W" <%=rbWar00001WW%> />Others 
															<input id="crCustLengthWarranty00001" name="crCustLengthWarranty00001" type="text" class="wordfieldnormal" size="5" maxlength="2" onKeyPress="return gjs_bilangan(event.KeyCode)" value="<%=valWar00001WW%>" />Month
														</td>	
													</tr>
												</table>
											</td>
										</tr>
                                        <tr height="5">
											<td></td>
										</tr>
                                        <tr>
											<td class="wordfield" valign="top" style="border-left:1 solid silver;border-top:1 solid silver;border-right:1 solid silver;">
                                            	&nbsp;&nbsp;Service:
											</td>
										</tr>
                                        <tr>
											<td class="wordfield" style="border-left:1 solid silver;border-bottom:1 solid silver;border-right:1 solid silver;">
												<table> 	
													<tr>		
														<td class="wordfield">
                                                        	<input id="crTypeCustWar00002" name="crTypeCustWar00002" type="radio" value="N" <%=rbWar00002NN%> />No Warranty
														</td>		
														<td class="wordfield">
                                                        	<input type="radio" name="crTypeCustWar00002" value="12" <%=rbWar0000212%> />12 Month
														</td>		
														<td class="wordfield">
                                                        	<input id="crTypeCustWar00002" name="crTypeCustWar00002" type="radio" value="36" <%=rbWar0000236%> />36 Month
														</td>	
													</tr>	
													<tr>		
														<td class="wordfield">
															<input id="crTypeCustWar00002" name="crTypeCustWar00002" type="radio" value="L" <%=rbWar00002LL%> />Life Time
														</td>		
														<td class="wordfield" colspan="2" >
															<input id="crTypeCustWar00002" name="crTypeCustWar00002" type="radio" value="W" <%=rbWar00002WW%> />Others 
															<input id="crCustLengthWarranty00002" name="crCustLengthWarranty00002" type="text" class="wordfieldnormal" size="5" maxlength="2" onKeyPress="return gjs_bilangan(event.KeyCode)" value="<%=valWar00002WW%>" />Month		
														</td>	
													</tr>
												</table>
											</td>
										</tr>
                                        <tr height="5">
											<td></td>
										</tr>
                                    </table>
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td width="5">&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Extended Text Warranty
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="crExtTxtWarr" name="crExtTxtWarr" type="text" class="wordfieldnormal" value="<%=escapeVal(currExtTxtWarranty)%>" size="50" maxlength="50" />
									<input id="crExtTxtWarrTemp" name="crExtTxtWarrTemp" type="text" value="<%=escapeVal(currExtTxtWarranty)%>" style="display:none" size="50" maxlength="50" />
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									<u>M</u>arketing Info
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<textarea id="crMarketingInfo" name="crMarketinginfo" cols="38" rows="3" class="wordfieldnormal" accesskey="m" readonly><%=escapeVal(currMarketinginfo)%></textarea>
									<img id="btnMarketingInfo" name="btnMarketingInfo" <%=xMouseEdit%> style="vertical-align:top;cursor:hand;" alt="Edit Marketing Word" onClick="fjs_pupMarketing()"/>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									N<u>o</u>te
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<textarea id="crNote" name="crNote" cols="38" rows="3" class="wordfieldnormal" accesskey="o"><%=escapeVal(currNote)%></textarea>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Manu<u>f</u>acturer
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									http://<input id="crManufacturer" name="crManufacturer" type="text" class="wordfieldnormal" value="<%=escapeVal(currManufacturer)%>" accesskey="f" size="33" maxlength="100" />
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Status
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
								  	<select id="crStatusID" name="crStatusID" class="wordfieldnormal"  onchange="fjs_displayDate()" style="border:1 solid #0A4F9A;border-width: 1;">
                                    	<option value="">--Not Select--</option>
                                        <option value="1" <%=selStatusIDNew%>>New</option>
                                    	<option value="2" <%=selStatusIDSales%>>Sales</option>
                                    	<option value="3" <%=selStatusIDHot%>>Hot</option>
                                    	<option value="4" <%=selStatusIDLimited%>>Limited Stock</option>
                                    </select>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr id="tr_periode" name="tr_periode" height="25" <%=strDisplayPeriode%>>
								<td width="5">&nbsp;
								</td>
								<td class="wordfield" valign="top" width="25%">&nbsp;
								</td>
								<td class="wordfield" valign="top" width="1%">&nbsp;
								</td>
								<td class="wordfield">
									<table>
										<tr>
											<td class="wordfield">From</td>
											<td class="wordfield">
                                            	<table cellpadding="0" cellspacing="0" border="0">
													<tr height="25">
														<td valign="top" style="border:0px">
															<input id="CrPeriodStart" name="CrPeriodStart" type="hidden" value="<%=FormatDateTime(currStartStatus,2)%>" class="wordfieldnormal" readonly />
															<input id="CrPeriodStartView" name="CrPeriodStartView" type="text" value="<%=FormatDateView(currStartStatus)%>" class="wordfieldnormal" size="15" maxlength="20" readonly />
														</td>
														<td valign="top" style="border:0px;padding-left:5px" style="cursor:hand;">
															<a onMouseOver="document['PeriodStart'].imgRolln=document['PeriodStart'].src;document['PeriodStart'].src=document['PeriodStart'].lowsrc;" onMouseOut="document['PeriodStart'].src=document['PeriodStart'].imgRolln" onClick="javascript:getCalendarFor('model1','0','fjs_setlastupdateSpecialPrice()','<%=now()%>',document.frm.CrPeriodStart,null,document.frm.CrPeriodStart.value,document.frm.CrPeriodStartView,'d s Y',null,null,false)" />
																<img id="PeriodStart" name="PeriodStart" border="0" src="../Image/btn_clock_000.gif" dynamicanimation="PeriodStart" lowsrc="../Image/btn_clock_100.gif" align="middle" >
															</a>
														</td>
													</tr>
												</table>
                                            </td>
										</tr>
										<tr>
											<td class="wordfield">To</td>
											<td class="wordfield">
                                            	<table cellpadding="0" cellspacing="0" border="0">
													<tr height="25">
														<td valign="top" style="border:0px">
															<input id="CrPeriodEnd" name="CrPeriodEnd" type="hidden" value="<%=FormatDateTime(currEndStatus,2)%>" class="wordfieldnormal" readonly />
															<input id="CrPeriodEndView" name="CrPeriodEndView" type="text" size="15" maxlength="15" value="<%=FormatDateView(currEndStatus)%>" class="wordfieldnormal" readonly  />
														</td>
														<td valign=top style='border:0px;padding-left:5px' style="cursor:hand;">
															<a onMouseOver="document['PeriodEnd'].imgRolln=document['PeriodEnd'].src;document['PeriodEnd'].src=document['PeriodEnd'].lowsrc;" onMouseOut="document['PeriodEnd'].src=document['PeriodEnd'].imgRolln" onClick="javascript:getCalendarFor('model1','','','<%=now()%>',document.frm.CrPeriodEnd,null,document.frm.CrPeriodEnd.value,document.frm.CrPeriodEndView,'d s Y',null,null,false)">
																<img border="0" src="../Image/btn_clock_000.gif" id="PeriodEnd" name="PeriodEnd" dynamicanimation="PeriodEnd" lowsrc="../Image/btn_clock_100.gif" align="middle" >
															</a>
														</td>
													</tr>
												</table>
                                           </td>
										</tr>
									</table>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" valign="top" width="25%">
									Jaminan Murah
								</td>
								<td class="wordfield" valign="top" width="1%">
									:
								</td>
								<td class="wordfield" valign="top">
									<input id="crJaminanMurah" name="crJaminanMurah" type="checkbox" <%=chkJaminanMurah%> />
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>									
               		 </table>
             	</td>
					<td width="100%" valign="top" style="border: 1px solid <%=xWarnaLine%>">
						<table width="100%" >
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									VPL From
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="crPVndID" name="crPVndID" type="hidden" class="wordfieldnormal" value="<%=escapeVal(currVPLFromID)%>" size="20" maxlength="15" readonly />
									<input id="crPVndName" name="crPVndName" type="text" class="wordfieldnormalmust" value="<%=escapeVal(currVPLFromName)%>" size="20" maxlength="15" readonly />
									<img id="FindPVnd" name="FindPVnd" align="absmiddle" STYLE="cursor:hand;" <%=xMouseFind%> border="0" alt="FindPVnd" dynamicanimation="FindPVnd" onClick="fjs_getPVnd();">
									<img id="DelPVnd" name="DelPVnd" align="absmiddle" STYLE="cursor:hand;" <%=xMouseRemove%> border="0" alt="DelPVnd" dynamicanimation="DelPVnd" onClick="fjs_delPVnd();">
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									VPL Before Tax
                                </td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
                                	<table class="wordfield">
                                    	<tr>
                                			<td id="trCurrencyVPL" name="trCurrencyVPL" class="wordfieldnormalmust" valign="top" style="border-left:1 solid Silver;border-top:1 solid Silver;border-right:1 solid Silver; border-bottom:1 solid silver;">Currency :<br />
												<input id="crVPLPriceIDIDR" name="crVPLPriceID" type="radio" value="CUR01" onClick="fjs_VPLPriceID(1);fjs_setlastupdateprcVPL();" <%=rbVPLPriceIDIDR%> />IDR
                                        		<br /> 
                                        		<input id="crVPLPriceIDUSD" name="crVPLPriceID" type="radio" value="CUR02" onClick="fjs_VPLPriceID(2);fjs_setlastupdateprcVPL();" <%=rbVPLPriceIDUSD%> />USD                                            
                                        		<br />                                                                      
                                       			<input id="crVPLPriceIDJPY" name="crVPLPriceID" type="radio" value="CUR03" onClick="fjs_VPLPriceID(3);fjs_setlastupdateprcVPL();" <%=rbVPLPriceIDJPY%> />JPY
                                        		<br />
                                    		</td>
                                    		<td>
                                    			<input id="crVPLPrice" name="crVPLPrice" type="text" class="wordfieldnormal" value="<%=FormatNumber(currVPLPrice,2)%>" style="text-align:right" size="20" maxlength="15" readonly />
                                    			<input id="hidVPLPrice" name="hidVPLPrice" type="text" class="wordfieldnormal" value="<%=FormatNumberCustom(currVPLPrice,6)%>" style="text-align:right" size="20" maxlength="15" readonly />
												<input id="hidVPLPriceID" name="hidVPLPriceID" type="text" class="wordfieldnormal" value="<%=currVPLPriceID%>" size="5" maxlength="5" readonly />
                                            </td>
                                    	</tr>
                                    </table>
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr> 
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									VPL After Tax
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
                                <td class="wordfield">
                                    <input id="crVPLPriceTax" name="crVPLPriceTax" type="text" class="wordfieldnormalmust" value="<%=FormatNumberRoundPrice(currVPLPriceTax,currVPLPriceID,2)%>" accesskey="i" onKeyPress="return gjs_bilangan(event.KeyCode)" style="text-align:right" onChange="fjs_updateVPLPriceTax();fjs_setlastupdateprcVPL();" size="20" maxlength="15" />
                                    <input id="hidVPLPriceTax" name="hidVPLPriceTax" type="text" class="wordfieldnormal" value="<%=FormatNumberCustom(currVPLPriceTax,6)%>" style="text-align:right" size="20" maxlength="15" readonly />
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									PPN
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="radInclude" name="radPPN" type="radio" value="1" onClick="fjs_updatePPNVPL(1);fjs_updatePPNWeb(1);fjs_setlastupdateWebPrice();" <%=rbPPNInclude%> <%''=disabledInclude%>/><span style="width:30%">Include</span>
                                    <input id="radExclude" name="radPPN" type="radio" value="0" onClick="fjs_updatePPNVPL(0);fjs_updatePPNWeb(0);fjs_setlastupdateWebPrice();" <%=rbPPNExclude%> <%''=disabledExclude%>/><span style="width:40%">Exclude</span>
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield">
									Web Price Exception
								</td>
								<td class="wordfield">
									:
								</td>
								<td class="wordfield">
                                    <input id="chkCall" name="chkCall" type="checkbox" onClick="fjs_Call(this);fjs_setlastupdateWebPrice();fjs_setlastupdateSpecialPrice();" <%=rbCall%> <%=disabledCall%>/><span style="width:30%">Call</span>
                                    <input id="chkOOS" name="chkOOS" type="checkbox"  onClick="fjs_OOS(this);fjs_setlastupdateWebPrice();fjs_setlastupdateSpecialPrice();" <%=rbOOS %> <%=disabledOOS%>/><span style="width:40%">Out Of Stock</span>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">&nbsp;</td>
								<td class="wordfield" width="1%" valign="top">&nbsp;</td>
								<td class="wordfield" >
									<font face="Georgia, Times New Roman, Times, serif">Jika kolom <font color="red">PPN</font> 
                                    adalah <font color="red">Include</font>, maka <font color="red">Harga Jual termasuk 10% PPN</font></font>&nbsp;
								</td>
							</tr>                          
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Margin
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
									<input id="crMarginPct" name="crMarginPct" type="text" class="wordfieldnormalmust" style="text-align:right" value="<%=FormatNumber(currMarginPct,2)%>" onChange="fjs_marginPct();fjs_setlastupdateWebPrice();" onKeyPress="return gjs_bilangan(event.KeyCode)" size="5" maxlength="5" <%=disabledMarginPct%> /><span style="width:24%">Percentage(%)</span>
									<input id="crMarginValue" name="crMarginValue" type="text" class="wordfieldnormalmust" value="<%=FormatNumber(currMarginValue,2)%>" style="text-align:right" onChange="fjs_marginValue();fjs_setlastupdateWebPrice();" onKeyPress="return gjs_bilangan(event.KeyCode)" size="20" maxlength="15" <%=disabledMarginValue%> /><span style="width:20%">Value</span>
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Web Pr<u>i</u>ce Before Tax
                                </td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
                                    <table>
                                    	<tr>
                               				<td id="tdWebPriceID" name="tdWebPriceID" class="wordfieldnormalmust" valign="top" style="border-left: 1 solid Silver;border-top: 1 solid Silver;border-right: 1 solid Silver; border-bottom: 1 solid silver;" <%=disabledTdWebPriceID%>>Currency :<br>
												<input id="crWebPriceIDIDR" name="crWebPriceID" type="radio" onClick="fjs_WebPriceID(1);fjs_setlastupdateWebPrice();" value="CUR01" <%=rbWebPriceIDIDR%> />IDR
                                                <br />                                                                                                                
                                            	<input id="crWebPriceIDUSD" name="crWebPriceID" type="radio" onClick="fjs_WebPriceID(2);fjs_setlastupdateWebPrice();" value="CUR02" <%=rbWebPriceIDUSD%> />USD
                                                <br />
                                                <input id="crWebPriceIDJPY" name="crWebPriceID" type="radio" onClick="fjs_WebPriceID(3);fjs_setlastupdateWebPrice();" value="CUR03" <%=rbWebPriceIDJPY%> />JPY
                                                <br />
                                         	</td>
                                            <td class="wordfield">	
                                            	<input id="crWebPrice" name="crWebPrice" type="text" class="wordfieldnormal" value="<%=FormatNumber(currWebPrice)%>" onFocus="this.select();"  tabindex=15 accesskey="i" onKeyPress="return gjs_bilangan(event.KeyCode)" style="text-align:right" onChange="javascript:fjs_chgWebPrice(this);" size="20" readonly <%=disabledWebPrice%> />
                                                <input id="hidWebPrice" name="hidWebPrice" type="text" class="wordfieldnormal" value="<%=FormatNumberCustom(currWebPrice,6)%>" size="20" style="text-align:right" maxlength="15" readonly />
                                                <input id="hidWebPriceID" name="hidWebPriceID" type="text" class="wordfieldnormal" value="<%=currWebPriceID%>" size="5" maxlength="5" readonly />
                                            </td>
                                      	</tr>
                              		</table>
								</td>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Web Price After Tax
                                </td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">	
                                	<input id="crWebPriceTax" name="crWebPriceTax" type="text" class="wordfieldnormalmust" onChange="fjs_updateWebPriceTax();fjs_setlastupdateWebPrice();" value="<%=FormatNumberRoundPrice(currWebPriceTax,currWebPriceID,2)%>" style="text-align:right" size="20" maxlength="15" <%=disabledWebPriceTax%> />
									<input id="hidWebPriceTax" name="hidWebPriceTax" type="text" class="wordfieldnormal" value="<%=FormatNumberCustom(currWebPriceTax,6)%>" style="text-align:right" size="20" maxlength="15" readonly />
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Sp<u>e</u>cial Price Before Tax
								</td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
							  	<td class="wordfield">
                                    <table>
                                    	<tr>
                               				<td id="tdSPriceID" name="tdSPriceID" class="wordfieldnormalmust" valign="top" style="border-left: 1 solid Silver;border-top: 1 solid Silver;border-right: 1 solid Silver; border-bottom: 1 solid silver;" <%=disabledTdSPriceID%>>Currency :<br>
												<input id="crSPriceIDIDR" name="crSPriceID" type="radio" value="CUR01" onClick="fjs_SPriceID(1);fjs_setlastupdateSpecialPrice();" <%=rbSPriceIDIDR%> />IDR
                                                <br />                                                                                      
                                                <input id="crSPriceIDUSD" name="crSPriceID" type="radio" value="CUR02" onClick="fjs_SPriceID(2);fjs_setlastupdateSpecialPrice();" <%=rbSPriceIDUSD%> />USD
                                                <br />
                                                <input id="crSPriceIDJPY" name="crSPriceID" type="radio" value="CUR03" onClick="fjs_SPriceID(3);fjs_setlastupdateSpecialPrice();" <%=rbSPriceIDJPY%> />JPY
                                                <br />
                                         	</td>
                                            <td>
                                    			<input id="crSPrice" name="crSPrice" type="text" value="<%=FormatNumber(currSPrice)%>" class="wordfieldnormal" accesskey="e" onKeyPress="return gjs_bilangan(event.KeyCode)" style="text-align:right" size="20" maxlength="15" readonly <%=disabledSPrice%> />
												<input id="hidSPrice" name="hidSPrice" type="hidden" class="wordfieldnormal" value="<%=FormatNumberCustom(currSPrice,4)%>" size="20" style="text-align:right"s maxlength="15" readonly />
                                                <input id="hidSPriceID" name="hidSPriceID" type="hidden" value="<%=currSPriceID%>" class="wordfieldnormal" size="5" maxlength="5" readonly />                                             
                                            </td>
                                     	</tr>
                                	</table>
								<td valign="bottom">&nbsp;</td>
							</tr>
                            <tr height="25">
								<td>&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">
									Special Price After Tax
                                </td>
								<td class="wordfield" width="1%" valign="top">
									:
								</td>
								<td class="wordfield">
                                   	<input id="crSPriceTax" name="crSPriceTax" type="text" class="wordfieldnormalmust" onFocus="this.select();" onChange="fjs_updateSPriceTax();fjs_setlastupdateSpecialPrice();" value="<%=FormatNumberRoundPrice(currSPriceTax,currSPriceID,2)%>" accesskey="f" onKeyPress="return gjs_bilangan(event.KeyCode)" style="text-align:right" size="20" maxlength="15" <%=disabledSPriceTax%>>
                 					<input id="hidSPriceTax" name="hidSPriceTax" type="hidden" class="wordfieldnormal" value="<%=FormatNumberCustom(currSPriceTax,4)%>" style="text-align:right" size="20" maxlength="15" readonly />
                                </td>
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td width="5">&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">&nbsp;
								</td>
								<td class="wordfield" width="1%" valign="top">&nbsp;
								</td>
								<td class="wordfield">
									<table>
										<tr>
											<td class="wordfield">From</td>
											<td class="wordfield">
                                            	<table cellpadding="0" cellspacing="0" border="0">
													<tr height="25">
														<td valign="top" style="border:0px">
															<input id="CrSPValidStart" name="CrSPValidStart" type="hidden" value="<%=FormatDateTime(currStartSPrice,2)%>" class="wordfieldnormal" size="15" maxlength="15" readonly />
															<input id="CrSPValidStartView" name="CrSPValidStartView" type="text" value="<%=FormatDateView(currStartSPrice)%>" class="wordfieldnormal" size="15" maxlength="15" readonly <%=disabledValidStart%> />
														</td>
														<td id="tdSPriceStartCalendar" name="tdSPriceStartCalendar" valign="top" style="border:0px;padding-left:5px" style="cursor:hand;" <%=disabledSPriceStartCalendar%>>
															<a onMouseOver="document['SPValidStart'].imgRolln=document['SPValidStart'].src;document['SPValidStart'].src=document['SPValidStart'].lowsrc;" onMouseOut="document['SPValidStart'].src=document['SPValidStart'].imgRolln" onClick="javascript:getCalendarFor('model1','0','fjs_setlastupdateSpecialPrice()','<%=now()%>',document.frm.CrSPValidStart,null,document.frm.CrSPValidStart.value,document.frm.CrSPValidStartView,'d s Y',null,null,false)" />
																<img id="SPValidStart" name="SPValidStart" border="0" src="../Image/btn_clock_000.gif" dynamicanimation="SPValidStart" lowsrc="../Image/btn_clock_100.gif" align="middle" >
															</a>
														</td>
													</tr>
												</table>
                                			</td>
                                        </tr>
                                        <tr>
											<td class="wordfield">To</td>
											<td class="wordfield">
                                            	<table cellpadding="0" cellspacing="0" border="0">
													<tr height="25">
														<td valign="top" style="border:0px">
															<input id="CrSPValidEnd" name="CrSPValidEnd" type="hidden" value="<%=FormatDateTime(currEndSPrice,2)%>" class="wordfieldnormal" size="15" maxlength="15" readonly />
															<input id="CrSPValidEndView" name="CrSPValidEndView" type="text" size="15" maxlength="15" value="<%=FormatDateView(currEndSPrice)%>" class="wordfieldnormal" readonly <%=disabledValidEnd%> />
														</td>
														<td id="tdSPriceEndCalendar" name="tdSPriceEndCalendar" valign="top" style="border:0px;padding-left:5px" style="cursor:hand;" <%=disabledSPriceEndCalendar%> >
															<a onMouseOver="document['SPValidEnd'].imgRolln=document['SPValidEnd'].src;document['SPValidEnd'].src=document['SPValidEnd'].lowsrc;" onMouseOut="document['SPValidEnd'].src=document['SPValidEnd'].imgRolln" onClick="javascript:getCalendarFor('model1','0','fjs_setlastupdateSpecialPrice()','<%=now()%>',document.frm.CrSPValidEnd,null,document.frm.CrSPValidEnd.value,document.frm.CrSPValidEndView,'d s Y',null,null,false)">
																<img id="SPValidEnd" name="SPValidEnd" border="0" src="../Image/btn_clock_000.gif" dynamicanimation="SPValidEnd" lowsrc="../Image/btn_clock_100.gif" align="middle" />
															</a>
														</td>
													</tr>
												</table>
                                        	</td>
										</tr>
                                    </table>		
								<td valign="bottom">&nbsp;</td>
							</tr>
							<tr height="25">
								<td width="5">&nbsp;</td>
								<td class="wordfield" width="25%" valign="top">&nbsp;
								</td>
								<td class="wordfield" width="1%" valign="top">&nbsp;
								</td>
								<td class="wordfield" colspan="2">
									<input id="crChkStsSameSPrice" name="crChkStsSameSPrice" type="checkbox" value="1" accesskey="o" onClick="javascript:fjs_samePrdSPrice(this);fjs_setlastupdateSpecialPrice()" <%=disabledSamePeriod%> />Same with Status Period
								</td>
							</tr>
						</table>
					</td>
				</tr>				
			</table>
			<table height="5px">
            	<tr>
                	<td>&nbsp;</td>
                </tr>
            </table>
            <table border=0 width="95%" align=center cellpadding=2 cellspacing=2>
				<tr>
					<td class="headerdata" height="20" align="center">
						Bhinneka Promo
					</td>
				</tr>
                <tr>
					<td valign=top  style="border: 1px solid <%=xWarnaLine%>">
                    	<!-- BP->Bhinneka Promo -->
                		<table border="0" width="100%" cellspacing="0" cellpadding="0">
                            <tr>
                                <td class="wordfield">
                                    <table border="0" width="100%" id="tableBhinnekaPromo" cellspacing="0" cellpadding="0" >
                                        <tr style="border: 1px solid <%=xWarnaLine%>" id="trHeaderBP">
                                            <td <%=bgnoact%> width="10%" align="center" title="Type"><font class="wordblackmedium"><nobr><strong>Type</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="10%" align="center" title="SKU"><font class="wordblackmedium"><nobr><strong>SKU</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="40%" align="center" title="ItemDesc"><font class="wordblackmedium"><nobr><strong>Item Description</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="3%" align="center" title="Qty"><font class="wordblackmedium"><nobr><strong>Qty</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="12%" align="center" title="StartDate"><font class="wordblackmedium"><nobr><strong>Start Date</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="12%" align="center" title="EndDate"><font class="wordblackmedium"><nobr><strong>End Date</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="3%" align="center" title="TagDate"><font class="wordblackmedium"><nobr><strong>Tag</strong></nobr></font></td>
											<td <%=bgnoact%> width="10%" align="center" title="Action"><font class="wordblackmedium"><nobr><strong>Action</strong></nobr></font></td>
                                        </tr>
									<%
										if adaPromo = 1 then
											set rsPromoBP = Server.CreateObject("ADODB.Recordset")
											queryBP = "select * from trx_InveCPBundleDetail where vTrxID = '"&currBundleID&"' and vPromo = 0 order by vNoUrut ASC "
											rsPromoBP.open queryBP, conn, 3, 1, 0
														
											for bp = 0 to rsPromoBP.RecordCount - 1
											currBundlePartID = rsPromoBP("vBundlePartID")
									%>
											<tr id="trBP_<%=bp%>" name="trBP">
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                    <select name="crTypeBP" id="crTypeBP_<%=bp%>" disabled>
                                                        <option value="0" <%if rsPromoBP("vType") = 0 then Response.Write("selected") end if%> >Free</option>
                                                        <option value="1" <%if rsPromoBP("vType") = 1 then Response.Write("selected") end if%> >Bundle</option>
                                                    </select> 
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                                    <input type="text" name="crSKUBP" id="crSKUBP_<%=bp%>" size="12" onClick="this.select();" onChange="checkSKUBP(this);" onKeyPress="disableEnter()" value="<%=currBundlePartID%>" disabled>
                                                    <input type=hidden name="hidSKUBP" id="hidSKUBP_<%=bp%>" value="<%=currBundlePartID%>" >
                                                </td>
                                                <td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                                	<%
													set rsBrandSeriDesc = server.CreateObject("ADODB.Recordset")
													queryBrandSeriDesc = "select vBrandID, vSeri, vDesc from trx_InveComputer where vPartID = '"&currBundlePartID&"' "
													rsBrandSeriDesc.open queryBrandSeriDesc, conn, 3, 1, 0
													
													if not rsBrandSeriDesc.eof then
														tempBrandID = rsBrandSeriDesc("vBrandID")
														tempSeri = rsBrandSeriDesc("vSeri")
														tempDesc = rsBrandSeriDesc("vDesc")
														
														set rsBrandName = server.CreateObject("ADODB.Recordset")
														queryBrandName = "select vName from tlu_InveCPBrand where vBrandID = '"&tempBrandID&"' "
														rsBrandName.open queryBrandName, conn, 3, 1, 0
														
														if not rsBrandName.eof then
															tempBrandName = rsBrandName("vName")
														end if
													end if
													%>
                                                    <div align="justify" id="divBrandSeriDescBP_<%=bp%>" name="divBrandSeriDescBP"><%=tempBrandName%><br><%=tempSeri%><br><%=tempDesc%></div>
                                                    <input type=hidden id="hidBrandBP_<%=bp%>" name="hidBrandBP" value="<%=tempBrandName%>">
                                                    <input type=hidden id="hidSeriBP_<%=bp%>" name="hidSeriBP" value="<%=tempSeri%>">
                                                    <input type=hidden id="hidDescBP_<%=bp%>" name="hidDescBP" value="<%=tempDesc%>">
													<%
														tempBrandID = ""
														tempSeri = ""
														tempDesc = ""
														tempBrandName = ""
														rsBrandName.close
														rsBrandSeriDesc.close
														set rsBrandName = nothing
														set rsBrandSeriDesc = nothing
													%>                                      
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                                    <input align="center" type="text" name="crQtyBP" id="crQtyBP_<%=bp%>" size=1 maxlength=1 onKeyPress="crCekNumber();" value="<%=rsPromoBP("vQty")%>" disabled>
												</td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanStartDateBP_<%=bp%>" name="spanStartDateBP">
                                                        <input type=hidden id="CrBPValidStart_<%=bp%>" name="CrBPValidStart" value="<%=FormatDateTime(CDate(rsPromoBP("vStartDateTime")),2)%>" >
                                                        <input name="CrBPValidStartView" type="text" size=15 id="CrBPValidStart_<%=bp%>View" class=wordfieldnormaldisabled readonly value="<%=Day(rsPromoBP("vStartDateTime"))&" "&MonthName(Month(rsPromoBP("vStartDateTime")),1) &" "&Year(rsPromoBP("vStartDateTime"))%>" onKeyPress="disableEnter()" >
                                                    </span>
                                                    <span id="spanImgStartDateBP_<%=bp%>" name="spanImgStartDateBP">
                                                        <a id="BPValidStartImg_<%=bp%>" name ="BPValidStartImg" style="cursor:hand;" >
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='BPValidStart_<%=bp%>' name='BPValidStart' align='middle' >
                                                        </a>
                                                    </span>                                                
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanEndDateBP_<%=bp%>" name="spanEndDateBP"> 
                                                        <input type=hidden id="CrBPValidEnd_<%=bp%>" name="CrBPValidEnd" value="<%=FormatDateTime(CDate(rsPromoBP("vEndDateTime")),2)%>">
                                                        <input name="CrBPValidEndView" type="text" size=15 id="CrBPValidEnd_<%=bp%>View" class=wordfieldnormaldisabled readonly value="<%=Day(rsPromoBP("vEndDateTime"))&" "&MonthName(Month(rsPromoBP("vEndDateTime")),1) &" "&Year(rsPromoBP("vEndDateTime"))%>" onKeyPress="disableEnter()">
                                                    </span>
                                                    <span id="spanImgEndDateBP_<%=bp%>" name="spanImgEndDateBP">
                                                        <a id="BPValidEndImg_<%=bp%>" name ="BPValidEndImg" style="cursor:hand;" >                                          																																																								
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='BPValidEnd_<%=bp%>' name='BPValidEnd' align='middle' >
                                                        </a>
                                                    </span>
                                                </td>
												<%
												''add queryTagBP
												set rsTagBP = Server.CreateObject("ADODB.Recordset")
												queryTagBP = "select * from trx_InveCPBundleDetail where vBundlePartID = '"&currBundlePartID&"' and vTrxID = '"&currBundleID&"' and vPromo = 0 order by vNoUrut ASC "
												rsTagBP.open queryTagBP, conn, 3, 1, 0
												IF NOT rsTagBP.eof Then
													cekTagBP = rsTagBP("vTagDate")
												END IF
												rsTagBP.Close
												SET rsTagBP = NOTHING
												%>
												<td align="center" valign="top">   
                                                	<input type="checkbox" id="tagBP_<%=bp%>" name="tagBP" <%IF cekTagBP = TRUE THEN response.write "checked"  END IF%> disabled>                                               	
                                                </td>
                                                <td <%=xLineGrid%> class="wordingrid" align="center" title="Action" valign="top">
                                                    <a id="lockBP_<%=bp%>" name="lockBP" style="visibility:block" class="cursorHand" onClick="javascript:lockBP(this);">Lock</a>
                                                    &nbsp;||&nbsp;
                                                    <a id="deleteBP_<%=bp%>" name="deleteBP" style="visibility:block" class="cursorHand" onClick="javascript:deleteBP(this);">Del</a>
                                                </td>
                                                <td>
                                                	<input type="hidden" id="hidStatusBP_<%=bp%>" name="hidStatusBP" value="<%=rsPromoBP("vStsActive")%>">
												</td>																										
                                            </tr>
									<%
												rsPromoBP.MoveNext
											next
											rsPromoBP.close
											set rsPromoBP = nothing
									%>
											<tr id="trBP_<%=bp%>" name="trBP">
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                    <select name="crTypeBP" id="crTypeBP_<%=bp%>">
                                                        <option value="0" >Free</option>
                                                        <option value="1" >Bundle</option>
                                                    </select> 
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                                    <input type="text" name="crSKUBP" id="crSKUBP_<%=bp%>" size="12" onClick="this.select();" onChange="checkSKUBP(this);" onKeyPress="disableEnter()" >
                                                    <input type="hidden" name="hidSKUBP" id="hidSKUBP_<%=bp%>" >
                                                </td>
                                                <td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                                    <div id="divBrandSeriDescBP_<%=bp%>" align="justify" name="divBrandSeriDescBP"></div>
                                                    <input type="hidden" id="hidBrandBP_<%=bp%>" name="hidBrandBP">
                                                    <input type="hidden" id="hidSeriBP_<%=bp%>" name="hidSeriBP">
                                                    <input type="hidden" id="hidDescBP_<%=bp%>" name="hidDescBP">                                                
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                                    <input type="text" name="crQtyBP" id="crQtyBP_<%=bp%>" size=1 maxlength=1 onKeyPress="crCekNumber();" >
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanStartDateBP_<%=bp%>" name="spanStartDateBP">
                                                        <input type=hidden id="CrBPValidStart_<%=bp%>" name="CrBPValidStart" >
                                                        <input type=text size=15 id="CrBPValidStart_<%=bp%>View" class=wordfieldnormaldisabled readonly name="CrBPValidStartView">
                                                    </span>
                                                    <span id="spanImgStartDateBP_<%=bp%>" name="spanImgStartDateBP">
                                                        <a id="BPValidStartImg_<%=bp%>" name ="BPValidStartImg" style="cursor:hand;" >
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='BPValidStart_<%=bp%>' name='BPValidStart' align='middle' >
                                                        </a>
                                                    </span>                                                
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanEndDateBP_<%=bp%>" name="spanEndDateBP"> 
                                                        <input type=hidden id="CrBPValidEnd_<%=bp%>" name="CrBPValidEnd" >
                                                        <input name="CrBPValidEndView" type=text size=15 id="CrBPValidEnd_<%=bp%>View" class=wordfieldnormaldisabled readonly onKeyPress="disableEnter()">
                                                    </span>
                                                    <span id="spanImgEndDateBP_<%=bp%>" name="spanImgEndDateBP">
                                                        <a id="BPValidEndImg_<%=bp%>" name = "BPValidEndImg" style="cursor:hand;" >                                          																																																								
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='BPValidEnd_<%=bp%>' name='BPValidEnd' align='middle' >
                                                        </a>
                                                    </span>
                                                </td>
												<td align="center" valign="top">
                                                	<input type="checkbox" id="tagBP_<%=bp%>" name="tagBP" >
                                                </td>
                                                <td <%=xLineGrid%> class="wordingrid" align=center title="Action">
                                                    <a id="lockBP_<%=bp%>" name="lockBP" style="visibility:block" class="cursorHand" onClick="javascript:lockBP(this);">Lock</a>
                                                    &nbsp;||&nbsp;
                                                    <a id="deleteBP_<%=bp%>" name="deleteBP" style="visibility:block" class="cursorHand" onClick="javascript:deleteBP(this);">Del</a>
                                                </td>
                                                <td>
                                                	<input type="hidden" id="hidStatusBP_<%=bp%>" name="hidStatusBP">
                                                </td>																										
                                            </tr>
									<%
										else
                                    %>
                                        	<tr id="trBP_0" name="trBP">
                                            	<td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                	<select name="crTypeBP" id="crTypeBP_0">
                                                    	<option value="0">Free</option>
                                                    	<option value="1">Bundle</option>
                                                	</select> 
                                            	</td>
                                            	<td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                                	<input type="text" name="crSKUBP" id="crSKUBP_0" size="12" onClick="this.select();" onChange="checkSKUBP(this);" onKeyPress="disableEnter()">
                                                	<input type="hidden" name="hidSKUBP" id="hidSKUBP_0">
                                            	</td>
                                            	<td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                                	<div id="divBrandSeriDescBP_0" align="justify" name="divBrandSeriDescBP"></div>
                                                	<input type="hidden" id="hidBrandBP_0" name="hidBrandBP">
                                                	<input type="hidden" id="hidSeriBP_0" name="hidSeriBP">
                                                	<input type="hidden" id="hidDescBP_0" name="hidDescBP">                                                
                                            	</td>
                                            	<td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                                	<input type="text" name="crQtyBP" id="crQtyBP_0" size=1 maxlength=1 onKeyPress="crCekNumber();">
                                            	</td>
                                            	<td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                               		<span id="spanStartDateBP_0" name="spanStartDateBP">
                                                    	<input type=hidden id="CrBPValidStart_0" name="CrBPValidStart" value="" >
                                                    	<input type=text size=15 id="CrBPValidStart_0View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                	</span>
                                                	<span id="spanImgStartDateBP_0" name="spanImgStartDateBP">
                                                    	<a id="BPValidStartImg_0" name ="BPValidStartImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrBPValidStart_0'),null,document.getElementById('CrBPValidStart_0').value,document.getElementById('CrBPValidStart_0View'),'d s Y',null,null,false);">
                                                        	<img border='0' src="../Image/btn_clock_000.gif" id='BPValidStart_0' name='BPValidStart' align='middle' >
                                                    	</a>
                                                	</span>                                                
                                            	</td>
                                            	<td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                	<span id="spanEndDateBP_0" spanEndDateBP>
                                                    	<input type=hidden id="CrBPValidEnd_0" name="CrBPValidEnd" value="" >
                                                    	<input type=text size=15 id="CrBPValidEnd_0View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                	</span>
                                                	<span id="spanImgEndDateBP_0" name="spanImgEndDateBP">
                                                    	<a id="BPValidEndImg_0" name = "BPValidEndImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrBPValidEnd_0'),null,document.getElementById('CrBPValidEnd_0').value,document.getElementById('CrBPValidEnd_0View'),'d s Y',null,null,false)">                                          																																																								
                                                        	<img border='0' src="../Image/btn_clock_000.gif" id='BPValidEnd_0' name='BPValidEnd' align='middle' >
                                                    	</a>
                                                	</span>
                                            	</td>
												<td align="center" valign="top">
                                                	<input type="checkbox" id="tagBP_0" name="tagBP" >
                                            	</td>
                                           		<td <%=xLineGrid%> class="wordingrid" align=center title="Action" valign="top">
                                                	<a id="lockBP_0" name="lockBP" style="visibility:block" class="cursorHand" onClick="javascript:lockBP(this);">Lock</a>
                                                	&nbsp;||&nbsp;
                                                	<a id="deleteBP_0" name="deleteBP" style="visibility:block" class="cursorHand" onClick="javascript:deleteBP(this);">Del</a>
                                            	</td>
                                            	<td>
                                                	<input type="hidden" id="hidStatusBP_0" name="hidStatusBP" >
                                            	</td>																									
                                        	</tr>
                                    	<%
										end if
                                    	%>
                                    	<tr id="rowlastBP">
                                   			<td width=100% colspan="7" align="right" >
                                        	</td>
                                    	</tr>
                                	</table>
                             	</td>
                        	</tr>
                   		</table>
                 	</td>
             	</tr>
            </table>
            
            <table height="5px">
            	<tr>
                	<td>&nbsp;</td>
                </tr>
            </table>
            
            <table border=0 width="95%" align=center cellpadding=2 cellspacing=2>
				<tr>
					<td class="headerdata" height="20" align="center">
						Special Deal
					</td>
				</tr>
                <tr>
                	<td valign=top  style="border: 1px solid <%=xWarnaLine%>">
                    	<table border="0" width="100%" cellspacing="0" cellpadding="0">
                            <tr>
                                <td class="wordfield">
                                	<!-- SD->Special Deal-->
                                    <table border="0" width="100%" id="tableSpecialDeal" cellspacing="0" cellpadding="0" >
                                        <tr style="border: 1px solid <%=xWarnaLine%>">
                                            <td <%=bgnoact%> width="10%" align="center" title="Type"><font class="wordblackmedium"><nobr><strong>Type</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="10%" align="center" title="SKU"><font class="wordblackmedium"><nobr><strong>SKU</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="40%" align="center" title="ItemDesc"><font class="wordblackmedium"><nobr><strong>Item Description</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="3%" align="center" title="Qty"><font class="wordblackmedium"><nobr><strong>Qty</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="12%" align="center" title="Start Date"><font class="wordblackmedium"><nobr><strong>Start Date</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="12%" align="center" title="End Date"><font class="wordblackmedium"><nobr><strong>End Date</strong></nobr></font></td>
                                            <td <%=bgnoact%> width="3%" align="center" title="Tag Date"><font class="wordblackmedium"><nobr><strong>Tag</strong></nobr></font></td>
											<td <%=bgnoact%> width="10%" align="center" title="Action"><font class="wordblackmedium"><nobr><strong>Action</strong></nobr></font></td>
                                        </tr>
									<%
										if adaPromo = 1 then
											set rsPromoSD = Server.CreateObject("ADODB.Recordset")
											querySD = "select * from trx_InveCPBundleDetail where vTrxID = '"&currBundleID&"' and vPromo = 1 order by vNoUrut ASC "
											rsPromoSD.open querySD, conn, 3, 1, 0
											for sd = 0 to rsPromoSD.RecordCount - 1
											currBundlePartID = rsPromoSD("vBundlePartID")
									%>
                                            <tr id="trSD_<%=sd%>" name="trSD">
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                    <select name="crTypeSD" id="crTypeSD_<%=sd%>" disabled>
                                                        <option value="0" <%if rsPromoSD("vType") = 0 then Response.Write("selected") end if%> >Free</option>
                                                        <option value="1" <%if rsPromoSD("vType") = 1 then Response.Write("selected") end if%> >Bundle</option>
                                                    </select> 
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                                    <input type="text" name="crSKUSD" id="crSKUSD_<%=sd%>" size="12" onClick="this.select();" onChange="checkSKUSD(this);" onKeyPress="disableEnter()" value="<%=currBundlePartID%>" disabled>
                                                    <input type="hidden" name="hidSKUSD" id="hidSKUSD_<%=sd%>" value="<%=currBundlePartID%>">
                                                </td>
                                                <td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                                	<%
														set rsBrandSeriDesc = server.CreateObject("ADODB.Recordset")
														queryBrandSeriDesc = "select vBrandID, vSeri, vDesc from trx_InveComputer where vPartID = '"&currBundlePartID&"' "
														rsBrandSeriDesc.open queryBrandSeriDesc, conn, 3, 1, 0
														if not rsBrandSeriDesc.eof then
															tempBrandID = rsBrandSeriDesc("vBrandID")
															tempSeri = rsBrandSeriDesc("vSeri")
															tempDesc = rsBrandSeriDesc("vDesc")
															set rsBrandName = server.CreateObject("ADODB.Recordset")
															queryBrandName = "select vName from tlu_InveCPBrand where vBrandID = '"&tempBrandID&"' "
															rsBrandName.open queryBrandName, conn, 3, 1, 0
															if not rsBrandName.eof then
																tempBrandName = rsBrandName("vName")
															end if
														end if
													%>
                                                    <div id="divBrandSeriDescSD_<%=sd%>" align="justify" name="divBrandSeriDescSD"><%=tempBrandName%><br><%=tempSeri%><br><%=tempDesc%></div>
                                                    <input type="hidden" id="hidBrandSD_<%=sd%>" name="hidBrandSD" value="<%=tempBrandName%>">
                                                    <input type="hidden" id="hidSeriSD_<%=sd%>" name="hidSeriSD" value="<%=tempSeri%>">
                                                    <input type="hidden" id="hidDescSD_<%=sd%>" name="hidDescSD" value="<%=tempDesc%>">
                                                    <%
														tempBrandID = ""
														tempSeri = ""
														tempDesc = ""
														tempBrandName = ""
														rsBrandName.close
														rsBrandSeriDesc.close
														set rsBrandName = nothing
														set rsBrandSeriDesc = nothing
													%>                                                
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                                    <input type="text" name="crQtySD" id="crQtySD_<%=sd%>" size=1 maxlength=1 onKeyPress="crCekNumber();" value="<%=rsPromoSD("vQty")%>" disabled>
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanStartDateSD_<%=sd%>" name="spanStartDateSD">
                                                        <input type=hidden id="CrSDValidStart_<%=sd%>" name="CrSDValidStart" value="<%=FormatDateTime(CDate(rsPromoSD("vStartDateTime")),2)%>" >
                                                        <input type=text size=15 id="CrSDValidStart_<%=sd%>View" class=wordfieldnormaldisabled readonly value="<%=Day(rsPromoSD("vStartDateTime"))&" "&MonthName(Month(rsPromoSD("vStartDateTime")),1) &" "&Year(rsPromoSD("vStartDateTime"))%>" onKeyPress="disableEnter()">
                                                    </span>
                                                    <span id="spanImgStartDateSD_<%=sd%>" name="spanImgStartDateSD">
                                                        <a id="SDValidStartImg_<%=sd%>" name ="SDValidStartImg" style="cursor:hand;" >
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='SDValidStart_<%=sd%>' name='SDValidStart' align='middle' >
                                                        </a>
                                                    </span>                                                
                                                </td>
                                                <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                    <span id="spanEndDateSD_<%=sd%>" name="spanEndDateSD">
                                                        <input type=hidden id="CrSDValidEnd_<%=sd%>" name="CrSDValidEnd" value="<%=FormatDateTime(CDate(rsPromoSD("vEndDateTime")),2)%>" >
                                                        <input type=text size=15 id="CrSDValidEnd_<%=sd%>View" class=wordfieldnormaldisabled readonly value="<%=Day(rsPromoSD("vEndDateTime"))&" "&MonthName(Month(rsPromoSD("vEndDateTime")),1) &" "&Year(rsPromoSD("vEndDateTime"))%>" onKeyPress="disableEnter()">
                                                    </span>
                                                    <span id="spanImgEndDateSD_<%=sd%>" name="spanImgEndDateSD">
                                                        <a id="SDValidEndImg_<%=sd%>" name = "SDValidEndImg" style="cursor:hand;" >                                          																																																								
                                                            <img border='0' src="../Image/btn_clock_000.gif" id='SDValidEnd_<%=sd%>' name='SDValidEnd' align='middle' >
                                                        </a>
                                                    </span>
                                                </td>
												<%
												''add queryTagSD
												set rsTagSD = Server.CreateObject("ADODB.Recordset")
												queryTagSD = "select * from trx_InveCPBundleDetail where vBundlePartID = '"&currBundlePartID&"' and vTrxID = '"&currBundleID&"' and vPromo = 1 order by vNoUrut ASC "
												rsTagSD.open queryTagSD, conn, 3, 1, 0
												IF NOT rsTagSD.eof Then
													cekTagSD = rsTagSD("vTagDate")
												END IF
												rsTagSD.Close
												SET rsTagSD = NOTHING
												%>
												<td align="center" valign="top">   
                                                	<input type="checkbox" id="tagSD_<%=sd%>" name="tagSD" <%IF cekTagSD = TRUE THEN Response.Write "Checked" END IF%> disabled>                                               	
                                                </td>
                                                <td <%=xLineGrid%> class="wordingrid" align=center title="Action">
                                                    <a id="lockSD_<%=sd%>" name="lockSD" style="visibility:block" class="cursorHand" onClick="javascript:lockSD(this);">Lock</a>
                                                    &nbsp;||&nbsp;
                                                    <a id="deleteSD_<%=sd%>" name="deleteSD" style="visibility:block" class="cursorHand" onClick="javascript:deleteSD(this);">Del</a>
                                                </td>
                                                <td>
                                                	<input type="hidden" id="hidStatusSD_<%=sd%>" name="hidStatusSD" value="<%=rsPromoSD("vStsActive")%>">
                                                </td>																										
                                            </tr> 
									<%
												rsPromoSD.MoveNext
											next
											rsPromoSD.close
											set rsPromoSD = nothing
									%>
                                    		<tr id="trSD_<%=sd%>" name="trSD">
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                <select name="crTypeSD" id="crTypeSD_<%=sd%>">
                                                    <option value="0">Free</option>
                                                    <option value="1">Bundle</option>
                                                </select> 
                                            </td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                            	<input type="text" name="crSKUSD" id="crSKUSD_<%=sd%>" size="12" onClick="this.select();" onChange="checkSKUSD(this);" onKeyPress="disableEnter()">
                                                <input type="hidden" name="hidSKUSD" id="hidSKUSD_<%=sd%>">
                                            </td>
                                            <td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                            	<div id="divBrandSeriDescSD_<%=sd%>" align="justify" name="divBrandSeriDescSD"></div>
                                                <input type="hidden" id="hidBrandSD_<%=sd%>" name="hidBrandSD">
                                                <input type="hidden" id="hidSeriSD_<%=sd%>" name="hidSeriSD">
                                                <input type="hidden" id="hidDescSD_<%=sd%>" name="hidDescSD">                                                
                                            </td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                            	<input type="text" name="crQtySD" id="crQtySD_<%=sd%>" size=1 maxlength=1 onKeyPress="crCekNumber();">
											</td>
											<td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                <span id="spanStartDateSD_<%=sd%>">
                                                    <input type=hidden id="CrSDValidStart_<%=sd%>" name="CrSDValidStart" value="" >
                                                    <input type=text size=15 id="CrSDValidStart_<%=sd%>View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                </span>
                                                <span id="spanImgStartDateSD_<%=sd%>" name="spanImgStartDateSD">
                                                    <a id="SDValidStartImg_<%=sd%>" name ="SDValidStartImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrSDValidStart_<%=sd%>'),null,document.getElementById('CrSDValidStart_<%=sd%>').value,document.getElementById('CrSDValidStart_<%=sd%>View'),'d s Y',null,null,false);">
                                                        <img border='0' src="../Image/btn_clock_000.gif" id='SDValidStart_<%=sd%>' name='SDValidStart' align='middle' >
                                                    </a>
                                                </span>                                                
											</td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                <span id="spanEndDateSD_<%=sd%>" name="spanEndDateSD">
                                                    <input type=hidden id="CrSDValidEnd_<%=sd%>" name="CrSDValidEnd" value="" >
                                                    <input type=text size=15 id="CrSDValidEnd_<%=sd%>View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                </span>
                                                <span id="spanImgEndDateSD_<%=sd%>" name="spanImgEndDateSD">
                                                    <a id="SDValidEndImg_<%=sd%>" name = "SDValidEndImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrSDValidEnd_<%=sd%>'),null,document.getElementById('CrSDValidEnd_<%=sd%>').value,document.getElementById('CrSDValidEnd_<%=sd%>View'),'d s Y',null,null,false)">                                          																																																								
                                                        <img border='0' src="../Image/btn_clock_000.gif" id='SDValidEnd_<%=sd%>' name='SDValidEnd' align='middle' >
                                                    </a>
                                                </span>
											</td>
											<td align="center" valign="top">   
                                                	<input type="checkbox" id="tagSD_<%=sd%>" name="tagSD" >                                               	
                                                </td>
                                            <td <%=xLineGrid%> class="wordingrid" align=center title="Action">
                                            	<a id="lockSD_<%=sd%>" name="lockSD" style="visibility:block" class="cursorHand" onClick="javascript:lockSD(this);">Lock</a>
                                                &nbsp;||&nbsp;
                                                <a id="deleteSD_<%=sd%>" name="deleteSD" style="visibility:block" class="cursorHand" onClick="javascript:deleteSD(this);">Del</a>
                                            </td>	
                                            <td>
                                                <input type="hidden" id="hidStatusSD_<%=sd%>" name="hidStatusSD" >
                                            </td>																									
                                        </tr>
                                    <%
                                        else
                                    %>                                   
                                        <tr id="trSD_0" name="trSD">
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Promo">
                                                <select name="crTypeSD" id="crTypeSD_0">
                                                    <option value="0">Free</option>
                                                    <option value="1">Bundle</option>
                                                </select> 
                                            </td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="SKU">
                                            	<input type="text" name="crSKUSD" id="crSKUSD_0" size="12" onClick="this.select();" onChange="checkSKUSD(this);" onKeyPress="disableEnter()">
                                                <input type="hidden" name="hidSKUSD" id="hidSKUSD_0">
                                            </td>
                                            <td <%=xLineGrid%> class="wordInLineGrid" align="left" title="Product" valign="top">
                                            	<div id="divBrandSeriDescSD_0" align="justify" name="divBrandSeriDesc"></div>
                                                <input type="hidden" id="hidBrandSD_0" name="hidBrandSD">
                                                <input type="hidden" id="hidSeriSD_0" name="hidSeriSD">
                                                <input type="hidden" id="hidDescSD_0" name="hidDescSD">                                                
                                            </td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center title="Quantity">
                                            	<input type="text" name="crQtySD" id="crQtySD_0" size=1 maxlength=1 onKeyPress="crCekNumber();">
											</td>
											<td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                <span id="spanStartDateSD_0" name="spanStartDateSD">
                                                    <input type=hidden id="CrSDValidStart_0" name="CrSDValidStart" value="" >
                                                    <input type=text size=15 id="CrSDValidStart_0View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                </span>
                                                <span id="spanImgStartDateSD_0" name="spanImgStartDateSD">
                                                    <a id="SDValidStartImg_0" name ="SDValidStartImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrSDValidStart_0'),null,document.getElementById('CrSDValidStart_0').value,document.getElementById('CrSDValidStart_0View'),'d s Y',null,null,false);">
                                                        <img border='0' src="../Image/btn_clock_000.gif" id='SDValidStart_0' name='SDValidStart' align='middle' >
                                                    </a>
                                                </span>                                                
											</td>
                                            <td <%=xLineGrid%> class="wordfield" valign=top align=center>
                                                <span id="spanEndDateSD_0" name="spanEndDateSD">
                                                    <input type=hidden id="CrSDValidEnd_0" name="CrSDValidEnd" value="" >
                                                    <input type=text size=15 id="CrSDValidEnd_0View" class=wordfieldnormaldisabled readonly value="" onKeyPress="disableEnter()">
                                                </span>
                                                <span id="spanImgEndDateSD_0" name="spanImgEndDateSD">
                                                    <a id="SDValidEndImg_0" name = "SDValidEndImg" style="cursor:hand;" onClick="javascript:getCalendarFor('model1','0','','<%=now()%>',document.getElementById('CrSDValidEnd_0'),null,document.getElementById('CrSDValidEnd_0').value,document.getElementById('CrSDValidEnd_0View'),'d s Y',null,null,false)">                                          																																																								
                                                        <img border='0' src="../Image/btn_clock_000.gif" id='SDValidEnd_0' name='SDValidEnd' align='middle' >
                                                    </a>
                                                </span>
											</td>
											<td align="center" valign="top">   
												<input type="checkbox" id="tagSD_0" name="tagSD" >                                               	
                                            </td>
                                            <td <%=xLineGrid%> class="wordingrid" align=center title="Action">
                                            	<a id="lockSD_0" name="lockSD" style="visibility:block" class="cursorHand" onClick="javascript:lockSD(this);">Lock</a>
                                                &nbsp;||&nbsp;
                                                <a id="deleteSD_0" name="deleteSD" style="visibility:block" class="cursorHand" onClick="javascript:deleteSD(this);">Del</a>
                                            </td>	
                                            <td>
                                                <input type="hidden" id="hidStatusSD_0" name="hidStatusSD" >
                                            </td>																									
                                        </tr> 
									<%
                                        end if
                                    %>
                                        <tr id="rowlastSD">
                                            <td width=100% colspan="7" align="right" >
                                            </td>
                                        </tr> 
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            
            <table height="5px">
            	<tr>
                	<td>&nbsp;</td>
                </tr>
            </table>
             
			<table border=0 width="40%" align=center cellpadding=2 cellspacing=2>
				<tr>
					<td colspan="2" class="headerdata" height="20" align="center">
						Vendor
					</td>
				</tr>			
				<tr>
					<td align="center" valign="top"  style="border: 1px solid <%=xWarnaLine%>" id="div_vendor">
						<table border="0" cellpadding=0 cellspacing=0 width=100%>
							<tr><td height=2></td></tr>
							<tr>
								<td width=4% height=17 <%=bgnoact%> align="center" title="No"><font class="wordblackmedium"><nobr><strong>No</strong></nobr></font></td>
								<td width=15% height=17 <%=bgnoact%> align="center" title="Vendor ID"><font class="wordblackmedium"><nobr><strong>Vendor ID</strong></nobr></font></td>
								<td width=75% height=17 <%=bgnoact%> align="center" title="Vendor Name"><font class="wordblackmedium"><nobr><strong>Vendor Name</strong></nobr></font></td>
								<td <%=bgnoactright%> align="center" title="Action"><font class="wordblackmedium"><nobr><strong>Del</strong></nobr></font></td>
							</tr>
							<tr><td height=2></td></tr>
							<%
							currJmlVendorDef=5:currJmlVendorAdd=1
							if lcase(curract)="edit" then%>
						    <%
								xstring=" select b.vvndid,v.vname,b.vVndType,b.vVPLFromID from trx_invebuyingvendor b, trx_vendormain v "&_
										" where b.vvndid=v.vvndid and b.vpartid='"&currPartId&"' order by vVPLFromID DESC, vVndType DESC"
										'Response.Write(xstring)
								getRs.open xstring,conn,1,3
								currJmlVendor=getRs.RecordCount
								if getRs.eof then currJmlVendor=0
								if currJmlVendor<currJmlVendorDef then currJmlVendorAdd=currJmlVendorDef-currJmlVendor
								if not getRs.eof then
									for currCntVendor=1 to currJmlVendor
							%>
							<tr id=trObjVendor<%=currCntVendor%>>
								<td <%=xLineGrid%> class=wordingrid><div id=objNo><%=currCntVendor%></div></td>
								<td <%=xLineGrid%> class=wordingrid><div id=objID><input type="text" size="15" name="crVendorID<%=currCntVendor%>" class="wordfieldnormaldisabled" value="<%=escapeVal(replace(trim(getRs("vvndid")),",",""))%>" readonly>&nbsp;<img align=absmiddle STYLE="cursor:hand;" <%=xMouseFind%> border="0" alt="findVendor<%=currCntVendor%>" id="findVendor<%=currCntVendor%>" name="findVendor<%=currCntVendor%>" dynamicanimation="findVendor<%=currCntVendor%>" onClick="javascript:fjs_getVendor(<%=currCntVendor%>);"></div></td>
								<td <%=xLineGrid%> class=wordingrid><div id=objNm><input type=hidden name="crVendorName<%=currCntVendor%>" value="<%=escapeVal(trim(getrs("vname")))%>"><div id=objVendNmDet<%=currCntVendor%>><%=getrs("vname")%>&nbsp;</div><nobr></div></td>
								<td <%=xLineGrid%> class=wordingrid>
									<div id=objAct><img align=absmiddle STYLE="cursor:hand;" <%=xMouseRemove%> border="0" alt="DelVendor<%=currCntVendor%>" id="DelVendor<%=currCntVendor%>" name="DelVendor<%=currCntVendor%>" dynamicanimation="DelVendor<%=currCntVendor%>" onClick="javascript:fjs_clrVendor(<%=currCntVendor%>);"></div>
								</td>
							</tr>
							<%
									getRs.MoveNext
									if getRs.EOF then exit for
									Next
								end if
								getRs.close
							end if
							if currJmlVendorAdd>0 then
								for currCntVendor2=currJmlVendor+1 to currJmlVendor+currJmlVendorAdd
							%>
							<tr id=trObjVendor<%=currCntVendor2%>>
								<td <%=xLineGrid%> class=wordingrid><div id=objNo><%=currCntVendor2%></div></td>
								<td <%=xLineGrid%> class=wordingrid><div id=objID><input type="text" size="15" name="crVendorID<%=currCntVendor2%>" class="wordfieldnormaldisabled" readonly>&nbsp;<img align=absmiddle STYLE="cursor:hand;" <%=xMouseFind%> border="0" alt="findVendor<%=currCntVendor%>" id="findVendor<%=currCntVendor%>" name="findVendor<%=currCntVendor%>" dynamicanimation="findVendor<%=currCntVendor%>" onClick="javascript:fjs_getVendor(<%=currCntVendor2%>);"></div></td>
								<td <%=xLineGrid%> class=wordingrid><div id=objNm><input type=hidden name="crVendorName<%=currCntVendor2%>"><div id=objVendNmDet<%=currCntVendor2%>>&nbsp;</div><nobr></div></td>
								<td <%=xLineGrid%> class=wordingrid>
									<div id=objAct><img align=absmiddle STYLE="cursor:hand;" <%=xMouseRemove%> border="0" alt="DelVendor<%=currCntVendor2%>" id="DelVendor<%=currCntVendor2%>" name="DelVendor<%=currCntVendor2%>" dynamicanimation="DelVendor<%=currCntVendor2%>" onClick="javascript:fjs_clrVendor(<%=currCntVendor2%>);"></div>
								</td>
							</tr>
							<%
								Next
								currCntVendor=currCntVendor2-1
							end if
							%>
						</table>
						<br>
						<div align=right><img align=absmiddle STYLE="cursor:hand;" <%=xMouseAdd%> border="0" alt="AddVendor" id="AddVendor" name="AddVendor" dynamicanimation="AddVendor" onClick="javascript:fjs_addObjVendor();"></div>
					</td>
				</tr>
			</table>
			
			<table border=0 width="95%" align=center cellpadding=2 cellspacing=2>				
				<tr>
					<td colspan=2 align=center>
						<table width=100% border=0>
							<tr>
								<td><!-- #include file="../include/glob_footer_editor.asp" --></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2" align="center">
						<br>
						<%if currAct="Edit" then
						%>
                        <img STYLE="cursor:hand;" onClick="fjs_cancel();" <%=xMouseCancel%> border="0" alt="Cancel" id="Cancel" name="Cancel" dynamicanimation="Cancel">
						<%
						end if
                        %>
                        <%if currAct="add" then
						%>
                        <img STYLE="cursor:hand;" onClick="javascript:window.close();" <%=xMouseCancel%> border="0" alt="Cancel" id="Cancel" name="Cancel" dynamicanimation="Cancel">
						<%
						end if
                        %>
                        <img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='save';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSave%> border="0" alt="Save" id="Save" name="Save" dynamicanimation="Save">
						<img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='savenext';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSaveNext%> border="0" alt="SaveNext" id="SaveNext" name="SaveNext" dynamicanimation="SaveNext">
						<img STYLE="cursor:hand;" onClick="javascript:document.frm.crtypesave.value='savenew';if(fjs_chkdt()){document.frm.submit()}" <%=xMouseSaveNew%> border="0" alt="SaveNew" id="SaveNew" name="SaveNew" dynamicanimation="SaveNew">
						<div id="pp" width="100%" height="100%" style="position: absolute; left: 0; top: 0"></div>
						<input type="image" src="../image/btn_blank.gif" />
                        <input id="crValueIDR" name="crValueIDR" type="hidden" size="20" maxlength="15" value="<%=currValueIDR%>" readonly />
                        <input id="crValueUSD" name="crValueUSD" type="hidden" size="20" maxlength="15" value="<%=currValueUSD%>" readonly />
                       	<input id="crValueJPY" name="crValueJPY" type="hidden" size="20" maxlength="15" value="<%=currValueJPY%>" readonly />
                        <input id="crLastUpdateVPL" name="crLastUpdateVPL" type="hidden" size="2" maxlength="1" value="0" readonly />
                        <input id="crLastUpdateWebPrice" name="crLastUpdateWebPrice" type="hidden" size="2" maxlength="1" value="0" readonly />
                        <input id="crLastUpdateSpecialPrice" name="crLastUpdateSpecialPrice" type="hidden" size="2" maxlength="1" value="0" readonly />
                        <input type=hidden name=crJmlCatSec value="<%=currCntCatSec%>">
						<input type=hidden name=crJmlVendor value="<%=currCntVendor%>">
						
                        <input type="hidden" name="crPartID" value="<%=currPartID%>" />
                        <input type="hidden" name="crBrandNameOri" value="<%=currBrandName%>" />
                        <input type="hidden" name="crGoodDesc" value="<%=currGoodDesc%>" />
                        <input type="hidden" name="crStatusSN" value="<%=currNeedSN%>" />
						<input type=hidden name=crAct value="<%=currAct%>">
                        <input type=hidden name=crCopy value="<%=currCopy%>">
						<input type=hidden name=crVndTypeh>
						<input type=hidden name=crtypesave>
                        <input type=hidden name=flagCat id=flagCat value="<%=flagCat%>">
                        <input type=hidden name=flagBrand id=flagBrand value="<%=flagBrand%>">
                        
                        <input type=hidden name="nTypeBP" id="nTypeBP">
						<input id="nSKUBP" name="nSKUBP" type="hidden">
                        <input type=hidden name="nBrandBP" id="nBrandBP">
                        <input type=hidden name="nSeriBP" id="nSeriBP">
                        <input type=hidden name="nDescBP" id="nDescBP">
                        <input type=hidden name="nQtyBP" id="nQtyBP">
                        <input type=hidden name="nStartDateBP" id="nStartDateBP">
                        <input type=hidden name="nEndDateBP" id="nEndDateBP">
                        <input type=hidden name="nTagBP" id="nTagBP">
                        
                        <input id="nTypeSD" name="nTypeSD" type="hidden">
						<input type=hidden name="nSKUSD" id="nSKUSD">
                        <input type=hidden name="nBrandSD" id="nBrandSD">
                        <input type=hidden name="nSeriSD" id="nSeriSD">
                        <input type=hidden name="nDescSD" id="nDescSD">
                        <input type=hidden name="nQtySD" id="nQtySD">
                        <input type=hidden name="nStartDateSD" id="nStartDateSD">
                        <input type=hidden name="nEndDateSD" id="nEndDateSD">
                        <input type=hidden name="nTagSD" id="nTagSD">
						<input type="hidden" name="lastMargin" id="lastMargin" value="<%if lcase(currAct)="edit" then Response.Write("crMarginValue") else Response.Write("") end if%>">
					</td>
				 </tr>
                 <tr><td>&nbsp;</td></tr>
                 <tr>
					<td colspan=2 align=right class=""wordfield"" >
						<font size=2><strong>What To do?</strong>
                		[&nbsp;<a href="digoff_inve_prodcatalog.asp?crAct=Add&crCopy=Yes&crPartId=<%=currPartId%>&crLoad=1">Copy</a>&nbsp;]
                			&nbsp;&nbsp;
						</font>
            		</td>
				</tr>
			</table>
            
            </form>			
			<!--#include file="../include/glob_tab_bottom.asp"-->
		</td>
	</tr>
	<!-- End Table disini -->
	<%
		''if escapeVal(currCntnPrice) <> "" then
		''	Response.Write("<script language = 'javascript'> ")
		''	Response.Write("fjs_setlastupdateprcVPL();")
		''	Response.Write("fjs_inputHidCntnPriceCurrID(document.frm.crCntnPrcCurrID);")
		''	Response.Write("fjs_formatCurrencyCntn(document.frm.crCntnPrcCurrID,document.frm.crCntnPrice);")
		''	Response.Write("<script>")														
		''end if
		
		''if escapeVal(currPrice) <> "" then
		''	Response.Write("<script language = 'javascript'> ")
		''	Response.Write("fjs_setlastupdateWebPrice();")
		''	Response.Write("fjs_inputHidWebPriceCurrID(document.frm.crPrcCurrID);")
		''	Response.Write("fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice)")
		''	Response.Write("<script>")														
		''end if

		''if escapeVal(currSPrice) <> "" then
		''	Response.Write("<script language = 'javascript'> ")
		''	Response.Write("fjs_setlastupdateSpecialPrice();")
		''	Response.Write("fjs_inputHidSpPriceCurrID(document.frm.crSPrcCurrID);")
		''	Response.Write("fjs_formatCurrencySp(document.frm.crSPrcCurrID,document.frm.crSPrice);")
		''	Response.Write("<script>")														
		''end if
		
	''set List = new CListPage
	''response.write List.GenerateCloseContent
	''response.write List.GenerateFooter
	''set List = nothing
%>
</table>
<!--#include file="../include/gen_calendar.asp" -->
</body>
<LINK href="../css/inve_prodcatalog.css" rel=STYLESHEET type="text/css">
<link rel="stylesheet" type="text/css" href="../css/basic.css">
<LINK href="../css/gnr_all.css" rel=STYLESHEET type="text/css">
<link type="text/css" href="../css/flick/jquery-ui-1.8.6.custom.css" rel="stylesheet" />
<script language="javascript" src="../js/jquery-1.4.2.min.js"></script>
<script language="javascript" src="../js/jquery.bgiframe-2.1.1.js"></script>
<script type="text/javascript" src="../js/jquery.ui.core.js"></script>
<script type="text/javascript" src="../js/jquery.ui.widget.js"></script>
<script type="text/javascript" src="../js/jquery.ui.mouse.js"></script>
<script type="text/javascript" src="../js/jquery.ui.button.js"></script>
<script type="text/javascript" src="../js/jquery.ui.draggable.js"></script>
<script type="text/javascript" src="../js/jquery.ui.position.js"></script>
<script type="text/javascript" src="../js/jquery.ui.dialog.js"></script>
<script type="text/javascript" src="../js/jquery.ui.autocomplete.js"></script>
<script type="text/javascript" src="../js/jquery.ui.datepicker.js"></script>
<script type="text/javascript" src="../js/jquery.effects.core.js"></script>
<script type="text/javascript" src="../js/jquery.effects.drop.js"></script>  
<link href="../css/accordion.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../js/jquery.simplemodal.js"></script>
<script type="text/javascript" src="../js/gjs_inve_prodcatalog.js"></script> 
<SCRIPT type="text/javascript" src="../js/gjs_winopen.js"></SCRIPT>
<%
set getRs= nothing
set RSCurrency= nothing
%>
<SCRIPT type="text/javascript">
function fjs_clkCat(currCatID,infoSN){
	s = '<iframe src="digoff_inveprodcatalog_getCat.asp?crCatID='+currCatID+'&crSN='+infoSN+'" width=0 height=0 style="display:block"></iframe>';
	pp.innerHTML = s;
}
function fjs_clkBrand(currBrandID){
	s = '<iframe src="digoff_inveprodcatalog_getBrandNew.asp?crBrandID='+currBrandID+'" width=0 height=0 style="display:block"></iframe>';
	pp.innerHTML = s;
}
function chk_brandCode(){
	var upper;
	var res;
	var objRes;
	
	objRes = document.getElementById("crBrandNewCode");
	res = document.getElementById("crBrandNewCode").value;	
	upper = res.toUpperCase();
	objRes.value = upper;
}
function fjs_newBrand(){
	var objBrand = document.getElementById("crBrandID");
	var objBrandNewCode = document.getElementById("crBrandNewCode");
	var objBrandNewName = document.getElementById("crBrandNewName");
	var optBrand = objBrand.options;
	if(objBrand.selectedIndex!=0){
		indexLama = objBrand.selectedIndex;
	}
	if(document.getElementById("cbNewBrand").checked){
		objBrand.selectedIndex = 0;
		objBrand.className = "wordfield";
		objBrandNewCode.disabled = false;
		objBrandNewName.disabled = false;
		document.getElementById("crBrandNewCode").className = "wordfieldnormalmust";
		document.getElementById("crBrandNewName").className = "wordfieldnormalmust";				
	}else{
		objBrand.selectedIndex = indexLama;
		objBrand.className = "wordfieldnormalmust";
		objBrandNewCode.value = "";
		objBrandNewName.value = "";
		objBrandNewCode.className = "wordfield";
		objBrandNewName.className = "wordfield";
		objBrandNewCode.disabled = true;
		objBrandNewName.disabled = true;
	}
}
function fjs_newSeri(){
	var valSeri;
	
	valSeri = document.getElementById("crSeri").value;
	if (valSeri.length>100){
		alert("You can not insert seri more than 100 characters");	
	}
}
function fjs_changenew(){
	var tempChange = document.getElementsByName("crSeri").length;
	for(q=0;q<tempChange;q++){
		var frmBrand=fjs_getnamebrand();
		var frmSeri=document.getElementsByName("crSeri")[q].value;
		var s="";
		if(frmBrand!=""){
			s=frmBrand.substr(0,3);
		}
		if(s.length<3){
			s+=replicatestr(" ",3-s.length);
		}
		if(frmSeri!=""){
			s+=frmSeri.substr(frmSeri.length-5,5);
		}
	}
}
function gjs_string(evt) {
	evt = (evt) ? evt : window.event;
	var charCode = (evt.which) ? evt.which : evt.keyCode;
	if (charCode>64) {
   		return true;
	}else{
	    return false;
	}
}
function gjs_bilangan(evt) {
	evt = (evt) ? evt : window.event;
	var tombol = (evt.which) ? evt.which : evt.keyCode;
	st = '0123456789.';
	key = String.fromCharCode(tombol);
	if(st.indexOf(key)==-1){
		return false;
	}else{
		return true;
	}		
}
function chk_brandName(){
	var jmlIndex;
	var objBrandID = document.getElementById("crBrandID");
	jmlIndex = objBrandID.length;
	var strNewBrand;
	var varIndex;
	var varOpt;
	var cek = true;
	strNewBrand = document.getElementById("crBrandNewName").value;
	strNewBrand = strNewBrand.replace(" ","");
	strNewBrand = strNewBrand.toLowerCase();
			
	for(x=1;x<jmlIndex;x++){
		varIndex = document.getElementById("crBrandID").selectedIndex=x;
		varOpt = document.getElementById("crBrandID").options;
		if (strNewBrand==varOpt[varIndex].text.toLowerCase()){
			document.getElementById("crBrandID").selectedIndex=0;
			cek = false;
			alert("Your Brand has been existed in Database!!!");
			break;	
		}
		
	}
	if(cek==true){
		document.getElementById("crBrandID").selectedIndex=0;
		cekNewBrand = true;
		return true;
	}else{
		document.getElementById("crBrandID").selectedIndex=0;
		cekNewBrand = false;
		return false;
	}
}
function fjs_descAll(){
	var str;
	var strResult;
	var strDesc;
	var extStr;
	var indxStr;
			
	str = document.getElementById("crDesc").value;
	strDesc = document.getElementById("crDescOnly").value;
	indxStr = str.indexOf("<br>")
	if (str =='' || indxStr == -1){
		document.getElementById("crDesc").value=strDesc;
	}else if (indxStr != -1){
		extStr = str.substring(0,indxStr);
		strResult = str.replace(extStr,strDesc);
		document.getElementById("crDesc").value=strResult;
	}
}
function fjs_pupDescription(){
	var content;
			
	content = document.getElementById("crDesc").value;
	content = content.replace(/["]+/gim,"symDQ");
	gjs_winopen('Description','gnrt_descriptionInfo.asp?crParam='+content,750,150,550,200,true);		
}
function fjs_Warranty(){
	var authWarranty;
	var mchWarranty;
	var extWarranty;
	
	mchWarranty = document.getElementById("crMchWarranty").checked;
	authWarranty = document.getElementById("crAuthWarranty").checked;
	extWarranty = document.getElementById("crExtTxtWarrTemp").value
	cekRegex = /Merchant/gi;
	
	if (mchWarranty==true){
		document.getElementById("crAuthWarranty").disabled = true;
		if (extWarranty.match(cekRegex)){
			document.getElementById("crExtTxtWarr").value = extWarranty
		}else{
			document.getElementById("crExtTxtWarr").value = 'dari Merchant'
		}
	}else if (authWarranty==true){
		document.getElementById("crMchWarranty").disabled = true;
		if (extWarranty.match(cekRegex)){
			document.getElementById("crExtTxtWarr").value = '';
		}else{
			document.getElementById("crExtTxtWarr").value = extWarranty;
		}
	}else{
		document.getElementById("crAuthWarranty").disabled = false;
		document.getElementById("crMchWarranty").disabled = false;
		if (extWarranty.match(cekRegex)){
			document.getElementById("crExtTxtWarr").value = '';
		} else {
			document.getElementById("crExtTxtWarr").value = extWarranty;
		}
	}
}
function fjs_pupMarketing(){
	var content;
			
	content = document.getElementById("crMarketingInfo").value;
	gjs_winopen('Editor Marketing Info','gnrt_marketingInfo.asp?crParam='+content,980,550,350,0,true);	
}
function SetMarketingInfo(cparam){
	document.getElementById("crMarketingInfo").value=cparam;
	alert("Marketing Info Has Been Saved Successfully");
	tutupwindow();
}
function fjs_displayDate(val){
	if(document.frm.crStatusID.value==""){
		document.all.tr_periode.style.display="none";
	}else{
		document.all.tr_periode.style.display="inline";
	}
}
function fjs_getPVnd(){
	gjs_winopen('','gnrt_vendorAdd.asp',610,300,0,0,true);
}
function fjs_delPVnd(){
	document.frm.crPVndID.value="";
	document.frm.crPVndName.value="";
}			
function fjs_getVendor(cIdx){
	gjs_winopen('','gnrt_vendor.asp?crParam='+cIdx,610,300,0,0,true);
}		
function SetVendorAdd(id,name,caddress,cidx){
	if(cidx!=""){
		eval("document.frm.crVendorID"+cidx+".value=id");
		eval("document.frm.crVendorID"+cidx+".focus()");
		eval("document.frm.crVendorName"+cidx+".value=name");
		eval("document.all.objVendNmDet"+cidx+".innerText=name");
	}else{
		document.frm.crPVndID.value=id;
		document.frm.crPVndName.value=name;
	}
	tutupwindow();
}
function fjs_setlastupdateprc(){
	document.frm.crLastUpdate.value="1";
}		
function fjs_setlastupdateprcVPL(){
	document.frm.crLastUpdateVPL.value="1";
}		
function fjs_setlastupdateWebPrice(){
	document.frm.crLastUpdateWebPrice.value="1";
}		
function fjs_setlastupdateSpecialPrice(){
	document.frm.crLastUpdateSpecialPrice.value="1";
}
function fjs_inputHidCntnPriceCurrID(objCurr){
	var myOption2;
	myOption2 = -1;
	for (i=objCurr.length-1;i>-1;i--) {
		if (objCurr[i].checked) {
			myOption2 = i;
			i = -1;
		}
	}
	document.frm.hidCntnPriceCurrID.value = objCurr[myOption2].value;
}
function fjs_formatCurrencyCntn(objCur,obj){
	var myOption0;
	myOption0 = -1;			
	for (i=objCur.length-1; i > -1; i--) {
		if (objCur[i].checked) {
			myOption0 = i; i = -1;
		}
	}		
	if(myOption0 != -1 ){
		if(objCur[myOption0].value == "CUR01"){
			CntnPriceCurrID = "IDR";
		}else if(objCur[myOption0].value == "CUR02"){
			CntnPriceCurrID = "USD";
		}else{
			CntnPriceCurrID = "JPY";
		}
	}
	if (document.frm.hidCntnPriceCurrID.value!=""){
		var num = new NumberFormat();
		num.setInputDecimal('.');
		num.setCurrency(true);
		num.setCurrencyPosition(num.LEFT_OUTSIDE);
		num.setCurrencyValue('');
		num.setNegativeFormat(num.LEFT_DASH);
		num.setNegativeRed(false);
		num.setSeparators(true, ',', '<%=chr(0159)%>');
		if(CntnPriceCurrID.toLowerCase()=="idr"){
			num.setPlaces('0');
		}else{
			num.setPlaces('2');
		}
		num.setNumber(obj.value);
		obj.value=num.toFormatted().replace(/<%=chr(0159)%>/gi, ".");
		fjs_chklimitcurrency(CntnPriceCurrID.toLowerCase(),obj.value.replace(/,/gi,""));
	}else{
		alert("VPL Currency must selected before");
		obj.value="";
	}
}
function set_checkedWP() {
	if(document.getElementById("crMarginPct").value != "" || document.getElementById("chkCall").checked == true || document.getElementById("chkStock").checked == true) {
		for(x=0;x<=2;x++) {
			if(document.getElementsByName("crCntnPrcCurrID").item(x).checked == true) {
				document.getElementsByName("crPrcCurrID").item(x).checked = true;
				fjs_setlastupdateWebPrice();
				break;
			}
		}
		fjs_inputHidWebPriceCurrID(document.frm.crPrcCurrID);
		fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice);
	}
	//alert(document.getElementById("hidWebPriceCurrID").value);
	set_Comma();
	set_HBWP();
}
function set_HBWP() {
	VPList = document.getElementById("crCntnPrice").value.replace(/,/gi,"");
	currVPL = false;
	for(i=0;i<=2;i++) {
		if(document.getElementsByName("crCntnPrcCurrID").item(i).checked == true) {
			currVPL = true;
			break;
		}
	}
	if(currVPL) {
		if(VPList <= "0" || VPList <= "0.00") {
			alert("Vendor Price must more than 0");
			document.getElementById("crCntnPrice").focus();
			document.getElementById("radInclude").checked = false;
			document.getElementById("radExclude").checked = false;
			document.getElementById("txtHB").value = "";
			document.getElementById("hidHB").value = "";
	
			if(document.getElementById("crMarginPct").value != "") {
				document.getElementById("crMarginPct").value = "";
				document.getElementById("crMarginValue").value = "";
				for(i=0;i<=2;i++) {
					document.getElementsByName("crPrcCurrID").item(i).checked = false;	
				}
				document.getElementById("crPrice").value = "";
				document.getElementById("crWebPrice").value = "";
				document.frm.crLastUpdateWebPrice.value="";
				document.getElementById("hidWebPriceCurrID").value = "";
				document.getElementById("chkCall").disabled = false; 
				document.getElementById("chkStock").disabled = false;
			}
		} else {	
			include = document.getElementById("radInclude").checked;
			exclude = document.getElementById("radExclude").checked;
			if(include == true || exclude == true) {
				if(include) {
					HB = VPList.replace(/,/gi,"")/1.1;
					document.getElementById("txtHB").value = HB;
					document.getElementById("hidHB").value = HB;
				} else if(exclude) {
					VPList = VPList.replace(/,/gi,"");
					document.getElementById("txtHB").value = VPList;
					document.getElementById("hidHB").value = VPList;
				}
				changeFormat(document.getElementById("txtHB"));
			}
	
			lastMargin = document.getElementById("lastMargin").value;
			if(lastMargin != "" && document.getElementById("crMarginPct").value != "") {
				HB = document.getElementById("hidHB").value.replace(/,/gi,"");
				if(lastMargin == "crMarginPct") {
					MarginPct = document.getElementById(lastMargin).value.replace(/,/gi,"");
					MarginVal = (parseFloat(HB)*parseFloat(MarginPct))/100;
					WP = parseFloat(HB)+((HB)*(parseFloat(MarginPct)/100));
					document.getElementById("crPrice").value = WP;
					document.getElementById("crMarginValue").value = MarginVal;
					for(x=0;x<=2;x++) {
						if(document.getElementsByName("crCntnPrcCurrID").item(x).checked == true) {
							document.getElementsByName("crPrcCurrID").item(x).checked = true;
							fjs_setlastupdateWebPrice();
							break;
						}
					}
					changeFormat(document.getElementById("crMarginPct"));
					changeFormat(document.getElementById("crMarginValue"));
					fjs_inputHidWebPriceCurrID(document.frm.crPrcCurrID);
					fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice);
				} else if(lastMargin == "crMarginValue") {
					MarginVal = document.getElementById(lastMargin).value.replace(/,/gi,"");
					MarginPct = (parseFloat(MarginVal)/(HB)*100);
					WP = parseFloat(HB)+parseFloat(MarginVal);
					document.getElementById("crPrice").value = WP;
					document.getElementById("crMarginPct").value = MarginPct;
					for(x=0;x<=2;x++) {
						if(document.getElementsByName("crCntnPrcCurrID").item(x).checked == true) {
							document.getElementsByName("crPrcCurrID").item(x).checked = true;
							fjs_setlastupdateWebPrice();
							break;
						}
					}
					changeFormat(document.getElementById("crMarginPct"));
					changeFormat(document.getElementById("crMarginValue"));
					fjs_inputHidWebPriceCurrID(document.frm.crPrcCurrID);
					fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice);
					document.getElementById("lastMargin").value = "crMarginValue";
				}
				set_WPAT();
			}
		}
	}
	//alert(document.getElementById("hidWebPriceCurrID").value);
}
function set_Comma() {
	chkComma = document.getElementById("allowComma");
	VPLBT = document.getElementById("hidHB").value.replace(/,/gi,"");
	marginVal = document.getElementById("crMarginValue").value.replace(/,/gi,"");;
	marginPct = document.getElementById("crMarginPct").value.replace(/,/gi,"");;

	if(chkComma.checked == true) {
		if(document.getElementById("hidWebPriceCurrID").value == "CUR01") {
			if(marginVal == "" || marginPct == "") {
				alert("Margin must be filled first !");
				document.getElementById("crMarginPct").focus();
				chkComma.checked = false;
			} else {
				WPBT = parseFloat(VPLBT) + parseFloat(marginVal);
				document.getElementById("crPrice").value = WPBT;		
			}
		} else {
			alert("Currency must IDR to use 'Allow Comma to IDR' !");	
			chkComma.checked = false;
		}
	} else {
		WPBT = parseFloat(VPLBT) + parseFloat(marginVal);
		document.getElementById("crPrice").value = WPBT;
		fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice);
	}
	changeFormat(document.getElementById("crPrice"));
	set_WPAT();	
}
function set_WPAT() {
	if(document.getElementById("hidWebPriceCurrID").value == "CUR01") {
		WPAT = parseFloat(document.getElementById("crPrice").value.replace(/,/gi,"")) * 1.1;
		WPAT = Math.ceil(WPAT);
	} else {
		WPAT = parseFloat(document.getElementById("crPrice").value.replace(/,/gi,"")) * 1.1;
	}
	document.getElementById("crWebPrice").value = WPAT;
	changeFormat(document.getElementById("crWebPrice"));	
}
/*function fjs_setlastupdateSpecialPrice(){
	document.frm.crLastUpdateSpecialPrice.value="1";
}*/
function changeFormat(th){
    var TempValue = gjs_getUnformatedCurrency(th.value);
    if (!isNaN(TempValue)){
        th.value=gjs_getFormatCurrency(TempValue);
    }
}
function fjs_inputHidWebPriceCurrID(objCurr){
	var myOption4;
	myOption4 = -1;
	for (i=objCurr.length-1; i > -1; i--){
		if (objCurr[i].checked){
			myOption4 = i; i = -1;
		}
	}
	document.frm.hidWebPriceCurrID.value = objCurr[myOption4].value;
}
function fjs_VPLPriceID(num){
	var objVPLPrice;
	var objHidVPLPrice;
	var objVPLPriceTax;
	var objHidVPLPriceTax;
	var objHidVPLPriceID;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objHidWebPriceID;
	var objValueIDR;
	var objValueUSD;
	var objValueJPY;
	var objPPN;
	var objMarginPct;
	var objMarginValue;
	var tempValue;
	var bool;
	var tempDiff;
	var tempVPLPriceTax;
	objVPLPrice =  document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objWebPrice =  document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objValueIDR = document.getElementById("crValueIDR");
	objValueUSD = document.getElementById("crValueUSD");
	objValueJPY = document.getElementById("crValueJPY");
	objPPN = document.getElementsByName("radPPN");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	bool = false;
	if (num==1 && objHidVPLPriceID.value!='CUR01'){
		if (objHidVPLPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidVPLPrice.value)*parseFloat(objValueUSD.value);
		}else if(objHidVPLPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidVPLPrice.value)*parseFloat(objValueJPY.value);
		}
		bool = true;
		objHidVPLPriceID.value = 'CUR01';
	}else if(num==2 && objHidVPLPriceID.value!='CUR02'){
		if (objHidVPLPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidVPLPrice.value)/parseFloat(objValueUSD.value);
		}else if(objHidVPLPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidVPLPrice.value)*parseFloat(objValueJPY.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueUSD.value);
		}
		bool = true;
		objHidVPLPriceID.value = 'CUR02';
	}else if(num==3 && objHidVPLPriceID.value!='CUR03'){
		if (objHidVPLPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidVPLPrice.value)/parseFloat(objValueJPY.value);
		}else if(objHidVPLPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidVPLPrice.value)*parseFloat(objValueUSD.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueJPY.value);
		}
		bool = true;
		objHidVPLPriceID.value = 'CUR03';
	}
	if((parseFloat(fjs_cekMoney(objMarginValue.value))==0) && (objHidWebPriceID.value==objHidVPLPriceID.value)){ 
		objHidVPLPrice.value = objHidWebPrice.value;
		objVPLPrice.value = objWebPrice.value;
		objHidVPLPriceTax.value = objHidWebPriceTax.value;
		objVPLPriceTax.value = objWebPriceTax.value;		
	}else{
		if (bool){
			objHidVPLPrice.value = tempValue.toFixed(6);
			objVPLPrice.value = tempValue.format(2,3,',','.');
			if(objHidVPLPriceID.value=='CUR01'){
				if(objPPN[0].checked){
					objHidVPLPriceTax.value = (tempValue*1.1).toFixed(6);
					objVPLPriceTax.value = ((tempValue*1.1)+0.5).format(0,3,',','.');
					tempVPLPriceTax = parseFloat(fjs_cekMoney(objVPLPriceTax.value));
					if((tempVPLPriceTax%100)>0){
						tempDiff = parseFloat(tempVPLPriceTax%100);
						tempValue = (((parseFloat(tempVPLPriceTax-tempDiff)-0.5))/1.1).toFixed(6);
						objHidVPLPrice.value = parseFloat(tempValue).toFixed(6);
						objVPLPrice.value = parseFloat(tempValue).format(2,3,',','.');
						objHidVPLPriceTax.value = parseFloat(tempValue*1.1).toFixed(6);
						objVPLPriceTax.value = parseFloat((tempValue*1.1)+0.5).format(0,3,',','.');
					}
				}else{
					objHidVPLPriceTax.value = (tempValue).toFixed(6);
					objVPLPriceTax.value = (tempValue+0.5).format(0,3,',','.');
				}
			}else{
				if(objPPN[0].checked){
					objHidVPLPriceTax.value = (tempValue*1.1).toFixed(6);
					objVPLPriceTax.value = (tempValue*1.1).format(2,3,',','.');
				}else{
					objHidVPLPriceTax.value = (tempValue).toFixed(6);
					objVPLPriceTax.value = (tempValue).format(2,3,',','.');
				}
			}
			//fjs_updateVPLPriceTax();
		}
	}
}
function fjs_WebPriceID(num){
	var objHidVPLPrice;
	var objVPLPrice;
	var objVPLPriceID;
	var objHidVPLPriceTax;
	var objVPLPriceTax;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objHidWebPriceID;
	var objValueIDR;
	var objValueUSD;
	var objValueJPY;
	var objMarginPct;
	var objMarginValue;
	var objPPN;
	var tempValue;
	var bool;
	objVPLPrice =  document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objWebPrice =  document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objPPN = document.getElementsByName("radPPN");
	objValueIDR = document.getElementById("crValueIDR");
	objValueUSD = document.getElementById("crValueUSD");
	objValueJPY = document.getElementById("crValueJPY");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	bool = false;
	if (num==1 && objHidWebPriceID.value!='CUR01'){
		if (objHidWebPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidWebPrice.value)*parseFloat(objValueUSD.value);
		}else if(objHidWebPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidWebPrice.value)*parseFloat(objValueJPY.value);
		}
		bool = true;
		objHidWebPriceID.value = 'CUR01';
	}else if(num==2 && objHidWebPriceID.value!='CUR02'){
		if (objHidWebPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidWebPrice.value)/parseFloat(objValueUSD.value);
		}else if(objHidWebPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidWebPrice.value)*parseFloat(objValueJPY.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueUSD.value);
		}
		bool = true;
		objHidWebPriceID.value = 'CUR02';
	}else if(num==3 && objHidWebPriceID.value!='CUR03'){
		if (objHidWebPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidWebPrice.value)/parseFloat(objValueJPY.value);
		}else if(objHidWebPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidWebPrice.value)*parseFloat(objValueUSD.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueJPY.value);
		}
		bool = true;
		objHidWebPriceID.value = 'CUR03';
	}
	if((parseFloat(fjs_cekMoney(objMarginValue.value))==0) && (objHidWebPriceID.value==objHidVPLPriceID.value)){ 
		objHidWebPrice.value = objHidVPLPrice.value;
		objWebPrice.value = objVPLPrice.value;
		objHidWebPriceTax.value = objHidVPLPriceTax.value;
		objWebPriceTax.value = objVPLPriceTax.value;		
	}else{
		if(bool){
			objHidWebPrice.value = tempValue.toFixed(6);
			objHidWebPriceTax.value = (tempValue*1.1).toFixed(6);
			objWebPrice.value = tempValue.format(2,3,',','.');
			if(objHidWebPriceID.value=='CUR01'){
				if(objPPN[0].checked){
					objHidWebPriceTax.value = ((parseFloat(tempValue)*1.1)).toFixed(6);
					if((parseFloat(objHidWebPriceTax.value)%1)>0){
						objWebPriceTax.value = ((parseFloat(tempValue)*1.1)+0.5).format(0,3,',','.');
						if((parseFloat(fjs_cekMoney(objWebPriceTax.value))%100)>0){
							tempValue = ((((tempValue.toFixed(6))*1.1)-0.5)/1.1).toFixed(6);
							objHidWebPrice.value = parseFloat(tempValue).toFixed(6);
							objWebPrice.value = parseFloat(tempValue).format(2,3,',','.');
							objHidWebPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
							objWebPriceTax.value = ((parseFloat(tempValue)*1.1)+0.5).format(0,3,',','.');
						}
					}else{
						objWebPriceTax.value = ((parseFloat(tempValue)*1.1)).format(0,3,',','.');
					}
				}else{
					objWebPriceTax.value = (parseFloat(tempValue)+0.5).format(0,3,',','.');
					objHidWebPriceTax.value = (parseFloat(tempValue)).toFixed(6);	
				}
			}else{
				if(objPPN[0].checked){
					objWebPriceTax.value = (parseFloat(tempValue)*1.1).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
				}else{
					objWebPriceTax.value = (parseFloat(tempValue)).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(tempValue)).toFixed(6);
				}
			}
		}
	}
}
function fjs_updateVPLPriceTax(){
	var objVPLPriceID;
	var objVPLPrice;
	var objHidVPLPrice;
	var objHidVPLPriceID;
	var objVPLPriceTax;
	var objHidVPLPriceTax;
	var objCall;
	var objOOS;
	var marginPct;
	var marginValue;
	var tempValueTax;
	var objPPN;
	var objMarginPct;
	var objMarginValue;
	var objTdWebPriceID;
	var objWebPriceID;
	var objHidWebPriceID;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objTdSPriceID;
	var objSPriceID;
	var objHidSPriceID;
	var objSPrice;
	var objHidSPrice;
	var objSPriceTax;
	var objHidSPriceTax;
	var objSPriceStart;
	var objSPriceStartView;
	var objSPriceEnd;
	var objSPriceEndView;
	var objSPriceStartCalendar;
	var objSPriceEndCalendar;
	var objChkStsSameSPrice;
	var tempVPLPriceTax;
	var tempDiff;
	var tempValue;
	objVPLPriceID = document.getElementsByName("crVPLPriceID");
	objVPLPrice =  document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objOOS = document.getElementById("chkOOS");
	objCall =  document.getElementById("chkCall");
	objPPN = document.getElementsByName("radPPN");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objTdWebPriceID = document.getElementById("tdWebPriceID");
	objWebPriceID = document.getElementsByName("crWebPriceID");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objWebPrice = document.getElementById("crWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objTdSPriceID = document.getElementById("tdSPriceID");
	objSPriceID = document.getElementsByName("crSPriceID");
	objSPrice = document.getElementById("crSPrice");
	objSPriceTax = document.getElementById("crSPriceTax");
	objSPriceStart = document.getElementById("CrSPValidStart");
	objSPriceStartView = document.getElementById("CrSPValidStartView");
	objSPriceEnd = document.getElementById("CrSPValidEnd");
	objSPriceEndView = document.getElementById("CrSPValidEndView");
	objSPriceStartCalendar = document.getElementById("tdSPriceStartCalendar");
	objSPriceEndCalendar = document.getElementById("tdSPriceEndCalendar");
	objChkStsSameSPrice = document.getElementById("crChkStsSameSPrice");
	tempVPLPriceTax = parseFloat(fjs_cekMoney(objVPLPriceTax.value)).toFixed(6);
	if(objHidVPLPriceID.value=='CUR01'){
		objVPLPriceTax.value = parseFloat(tempVPLPriceTax).format(0,3,',','.');
		if(objPPN[0].checked){
			objVPLPrice.value = parseFloat(tempVPLPriceTax/1.1).format(2,3,',','.');
			objHidVPLPrice.value = parseFloat(tempVPLPriceTax/1.1).toFixed(6);
			tempValue = (parseFloat(objHidVPLPrice.value*1.1)+0.5).toFixed(6);
			if(tempValue>tempVPLPriceTax){
				tempValue = ((parseFloat(tempVPLPriceTax)-0.5)/1.1).toFixed(6);
				objHidVPLPrice.value = parseFloat(tempValue).toFixed(6);
				objVPLPrice.value = parseFloat(tempValue).format(2,3,',','.');
				objHidVPLPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
			}else{
				tempValue = parseFloat(objHidVPLPrice.value).toFixed(6);
				objHidVPLPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
			}
		}else{
			objHidVPLPrice.value = parseFloat(tempVPLPriceTax).toFixed(6);
			objVPLPrice.value = parseFloat(tempVPLPriceTax).format(2,3,',','.');
			objHidVPLPriceTax.value = parseFloat(tempVPLPriceTax).toFixed(6);
		}
	}else{
		objVPLPriceTax.value = parseFloat(tempVPLPriceTax).format(2,3,',','.');
		if(objPPN[0].checked){
			objHidVPLPrice.value = (parseFloat(tempVPLPriceTax)/1.1).toFixed(6);
			objVPLPrice.value = (parseFloat(tempVPLPriceTax)/1.1).format(2,3,',','.');
			tempValue = (parseFloat(objHidVPLPrice.value*1.1)+0.5).toFixed(6);
			if(tempValue>tempVPLPriceTax){
				tempValue = ((parseFloat(tempVPLPriceTax)-0.5)/1.1).toFixed(6);
				objHidVPLPrice.value = parseFloat(tempValue).toFixed(6);
				objVPLPrice.value = parseFloat(tempValue).format(2,3,',','.');
				objHidVPLPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
			}else{
				tempValue = parseFloat(objHidVPLPrice.value).toFixed(6);
				objHidVPLPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
			}
		}else{
			objHidVPLPrice.value = parseFloat(tempVPLPriceTax).toFixed(6);
			objVPLPrice.value = parseFloat(tempVPLPriceTax).format(2,3,',','.');
			objHidVPLPriceTax.value = parseFloat(tempVPLPriceTax).toFixed(6);
		}
	}
	if(objVPLPriceID[0].checked){
		objWebPriceID[0].checked = true;
		objHidWebPriceID.value = 'CUR01';
	}else if(objVPLPriceID[1].checked){
		objWebPriceID[1].checked = true;
		objHidWebPriceID.value = 'CUR02';
	}else if(objVPLPriceID[2].checked){
		objWebPriceID[2].checked = true;
		objHidWebPriceID.value = 'CUR03';
	}else{
		objVPLPriceID[0].click();
		objWebPriceID[0].checked = true;
		objHidWebPriceID.value = 'CUR01';
	}
	//tempValueTax = parseFloat(fjs_cekMoney(objVPLPriceTax.value));
	//objVPLPriceTax.value = tempValueTax.format(0,3,',','.');
	//objHidVPLPriceTax.value = tempValueTax.toFixed(6);
	//objVPLPrice.value = (tempValueTax/1.1).format(2,3,',','.');
	//objHidVPLPrice.value = (tempValueTax/1.1).toFixed(6);
	
	if((objCall.checked==false) && (objOOS.checked==false)){
		objCall.checked = false;
		objCall.disabled = false;
		objOOS.checked = false;
		objOOS.disabled = false;
		objPPN[0].disabled = false;
		objPPN[1].disabled = false;
		//objPPN[0].checked = true;
		objMarginPct.readOnly = false;
		objMarginValue.readOnly = false;
		objTdWebPriceID.disabled = false;
		objWebPrice.disabled = false;
		objWebPriceTax.disabled = false;
		objTdSPriceID.disabled = false;
		objSPrice.disabled = false;
		objSPriceTax.disabled = false;
		objSPriceStartView.disabled = false;
		objSPriceEndView.disabled = false;
		objSPriceStartCalendar.disabled = false;
		objSPriceEndCalendar.disabled = false;
		objChkStsSameSPrice.disabled = false;
		fjs_marginValue();
	}
}
function fjs_updatePPNVPL(num){
	var objHidVPLPriceID;
	var objVPLPrice;
	var objHidVPLPrice;
	var objVPLPriceTax;
	var objHidVPLPriceTax;
	var tempVPLPriceTax;
	var tempDiff;
	var tempValue;
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	if(num==1 && (parseFloat(objHidVPLPrice.value)==parseFloat(objHidVPLPriceTax.value)) && (parseFloat(objHidVPLPrice.value)>0)){
		if(objHidVPLPriceID.value=='CUR01'){
			objHidVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).toFixed(6);
			objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)+0.5).format(0,3,',','.');	
			tempVPLPriceTax = parseFloat(fjs_cekMoney(objVPLPriceTax.value));
			if((tempVPLPriceTax%100)>0){
				tempDiff = parseFloat(tempVPLPriceTax%100);
				tempValue = (((parseFloat(tempVPLPriceTax-tempDiff)-0.5))/1.1).toFixed(6);
				objHidVPLPrice.value = parseFloat(tempValue).toFixed(6);
				objVPLPrice.value = parseFloat(tempValue).format(2,3,',','.');
				objHidVPLPriceTax.value = parseFloat(tempValue*1.1).toFixed(6);
				objVPLPriceTax.value = parseFloat((tempValue*1.1)+0.5).format(0,3,',','.');
			}
		}else{
			objHidVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).toFixed(6);
			objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).format(2,3,',','.');	
		}
	}else if(num==0 && (parseFloat(objHidVPLPrice.value)!=parseFloat(objHidVPLPriceTax.value)) && (parseFloat(objHidVPLPrice.value)>0)){
		if(objHidVPLPriceID.value=='CUR01'){
			objHidVPLPriceTax.value = parseFloat(objHidVPLPrice.value).toFixed(6);
			objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)+0.5).format(0,3,',','.');	
		}else{
			objHidVPLPriceTax.value = parseFloat(objHidVPLPrice.value).toFixed(6);
			objVPLPriceTax.value = parseFloat(objHidVPLPrice.value).format(2,3,',','.');	
		}
	}
}
function fjs_updatePPNWeb(num){
	var objHidWebPriceID;
	var objWebPrice;
	var objWebVPLPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var tempWebPriceTax;
	var tempDiff;
	var tempValue;
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	if(num==1 && (parseFloat(objHidWebPrice.value)==parseFloat(objHidWebPriceTax.value)) && (parseFloat(objHidWebPrice.value)>0)){
		if(objHidWebPriceID.value=='CUR01'){
			objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).toFixed(6);
			objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)+0.5).format(0,3,',','.');	
			tempWebPriceTax = parseFloat(fjs_cekMoney(objWebPriceTax.value));
			if((tempWebPriceTax%100)>0){
				tempDiff = parseFloat(tempWebPriceTax%100);
				tempValue = (((parseFloat(tempWebPriceTax-tempDiff)-0.5))/1.1).toFixed(6);
				objHidWebPrice.value = parseFloat(tempValue).toFixed(6);
				objWebPrice.value = parseFloat(tempValue).format(2,3,',','.');
				objHidWebPriceTax.value = parseFloat(tempValue*1.1).toFixed(6);
				objWebPriceTax.value = parseFloat((tempValue*1.1)+0.5).format(0,3,',','.');
			}
		}else{
			objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).toFixed(6);
			objWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).format(2,3,',','.');	
		}
	}else if(num==0 && (parseFloat(objHidWebPrice.value)!=parseFloat(objHidWebPriceTax.value)) && (parseFloat(objHidWebPrice.value)>0)){
		if(objHidWebPriceID.value=='CUR01'){
			objHidWebPriceTax.value = parseFloat(objHidWebPrice.value).toFixed(6);
			objWebPriceTax.value = (parseFloat(objHidWebPrice.value)+0.5).format(0,3,',','.');	
		}else{
			objHidWebPriceTax.value = parseFloat(objHidWebPrice.value).toFixed(6);
			objWebPriceTax.value = parseFloat(objHidWebPrice.value).format(2,3,',','.');	
		}
	}
}

function fjs_updateWebPriceTax(){
	var objPPN;
	var objHidVPLPriceID;
	var objVPLPrice;
	var objHidVPLPrice;
	var objHidWebPriceID;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objMarginPct;
	var objMarginValue;
	var tempVPLPrice;
	var tempWebPriceTax;
	var marginValue;
	var objValueIDR;
	var objValueUSD;
	var objValueJPY;
	var num;
	var tempValue;
	objPPN = document.getElementsByName("radPPN");
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objValueIDR = document.getElementById("crValueIDR");
	objValueUSD = document.getElementById("crValueUSD");
	objValueJPY = document.getElementById("crValueJPY");
	if(objHidWebPriceID.value=='CUR01'){
		num = 1;
	}else if(objHidWebPriceID.value=='CUR02'){
		num = 2;
	}else if(objHidWebPriceID.value=='CUR03'){
		num = 3;	
	}
	if (num==1 && objHidVPLPriceID.value!='CUR01'){
		if (objHidVPLPriceID.value=='CUR02'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)*parseFloat(objValueUSD.value);
		}else if(objHidVPLPriceID.value=='CUR03'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)*parseFloat(objValueJPY.value);
		}
	}else if(num==2 && objHidVPLPriceID.value!='CUR02'){
		if (objHidVPLPriceID.value=='CUR01'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)/parseFloat(objValueUSD.value);
		}else if(objHidVPLPriceID.value=='CUR03'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)*parseFloat(objValueJPY.value);
			tempVPLPrice = parseFloat(tempVPLPrice)/parseFloat(objValueUSD.value);
		}
	}else if(num==3 && objHidVPLPriceID.value!='CUR03'){
		if (objHidVPLPriceID.value=='CUR01'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)/parseFloat(objValueJPY.value);
		}else if(objHidVPLPriceID.value=='CUR02'){
			tempVPLPrice = parseFloat(objHidVPLPrice.value)*parseFloat(objValueUSD.value);
			tempVPLPrice = parseFloat(tempVPLPrice)/parseFloat(objValueJPY.value);
		}
	}else{
		tempVPLPrice = parseFloat(objHidVPLPrice.value);
	}
	
	tempWebPriceTax = parseFloat(fjs_cekMoney(objWebPriceTax.value));
	objWebPriceTax.value = parseFloat(tempWebPriceTax).format(2,3,',','.');
	objHidWebPriceTax.value = parseFloat(tempWebPriceTax).toFixed(6);
	//alert(tempWebPriceTax);
	if(tempWebPriceTax>tempVPLPrice){
		if(objPPN[0].checked){
			marginValue = parseFloat(((tempWebPriceTax/1.1)-tempVPLPrice).toFixed(2));
			if(objHidWebPriceID.value=='CUR01'){
				tempValue = ((tempVPLPrice+marginValue)*1.1);
				if(tempValue>tempWebPriceTax){
					marginValue = parseFloat((((tempWebPriceTax/1.1)-0.5)-tempVPLPrice).toFixed(2));
				}
			}
		}else if(objPPN[1].checked){
			marginValue = tempWebPriceTax-tempVPLPrice;
		}
		objMarginValue.value = marginValue.format(2,3,',','.');
		objMarginPct.value = parseFloat((marginValue*100)/tempVPLPrice).format(2,3,',','.');
		objWebPrice.value = (tempVPLPrice+marginValue).format(2,3,',','.');
		objHidWebPrice.value = (tempVPLPrice+marginValue).toFixed(6);
	}
}
function fjs_cekVPL(){
	var objHidVPLPriceTax;
	var tempValue;
	var result;
	result = 0;
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	tempValue = parseFloat(fjs_cekMoney(objHidVPLPriceTax.value));
	if(tempValue>0){
		result = 1;
	}else{
		alert("Cek Your VPL Price!!!");	
	}
	return result;
}
function fjs_cekMargin(){
	var objMarginPct;
	var objMarginValue;
	var tempValue;
	var result;
	result = 0;
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	tempValue = parseFloat(fjs_cekMoney(objMarginPct.value));
	if(tempValue>0){
		tempValue = parseFloat(fjs_cekMoney(objMarginValue.value));
		if(tempValue>0){
			result = 1;
		}
	}
	return result;
}
//function fjs_PPN(num){
//	fjs_marginValue();
	/*var objHidVPLPriceTax;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objMarginValue;
	var bool;
	var margin;
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objWebPrice =  document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objMarginValue = document.getElementById("crMarginValue");
	bool = fjs_cekVPL();
	if (bool==1){
		bool = fjs_cekMargin();
		if(bool==1){
			margin = parseFloat(fjs_cekMoney(objMarginValue.value));
		}
	}*/
	/*if(parseFloat(objHidWebPrice.value)>0){
		if(num==1){
			objHidWebPriceTax.value = parseFloat(objHidWebPrice.value*1.1).toFixed(6);
			objWebPriceTax.value = parseFloat(objHidWebPrice.value*1.1).format(2,3,',','.');	
		}else{
			objHidWebPriceTax.value = parseFloat(objHidWebPrice.value).toFixed(6);
			objWebPriceTax.value = parseFloat(objHidWebPrice.value).format(2,3,',','.');	
		}
	}*/
//}
function fjs_Call(objCall){
	var objHidVPLPriceID;
	var objVPLPrice;
	var objHidVPLPrice;
	var objVPLPriceTax;
	var objHidVPLPriceTax;
	var objOOS;
	var objPPN;
	var objMarginPct;
	var objMarginValue;
	var objTdWebPriceID;
	var objWebPriceID;
	var objHidWebPriceID;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objTdSPriceID;
	var objSPriceID;
	var objHidSPriceID;
	var objSPrice;
	var objHidSPrice;
	var objSPriceTax;
	var objHidSPriceTax;
	var objSPriceStart;
	var objSPriceStartView;
	var objSPriceEnd;
	var objSPriceEndView;
	var objSPriceStartCalendar;
	var objSPriceEndCalendar;
	var objChkStsSameSPrice;
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objOOS = document.getElementById("chkOOS");
	objPPN = document.getElementsByName("radPPN");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objTdWebPriceID = document.getElementById("tdWebPriceID");
	objWebPriceID = document.getElementsByName("crWebPriceID");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objTdSPriceID = document.getElementById("tdSPriceID");
	objSPriceID = document.getElementsByName("crSPriceID");
	objHidSPriceID = document.getElementById("hidSPriceID");
	objSPrice = document.getElementById("crSPrice");
	objHidSPrice = document.getElementById("hidSPrice");
	objSPriceTax = document.getElementById("crSPriceTax");
	objHidSPriceTax = document.getElementById("hidSPriceTax");
	objSPriceStart = document.getElementById("CrSPValidStart");
	objSPriceStartView = document.getElementById("CrSPValidStartView");
	objSPriceEnd = document.getElementById("CrSPValidEnd");
	objSPriceEndView = document.getElementById("CrSPValidEndView");
	objSPriceStartCalendar = document.getElementById("tdSPriceStartCalendar");
	objSPriceEndCalendar = document.getElementById("tdSPriceEndCalendar");
	objChkStsSameSPrice = document.getElementById("crChkStsSameSPrice");
	if (objCall.checked==true){
		//objPPN[0].checked = true;
		objMarginPct.value = '0.00';
		objMarginValue.value = '0.00';
		objWebPriceID[0].checked = true;
		objHidWebPriceID.value =  'CUR01';
		objWebPrice.value = -1;
		objHidWebPrice.value = -1;
		objWebPriceTax.value = -1;
		objHidWebPriceTax.value = -1;
		objSPriceID[0].checked = true;
		objHidSPriceID.value =  'CUR01';
		objSPrice.value = '0.00';
		objHidSPrice.value = '0.000000';
		objSPriceTax.value = '0.00';
		objHidSPriceTax.value = '0.000000';
		objSPriceStart.value = '';
		objSPriceStartView.value = '';
		objSPriceEnd.value = '';
		objSPriceEndView.value = '';
		objChkStsSameSPrice.checked = false;
		objOOS.disabled = true;
		//objPPN[1].disabled = true;
		objMarginPct.readOnly = true;
		objMarginValue.readOnly = true;
		objTdWebPriceID.disabled = true;
		objWebPrice.disabled = true;
		objWebPriceTax.disabled = true;
		objTdSPriceID.disabled = true;
		objSPrice.disabled = true;
		objSPriceTax.disabled = true;
		objSPriceStartView.disabled = true;
		objSPriceEndView.disabled = true;
		objSPriceStartCalendar.disabled = true;
		objSPriceEndCalendar.disabled = true;
		objChkStsSameSPrice.disabled = true;
	}else{
		objOOS.disabled = false;
		//objPPN[1].disabled = false;
		objMarginPct.readOnly = false;
		objMarginValue.readOnly = false;
		objMarginPct.disabled = false;
		objMarginValue.disabled = false;
		objTdWebPriceID.disabled = false;
		objWebPrice.disabled = false;
		objWebPriceTax.disabled = false;
		objTdSPriceID.disabled = false;
		objSPrice.disabled = false;
		objSPriceTax.disabled = false;
		objSPriceStartView.disabled = false;
		objSPriceEndView.disabled = false;
		objSPriceStartCalendar.disabled = false;
		objSPriceEndCalendar.disabled = false;
		objChkStsSameSPrice.disabled = false;
		if(objHidVPLPriceID.value=='CUR01'){
			document.getElementById("crWebPriceIDIDR").checked = true;
		}else if(objHidVPLPriceID.value=='CUR02'){
			document.getElementById("crWebPriceIDUSD").checked = true;
		}else if(objHidVPLPriceID.value=='CUR03'){
			document.getElementById("crWebPriceIDJPY").checked = true;
		}
		objHidWebPriceID.value = objHidVPLPriceID.value;
		objWebPrice.value = objVPLPrice.value;
		objHidWebPrice.value = objHidVPLPrice.value;
		objWebPriceTax.value = objVPLPriceTax.value;
		objHidWebPriceTax.value = objHidVPLPriceTax.value;
	}
}
function fjs_OOS(objOOS){
	var objHidVPLPriceID;
	var objVPLPrice;
	var objHidVPLPrice;
	var objVPLPriceTax;
	var objHidVPLPriceTax;
	var objCall;
	var objPPN;
	var objMarginPct;
	var objMarginValue;
	var objTdWebPriceID;
	var objWebPriceID;
	var objHidWebPriceID;
	var objWebPrice;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var objTdSPriceID;
	var objSPriceID;
	var objHidSPriceID;
	var objSPrice;
	var objHidSPrice;
	var objSPriceTax;
	var objHidSPriceTax;
	var objSPriceStart;
	var objSPriceStartView;
	var objSPriceEnd;
	var objSPriceEndView;
	var objSPriceStartCalendar;
	var objSPriceEndCalendar;
	var objChkStsSameSPrice;
	objHidVPLPriceID = document.getElementById("hidVPLPriceID");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objCall = document.getElementById("chkCall");
	objPPN = document.getElementsByName("radPPN");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objTdWebPriceID = document.getElementById("tdWebPriceID");
	objWebPriceID = document.getElementsByName("crWebPriceID");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objTdSPriceID = document.getElementById("tdSPriceID");
	objSPriceID = document.getElementsByName("crSPriceID");
	objHidSPriceID = document.getElementById("hidSPriceID");
	objSPrice = document.getElementById("crSPrice");
	objHidSPrice = document.getElementById("hidSPrice");
	objSPriceTax = document.getElementById("crSPriceTax");
	objHidSPriceTax = document.getElementById("hidSPriceTax");
	objSPriceStart = document.getElementById("CrSPValidStart");
	objSPriceStartView = document.getElementById("CrSPValidStartView");
	objSPriceEnd = document.getElementById("CrSPValidEnd");
	objSPriceEndView = document.getElementById("CrSPValidEndView");
	objSPriceStartCalendar = document.getElementById("tdSPriceStartCalendar");
	objSPriceEndCalendar = document.getElementById("tdSPriceEndCalendar");
	objChkStsSameSPrice = document.getElementById("crChkStsSameSPrice");
	if (objOOS.checked==true){
		//objPPN[0].checked = true;
		objMarginPct.value = '0.00';
		objMarginValue.value = '0.00';
		objWebPriceID[0].checked = true;
		objHidWebPriceID.value =  'CUR01';
		objWebPrice.value = '0.00';
		objHidWebPrice.value = '0.000000';
		objWebPriceTax.value = '0.00';
		objHidWebPriceTax.value = '0.000000';
		objSPriceID[0].checked = true;
		objHidSPriceID.value =  'CUR01';
		objSPrice.value = '0.00';
		objHidSPrice.value = '0.000000';
		objSPriceTax.value = '0.00';
		objHidSPriceTax.value = '0.000000';
		objSPriceStart.value = '';
		objSPriceStartView.value = '';
		objSPriceEnd.value = '';
		objSPriceEndView.value = '';
		objChkStsSameSPrice.checked = false;
		objCall.disabled = true;
		//objPPN[1].disabled = true;
		objMarginPct.readOnly = true;
		objMarginValue.readOnly = true;
		objTdWebPriceID.disabled = true;
		objWebPrice.disabled = true;
		objWebPriceTax.disabled = true;
		objTdSPriceID.disabled = true;
		objSPrice.disabled = true;
		objSPriceTax.disabled = true;
		objSPriceStartView.disabled = true;
		objSPriceEndView.disabled = true;
		objSPriceStartCalendar.disabled = true;
		objSPriceEndCalendar.disabled = true;
		objChkStsSameSPrice.disabled = true;
	}else{
		objCall.disabled = false;
		//objPPN[1].disabled = false;
		objMarginPct.readOnly = false;
		objMarginValue.readOnly = false;
		objMarginPct.disabled = false;
		objMarginValue.disabled = false;
		objTdWebPriceID.disabled = false;
		objWebPrice.disabled = false;
		objWebPriceTax.disabled = false;
		objTdSPriceID.disabled = false;
		objSPrice.disabled = false;
		objSPriceTax.disabled = false;
		objSPriceStartView.disabled = false;
		objSPriceEndView.disabled = false;
		objSPriceStartCalendar.disabled = false;
		objSPriceEndCalendar.disabled = false;
		objChkStsSameSPrice.disabled = false;
		if(objHidVPLPriceID.value=='CUR01'){
			document.getElementById("crWebPriceIDIDR").checked = true;
		}else if(objHidVPLPriceID.value=='CUR02'){
			document.getElementById("crWebPriceIDUSD").checked = true;
		}else if(objHidVPLPriceID.value=='CUR03'){
			document.getElementById("crWebPriceIDJPY").checked = true;
		}
		objHidWebPriceID.value = objHidVPLPriceID.value;
		objWebPrice.value = objVPLPrice.value;
		objHidWebPrice.value = objHidVPLPrice.value;
		objWebPriceTax.value = objVPLPriceTax.value;
		objHidWebPriceTax.value = objHidVPLPriceTax.value;
	}
}
function fjs_marginPct(){
	var objHidVPLPrice;
	var objVPLPrice;
	var objHidVPLPriceTax;
	var objVPLPriceTax;
	var objMarginPct;
	var objMarginValue;
	var objWebPrice;
	var objHidWebPriceID;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var marginPct;
	var marginValue;
	var bool;
	var tempVPLPrice;
	var objPPN;
	var objCall;
	var objOOS;
	var tempRound;
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objPPN = document.getElementsByName("radPPN");
	objCall = document.getElementById("chkCall");
	objOOS = document.getElementById("chkOOS");
	bool = fjs_cekVPL();
	if(bool==1){
		if(parseFloat(fjs_cekMoney(objMarginPct.value))>0){
			marginPct = parseFloat(fjs_cekMoney(objMarginPct.value));
			marginValue = parseFloat((marginPct/100)*parseFloat(objHidVPLPrice.value)).toFixed(2);
			tempVPLPrice = parseFloat(objHidVPLPrice.value);
			//marginValue = parseFloat(fjs_cekMoney(objMarginValue.value));
			objMarginValue.value = parseFloat(marginValue).format(2,3,',','.');
			objHidWebPrice.value = parseFloat(tempVPLPrice+marginValue).toFixed(6);	
			objWebPrice.value = parseFloat(tempVPLPrice+marginValue).format(2,3,',','.');
			/*if(objHidWebPriceID.value=='CUR01'){
				if(objPPN[0].checked){*/
					/*objHidVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)).toFixed(6);
					if(parseFloat(objHidVPLPriceTax.value)%1>0){
						tempRound = parseFloat((parseFloat(objHidVPLPrice.value)*1.1)+0.5).format(0,3,',','.');
						if((parseFloat(fjs_cekMoney(tempRound))%100)>0){
							objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)-0.5).format(0,3,',','.');
						}else{
							objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)+0.5).format(0,3,',','.');
						}
					}else{
						objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)).format(0,3,',','.');
					}*/
					/*objHidWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)).toFixed(6);
					if(parseFloat(objHidWebPriceTax.value)%1>0){
						tempRound = parseFloat((parseFloat(objHidWebPrice.value)*1.1)+0.5).format(0,3,',','.');
						if((parseFloat(fjs_cekMoney(tempRound))%100)>0){
							objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)-0.5).format(0,3,',','.');
						}else{
							objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)+0.5).format(0,3,',','.');
						}
					}else{
						objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)).format(0,3,',','.');
					}
				}else{
					//objHidVPLPriceTax.value = objHidVPLPrice.value;
					//objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)+0.5).format(0,3,',','.');
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)+0.5).format(0,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)).toFixed(6);	
				}
			}else{
				if(objPPN[0].checked){
					//objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).format(2,3,',','.');
					//objHidVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).toFixed(6);
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).toFixed(6);
				}else{
					//objHidVPLPriceTax.value = objHidVPLPrice.value;
					//objVPLPriceTax.value = objVPLPrice.value;
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)).toFixed(6);
				}
			}*/
		}else{
			objMarginPct.value = '0.00';
			objMarginValue.value = '0.00';
		}
	}
}
function fjs_marginValue(){
	var objHidVPLPrice;
	var objVPLPrice;
	var objHidVPLPriceTax;
	var objVPLPriceTax;
	var objMarginPct;
	var objMarginValue;
	var objWebPrice;
	var objHidWebPriceID;
	var objHidWebPrice;
	var objWebPriceTax;
	var objHidWebPriceTax;
	var marginValue;
	var bool;
	var tempValue;
	var objPPN;
	var objCall;
	var objOOS;
	var tempRound;
	objHidVPLPrice = document.getElementById("hidVPLPrice");
	objVPLPrice = document.getElementById("crVPLPrice");
	objHidVPLPriceTax = document.getElementById("hidVPLPriceTax");
	objVPLPriceTax = document.getElementById("crVPLPriceTax");
	objWebPrice = document.getElementById("crWebPrice");
	objHidWebPriceID = document.getElementById("hidWebPriceID");
	objHidWebPrice = document.getElementById("hidWebPrice");
	objWebPriceTax = document.getElementById("crWebPriceTax");
	objHidWebPriceTax = document.getElementById("hidWebPriceTax");
	objMarginPct = document.getElementById("crMarginPct");
	objMarginValue = document.getElementById("crMarginValue");
	objPPN = document.getElementsByName("radPPN");
	objCall = document.getElementById("chkCall");
	objOOS = document.getElementById("chkOOS");
	bool = fjs_cekVPL();
	if((bool==1) && (objCall.checked==false) && (objOOS.checked==false)){
		if(objMarginValue.value!=''){
			tempValue = parseFloat(objHidVPLPrice.value);
			marginValue = parseFloat(fjs_cekMoney(objMarginValue.value));
			objMarginPct.value = parseFloat((marginValue*100)/tempValue).format(2,3,',','.');	
			objWebPrice.value = (tempValue+marginValue).format(2,3,',','.');
			objHidWebPrice.value = (tempValue+marginValue).toFixed(6);
			if(objHidWebPriceID.value=='CUR01'){
				if(objPPN[0].checked){
					/*objHidVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)).toFixed(6);
					if(parseFloat(objHidVPLPriceTax.value)%1>0){
						tempRound = parseFloat((parseFloat(objHidVPLPrice.value)*1.1)+0.5).format(0,3,',','.');
						if((parseFloat(fjs_cekMoney(tempRound))%100)>0){
							objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)-0.5).format(0,3,',','.');
						}else{
							objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)+0.5).format(0,3,',','.');
						}
					}else{
						objVPLPriceTax.value = ((parseFloat(objHidVPLPrice.value)*1.1)).format(0,3,',','.');
					}*/
					objHidWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)).toFixed(6);
					if(parseFloat(objHidWebPriceTax.value)%1>0){
						tempRound = parseFloat((parseFloat(objHidWebPrice.value)*1.1)+0.5).format(0,3,',','.');
						if((parseFloat(fjs_cekMoney(tempRound))%100)>0){
							objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)-0.5).format(0,3,',','.');
						}else{
							objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)+0.5).format(0,3,',','.');
						}
					}else{
						objWebPriceTax.value = ((parseFloat(objHidWebPrice.value)*1.1)).format(0,3,',','.');
					}
				}else{
					//objHidVPLPriceTax.value = objHidVPLPrice.value;
					//objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)+0.5).format(0,3,',','.');
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)+0.5).format(0,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)).toFixed(6);	
				}
			}else{
				if(objPPN[0].checked){
					//objVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).format(2,3,',','.');
					//objHidVPLPriceTax.value = (parseFloat(objHidVPLPrice.value)*1.1).toFixed(6);
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)*1.1).toFixed(6);
				}else{
					//objHidVPLPriceTax.value = objHidVPLPrice.value;
					//objVPLPriceTax.value = objVPLPrice.value;
					objWebPriceTax.value = (parseFloat(objHidWebPrice.value)).format(2,3,',','.');
					objHidWebPriceTax.value = (parseFloat(objHidWebPrice.value)).toFixed(6);
				}
			}
		}else{
			objMarginPct.value = '0.00';
			objMarginValue.value = '0.00';
		}
	}	
}
function fjs_SPriceID(num){
	var objSPrice;
	var objHidSPrice;
	var objSPriceTax;
	var objHidSPriceTax;
	var objHidSPriceID;
	var objValueIDR;
	var objValueUSD;
	var objValueJPY;
	var objPPN;
	var tempValue;
	var bool;
	objSPrice =  document.getElementById("crSPrice");
	objHidSPrice = document.getElementById("hidSPrice");
	objSPriceTax = document.getElementById("crSPriceTax");
	objHidSPriceTax = document.getElementById("hidSPriceTax");
	objHidSPriceID = document.getElementById("hidSPriceID");
	objPPN = document.getElementsByName("radPPN");
	objValueIDR = document.getElementById("crValueIDR");
	objValueUSD = document.getElementById("crValueUSD");
	objValueJPY = document.getElementById("crValueJPY");
	bool = false;
	if (num==1 && objHidSPriceID.value!='CUR01'){
		if (objHidSPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidSPrice.value)*parseFloat(objValueUSD.value);
		}else if(objHidSPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidSPrice.value)*parseFloat(objValueJPY.value);
		}else{
			if(parseFloat(fjs_cekMoney(objSPriceTax.value))>0){
				tempValue = parseFloat(fjs_cekMoney(objSPriceTax.value))/1.1;
			}else{
				tempValue = 0;
			}
		}
		bool = true;
		objHidSPriceID.value = 'CUR01';
	}else if(num==2 && objHidSPriceID.value!='CUR02'){
		if (objHidSPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidSPrice.value)/parseFloat(objValueUSD.value);
		}else if(objHidSPriceID.value=='CUR03'){
			tempValue = parseFloat(objHidSPrice.value)*parseFloat(objValueJPY.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueUSD.value);
		}else{
			if(parseFloat(fjs_cekMoney(objSPriceTax.value))>0){
				tempValue = parseFloat(fjs_cekMoney(objSPriceTax.value))/1.1;
			}else{
				tempValue = 0;
			}
		}
		bool = true;
		objHidSPriceID.value = 'CUR02';
	}else if(num==3 && objHidSPriceID.value!='CUR03'){
		if (objHidSPriceID.value=='CUR01'){
			tempValue = parseFloat(objHidSPrice.value)/parseFloat(objValueJPY.value);
		}else if(objHidSPriceID.value=='CUR02'){
			tempValue = parseFloat(objHidSPrice.value)*parseFloat(objValueUSD.value);
			tempValue = parseFloat(tempValue)/parseFloat(objValueJPY.value);
		}else{
			if(parseFloat(fjs_cekMoney(objSPriceTax.value))>0){
				tempValue = parseFloat(fjs_cekMoney(objSPriceTax.value))/1.1;
			}else{
				tempValue = 0;
			}
		}
		bool = true;
		objHidSPriceID.value = 'CUR03';
	}
	if(bool){
		objHidSPrice.value = tempValue.toFixed(6);
		objHidSPriceTax.value = (tempValue*1.1).toFixed(6);
		objSPrice.value = tempValue.format(2,3,',','.');
		if(objHidSPriceID.value=='CUR01'){
			objHidSPriceTax.value = ((parseFloat(tempValue)*1.1)).toFixed(6);
			if((parseFloat(objHidSPriceTax.value)%1)>0){
				objSPriceTax.value = ((parseFloat(tempValue)*1.1)+0.5).format(0,3,',','.');
				if((parseFloat(fjs_cekMoney(objSPriceTax.value))%100)>0){
					tempValue = ((((tempValue.toFixed(6))*1.1)-0.5)/1.1).toFixed(6);
					objHidSPrice.value = parseFloat(tempValue).toFixed(6);
					objSPrice.value = parseFloat(tempValue).format(2,3,',','.');
					objHidSPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
					objSPriceTax.value = ((parseFloat(tempValue)*1.1)+0.5).format(0,3,',','.');
				}
			}else{
				objSPriceTax.value = ((parseFloat(tempValue)*1.1)).format(0,3,',','.');
			}
		}else{
			objSPriceTax.value = (parseFloat(tempValue)*1.1).format(2,3,',','.');
			objHidSPriceTax.value = (parseFloat(tempValue)*1.1).toFixed(6);
		}
	}
}
function fjs_updateSPriceTax(){
	var objSPrice;
	var objHidSPrice;	
	var objHidSPriceID;
	var objSPriceTax;
	var objHidSPriceTax;
	var tempValue;
	var tempSPriceTax;
	objSPrice = document.getElementById("crSPrice");
	objHidSPrice = document.getElementById("hidSPrice");
	objHidSPriceID = document.getElementById("hidSPriceID");
	objSPriceTax = document.getElementById("crSPriceTax");
	objHidSPriceTax = document.getElementById("hidSPriceTax");
	if(objHidSPriceID.value==""){
		document.getElementById("crSPriceIDIDR").checked = true;
		document.getElementById("crSPriceIDIDR").click();
	}
	tempValue = parseFloat(fjs_cekMoney(objSPriceTax.value));
	if(tempValue>0){
		objHidSPriceTax.value = tempValue.toFixed(6);
		objHidSPrice.value = (tempValue/1.1).toFixed(6);
		if(objHidSPriceID.value=='CUR01'){
			//alert('OK');
			tempSPriceTax = ((parseFloat(objHidSPrice.value)*1.1)+0.5).toFixed(6);
			//alert(tempSPriceTax);
			//alert(tempSPriceTax);
			if(tempSPriceTax>parseFloat(objHidSPriceTax.value)){
				objHidSPrice.value = (((parseFloat(objHidSPrice.value)*1.1)-0.5)/1.1).toFixed(6);
				//alert(objHidSPrice.value);
			}
		}
		objSPrice.value = parseFloat(objHidSPrice.value).format(2,3,',','.');
		objSPriceTax.value = parseFloat(objHidSPriceTax.value).format(0,3,',','.');
	}else{
		objHidSPriceTax.value = 0;
		objHidSPrice.value = 0;
		objSPrice.value = 0;
		objSPriceTax.value = 0;
	}
}
Number.prototype.format = function(n, x, s, c) {
    var re = '\\d(?=(\\d{' + (x || 3) + '})+' + (n > 0 ? '\\D' : '$') + ')',
        num = this.toFixed(Math.max(0, ~~n));

    return (c ? num.replace('.', c) : num).replace(new RegExp(re, 'g'), '$&' + (s || ','));
};
function fjs_cekMoney(inputtxt){  
	var cekRegex = /[,]/g;
	var str;
	str = inputtxt;
	if(inputtxt.match(cekRegex)){
		str = inputtxt.replace(cekRegex,'');
	}
	return str;
}
function fjs_samePrdSPrice(obj){
	if(obj.checked){
		document.frm.CrSPValidStart.value=document.frm.CrPeriodStart.value;
		document.frm.CrSPValidEnd.value=document.frm.CrPeriodEnd.value;
		document.frm.CrSPValidStartView.value=document.frm.CrPeriodStartView.value;
		document.frm.CrSPValidEndView.value=document.frm.CrPeriodEndView.value;
	}else{
		document.frm.CrSPValidStart.value="";
		document.frm.CrSPValidEnd.value="";
		document.frm.CrSPValidStartView.value="";
		document.frm.CrSPValidEndView.value="";
	}
}
function fjs_cancel(){
	var tanya = confirm('Quit from this form?');
	var partID;
	
	partID = document.getElementById("crPartID").value;
	if ((tanya) && (partID!='')){
		document.location.replace('digoff_inve_prodcatalog_view.asp?crPartID='+partID);
	}else{
		window.close();	
	}
}
</SCRIPT>
<SCRIPT type="text/javascript">
	var clrWajibIsi = '<%=warnamustinput%>';
	var glob_submitClick = 0;
	var glob_currency_lmtdown_idr = 10000;
	var glob_currency_lmtup_idr = 50000000;
	var glob_currency_lmtdown_usd = 5;
	var glob_currency_lmtup_usd = 5000;
	var glob_currency_lmtdown_jpy =  100;
	var glob_currency_lmtup_jpy =  1000000;
	var indexLama;
		
	function fjs_chkdt(){
		var tipe = document.frm.crtypesave.value;
		var sts,cErrMsg,cekNewBrand;
		var cekDescOnly = 0;
		if((tipe=="save") || (tipe=="savenext") || (tipe=="savenew") || (tipe=="savecopy")){
			if(confirm("Save Data?")){
				cErrMsg = "";
				sts = true;
				if(typeof(document.frm.crActivation)=="object"){
					if(document.frm.crActivation.value==""){
						cErrMsg = cErrMsg + '- Activation Status \n';
					}
				}
				if(typeof(document.frm.crCatPrimaryID)=="object"){
					if(document.frm.crCatPrimaryID.value==""){
						cErrMsg = cErrMsg + "- Category Primary \n";
					}
				}
				if(typeof(document.frm.crBrandID)=="object"){
					if(document.frm.crBrandID.value==""){
						if(typeof(document.frm.crBrandNewCode)=="object"){
							if(document.frm.crBrandNewCode.value.length<=2){
								cErrMsg = cErrMsg + "- Brand Code Must be Three Characters \n";
							}
							if(document.frm.crBrandNewCode.value==""){
								cErrMsg = cErrMsg + "- Brand Code \n";
							}
						}
						if(typeof(document.frm.crBrandNewName)=="object"){
							if(document.frm.crBrandNewName.value.length<=3){
								cErrMsg = cErrMsg + "- Brand Name \n";
							}
							cekNewBrand = chk_brandName();
							if(cekNewBrand==false){
								cErrMsg = cErrMsg + "- Brand Name \n";
							}
						}
					}
				}
				if(typeof(document.frm.crSeri)=="object"){
					fjs_cleanSeri();
					if(document.frm.crSeri.value==""){
						cErrMsg = cErrMsg + "- Seri \n";
					}
				}
				if(typeof(document.frm.crDescOnly)=="object"){	
					objDesc = document.frm.crDescOnly.value;
					if(objDesc==""){
						cErrMsg = cErrMsg + "- Description \n";
					}else{	
						cekDesc = fjs_cekStringDesc(objDesc);
						if(cekDesc==1){
							cErrMsg = cErrMsg + "- Description \n";
						}else{
							cekSpaceDesc = fjs_cekSpaceDesc(objDesc);
							if(cekSpaceDesc==1){
								cekDescOnly = 1;
							}	
						}
					}
				}			
				if(typeof(document.frm.crShipWeight)=="object"){
					if((document.frm.crShipWeight.value=="") && (parseFloat(document.frm.crShipWeight.value)>0)){
						cErrMsg = cErrMsg + "- Ship Weight \n";
					}	
				}
				if(typeof(document.frm.crManufacturer)=="object"){
					if(document.frm.crManufacturer.value.length>=151){
						cErrMsg = cErrMsg + "- Manufacture Length more than 150 \n";
					}
				}
				if(typeof(document.frm.crMarketingInfo)=="object"){
					if(document.frm.crMarketingInfo.value.length>=601){
						cErrMsg = cErrMsg + "- Marketing Info Length more than 600 \n";
					}
				}	
				if(typeof(document.frm.crStatusID)=="object"){	
					if(document.frm.crStatusID.value!=""){
						var startDateStatus;
						var endDateStatus;
						startDateStatus = new Date(document.getElementsByName("CrPeriodStart")[0].value);
						endDateStatus = new Date(document.getElementsByName("CrPeriodEnd")[0].value);
						if (document.getElementsByName("CrPeriodStart")[0].value==""){
							cErrMsg = cErrMsg + "- Period Start \n";
						}
						if (document.getElementsByName("CrPeriodEnd")[0].value==""){
							cErrMsg = cErrMsg + "- Period End \n";
						}
						if(startDateStatus>endDateStatus){
							cErrMsg = cErrMsg + "- Period Start Date Must Less Than Period End Date \n";
						}
					}
				}
				if(typeof(document.frm.crPVndName)=="object"){
					if(document.getElementsByName("crPVndName")[0].value==""){
						cErrMsg = cErrMsg + "- VPL FROM \n";
					}
				}
				if(typeof(document.frm.crVPLPrice)=="object"){
					if(parseFloat(document.frm.hidVPLPrice.value)<=0){
						cErrMsg = cErrMsg +"- VPL Price > 0 \n";
					}	
				}
				if((typeof(document.frm.chkCall)=="object") && (typeof(document.frm.chkOOS)=="object")){
					if((document.getElementById("chkCall").checked==true) || (document.getElementById("chkOOS").checked==true)){
						if((parseFloat(document.getElementById("crMarginPct").value)>0) || (parseFloat(document.getElementById("crMarginValue").value)>0)){
							cErrMsg = cErrMsg + "- Margin \n";
						}
						if(parseFloat(document.getElementById("hidWebPrice").value)>0){
							cErrMsg = cErrMsg + "- Web Price \n";
						}
					}else{
						if((parseFloat(document.getElementById("crMarginPct").value)<=0) || (parseFloat(document.getElementById("crMarginValue").value)<=0)){
							cErrMsg = cErrMsg + "- Margin must be more than 0 \n";
						}
					}
				}
				if(typeof(document.frm.crSPrice)=="object"){
					if(parseFloat(document.frm.hidSPrice.value)>0){
						var startDateSP;
						var endDateSP;
						var currentDate;
						startDateSP = new Date(document.frm.CrSPValidStart.value);
						endDateSP = new Date(document.frm.CrSPValidEnd.value);
						currentDate = new Date();
						if(typeof(document.frm.hidSPriceID)=="object"){
							if(document.getElementsByName("hidSPriceID")[0].value==""){
								cErrMsg = cErrMsg + "- Special Price Currency \n";
							}
						}
						if(document.frm.CrSPValidStart.value==""){
							cErrMsg = cErrMsg + "- Special Price From Date \n";
						}
						if(document.frm.CrSPValidEnd.value==""){
							cErrMsg = cErrMsg + "- Special Price To Date \n";
						}
						if(startDateSP>endDateSP){
							cErrMsg = cErrMsg + "- Special Price From Date Must Less Than Special Price To Date \n";
						}
					}					
				}
				if(typeof(document.frm.CrSPValidStart)=="object"){
					var startDateSP;
					var endDateSP;
					var currentDate;
					startDateSP = new Date(document.frm.CrSPValidStart.value);
					endDateSP = new Date(document.frm.CrSPValidEnd.value);
					currentDate = new Date();
					if((document.frm.CrSPValidStart.value!="") || (document.frm.CrSPValidEnd.value!="")){
						if((parseFloat(document.frm.crSPrice.value)<=0) && (startDateSP<endDateSP) && (currentDate>=startDateSP) && (currentDate<=endDateSP)){
							cErrMsg = cErrMsg + "- Special Price \n";	
						}
					}			
				}
				if(cErrMsg!=""){
					sts = false;
					cErrMsg = "These field must fill before \n" + cErrMsg;
					alert(cErrMsg);			
				}else{
					generateStringPromo();
				}
				return sts;
			}
		}
	}
	
	function fjs_cleanSeri(){
		var strInput;
		var strInputNew;	
		strInput = document.getElementById("crSeri").value;
		strInputNew = fjs_cleanString(strInput);
		document.getElementById("crSeri").value = strInputNew;		
	}
	
	function fjs_cleanString(strInput){
		var strOutput;
		
		strOutput = strInput.replace(/^\s+|\s+$/gim,"");
		return strOutput;
	}
		
	function convertHtmlToText(strInput) {
		var inputText = strInput;
		var returnText = "" + inputText;
		returnText=returnText.replace(/\&#160;/gi,"");
		returnText=returnText.replace(/\&amp;/gi,"&");
		returnText=returnText.replace(/\&quot;/gi,'"');
		returnText=returnText.replace(/\&lt;/gi,'<');
		returnText=returnText.replace(/\&gt;/gi,'>');
		
		return returnText;
	}
	
	/*function fjs_catActiv(element){
		var text = element.options[element.selectedIndex].text;
		var active = document.getElementById ("crActive").value;
		cekRegex = /Dump/ig;
		if(text.match(cekRegex)){
			document.getElementsByName ("crActivation")[0].value = "3";
		}else{
			document.getElementsByName ("crActivation")[0].value = active;
		}
	}*/
	
	function fjs_cekStringDesc(inputtxt){  
		var cekRegex = /[<>$^*`~]/g;
		var hsl = 0 ;
		var str;
		if(inputtxt.match(cekRegex)){
			alert('Deleting illegal character' + inputtxt.match(cekRegex));
			str = inputtxt.replace(cekRegex,'');
			document.getElementById ("crDescOnly").value = str;
			fjs_descAll();
			hsl = 1;
		} 
		return hsl;
	}	
	
		function fjs_displayDate(val){
			if(document.frm.crStatusID.value==""){
				document.all.tr_periode.style.display="none";
			}
			else{
				document.all.tr_periode.style.display="block";
			}
		}
		
		
		function fjs_samePrdSPrice(obj){
			if(obj.checked){
				document.frm.CrSPValidStart.value=document.frm.CrPeriodStart.value;
				document.frm.CrSPValidEnd.value=document.frm.CrPeriodEnd.value;
				document.frm.CrSPValidStartView.value=document.frm.CrPeriodStartView.value;
				document.frm.CrSPValidEndView.value=document.frm.CrPeriodEndView.value;
			}
			else{
				document.frm.CrSPValidStart.value="";
				document.frm.CrSPValidEnd.value="";
				document.frm.CrSPValidStartView.value="";
				document.frm.CrSPValidEndView.value="";
			}
		}
		
		function fjs_getPVnd(){
			gjs_winopen('','gnrt_vendor.asp',610,300,0,0,true);
		}
		
		function fjs_getPVnd(){
			gjs_winopen('','gnrt_vendorAdd.asp',610,300,0,0,true);
		}
		
		function fjs_delPVnd(){
			document.frm.crPVndID.value="";
			document.frm.crPVndName.value="";
		}
		
		function fjs_clrVendor(cIdx){
			eval("document.frm.crVendorID"+cIdx+".value=''")
			eval("document.frm.crVendorName"+cIdx+".value=''")
			eval("document.all.objVendNmDet"+cIdx+".innerText=''")
		}
		
		function fjs_delSVndAdd(){
			document.frm.crPVndID.value="";
			document.frm.crPVndName.value="";
		}
		
		function fjs_getVendor(cIdx){
			gjs_winopen('','gnrt_vendor.asp?crParam='+cIdx,610,300,0,0,true);
		}
		
		function SetVendor(id,name,caddress,cidx){
			if(cidx!=""){
				eval("document.frm.crVendorID"+cidx+".value=id");
				eval("document.frm.crVendorID"+cidx+".focus()");
				eval("document.frm.crVendorName"+cidx+".value=name");
				eval("document.all.objVendNmDet"+cidx+".innerText=name");
			}
			else{
				document.frm.crPVndID.value=id;
				document.frm.crPVndName.value=name;

			}
			tutupwindow();
		}
		
		function SetVendorAdd(id,name,caddress,cidx){
			if(cidx!=""){
				eval("document.frm.crVendorID"+cidx+".value=id");
				eval("document.frm.crVendorID"+cidx+".focus()");
				eval("document.frm.crVendorName"+cidx+".value=name");
				eval("document.all.objVendNmDet"+cidx+".innerText=name");
			}
			else{
				document.frm.crPVndID.value=id;
				document.frm.crPVndName.value=name;
			}
			tutupwindow();
		}
		//add description
		function fjs_pupDescription(){
			var content;
			
			content = document.getElementById("crDesc").value;
			content = content.replace(/["]+/gim,"symDQ");
			gjs_winopen('Description','gnrt_descriptionInfo.asp?crParam='+content,750,150,550,200,true);
		
		}
		//end description
		//edit marketing
		function fjs_pupMarketing(){
			var content;
			
			content = document.getElementById("crMarketingInfo").value;
			//content = content.replace(/["]+/gim,"symDQ");
			gjs_winopen('Editor Marketing Info','gnrt_marketingInfo.asp?crParam='+content,980,550,350,0,true);	
		}
		//end marketing
		function SetMarketingInfo(cparam){
			document.getElementById("crMarketingInfo").value=cparam;
			alert("Marketing Info Has Been Saved Successfully");
			tutupwindow();
		}
		
		//Currency untuk VPL
		function fjs_formatCurrencyCntn(objCur,obj){
			var myOption0;
			myOption0 = -1;			
			for (i=objCur.length-1; i > -1; i--) {
				if (objCur[i].checked) {
					myOption0 = i; i = -1;
				}
			}
			
			if(myOption0 != -1 ){
				if(objCur[myOption0].value == "CUR01"){
					CntnPriceCurrID = "IDR";
				}else if(objCur[myOption0].value == "CUR02"){
					CntnPriceCurrID = "USD";
				}else{
					CntnPriceCurrID = "JPY";
				}
			}

			if (document.frm.hidCntnPriceCurrID.value!=""){
				var num = new NumberFormat();
				num.setInputDecimal('.');
				num.setCurrency(true);
				num.setCurrencyPosition(num.LEFT_OUTSIDE);
				num.setCurrencyValue('');
				num.setNegativeFormat(num.LEFT_DASH);
				num.setNegativeRed(false);
				num.setSeparators(true, ',', '<%=chr(0159)%>');
				if(CntnPriceCurrID.toLowerCase()=="idr"){
					num.setPlaces('0');
				}else{
					num.setPlaces('2');
				}
				num.setNumber(obj.value);
				obj.value=num.toFormatted().replace(/<%=chr(0159)%>/gi, ".");
				fjs_chklimitcurrency(CntnPriceCurrID.toLowerCase(),obj.value.replace(/,/gi,""));
			}else{
				alert("VPL Currency must selected before");
				obj.value="";
			}
		}
		
		function fjs_inputHidCntnPriceCurrID(objCurr){
			var myOption2;
			myOption2 = -1;
			for (i=objCurr.length-1; i > -1; i--) {
				if (objCurr[i].checked) {
					myOption2 = i; i = -1;
				}
			}
			document.frm.hidCntnPriceCurrID.value = objCurr[myOption2].value;
		}
		
		//Currency untuk Web Price
		function fjs_formatCurrencyWeb(objCur,obj){
			var myOption1;
			myOption1 = -1;
			for (i=objCur.length-1; i > -1; i--) {
				if (objCur[i].checked) {
					myOption1 = i; i = -1;
				}
			}
			
			if(myOption1 != -1){
				if(objCur[myOption1].value == "CUR01"){
					WebPriceCurrID = "IDR";
				}
				else if(objCur[myOption1].value == "CUR02"){
					WebPriceCurrID = "USD";
				}
				else{
					WebPriceCurrID = "JPY";
				}
			}

			if (document.frm.hidWebPriceCurrID.value!=""){
				var num = new NumberFormat();
				num.setInputDecimal('.');
				num.setCurrency(true);
				num.setCurrencyPosition(num.LEFT_OUTSIDE);
				num.setCurrencyValue('');
				num.setNegativeFormat(num.LEFT_DASH);
				num.setNegativeRed(false);
				num.setSeparators(true, ',', '<%=chr(0159)%>');
				if(WebPriceCurrID.toLowerCase()=="idr"){
					num.setPlaces('0');
				}else{
					num.setPlaces('2');
				}
				num.setNumber(obj.value);
				obj.value=num.toFormatted().replace(/<%=chr(0159)%>/gi, ".");
				fjs_chklimitcurrency(WebPriceCurrID.toLowerCase(),obj.value.replace(/,/gi,""));
			}else{
				if (myOption1 == -1){
					alert("Web Currency must selected before");
					obj.value="";
				}
			}
		}		
		
		function fjs_inputHidWebPriceCurrID(objCurr){
			var myOption4;
			myOption4 = -1;
			for (i=objCurr.length-1; i > -1; i--){
				if (objCurr[i].checked){
					myOption4 = i; i = -1;
				}
			}
			document.frm.hidWebPriceCurrID.value = objCurr[myOption4].value;
		}
		
		//Currency untuk Special Price
		function fjs_formatCurrencySp(objCur,obj){
			var myOption5;
			myOption5 = -1;
			for (i=objCur.length-1; i > -1; i--) {
				if (objCur[i].checked){
					myOption5 = i; i = -1;
				}
			}
			
			if(myOption5 != -1){
				if(objCur[myOption5].value == "CUR01"){
					SpPriceCurrID = "IDR";
				}else if(objCur[myOption5].value == "CUR02"){
					SpPriceCurrID = "USD";
				}else{
					SpPriceCurrID = "JPY";
				}
			}
			
			if (document.frm.hidSpPriceCurrID.value!=""){
				var num = new NumberFormat();
				num.setInputDecimal('.');
				num.setCurrency(true);
				num.setCurrencyPosition(num.LEFT_OUTSIDE);
				num.setCurrencyValue('');
				num.setNegativeFormat(num.LEFT_DASH);
				num.setNegativeRed(false);
				num.setSeparators(true, ',', '<%=chr(0159)%>');
				if(SpPriceCurrID.toLowerCase()=="idr"){
					num.setPlaces('2');
				}else{
					num.setPlaces('2');
				}
				num.setNumber(obj.value);
				obj.value=num.toFormatted().replace(/<%=chr(0159)%>/gi, ".");
				fjs_chklimitcurrency(SpPriceCurrID.toLowerCase(),obj.value.replace(/,/gi,""));
			}else {
				if (myOption5 == -1){
					alert("Special Price Currency must selected before");
					obj.value="";
				}
			}
		}	
			
		function fjs_inputHidSpPriceCurrID(objCurr){
			var myOption6;
			myOption6 = -1;
			for (i=objCurr.length-1; i > -1; i--){
				if (objCurr[i].checked){
					myOption6 = i; i = -1;
				}
			}
			document.frm.hidSpPriceCurrID.value = objCurr[myOption6].value;
		}
		
		function fjs_chklimitcurrency(currency,val){
			if(val!=""&&typeof(eval("glob_currency_lmtdown_"+currency))=="number"&&typeof(eval("glob_currency_lmtup_"+currency))=="number"){
				var sts=true;
				//check limit down
				if(eval("glob_currency_lmtdown_"+currency)>0&&val>0&&val<eval("glob_currency_lmtdown_"+currency)){
					sts=false;
					alert("This value is too small");
				}
				//check limit up
				if(sts&eval("glob_currency_lmtup_"+currency)>0&&val>0&&val>eval("glob_currency_lmtup_"+currency)){
					sts=false;
					alert("This value is too high");
				}
			}
		}
		
		/*function fjs_cancel(){
			var tanya = confirm('Quit from this form?');
			if(tanya){
				document.location.replace('digoff_inve_prodcatalog_view.asp?crPartID=<%=currpartid%>');
			}
		}*/
		
		function fjs_chgCat(val){
			s='<iframe src="digoff_inveprodcatalog_getbrand.asp?crCatID='+val+'" width=200 height=200 style="display:none"></iframe>';
			pp.innerHTML=s;
		}
		
		function fjs_chkBarcodeSeri(val){
			s='<iframe src="digoff_inveprodcatalog_chkbrand.asp?crCatID='+val+'" width=200 height=200 style="display:none"></iframe>';
			pp.innerHTML=s;
		}
		
		function fjs_chk50(obj,a){
			if (obj.value.length >=50){
				obj.style.background='#FFFF00';
				if (a=='x'){
					document.getElementById('saran').style.display='block';
				}
			}else{
				obj.style.background='#C5DEF7';
				if (a=='x'){
					document.getElementById('saran').style.display='none';
				}
			}
		}
		
		function fjs_chk50new(obj,jmlh,a){
			if (obj.value.length >=50){
				obj.style.background='#FFFF00';
				if (a=='x'){
					document.getElementById('saran'+jmlh).style.display='block';
				}
			}else{
				obj.style.background='#C5DEF7';
				if (a=='x'){
					document.getElementById('saran'+jmlh).style.display='none';
				}
			}
		}
		
		function fjs_changenew(){
			var tempChange = document.getElementsByName("crSeri").length;
			for(q=0;q<tempChange;q++){
				var frmBrand=fjs_getnamebrand();
				var frmSeri=document.getElementsByName("crSeri")[q].value;
				//var frmSeri=document.frm.crSeri.value
				var s="";
					if(frmBrand!=""){
						s=frmBrand.substr(0,3);
					}
					if(s.length<3){
						s+=replicatestr(" ",3-s.length);
					}
					if(frmSeri!=""){
						s+=frmSeri.substr(frmSeri.length-5,5);
					}
					document.getElementsByName("crBarcodeDesc")[q].value=s;
			}
		}
		
		function fjs_newBrand(){
			var objBrand = document.getElementById("crBrandID");
			if(objBrand.selectedIndex!=0){
				indexLama = objBrand.selectedIndex;
			}
			var objBrandNewCode = document.getElementById("crBrandNewCode");
			var objBrandNewName = document.getElementById("crBrandNewName");
			var optBrand = objBrand.options;
			if(document.getElementById("cbNewBrand").checked){
				objBrand.selectedIndex = 0;
				objBrand.className = "wordfield";
				objBrandNewCode.disabled = false;
				objBrandNewName.disabled = false;
				document.getElementById("crBrandNewCode").className = "wordfieldnormalmust";
				document.getElementById("crBrandNewName").className = "wordfieldnormalmust";
					
			}else{
				objBrand.selectedIndex = indexLama;
				objBrand.className = "wordfieldnormalmust";
				objBrandNewCode.value = "";
				objBrandNewName.value = "";
				objBrandNewCode.className = "wordfield";
				objBrandNewName.className = "wordfield";
				objBrandNewCode.disabled = true;
				objBrandNewName.disabled = true;
			}
		}
		
		function fjs_delbrandinbarcode(){
			var bnyk = document.getElementsByName("crSeri").length;
			for(b=0;b<bnyk;b++){
				var frmSeri=document.getElementsByName("crSeri")[b].value.substr(document.getElementsByName("crSeri")[b].value.length-5,5);
				document.getElementsByName("crBarcodeDesc")[b].value=frmSeri;
			}
		}
		
		function fjs_cloneobjsel(objSrc,objDes){
			var p=eval("document.frm."+objSrc+".length");
			for(var i=0;i<p;i++){
				eval("document.frm."+objDes+".options[i]=new Option(document.frm."+objSrc+".options[i].text,document.frm."+objSrc+".options[i].value)");
				if(eval("document.frm."+objSrc+".options[i].selected")){
					eval("document.frm."+objDes+".options[i].selected=true");
				}
			}
		}
		
		function replicatestr(chr,len){
			var s="";
			for(var i=0;i<len;i++){
				s+=chr;
			}
			return s;
		}
		
		function fjs_getnamebrand(){
			var s="";
			if (typeof(document.frm.crBrandID)!="object"){
				for(var i=0;i<document.frm.crBrandID.length;i++){
					if(document.frm.crBrandID[i].selected){
						s=document.frm.crBrandID[i].text;
						break;
					}
				}
			}
			else{
				s=document.frm.crBrandNameOri.value;
			}
			return s
		}
		
		function fjs_chkVendor(){
			var sts=true;
			var tempVendor = "";
			var jmlVendor=document.frm.crJmlVendor.value;
			for(var i=1;i<=jmlVendor;i++){
				var tempJumlah = 0;
				if(typeof(eval("document.frm.crVendorID"+i))=="object" && eval("document.frm.crVendorID"+i+".value") != ""){
					tempVendor = eval("document.frm.crVendorID"+i+".value");
					for(var a=1;a<=jmlVendor;a++){
						if(eval("document.frm.crVendorID"+a+".value") != ""){
							if(tempVendor == eval("document.frm.crVendorID"+a+".value")){
								tempJumlah++;
							}
						}
					}
					if(tempJumlah > 1){
						sts=false;
						alert("There are double Vendor Name !");
						break;
					}
				}
			}
			return sts;
		}

		function fjs_setlastupdateprc(){
			document.frm.crLastUpdate.value="1";
		}
		
		function fjs_setlastupdateprcVPL(){
			document.frm.crLastUpdateVPL.value="1";
		}
		
		function fjs_setlastupdateWebPrice(){
			document.frm.crLastUpdateWebPrice.value="1";
		}
		
		function fjs_setlastupdateSpecialPrice(){
			document.frm.crLastUpdateSpecialPrice.value="1";
		}
		
		function fjs_chgCustWar(cModelID){		
			s=eval("document.frm.crTypeCustWar"+cModelID+".value");
			if(s=="w"){
				eval("document.all.crCustLengthWarranty"+cModelID+".style.display='block'");
				eval("document.all.cust"+cModelID+".style.display='block'");
			}
			else{
				eval("document.all.crCustLengthWarranty"+cModelID+".style.display='none'");
				eval("document.all.cust"+cModelID+".style.display='none'");
			}
		}
		
		function fjs_chgWebPrice(obj){
			if(obj.value!=""){
				if(parseFloat(obj.value)>-1){
					fjs_setlastupdateWebPrice();fjs_formatCurrencyWeb(document.frm.crPrcCurrID,document.frm.crPrice);
				}
				else{
					alert("Invalid Price");
					gjs_defSelObj(obj,"");
				}
			}
		}		
		//untuk bundle
		function fjs_getSKU(cIdx){
			var cPartIDs=fjs_getpartids();
			gjs_winopen('','gnrt_product_bundle.asp?crPartIDs='+cPartIDs+'&crIdx='+cIdx,610,300,0,0,true);
	    }
		
		function fjs_getpartids(){
			var s;
			s=document.frm.crSKUID1.value;
			return s;
		}
		
		function fjs_setProduct(cpartid, cqtyavail,cIdx,cbrandname,cseri){	
			eval("document.frm.crSKUID"+cIdx+".value='"+cpartid+"'");
			eval("document.all.dBrandName"+cIdx+".innerText='"+cbrandname+"'");
			eval("document.all.dSeri"+cIdx+".innerText='"+cseri+"'");
			tutupwindow();
		}
		
		function chk_brandCode(){
			var upper;
			var res;
			var objRes;
			objRes = document.getElementById("crBrandNewCode");
			res = document.getElementById("crBrandNewCode").value;	
			upper = res.toUpperCase();
			objRes.value = upper;
		}
		
		function chk_brandName(){
			var jmlIndex;
			var objBrandID = document.getElementById("crBrandID");
			jmlIndex = objBrandID.length;
			var strNewBrand;
			var varIndex;
			var varOpt;
			var cek = true;
			strNewBrand = document.getElementById("crBrandNewName").value;
			strNewBrand = strNewBrand.replace(" ","");
			strNewBrand = strNewBrand.toLowerCase();
			for(x=1;x<jmlIndex;x++){
				varIndex = document.getElementById("crBrandID").selectedIndex=x;
				varOpt = document.getElementById("crBrandID").options;
				if (strNewBrand==varOpt[varIndex].text.toLowerCase()){
					document.getElementById("crBrandID").selectedIndex=0;
					cek = false;
					alert("Your Brand has been existed in Database!!!");
					break;	
				}
				
			}
			if(cek==true){
				document.getElementById("crBrandID").selectedIndex=0;
				cekNewBrand = true;
				return true;
			}else{
				document.getElementById("crBrandID").selectedIndex=0;
				cekNewBrand = false;
				return false;
			}
		}
		
		function isString(evt) {
   			var jml;
    		evt = (evt) ? evt : window.event;
    		var charCode = (evt.which) ? evt.which : evt.keyCode;
    		if (charCode>64) {
        		return true;
    		}else{
    			return false;
			}
    
}
		function fjs_descAll() {
			var str;
			var strResult;
			var strDesc;
			var extStr;
			var indxStr;
			
			str = document.getElementById("crDesc").value;
			strDesc = document.getElementById("crDescOnly").value;
			indxStr = str.indexOf("<br>")
			if (str =='' || indxStr == -1) {
				document.getElementById("crDesc").value=strDesc;
			} else if (indxStr != -1) {
				extStr = str.substring(0,indxStr);
				strResult = str.replace(extStr,strDesc);
				document.getElementById("crDesc").value=strResult;
			}
		}
	
		function fjs_cekSpaceDesc(inputtxt){  
			var cekRegex = /\s\s+/g;
			var hsl = 0 ;
			var str;
			if(inputtxt.match(cekRegex)){
				str = inputtxt.replace(cekRegex,' ');
				document.getElementById ("crDescOnly").value = str;
				fjs_descAll();
				hsl = 1;
				} else { 
			} return hsl;
		} 
//end description
</script>
<SCRIPT type="text/javascript">
	document.onkeydown = function(event){	
	if (!event) { /* This will happen in IE */
		event = window.event;
	}
	var keyCode = event.keyCode;
	if (keyCode == 8 && event.srcElement.readOnly == true) { 
		alert("You Can Not Do This");
		event.returnValue = false;
		return false;
	}
};	
</SCRIPT>
<SCRIPT type="text/javascript">
	window.onload = function () { 
		var flagCat;
		var flagBrand;
		var flagWarranty;
		if(<%=adaPromo%> >=1){
			addNewEventAutoBP(<%=bp%>);
			defineOnClickGenCalenderBP(<%=bp%>);
			addNewEventAutoSD(<%=sd%>);
			defineOnClickGenCalenderSD(<%=sd%>);
		}
		flagCat = document.getElementById("flagCat").value;
		flagBrand = document.getElementById("flagBrand").value;
		if(flagCat==0){
			setTimeout(function(){ 
				document.getElementById("btn_clkCat").click();
			}, 4000);	
		}
		if(flagBrand>=0){
			setTimeout(function(){ 
				document.getElementById("btn_clkBrand").click();
			}, 2500);	
		}
    }
</SCRIPT>
<!--#include file="../include/glob_conn_close.asp" -->
</html>

