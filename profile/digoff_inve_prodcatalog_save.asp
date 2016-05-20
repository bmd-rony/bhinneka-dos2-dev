<%
'response.buffer = true
currUAModule = "DO005"
currUASubModule = "DO005024"
currUACategory = "DO"
%>
<!--#include file="../include/gen_include.asp" -->
<!--#include file="../include/alpha_numeric_generator.asp" -->
<%
'server di live itu BMD kalau di lokal NavDevSWD
''DBServer = "GSSQLCL03.BMD"
DBServer = "NavDevSWD"	
function gen_partid()
	dim xstring,getRs,currID,cLong,id,idt
	cLong = 6
	xstring = "SELECT MAX(Substring(vPartID,4,3)+Substring(vPartID,9,3)) AS vID "&_
			  "FROM trx_inveComputer WHERE Substring(vPartID,7,2)='"&Right(Year(Date),2)&"';"
	set getRs = conn.execute(xstring)
	IF NOT getRs.EOF THEN
		currID = getRs("vID")
		IF isNull(currID) THEN
			id = String(cLong-1,"0")&"1"
		ELSE
			idt = CStr(CLng(currID)+1)
			id = String(cLong-Len(idt),"0")&idt
		END IF
	ELSE
		id = String(cLong-1,"0")&"1"
	END IF
	getRs.Close
	set getRs = nothing
	id = "SKU"&Left(id,3)&Right(Year(Date),2)&Right(id,3)
	gen_partid = id
end function
	
function gen_id()
	dim xstring,getRs,cLong,id,idt,bln
	cLong=4
	xstring = " select max(substring(vtrxid,8,4)) vid "&_
			  " from trx_invecpbundle where substring(vtrxid,4,2)='"&right(year(date),2)&"' and convert(int,substring(vtrxid,6,2))="&month(date)&""
	set getRs = conn.execute(xstring)
	IF NOT getRs.eof THEN
		IF isnull(getRs(0)) THEN
			id = string(cLong-1,"0")&"1"
		ELSE
			idt = cstr(cint(getRs(0))+1)
			id = string(cLong-len(idt),"0")&idt
		END IF
	ELSE
		id = string(cLong-1,"0")&"1"
	END IF
	getRs.close
	set getRs = nothing
	bln = month(date)
	IF len(cstr(bln))<2 THEN 
		bln = "0"+cstr(bln)
	END IF
	id = "BND"&right(year(date),2)&bln&id
	gen_id = id
end function
	
function getkurs(cKurID)
	dim xstring,currRate
	xstring = " select top 1 vCurrValue from trx_currency  "&_
			  " where vCurrID='"&cKurID&"' order by vCurrDate desc"
	set getRsKurs = conn.execute(xstring)
	IF NOT getRsKurs.eof THEN
		currRate=getRsKurs("vCurrValue")
	END IF
	getRsKurs.close
	set getRsKurs = nothing
	getkurs=currRate
end function

FUNCTION stripHTML(strHTML)
'Strips the HTML tags from strHTML

  Dim objRegExp, strOutput
  SET objRegExp = New Regexp

  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|\n)+?>"

  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  SET objRegExp = NOTHING
  'Replace all < and > with &lt; and &gt;
  '''strOutput = Replace(strOutput, "<", "&lt;")
  '''strOutput = Replace(strOutput, ">", "&gt;")
  
  stripHTML = strOutput    'Return the value of strOutput
END FUNCTION
jam = Time()
waktu = Hour(jam)
conn.BeginTrans
conn.Execute "SET ARITHABORT ON"
	currNowDate = Year(Now)&"-"&Month(Now)&"-"&Day(Now)&" "&Hour(Now)&":"&Minute(Now)&":"&Second(Now)&".000"
	currCreatorNo = currLoginID
	currCreatorDateTime = currNowDate
	currCreatorIP = Request.ServerVariables("REMOTE_ADDR")
	currEditorNo = currLoginID
	currEditorTime = currNowDate
	currEditorIP = Request.ServerVariables("REMOTE_ADDR")
	
	currAct = Request.Form("crAct")
	currCopy = Request.Form("crCopy")
	currPartID = Request.Form("crPartID")
	currActivation = CInt(Request.Form("crActivation"))
	currCatPrimaryID = Request.Form("crCatPrimaryID")
	currBrandID = Request.Form("crBrandID")
	currBrandCode = Request.Form("crBrandNewCode")
	currBrandName = Request.Form("crBrandNewName")	
	currBrandNameOri = Request.Form("crBrandNameOri")
	currSeri = Trim(Request.Form("crSeri"))
	currDesc = Request.Form("crDesc")
	currDescPO = Replace(Replace(currDesc,"'",""),""""," inch")
	currShipWeight = CDbl(Request.Form("crShipWeight"))
	currAuthorisedWarranty = Request.Form("crAuthWarranty")
	currWarranty00001 = Request.Form("crTypeCustWar00001")
	currLengthWarranty00001 = Request.Form("crCustLengthWarranty00001")
	currWarranty00002 = Request.Form("crTypeCustWar00002")
	currLengthWarranty00002 = Request.Form("crCustLengthWarranty00002")
	currExtTxtWarranty = Request.Form("crExtTxtWarr")
	currMarketinginfo = Replace(Replace(Trim(Request.Form("crMarketinginfo")),"<b>",""),"</b>","")
	currNote = Request.Form("crNote")
	currManufacturer = Request.Form("crManufacturer")
	currStatusID = CStr(Request.Form("crStatusID"))
	currPeriodStart = Request.Form("CrPeriodStart")
	currPeriodEnd = Request.Form("CrPeriodEnd")
	currJaminanMurah = Request.Form("crJaminanMurah")
	
	currPVndID = Request.Form("crPVndID")
	currPVndName = Request.Form("crPVndName")
	currVPLPriceID = Request.Form("hidVPLPriceID")
	currVPLPrice = CDbl(Request.Form("hidVPLPrice"))
	currChkCall = Request.Form("chkCall")
	currChkOOS = Request.Form("chkOOS")
	currStsPPN = Request.Form("radPPN")
	currMarginPct = CDbl(Request.Form("crMarginPct"))
	currMarginValue = CDbl(Request.Form("crMarginValue"))
	currWebPriceID = Request.Form("hidWebPriceID")
	currWebPrice = CDbl(Request.Form("hidWebPrice"))
	currSPriceID = Request.Form("hidSPriceID")
	currSPrice = CDbl(Request.Form("hidSPrice"))
	currStartSPrice = Request.Form("CrSPValidStart")
	currEndSPrice = Request.Form("CrSPValidEnd")
	
	currNeedSN = Request.Form("crNeedSN")
	currGoodDesc = Request.Form("crGoodDesc")
	currJmlVendor = Request.Form("crJmlVendor")
	currVendorIDType = Request.Form("crVndTypeh")
	currTypeSave = Request.Form("crTypeSave")
	currValueIDR = CDbl(Request.Form("crValueIDR"))
	currValueUSD = CDbl(Request.Form("crValueUSD"))
	currValueJPY = CDbl(Request.Form("crValueJPY"))
	currTypeVndWar = "12"
	currVndModelID1 = "00001"
	currVndModelID2 = "00002"
	currVndLengthWarranty1 = 12
	currVndLengthWarranty2 = 12
	currLastUpdateVPL = CStr(Request.Form("crLastUpdateVPL"))
	currLastUpdateWeb = CStr(Request.Form("crLastUpdateWebPrice"))
	currLastUpdateSP = CStr(Request.Form("crLastUpdateSpecialPrice"))
	
	tempTypeBP = Split(Request.Form("nTypeBP"),"$%$")
	tempSKUBP = Split(Request.Form("nSKUBP"),"$%$")
	tempBrandBP = Split(Request.Form("nBrandBP"),"$%$")
	tempSeriBP = Split(Request.Form("nSeriBP"),"$%$")
	tempDescBP = Split(Request.Form("nDescBP"),"$%$")
	tempQtyBP = Split(Request.Form("nQtyBP"),"$%$")
	tempStartDateBP = Split(Request.Form("nStartDateBP"),"$%$")
	tempEndDateBP = Split(Request.Form("nEndDateBP"),"$%$")
	tempTagBP = Split(Request.Form("nTagBP"),"$%$")
	tempTypeSD = Split(Request.Form("nTypeSD"),"$%$")
	tempSKUSD = Split(Request.Form("nSKUSD"),"$%$")
	tempBrandSD = Split(Request.Form("nBrandSD"),"$%$")
	tempSeriSD = Split(Request.Form("nSeriSD"),"$%$")
	tempDescSD = Split(Request.Form("nDescSD"),"$%$")
	tempQtySD = Split(Request.Form("nQtySD"),"$%$")
	tempStartDateSD = Split(Request.Form("nStartDateSD"),"$%$")
	tempEndDateSD = Split(Request.Form("nEndDateSD"),"$%$")
	tempTagSD = Split(Request.Form("nTagSD"),"$%$")
	xSQL_DOS2 = ""
	xSQL_BMD = ""
	xSQL_WebPrice = ""
	xSQL_SPrice = ""
	xSQL_EditorPrice = ""
	
	SET rsNeedSN = Server.CreateObject("ADODB.Recordset")
	queryNeedSN = "SELECT * FROM tlu_InveCPCategory WHERE vCatID='"&currCatPrimaryID&"';"
	rsNeedSN.Open queryNeedSN,conn,3,1,0
	IF NOT rsNeedSN.EOF THEN
		currNeedSN = rsNeedSN("vNeedSN")
		currGoodDesc = rsNeedSN("vGoodsDescription")
	END IF
	IF (currNeedSN) THEN
		currNeedSN = 1
	ELSE
		currNeedSN = 0
	END IF
	rsNeedSN.close
	SET rsNeedSN = NOTHING
	
	IF (LCase(currChkCall)="on") THEN
		currStatusCall = 1
	ELSEIF (LCase(currChkOOS)="on") THEN
		currStatusOOS = 1
	ELSE
		currStatusCall = 0
		currStatusOOS = 0	
	END IF
	IF ((currStatusCall=0) AND (currStatusOOS=0)) THEN
		SELECT CASE currVPLPriceID
		CASE "CUR01"
			currVPLValue = currValueIDR
		CASE "CUR02"
			currVPLValue = currValueUSD
		CASE "CUR03"
			currVPLValue = currValueJPY
		END SELECT
		
		currVPLPriceValue = currVPLPrice*currVPLValue
			
		SELECT CASE currWebPriceID
		CASE "CUR01"
			currWebValue = currValueIDR
		CASE "CUR02"
			currWebValue = currValueUSD
		CASE "CUR03"
			currWebValue = currValueJPY
		END SELECT
		
		currWebPriceValue = currWebPrice*currWebValue
		
		SELECT CASE currSPriceID
		CASE "CUR01"
			currSPValue = currValueIDR
		CASE "CUR02"
			currSPValue = currValueUSD
		CASE "CUR03"
			currSPValue = currValueJPY
		END SELECT
		
		currSPriceValue = currSPrice*currSPValue
			
		IF (currWebPriceValue>0)  THEN
			IF currWebPriceValue<currVPLPriceValue THEN
				Response.Write "<script language=javascript>"
				Response.Write "alert('\can not save, Price is lower than Vendor Price.');"
				Response.Write "history.back();"
				Response.Write "</script>"
				Response.End
			END IF
		END IF
	END IF
	IF (currBrandID<>"") THEN
		currBrandCode = ""
		currBrandName = ""
		brandName = currBrandNameOri
	ELSE
		brandName = currBrandName
	END IF
	IF (currJaminanMurah="on") THEN
		currJaminanMurah = 1
	ELSE
		currJaminanMurah = 0
	END IF
	IF (currAuthorisedWarranty="on") THEN
		currAuthorisedWarranty = 1
	ELSE
		currAuthorisedWarranty = 0
	END IF
	SELECT CASE currWarranty00001
	CASE "W"
		currLengthWarranty00001 = currLengthWarranty00001
	CASE "N"
		currLengthWarranty00001 = 0
	CASE "L"
		currLengthWarranty00001 = 60
	CASE "12"
		currLengthWarranty00001 = 12
	CASE "36"
		currLengthWarranty00001 = 36
	END SELECT
	SELECT CASE currWarranty00002
	CASE "W"
		currLengthWarranty00002 = currLengthWarranty00002
	CASE "N"
		currLengthWarranty00002 = 0
	CASE "L"
		currLengthWarranty00002 = 60
	CASE "12"
		currLengthWarranty00002 = 12
	CASE "36"
		currLengthWarranty00002 = 36
	END SELECT
	IF (currCatSecondaryID<>"") THEN 
		currCatSecondaryID = Left(currCatSecondaryID,Len(currCatSecondaryID)-1) 
	END IF
	IF (currVendorID<>"") THEN 
		currVendorID = Left(currVendorID,Len(currVendorID)-1) 
	END IF
	IF ((currAct="") OR (currAct="add")) THEN 
		currAct="Add" 
	END IF
	IF ((currStatusID<>"") AND (currCopy<>"Yes")) THEN
		 currPeriodStart = Year(CDate(currPeriodStart))&"-"&Month(CDate(currPeriodStart))&"-"&Day(CDate(currPeriodStart))&" 00:00:00.000"
		 currPeriodEnd = Year(CDate(currPeriodEnd))&"-"&Month(CDate(currPeriodEnd))&"-"&Day(CDate(currPeriodEnd))&" 23:59:59.000"
	ELSE
		currPeriodStart = ""
		currPeriodEnd = ""
	END IF
	IF (currSPrice>0) THEN
		currStartSPrice = Year(CDate(currStartSPrice))&"-"&Month(CDate(currStartSPrice))&"-"&Day(CDate(currStartSPrice))&" 00:00:00.000"
		currEndSPrice = Year(CDate(currEndSPrice))&"-"&Month(CDate(currEndSPrice))&"-"&Day(CDate(currEndSPrice))&" 23:59:59.000"
	ELSE
		currStartSPrice = ""
		currEndSPrice = ""
	END IF
	IF ((LEN(currBrandCode)=3) AND (LEN(currBrandName)>=2)) THEN
		IF ISNumeric(Left(currBrandName,1)) THEN 
			alpha = "9" 
		ELSE 
			alpha = Left(currBrandName,1)
		END IF
		sqlCondition = " LEFT(vBrandID,1)='"&alpha&"' "
		currBrandID = alpha & Gen_AlphaNumeric("tlu_inveCPBrand","vBrandID",SqlCondition,conn,2)
		
		xSQL_DOS2 = xSQL_DOS2&" IF NOT EXISTS(SELECT top 1 vName FROM tlu_inveCPBrand WHERE vName="&replaceQuota(currBrandName,"str") & ") BEGIN" 
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO tlu_inveCPBrand(vBrandID,vCode,vName,vCreatorNo," & _
			   "vCreatorDateTime,vCreatorIP) VALUES ("&replaceQuota(currBrandID,"str")&"," & _
			   replaceQuota(currBrandCode,"str")&","&replaceQuota(currBrandName,"str") &_
			   ","&replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&"," & _
			   replaceQuota(currCreatorIP,"str")&");END"
		xSQL_BMD = xSQL_BMD&"IF NOT EXISTS(SELECT top 1 [Name] FROM "&DBServer&".[dbo].[BMD-Debug$Manufacturer] WHERE [Name]='"&replaceQuota(currBrandName,"str")&"') BEGIN" 
		xSQL_BMD = xSQL_BMD&" INSERT INTO "&DBServer&".[dbo].[BMD-Debug$Manufacturer Buffer]" & _
			   "([Code],[Name],[Kode Brand],[Nama Brand],[Creator ID],[Creator Date],[Creator IP],[Editor ID],[Editor Date],[Editor IP],[Import Status])" & _
			   " VALUES ((upper(''"&currBrandID&"'')),'"&replaceQuota(currBrandName,"str")&"',(upper(''"&currBrandCode&"'')),'''','''',''1753-01-01'','''','''',''1753-01-01'','''',1);END"
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO tlu_inveCPBrand_relation (vBrandID, vCatID, vRelasiID,vCreatorNo,vCreatorDateTime,vCreatorIP)"&_
			   " VALUES ('" & currBrandID & "'," & replaceQuota(currCatPrimaryID, "str")&", 1,"&_
			   " "&replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime, "str")&","&replaceQuota(currCreatorIP, "str")&");" 				
	ELSE
		xSQL_DOS2 = xSQL_DOS2&" IF NOT EXISTS(SELECT '1' FROM tlu_inveCPBrand_relation WHERE vCatID='"&currCatPrimaryID&"' AND vBrandID='"&currBrandID&"') "&_
			   "BEGIN "&_
			   "INSERT INTO tlu_inveCPBrand_relation(vBrandID,vCatID,vRelasiID,vCreatorNo,vCreatorDateTime,vCreatorIP) "&_
			   "VALUES ('"&currBrandID&"',"&replaceQuota(currCatPrimaryID,"str")&",1,"&_
			   replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&"); "&_
			   "END"
	END IF
	'''ImgName = brandName&"-"&Replace(currSeri," ","-")&"-"&currPartID&".jpg"
	strImgPath = "\\script\dos2_ho\data\inventory\thumbnail_product\"&brandName&"-"&Replace(currSeri," ","-")&"-"&currPartID&".jpg"
	sqlCategoryName = "SELECT vCatID,vName,vLevelPos FROM tlu_InveCPCategory WHERE vCatID='"&currCatPrimaryID&"';"
	SET CategoryNameValue = conn.execute(sqlCategoryName)
	IF NOT CategoryNameValue.EOF THEN
		levelID = CategoryNameValue("vCatID")
		levelCat = CategoryNameValue("vLevelPos")
		catName = CategoryNameValue("vName")
	END IF
	CategoryNameValue.Close:SET CategoryNameValue = Nothing	
	IF (levelCat=4) THEN
		sqlCategoryChildName = "SELECT (SELECT vName FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl3,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl3ID,"&_
			"(SELECT vName FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_ 
			"WHERE vCatID=cat.vParCatID)) AS lvl2,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_ 
			"WHERE vCatID=cat.vParCatID)) AS lvl2ID,"&_
			"(SELECT vName FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_ 
			"WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID))) AS lvl1,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_ 
			"WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID))) AS lvl1ID "&_
			"FROM tlu_InveCPCategory AS cat WHERE vCatID ='"&currCatPrimaryID&"';"
		SET CategoryChildNameValue = conn.execute(sqlCategoryChildName)	
		IF NOT CategoryChildNameValue.EOF THEN
			catLvl4Name	= "'"&catName&"'"
			catLvl4ID 	= "'"&levelID&"'"
			catLvl3Name	= "'"&CategoryChildNameValue("lvl3")&"'"
			catLvl3ID	= "'"&CategoryChildNameValue("lvl3ID")&"'"
			catLvl2Name	= "'"&CategoryChildNameValue("lvl2")&"'"
			catLvl2ID	= "'"&CategoryChildNameValue("lvl2ID")&"'"
			catLvl1Name	= "'"&CategoryChildNameValue("lvl1")&"'"
			catLvl1ID	= "'"&CategoryChildNameValue("lvl1ID")&"'"
		END IF
		CategoryChildNameValue.Close:SET CategoryChildNameValue = Nothing
	END IF
	IF (levelCat=3) THEN
		sqlCategoryChildName = "SELECT (SELECT vName FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl2,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl2ID,"&_
			"(SELECT vName FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_
			"WHERE vCatID=cat.vParCatID)) AS lvl1,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=(SELECT vParCatID FROM tlu_InveCPCategory "&_
			"WHERE vCatID=cat.vParCatID)) AS lvl1ID "&_
			"FROM tlu_InveCPCategory AS cat WHERE vCatID ='"&currCatPrimaryID&"';"
		SET CategoryChildNameValue=conn.execute(sqlCategoryChildName)
	
		IF NOT CategoryChildNameValue.EOF THEN
			catLvl4Name = "''"
			catLvl4ID 	= "''"
			catLvl3Name	= "'"&catName&"'"
			catLvl3ID	= "'"&levelID&"'"
			catLvl2Name	= "'"&CategoryChildNameValue("lvl2")&"'"
			catLvl2ID	= "'"&CategoryChildNameValue("lvl2ID")&"'"
			catLvl1Name	= "'"&CategoryChildNameValue("lvl1")&"'"
			catLvl1ID	= "'"&CategoryChildNameValue("lvl1ID")&"'"
		END IF
		CategoryChildNameValue.Close:SET CategoryChildNameValue = NOTHING
	END IF
	IF (levelCat=2) THEN
		sqlCategoryChildName = "SELECT (SELECT vName FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl1,"&_
			"(SELECT vCatID FROM tlu_InveCPCategory WHERE vCatID=cat.vParCatID) AS lvl1ID "&_
			"FROM tlu_InveCPCategory AS cat WHERE vCatID ='"&currCatPrimaryID&"';"
		SET CategoryChildNameValue = conn.execute(sqlCategoryChildName)
		IF NOT CategoryChildNameValue.EOF THEN
			catLvl4Name	= "''"
			catLvl4ID	= "''"
			catLvl3Name	= "''"
			catLvl3ID	= "''"
			catLvl2Name = "'"&catName&"'"
			catLvl2ID	= "'"&levelID&"'"
			catLvl1Name = "'"&CategoryChildNameValue("lvl1")&"'"
			catLvl1ID	= "'"&CategoryChildNameValue("lvl1ID")&"'"
		END IF
		CategoryChildNameValue.Close:SET CategoryChildNameValue = NOTHING
	END IF	
	IF (levelCat=1) THEN
		catLvl4Name	= "''"
		catLvl4ID 	= "''"
		catLvl3Name	= "''"
		catLvl3ID 	= "''"
		catLvl2Name	= "''"
		catLvl2ID 	= "''"
		catLvl1Name = "'"&catName&"'"
		catLvl1ID 	= "'"&levelID&"'"
	END IF
	seri1 = Mid(currSeri,1,30)	
	IF LEN(currSeri)>30 THEN
		seri2 = "'"&Mid(currSeri,31,30)&"'"
	ELSE
		seri2 = "''"
	END IF	
	replaceDescTagHTML = stripHTML(currDesc)
	desc1 = Mid(replaceDescTagHTML,1,250)
	IF LEN(replaceDescTagHTML)>250 THEN
		desc2 = "'"&Mid(replaceDescTagHTML,251,250)&"'"
	ELSE
		desc2 = "''"
	END IF	
	IF LEN(replaceDescTagHTML)>500 THEN
		desc3 = "'"&Mid(replaceDescTagHTML,501,250)&"'"
	ELSE
		desc3 = "''"
	END IF	
	IF LEN(replaceDescTagHTML)>750 THEN
		desc4 = "'"&Mid(replaceDescTagHTML,751,250)&"'"
	ELSE
		desc4 = "''"
	END IF	
	IF LEN(note1)>1 THEN
		note1 = "'"&Mid(currNote,1,250)&"'"
	ELSE
		note1 = "''"
	END IF	
	IF LEN(note2)>250 THEN
		note2 = Mid(currNote,251,250)
	ELSE
		note2 = "''"
	END IF	
	IF LEN(marketingInfo1)>1 THEN
		marketingInfo1 = "'"&Mid(currMarketinginfo,1,250)&"'"
	ELSE
		marketingInfo1 = "''"
	END IF	
	IF LEN(marketingInfo2)>250 THEN
		marketingInfo2 = "'"&Mid(currMarketinginfo,251,250)&"'"
	ELSE
		marketingInfo2 = "''"
	END IF	
	strImgPath1 = Mid(strImgPath,1,150)
	IF LEN(currGoodDesc<3) THEN
		currGoodDesc = "''"
	END IF
	IF (LCase(currAct)="add") THEN
		currPartID = gen_partid()	
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveComputer(vPartID,vBrandID,vSeri,vDesc,vStatusID,vStartPeriod,"&_
			"vEndPeriod,vShipWeight,vManufact,vActivation,vMarketingInfo,vNote,vBrandIDPO,vSeriPO,vDescPO,vCreatorNo,"&_
			"vCreatorDateTime,vCreatorIP,vJaminanMurah,vAuthorisedWarranty,vExtTxtWarranty) "&_
			"VALUES ("&replaceQuota(currPartID,"str")&","&replaceQuota(currBrandID,"str")&","&_
			replaceQuota(currSeri,"str")&","&replaceQuota(currDesc,"str")&","&_
			replaceQuota(currStatusID,"num")&","&replaceQuota(currPeriodStart,"str")&","&replaceQuota(currPeriodEnd,"str")&","&_
			replaceQuota(currShipWeight,"num")&","&replaceQuota(currManufacturer,"str")&","&replaceQuota(currActivation,"num")&","&_
			replaceQuota(currMarketinginfo,"str")&","&replaceQuota(currNote,"str")&","&replaceQuota(currBrandID,"str")&","&_
			replaceQuota(currSeri,"str")&","&replaceQuota(currDescPO,"str")&","&replaceQuota(currCreatorNo,"str")&","&_
			replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&","&replaceQuota(currJaminanMurah,"num")&","&_
			replaceQuota(currAuthorisedWarranty,"num")&",ISNULL("&replaceQuota(currExtTxtWarranty,"str")&",''));"
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveQty(vPartID,vCreatorNo,vCreatorDateTime,vCreatorIP) VALUES "&_ 
			"('"&currPartID&"',"&replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&_
			replaceQuota(currCreatorIP,"str")&");"
		
		xstring = "SELECT vWareHouseID FROM tlu_inveWareHouse;"
		SET getrsWHS = conn.execute(xstring)
		DO WHILE NOT getrsWHS.EOF 
			xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveQtyDetail(vPartID,vWareHouseID,vCreatorNo,vCreatorDateTime,vCreatorIP) "&_
				"VALUES ('"&currPartID&"','"&getrsWHS("vWareHouseID")&"',"&replaceQuota(currCreatorNo,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"
			getrsWHS.movenext
		LOOP			
		getrsWHS.close
		SET getrsWHS = NOTHING
		'PADA SAAT INSERT PRODUCT, CONTENT PRICE BOLEH UPDATE KE VENDOR PRICE
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_invePriceSetting(vPartID,vVndCurrID,vVndPrice,vMarginValue,vMarginPct,"&_
			"vStsPPN,vCreatorNo,vCreatorDateTime,vCreatorIP) "&_
			"VALUES ('"&currPartID&"',"&replaceQuota(currVPLPriceID,"str")&","&replaceQuota(currVPLPrice,"num")&","&_
			"ISNULL("&currMarginValue&",0),ISNULL("&currMarginPct&",0),ISNULL("&currStsPPN&",0),"&_
			replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"			
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveBuying(vPartID,vCntnPrcCurrID,vCntnPrc,vCurrCCurrID,vCurrCost,vCreatorNo,"&_
			"vCreatorDateTime,vCreatorIP) VALUES ('"&currPartID&"',"&replaceQuota(currVPLPriceID,"str")&","&_
			replaceQuota(currVPLPrice,"num")&","&replaceQuota(currVPLPriceID,"str")&","&replaceQuota(currVPLPrice,"num")&","&_
			replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"				
		IF (currSPrice>0) THEN 
			xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveSelling(vPartID,vPrcCurrID,vPrice,vSPrcCurrID,vSPrice,vStartSPrice,vEndSPrice,"&_
				"vMPrcCurrID,vMinPrice,vMinPrcRate,vSVndID,vLastUpdatePrc,vCreatorNo,vCreatorDateTime,vCreatorIP,vLastUpdateSpecialPrice) VALUES "&_
				"('"&currPartID&"',"&replaceQuota(currWebPriceID,"str")&","&replaceQuota(currWebPrice,"num")&","&_
				replaceQuota(currSPriceID,"str")&","&replaceQuota(currSPrice,"num")&","&replaceQuota(currStartSPrice,"str")&","&_
				replaceQuota(currEndSPrice,"str")&","&replaceQuota(currVPLPriceID,"str")&","&replaceQuota(currVPLPrice,"num")&","&_
				replaceQuota(getkurs(currVPLPriceID),"num")&","&replaceQuota(currPVndID,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorNo,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&");"
		ELSE
			xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveSelling(vPartID,vPrcCurrID,vPrice,"&_
				"vMPrcCurrID,vMinPrice,vMinPrcRate,vSVndID,vLastUpdatePrc,vCreatorNo,vCreatorDateTime,vCreatorIP) VALUES "&_
				"('"&currPartID&"',"&replaceQuota(currWebPriceID,"str")&","&replaceQuota(currWebPrice,"num")&","&_
				replaceQuota(currVPLPriceID,"str")&","&replaceQuota(currVPLPrice,"num")&","&_
				replaceQuota(getkurs(currVPLPriceID),"num")&","&replaceQuota(currPVndID,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorNo,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"
		END IF
		xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveCPCatRel WHERE vPartID='"&currPartID&"';"
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveCPCatRel(vPartID,vCatID,vTypeCat,vCreatorNo,vCreatorDateTime,vCreatorIP) VALUES "&_
			"('"&currPartID&"',"&replaceQuota(currCatPrimaryID,"str")&",'P',"&replaceQuota(currCreatorNo,"str")&","&_
			replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"
		
		xSQL_BMD = xSQL_BMD&" INSERT INTO "&DBServer&".dbo.[BMD-Debug$Item Buffer]"&_
			"([No_],[No_ 2],[Description],[Description 2],[Item Vendor],[Desc Produk],[Desc Produk2],[Desc Produk3],"&_
			"[Desc Produk4],[Manufacturer Code],[Berat Product],[Note 1],[Note 2],[Status Aktif],[Mkt Info],[Creator ID],"&_
			"[Creator Date],[Creator IP],[Editor ID],[Editor Date],[Editor IP],[Prod Kategori ID1],[Prod Kategori ID2],"&_
			"[Prod Kategori ID3],[Prod Kategori ID4],[Prod Kategori ID5],[Vendor Code],[Unit Price],[Purchase Price],[SN],"&_
			"[Length],[Width],[Height],[Pircture Link],[Mkt Info2],[Special Price],[Due Date Special Price],[Import Status],"&_
			"[Category Level 1],[Category Level 2],[Category Level 3],[Category Level 4],[Level Position],[Goods Description],"&_
			"[CurrPrice],[Price],[CurrSpecialPrice],[SpecialPrice]) "&_
			"VALUES "&_
			"('"&replaceQuota(currPartID,"str")&"','''',isnull('"&replaceQuota(seri1,"str")&"',''''),"&_
			"isnull('"&seri2&"',''''),'''',"&_
			"isnull('"&replaceQuota(desc1,"str")&"',''''),isnull('"&desc2&"',''''),"&_
			"isnull('"&desc3&"',''''),isnull('"&desc4&"',''''),"&_
			"isnull('"&replaceQuota(UCASE(currBrandID),"str")&"',''''),isnull("&replaceQuota(currShipWeight,"num")&",0),"&_
			"isnull('"&note1&"',''''),isnull('"&note2&"',''''),"&_
			"isnull("&replaceQuota(currActivation,"num")&",''''),isnull('"&marketingInfo1&"',''''),"&_
			"isnull('"&replaceQuota(currCreatorNo,"str")&"',''''),isnull('"&replaceQuota(currCreatorDateTime,"str")&"',''''),"&_
			"isnull('"&replaceQuota(currCreatorIP,"str")&"',''''),"&_
			" '''','''','''',"&_
			"isnull('"&catLvl1Name&"',''''),isnull('"&catLvl2Name&"',''''),"&_
			"isnull('"&catLvl3Name&"',''''),isnull('"&catLvl4Name&"',''''),''null'',"&_
			"isnull('"& replaceQuota(currPVndID,"str")&"',''''),isnull("&replaceQuota(currWebPrice,"num")&",0),"&_
			"isnull("&replaceQuota(currVPLPrice,"num")&",0),isnull("&replaceQuota(currNeedSN,"num")&",1),"&_
			"0,0,0,"&_
			"isnull('"&replaceQuota(strImgPath1,"str")&"',''''),"&_
			"isnull('"&marketingInfo2&"',''''),isnull("&replaceQuota(currSPrice,"num")&",0),"&_
			"isnull(''"&currEndSPrice&"'',''1753-01-01''),0,"&_
			"isnull('"&catLvl1ID&"',''''),isnull('"&catLvl2ID&"',''''),isnull('"&catLvl3ID&"',''''),isnull('"&catLvl4ID&"',''''),"&_
			"isnull("&levelCat&",0),isnull('"&currGoodDesc&"',''''),"&_
			"isnull('"&replaceQuota(currWebPriceID,"str")&"',''CUR01''),isnull("&replaceQuota(currWebPrice,"num")&",0),"&_
			"isnull(''"&currSPriceID&"'',''CUR01''),isnull("&replaceQuota(currSPrice,"num")&",0)"&_
			");"
		 
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveComputerExtInfo VALUES("&replaceQuota(currPartID,"str")&","&_
			replaceQuota(currNeedSN,"num")&",0,0,0);"			
	ELSEIF (LCase(currAct)="edit") THEN
		IF (currLastUpdateVPL="1") THEN 
			currLastUpdateVPLPrice = currNowDate 
		END IF
		IF (currLastUpdateWeb="1") THEN 
			currLastUpdateWebPrice = currNowDate 
			xSQL_EditorPrice = "vEditorNo="&replaceQuota(currEditorNo,"str")&","&_
							   "vEditorDateTime="&replaceQuota(currLastUpdatePrice,"str")&","&_
							   "vEditorIP="&replaceQuota(currEditorIP,"str")&","
		END IF
		IF ((currLastUpdateVPL="1") OR (currLastUpdateWeb="1")) THEN
			xSQL_WebPrice = "vPrcCurrID="&replaceQuota(currWebPriceID,"str")&","&_
							"vPrice="&currWebPrice&","&_
							"vLastUpdatePrc="&replaceQuota(currLastUpdateWebPrice,"str")&","
		END IF
		IF (currLastUpdateSP="1") THEN 
			currLastUpdateSPrice = currNowDate 
			xSQL_SPrice = "vSPrcCurrID="&replaceQuota(currSPriceID,"str")&","&_
						  "vSPrice="&currSPrice&","&_
						  "vStartSPrice="&replaceQuota(currStartSPrice,"str")&","&_
						  "vEndSPrice="&replaceQuota(currEndSPrice,"str")&","&_
						  "vLastUpdateSpecialPrice="&replaceQuota(currLastUpdateSPrice,"str")&","
			xSQL_EditorPrice = "vEditorNo="&replaceQuota(currEditorNo,"str")&","&_
					"vEditorDateTime="&replaceQuota(currLastUpdateSPrice,"str")&","&_
					"vEditorIP="&replaceQuota(currEditorIP,"str")&","
		END IF
		IF currActivation=3 THEN
			xString = " SELECT SUM([Remaining Quantity]) AS vOnHand FROM "&DBServer&".dbo.[BMD-Debug$Item Ledger Entry]" &_
			      	  " WHERE [Remaining Quantity]>0 AND [Item No_]='"&currPartID&"';"
			SET getRsChkQty = conn.execute(xString)
			IF NOT getRsChkQty.EOF THEN
				currOnHand = getRsChkQty("vOnHand")
			ELSE
				currOnHand = 0
			END IF
			getRsChkQty.Close
			SET getRsChkQty = NOTHING
			IF (currOnHand>0) THEN
				Response.Write "<script type=text/javascript language=javascript>"
				Response.Write "alert('"&currPartID&" "&currBrandName&" "&currSeri&" \n can not be inactive when quantity>0.');"
				Response.Write "history.back();"
				Response.Write "</script>"
				Response.End
			END IF
		END IF
		xSQL_DOS2 = xSQL_DOS2&" UPDATE trx_inveComputer SET "&_
					"vBrandID="&replaceQuota(currBrandID,"str")&","&_
					"vSeri="&replaceQuota(currSeri,"str")&","&_
					"vDesc="&replaceQuota(currDesc,"str")&","&_
					"vStatusID="&replaceQuota(currStatusID,"num")&","&_
					"vStartPeriod="&replaceQuota(currPeriodStart,"str")&","&_
					"vEndPeriod="&replaceQuota(currPeriodEnd,"str")&","&_
					"vShipWeight="&currShipWeight&","&_
					"vManufact="&replaceQuota(currManufacturer,"str")&","&_
					"vMarketingInfo="&replaceQuota(currMarketinginfo,"str")&","&_
					"vNote="&replaceQuota(currNote,"str")&","&_
					"vActivation="&replaceQuota(currActivation,"num")&","&_
					"vBrandIDPO="&replaceQuota(currBrandID,"str")&","&_
					"vSeriPO="&replaceQuota(currSeri,"str")&","&_
					"vDescPO="&replaceQuota(currDescPO,"str")&","&_
					"vEditorNo="&replaceQuota(currEditorNo,"str")&","&_
					"vEditorDateTime="&replaceQuota(currEditorTime,"str")&","&_
					"vEditorIP="&replaceQuota(currEditorIP,"str")&","&_
					"vJaminanMurah="&replaceQuota(currJaminanMurah,"num")&","&_
					"vAuthorisedWarranty="&replaceQuota(currAuthorisedWarranty,"num")&","&_
					"vExtTxtWarranty=ISNULL("&replaceQuota(currExtTxtWarranty,"str")&",'') "&_
					"WHERE vPartID="&replaceQuota(currPartID,"str")&";"
		IF ((currLastUpdateVPL="1") OR (currLastUpdateWeb="1") OR (currLastUpdateSP="1")) THEN
			xSQL_DOS2 = xSQL_DOS2&" UPDATE trx_inveSelling SET "&_
						"vMinPrice="&currVPLPrice&","&_
						xSQL_WebPrice&xSQL_SPrice&xSQL_EditorPrice&_
						"vSVndID="&replaceQuota(currPVndID,"str")&" "&_
						"WHERE vPartID="&replaceQuota(currPartID,"str")&";"
		END IF
		IF (currLastUpdateVPL="1") THEN
			xSQL_DOS2 = xSQL_DOS2&" IF NOT EXISTS(SELECT TOP 1 * FROM trx_inveBuying WHERE vPartID='"&currPartID&"') BEGIN "&_
				"INSERT INTO trx_inveBuying(vPartID,vCntnPrcCurrID,vCntnPrc,vCurrCCurrID,vCurrCost,vCreatorNo,vCreatorDateTime,"&_
				"vCreatorIP,vLastUpdateVPL) VALUES ('"&currPartID&"',"&replaceQuota(currVPLPriceID,"str")&","&_
				replaceQuota(currVPLPrice,"num")&","&replaceQuota(currVPLPriceID,"str")&","&_
				replaceQuota(currVPLPrice,"num")&","&replaceQuota(currCreatorNo,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&","&_
				replaceQuota(currCreatorDateTime,"str")&"); END ELSE BEGIN "&_
				"UPDATE trx_inveBuying SET "&_
				"vCntnPrcCurrID="&replaceQuota(currVPLPriceID,"str")&","&_
				"vCntnPrc="&replaceQuota(currVPLPrice,"num")&","&_
				"vLastUpdateVPL="&replaceQuota(currEditorDateTime,"str")&","&_
				"vEditorNo="&replaceQuota(currEditorNo,"str")&","&_
				"vEditorDateTime="&replaceQuota(currEditorTime,"str")&","&_
				"vEditorIP="&replaceQuota(currEditorIP,"str")&" "&_
				"WHERE vPartID =" & replaceQuota(currPartId,"str")&"; END"
		END IF
		xSQL_DOS2 = xSQL_DOS2&" EXECUTE sp_updMinPrice "&replacequota(currPartID,"str")&","&replacequota(currMarginPct,"num")&","&replacequota(currStsPPN,"str")&";"
		xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveCPCatRel WHERE vPartID='"&currPartID&"';"
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveCPCatRel(vPartID,vCatID,vTypeCat,vCreatorNo,vCreatorDateTime,"&_
			"vCreatorIP) VALUES ('"&currPartID&"',"&replaceQuota(currCatPrimaryID,"str")&",'P',"&_
			replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&replaceQuota(currCreatorIP,"str")&");"
		
		xSQL_BMD = xSQL_BMD&" IF NOT EXISTS(SELECT TOP 1 [No_] FROM "&DBServer&".dbo.[BMD-Debug$Item Buffer] "&_ 
			"WHERE [No_]='"&replaceQuota(currPartID,"str")&"') BEGIN "&_ 
			"INSERT INTO "&DBServer&".dbo.[BMD-Debug$Item Buffer]"&_
			"([No_],[No_ 2],[Description],[Description 2],[Item Vendor],[Desc Produk],[Desc Produk2],[Desc Produk3],[Desc Produk4],"&_
			"[Manufacturer Code],[Berat Product],[Note 1],[Note 2],[Status Aktif],[Mkt Info],[Creator ID],[Creator Date],[Creator IP],"&_
			"[Editor ID],[Editor Date],[Editor IP],[Prod Kategori ID1],[Prod Kategori ID2],[Prod Kategori ID3],[Prod Kategori ID4],"&_
			"[Prod Kategori ID5],[Vendor Code],[Unit Price],[Purchase Price],[SN],[Length],[Width],[Height],[Pircture Link],[Mkt Info2]," &_
			"[Special Price],[Due Date Special Price],[Import Status],[Category Level 1],[Category Level 2],[Category Level 3],"&_
			"[Category Level 4],[Level Position],[Goods Description],[CurrPrice],[Price],[CurrSpecialPrice],[SpecialPrice]) "&_
			"VALUES "&_
			"('"&replaceQuota(currPartID,"str")&"','''',isnull('"&replaceQuota(seri1,"str")&"',''''),"&_
			"isnull('"&seri2&"',''''),'''',"&_
			"isnull('"&replaceQuota(desc1,"str")&"',''''),"&_
			"isnull('"&desc2&"',''''),isnull('"&desc3&"',''''),"&_
			"isnull('"&desc4&"',''''),"&_
			"isnull('"&replaceQuota(UCASE(currBrandID),"str")&"',''''),isnull("&replaceQuota(currShipWeight,"num")&",0),"&_
			"isnull('"&note1&"',''''),"&_
			"isnull('"&note2&"',''''),isnull("&replaceQuota(currActivation,"num")&",1),"&_
			"isnull('"&marketingInfo1&"',''''),"&_
			"isnull('"&replaceQuota(currCreatorNo,"str")&"',''''),isnull('"&replaceQuota(currCreatorDateTime,"str")&"',''''),"&_
			"isnull('"&replaceQuota(currCreatorIP,"str")&"',''''),"&_
			"'''','''','''',"&_
			"isnull('"&catLvl1Name&"',''''),isnull('"&catLvl2Name&"',''''),"&_
			"isnull('"&catLvl3Name&"',''''),"&_
			"isnull('"&catLvl4Name&"',''''),''null'',isnull('"&replaceQuota(currPVndID,"str")&"',''''),"&_
			"isnull("&replaceQuota(currWebPrice,"num")&",0),isnull("&replaceQuota(currVPLPriceValue,"num")&",0),"&_
			"isnull("&replaceQuota(currNeedSN,"num")&",1),"&_
			"0,0,0,"&_
			"isnull('"&replaceQuota(strImgPath1,"str")&"',''''),isnull('"&marketingInfo2&"',''''),"&_
			"isnull("&replaceQuota(currSPrice,"num")&",0),"&_
			"isnull(''"&currEndSPrice&"'',''1753-01-01''),0,"&_
			"isnull('"&catLvl1ID&"',''''),isnull('"&catLvl2ID&"',''''),isnull('"&catLvl3ID&"',''''),isnull('"&catLvl4ID&"',''''),"&_
			"isnull("&levelCat&",0),isnull('"&currGoodDesc&"',''''),"&_
			"isnull('"&replaceQuota(currWebPriceID,"str")&"',''CUR01''),isnull("&replaceQuota(currWebPrice,"num")&",0),"&_
			"isnull('"&replaceQuota(currSPriceID,"str")&"',''CUR01''),isnull("&replaceQuota(currSPrice,"num")&",0)"&_
			"); END ELSE BEGIN UPDATE "&DBServer&".dbo.[BMD-Debug$Item Buffer] SET "&_ 
			"[Description]=isnull('"&replaceQuota(seri1,"str")&"',''''),"&_
			"[Description 2]=isnull('"&seri2&"',''''),[Item Vendor]='''',"&_
			"[Desc Produk]=isnull('"&replaceQuota(desc1,"str")&"',''''),[Desc Produk2]=isnull('"&desc2&"',''''),"&_
			"[Desc Produk3]=isnull('"&desc3&"',''''),[Desc Produk4]=isnull('"&desc4&"',''''),"&_
			"[Manufacturer Code]=isnull('"&replaceQuota(UCASE(currBrandID),"str")&"',''''),"&_
			"[Berat Product]=isnull("&replaceQuota(currShipWeight,"num")&",0),"&_
			"[Note 1]=isnull('"&note1&"',''''),[Note 2]=isnull('"&note2&"',''''),"&_
			"[Status Aktif]=isnull("&replaceQuota(currActivation,"num")&",1),"&_
			"[Mkt Info]=isnull('"&marketingInfo1&"',''''),"&_
			"[Editor ID]=isnull('"&replaceQuota(currCreatorNo,"str")&"',''''),"&_
			"[Editor Date]=isnull('"&replaceQuota(currCreatorDateTime,"str")&"',''''),"&_
			"[Editor IP]=isnull('"&replaceQuota(currCreatorIP,"str")&"',''''),"&_
			"[Prod Kategori ID1]=isnull('"&catLvl1Name&"',''''),"&_
			"[Prod Kategori ID2]=isnull('"&catLvl2Name&"',''''),"&_
			"[Prod Kategori ID3]=isnull('"&catLvl3Name&"',''''),"&_
			"[Prod Kategori ID4]=isnull('"&catLvl4Name&"',''''),"&_
			"[Vendor Code]=isnull('"&replaceQuota(currPVndID,"str")&"',''''),"&_
			"[Unit Price]=isnull("&replaceQuota(currWebPriceValue,"num")&",0),"&_
			"[Purchase Price]=isnull("&replaceQuota(currVPLPriceValue,"num")&",0),"&_
			"[SN]=isnull("&replaceQuota(currNeedSN,"num")&",1),[Length]=0,"&_
			"[Width]=0,[Height]=0,"&_
			"[Pircture Link]=isnull('"&replaceQuota(strImgPath1,"str")&"',''''),[Mkt Info2]=isnull('"&marketingInfo2&"',''''),"&_
			"[Special Price]=isnull("&replaceQuota(currSPriceValue,"num")&",0),"&_
			"[Due Date Special Price]=isnull(''"&currEndSPrice&"'',''1753-01-01''),"&_
			"[Category Level 1]=isnull('"&catLvl1ID&"',''''),[Category Level 2]=isnull('"&catLvl2ID&"',''''),"&_
			"[Category Level 3]=isnull('"&catLvl3ID&"',''''),[Category Level 4]=isnull('"&catLvl4ID&"',''''),"&_
			"[Level Position]=isnull("&levelCat&",0),[Goods Description]=isnull('"&currGoodDesc&"',''''),"&_
			"[CurrPrice]=isnull('"&replaceQuota(currWebPriceID,"str")&"',''CUR01''),[Price]=isnull("&replaceQuota(currWebPrice,"num")&",0),"&_
			"[CurrSpecialPrice]=isnull('"&replaceQuota(currSPriceID,"str")&"',''CUR01''),"&_
			"[SpecialPrice]=isnull("&replaceQuota(currSPrice,"num")&",0) "&_
			"WHERE [No_]='"&replaceQuota(currPartID,"str")&"';"&_
			"END "	
		xSQL_DOS2 = xSQL_DOS2&" UPDATE trx_inveComputerExtInfo SET bSN="&replaceQuota(currNeedSN,"num")&","&_
			"iLength=0,iWidth=0,iHeight=0 WHERE vPartID="&replaceQuota(currPartID,"str")&";" 
		'Start - generate Bhinneka Promo dan Special Deal
		SET rsPromo = Server.CreateObject("ADODB.Recordset")			
		queryPromo = "SELECT TOP 1 * FROM trx_InveCPBundle WHERE vPartID='"&currPartID&"';"
		rsPromo.open queryPromo,conn,3,1,0
		IF NOT rsPromo.EOF THEN
			bundleID = rsPromo("vTrxID")
			xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveCPBundleDetail WHERE vTrxID='"&bundleID&"';"
			xSQL_BMD = xSQL_BMD&" DELETE FROM "&DBServer&".dbo.[BMD-Debug$Item Free Line] WHERE [Trans Code]=''"&bundleID&"'';"
		END IF
	END IF
	currJmlVendor = CInt(currJmlVendor)
	xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveBuyingVendor WHERE vPartID='"&currPartID&"';"	
	FOR i=1 TO currJmlVendor
		IF Request.Form("crVendorID"&i)<>"" THEN			
			xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveBuyingVendor(vPartID,vVndID,vVndType,"&_
				"vCreatorNo,vCreatorDateTime,vCreatorIP,vVPLFromID) VALUES ('"&currPartID&"',"&_
				replaceQuota(Request.Form("crVendorID"&i),"str")&",'',"&_
				replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime, "str")&","&_
				replaceQuota(currCreatorIP,"str")&","&replaceQuota(currPVndID,"str")&");"
		END IF
	NEXT
	
	xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveWarranty WHERE vPartID='"&currPartID&"' AND vType=0;"
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveWarranty(vPartID,vType,vModelID,vLongWarranty,vCreatorNo,"&_
		"vCreatorDateTime,vCreatorIP) VALUES ('"&currPartID&"',0,'"&currVndModelID1&"',"&_
		replaceQuota(currVndLengthWarranty1,"num")&",'"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"');"
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveWarranty(vPartID,vType,vModelID,vLongWarranty,vCreatorNo,"&_
		"vCreatorDateTime,vCreatorIP) VALUES ('"&currPartID&"',0,'"&currVndModelID2&"',"&_
		replaceQuota(currVndLengthWarranty2,"num")&",'"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"');"
	xSQL_DOS2 = xSQL_DOS2&" DELETE FROM trx_inveWarranty WHERE vPartID='"&currPartID&"' AND vType=1;"
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveWarranty(vPartID,vType,vModelID,vLongWarranty,vCreatorNo,"&_
		"vCreatorDateTime,vCreatorIP) VALUES ('"&currPartID&"',1,'"&currVndModelID1&"',"&_
		replaceQuota(currLengthWarranty00001,"num")&",'"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"');"
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveWarranty(vPartID,vType,vModelID,vLongWarranty,vCreatorNo,"&_
		"vCreatorDateTime,vCreatorIP) VALUES ('"&currPartID&"',1,'"&currVndModelID2&"',"&_
		replaceQuota(currLengthWarranty00002,"num")&",'"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"');"
	
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_HistoryCatSKU(vPartID,vCatIDLvl1,vCatNameLvl1,vCatIDLvl2,vCatNameLvl2,"&_
		"vCatIDLvl3,vCatNameLvl3,vCatIDLvl4,vCatNameLvl4,vCreatorID,vCreatorDateTime,vCreatorIP) "&_
		"VALUES ('"&currPartID&"',ISNULL("&catLvl1ID&",''),ISNULL("&catLvl1Name&",''),ISNULL("&catLvl2ID&",''),"&_
		"ISNULL("&catLvl2Name&",''),ISNULL("&catLvl3ID&",''),ISNULL("&catLvl3Name&",''),ISNULL("&catLvl4ID&",''),"&_
		"ISNULL("&catLvl4Name&",''),"&replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&_
		replaceQuota(currCreatorIP,"str")&");"
	
	sqlCekAdaVend = "SELECT vVndID FROM trx_inveBuyingVendor WHERE vPartID='"&currPartID&"' AND vVndID='"&currPVndID&"';" 
	SET rsCekVend = conn.execute(sqlCekAdaVend)
	xAda = 0
	IF NOT rsCekVend.EOF THEN
		xAda = 1
	END IF
	rsCekVend.Close:SET rsCekVend = NOTHING
	
	IF (xAda=0) THEN
		xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveBuyingVendor(vPartID,vVndID,vVndType,vCreatorNo,vCreatorDateTime,vCreatorIP,"&_
			"vVPLFromID) VALUES ("&_
			"'"&currPartID&"',"&replaceQuota(currPVndID,"str")&",'P',"&_
			replaceQuota(currCreatorNo,"str")&","&replaceQuota(currCreatorDateTime,"str")&","&_
			replaceQuota(currCreatorIP,"str")&","&replaceQuota(currPVndID,"str")&")"
	ELSE
		xSQL_DOS2 = xSQL_DOS2&" UPDATE trx_InveBuyingVendor SET vVndType='P' WHERE vPartID='"&currPartID&"' AND "&_ 
		"vVndID='"&currPVndID&"';" 
	END IF 

	IF ((UBOUND(tempSKUBP)<>-1) OR (UBOUND(tempSKUSD)<>-1)) THEN
		IF (bundleID<>"") THEN
			xSQL_DOS2 = xSQL_DOS2&" UPDATE trx_inveCPBundle SET "&_
				"vEditorNo="&replaceQuota(currEditorNo,"str")&","&_
				"vEditorDateTime="&replaceQuota(currEditorTime,"str")&","&_
				"vEditorIP="&replaceQuota(currEditorIP,"str")&" "&_
				"WHERE vTrxID='"&bundleID&"';"
			xSQL_BMD = xSQL_BMD&" UPDATE "&DBServer&".dbo.[BMD-Debug$Item Free Header] SET "&_
				"[Editor ID]='"&replaceQuota(currEditorNo,"str")&"',"&_
				"[Editor Date]='"&replaceQuota(currEditorTime,"str")&"',"&_
				"[Editor IP]='"&replaceQuota(currEditorIP,"str")&"' "&_
				"WHERE [Trans Code]=''"&bundleID&"'';"
		ELSE
			bundleID = gen_id()
			xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_InveCPBundle(vTrxID,vType,vPartID,vStartDt,vEndDt,vStatus,vCreatorNo,vCreatorDateTime,vCreatorIP) "&_
				"VALUES ('"&bundleID&"','FRE','"&currPartID&"','','',1,'"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"');"
			xSQL_BMD = xSQL_BMD&" INSERT INTO "&DBServer&".dbo.[BMD-Debug$Item Free Header] ([Trans Code],"&_
				"[Trans Type],[Item Code],[Promo Start Date],[Promo End Date],[Active Status],[Start Date],[End Date],[Creator ID],"&_
				"[Creator Date],[Creator IP],[Editor ID],[Editor Date],[Editor IP]) "&_
				"VALUES (''"&bundleID&"'',''FRE'',''"&currPartID&"'','''','''',1,'''','''',''"&currCreatorNo&"'',''"&currCreatorDateTime&"'',"&_
				"''"&currCreatorIP&"'','''','''','''');"
		END IF			
		countPromo = 0
		IF UBOUND(tempSKUBP)<>-1 THEN
			currJmlhBP = UBOUND(tempSKUBP)
	
			FOR i=0 TO currJmlhBP
				currTypeBP = tempTypeBP(i)
				currSKUBP = tempSKUBP(i)
				currBrandBP = tempBrandBP(i)
				currSeriBP = tempSeriBP(i)
				currDescBP = tempDescBP(i) 
				currQtyBP = tempQtyBP(i)
				currStartDateBP = tempStartDateBP(i)
				currEndDateBP = tempEndDateBP(i)&" 23:59:59.000"
				currTagBP = tempTagBP(i)
				currStsActiveBP = 0
				
				IF ((DateValue(now())>=CDate(tempStartDateBP(i))) AND (DateValue(now())<=CDate(tempEndDateBP(i)))) THEN
					currStsActiveBP = 1				
				END IF
		
				xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveCPBundleDetail(vTrxID,vNoUrut,vBundlePartID,vQty,vStatus,vCreatorNo,"&_
					"vCreatorDateTime,vCreatorIP,vStsActive,vPromo,vType,vStartDateTime,vEndDateTime,vTagDate)"&_
					" VALUES ('"&bundleID&"','"&countPromo&"','"&currSKUBP&"','"&currQtyBP&"','1','"&currCreatorNo&"',"&_
					"'"&currCreatorDateTime&"','"&currCreatorIP&"','"&currStsActiveBP&"','0','"&currTypeBP&"',"&_
					"'"&currStartDateBP&"','"&currEndDateBP&"','"&currTagBP&"');"
				xSQL_BMD = xSQL_BMD&" INSERT INTO "&DBServer&".dbo.[BMD-Debug$Item Free Line]([Trans Code],"&_
					"[Line No_],[Item Code],[Qty],[Active Status],"&_
					"[Creator ID],[Creator Date],[Creator IP],[Editor ID],[Editor Date],[Editor IP],[Promo],[Type],"&_
					"[StartDateTime],[EndDateTime]) "&_
					"VALUES (''"&bundleID&"'',"&countPromo&",''"&currSKUBP&"'',"&currQtyBP&",''"&currStsActiveBP&"'',"&_
					"''"&currCreatorNo&"'',''"&currCreatorDateTime&"'',''"&currCreatorIP&"'','''','''','''',0,''"&currTypeBP&"'',"&_
					"''"&currStartDateBP&"'',''"&currEndDateBP&"'');"
				countPromo = countPromo+1
			NEXT
		END IF
		IF (UBOUND(tempSKUSD)<>-1) THEN
			currJmlhSD = UBOUND(tempSKUSD)
			FOR i=0 TO currJmlhSD
				currTypeSD = tempTypeSD(i)
				currSKUSD = tempSKUSD(i)
				currBrandSD = tempBrandSD(i)
				currSeriSD = tempSeriSD(i)
				currDescSD = tempDescSD(i) 
				currQtySD = tempQtySD(i)
				currStartDateSD = tempStartDateSD(i)
				currEndDateSD = tempEndDateSD(i)&" 23:59:59.000"
				currTagSD = tempTagSD(i)
				currStsActiveSD = 0
		
				IF ((DateValue(now())>=CDate(tempStartDateSD(i))) AND (DateValue(now())<=CDate(tempEndDateSD(i)))) THEN
					currStsActiveSD = 1				
				END IF
		
				xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveCPBundleDetail(vTrxID,vNoUrut,vBundlePartID,vQty,vStatus,vCreatorNo,"&_
					"vCreatorDateTime,vCreatorIP,vStsActive,vPromo,vType,vStartDateTime,vEndDateTime,vTagDate)"&_
					" VALUES ('"&bundleID&"','"&countPromo&"','"&currSKUSD&"','"&currQtySD&"','1','"&currCreatorNo&"',"&_
					"'"&currCreatorDateTime&"','"&currCreatorIP&"','"&currStsActiveSD&"','1','"&currTypeSD&"',"&_
					"'"&currStartDateSD&"','"&currEndDateSD&"','"&currTagSD&"');"
				xSQL_BMD = xSQL_BMD&"INSERT INTO "&DBServer&".dbo.[BMD-Debug$Item Free Line]"&_
					"([Trans Code],[Line No_],[Item Code],[Qty],[Active Status],"&_
					"[Creator ID],[Creator Date],[Creator IP],[Editor ID],[Editor Date],[Editor IP],[Promo],[Type],"&_
					"[StartDateTime],[EndDateTime]) "&_
					"VALUES (''"&bundleID&"'',"&countPromo&",''"&currSKUSD&"'',"&currQtySD&","&currStsActiveSD&","&_
					"''"&currCreatorNo&"'',''"&currCreatorDateTime&"'',''"&currCreatorIP&"'','''','''','''',1,"&currTypeSD&","&_
					"''"&currStartDateSD&"'',''"&currEndDateSD&"'');"
				countPromo = countPromo+1	
			NEXT
		END IF
		rsPromo.Close
		SET rsPromo = NOTHING
	END IF
	SET rsFindTmpltNo = Server.CreateObject("ADODB.Recordset")
	queryFindTmpltNo = "SELECT tlu_InveTmplt.vTmpltNo FROM tlu_InveTmpltRelation "&_
		"INNER JOIN tlu_InveTmplt ON tlu_InveTmplt.vTmpltID = tlu_InveTmpltRelation.vTmpltID "&_
		"WHERE tlu_InveTmpltRelation.vCatID = '"&currCatPrimaryID&"' AND tlu_InveTmplt.vId = '1092'"
	rsFindTmpltNo.Open queryFindTmpltNo,conn,3,1,0
	IF NOT rsFindTmpltNo.EOF THEN
		tempTmpltNo = rsFindTmpltNo("vTmpltNo")
		currLongWarranty = currLengthWarranty00001
		currWarrantyText = ""
		IF (currAuthorisedWarranty) THEN
			currWarrantyText = "dari Distributor Resmi di Indonesia"														
		END IF
		IF ((currLongWarranty>0) AND (currLongWarranty<=60)) THEN
			IF currWarrantyText<>"" THEN
				currWarrantyText =  currLongWarranty&" "&"Bulan"&" "&currWarrantyText&" "&currExtTxtWarranty
			ELSE
				currWarrantyText =  currLongWarranty&" "&"Bulan"&" "&currExtTxtWarranty
			END IF
		ELSEIF (currLongWarranty=0) THEN
			currWarrantyText = "Garansi tidak tersedia dari Distributor"
		ELSE
			currWarrantyText = "60 Bulan "&" "&currWarrantyText&" "&currExtTxtWarranty
		END IF
		currWarrantyText = TRIM(currWarrantyText)
		xSQL_DOS2 = xSQL_DOS2&" IF EXISTS(SELECT '1' FROM trx_InveTmpltDetail WHERE vPartID = '"&currPartID&"' AND vTmpltNo = "&tempTmpltNo&") BEGIN "&_
			"UPDATE trx_InveTmpltDetail SET vDesc = '"&currWarrantyText&"',vEditorNo='"&currCreatorNo&"',"&_
			"vEditorDateTime='"&currCreatorDateTime&"',vEditorIP='"&currCreatorIP&"' WHERE vPartID = '"&currPartID&"' AND "&_
			"vTmpltNo="&tempTmpltNo&" "&_
			"END ELSE BEGIN INSERT INTO trx_InveTmpltDetail(vPartID,vTmpltNo,vDesc,vCreatorNo,vCreatorDateTime,vCreatorIP) "&_
			"VALUES ('"&currPartID&"',"&tempTmpltNo&",'"&currWarrantyText&"','"&currCreatorNo&"','"&currCreatorDateTime&"',"&_
			"'"&currCreatorIP&"'); END"			
		rsFindTmpltNo.Close
		SET rsFindTmpltNo = NOTHING
	END IF	
	xSQL_DOS2 = xSQL_DOS2&" EXECUTE sp_vSeriValidator;"
	xSQL_DOS2 = xSQL_DOS2&" INSERT INTO trx_inveBufferDOS2(vStatement,vCreatorNo,vCreatorDateTime,vCreatorIP,vInformation) VALUES "&_
				"('"&xSQL_BMD&"','"&currCreatorNo&"','"&currCreatorDateTime&"','"&currCreatorIP&"','"&currPartID&"');"
	'''xSQL_DOS2 = xSQL_DOS2&" EXECUTE sp_executeBufferDOS2;"
	''Response.Write(xSQL_DOS2)
	''Response.End
	conn.execute(xSQL_DOS2)
	'''conn.execute(xSQL_BMD)
	'''Response.Write(xSQL_DOS2)
	'''Response.End
	IF (conn.errors.count=0) THEN
		conn.committrans
		'''IF (waktu>8 AND waktu<10) THEN
			'''xSQLSinkronisasi = "EXECUTE sp_sinkronisasiDos2HoToNav;"
			'''conn.execute(xSQLSinkronisasi)
			'''xSQLSKUError = "EXECUTE sp_errorSKUInactiveDump;"
			'''conn.execute(xSQLSKUError)
		'''END IF
		conn.execute("EXECUTE sp_executeBufferDOS2;")
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Data Successfully has been saved');"
		IF (currtypesave="save") THEN
			Response.Write "window.location='digoff_inve_prodcatalog_view.asp?crBhs=2&crPartID="&currPartID&"&crList=&crLoad=1';"
		ELSEIF (currtypesave="savenext") THEN
			Response.Write "window.location='digoff_inve_prodcatalogdet.asp?crAct="&currAct&"&crPartID=" & currPartID & xloadnext & "';"
		ELSEIF (currtypesave="savenew") THEN
			Response.Write "window.location='digoff_inve_prodcatalog.asp?crAct=Add"& xloadnext & "';"
		END IF
		Response.Write "</script>"
		Response.End
	ELSE
		FOR EACH item IN conn.errors
			Response.Write item.description
		NEXT
		conn.rollbacktrans
		Response.Write "<script language='javascript'>"
		Response.Write "history.back();"
		Response.Write "alert('attention!! \n\nsaving process is fail!');"
		Response.Write "</script>"
		Response.End
	END IF
%>
<!--#include file="../include/glob_conn_close.asp"-->