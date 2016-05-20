<%
currUAModule = "DO005"
currUASubModule = "DO005025"
currUACategory = "DO"
%>
<!--#include file="../include/glob_conn_open.asp"-->
<!--#include file="../include/global_value.asp" -->
<!--#include file="../include/gen_css.asp" -->
<!--#include file="../include/gen_data.asp" -->
<!--#include file="../include/glob_session_check.asp"-->
<!--#include file="../include/glob_akses_check.asp"-->
<html>
<%IF currUAAllAccess=0 and currUAView=0 THEN
	Response.Write glob_WriteBlockPage("")
%>
	<!--#include file="../include/glob_conn_close.asp" -->
	<%
	Response.end
END IF

set getRS = server.createobject("ADODB.recordset")

currPartID = Trim(Request.QueryString("crPartID"))
crLoad = Trim(Request.QueryString("crLoad"))
crBhs = 2
currBhs = crBhs
warna = 1
xWarnaLine = "#999"
xMouseBack= " src=""../image/btn_back_0" & warna &   crBhs & ".gif"" onmouseover=this.src=""../image/btn_back_1" & warna & crBhs & ".gif""; onmouseout=this.src=""../image/btn_back_0" & warna & crBhs & ".gif"";"
xLoadNext = "&crLoad=1"
adaPromo = 0

xstring=" select "&_
		" (select a.vnickname from trx_hrdemployment a where a.vnoemp=i.vcreatorno) creatorname,"&_
	    " (select a.vnickname from trx_hrdemployment a where a.vnoemp=i.veditorno) editorname,"&_
	    " i.vcreatordatetime, i.veditordatetime, "&_
		" i.vPartID,buy.vcntnprccurrid,buy.vcntnprc,isnull((select vcurrcode from tlu_currency where vcurrid=buy.vcntnprccurrid),'')vcntnprccurrname,"&_
		" s.vprccurrid,isnull((select vcurrcode from tlu_currency where vcurrid=s.vprccurrid),'')vPrcCurrName, s.vprice,"&_
		" s.vsprice, "&_
		" isnull((select vcurrcode from tlu_currency where vcurrid=s.vNominalCurrID),'')vMinPrcCurrName, "&_
		" isnull((select vcurrcode from tlu_currency where vcurrid=s.vsprccurrid),'')vSPrcCurrName, s.vsprccurrid,s.vsvndid,(select vname from trx_vendormain where vvndid=s.vsvndid) vsvndname, "&_
		" s.vStartSPrice,s.vEndSPrice,s.vMPrcCurrID,"&_
		" (select vcurrcode from tlu_currency where vcurrid=s.vMPrcCurrID)vMPrcCurrName, s.vminprice,"&_
		" (select vcatid from trx_inveCPCatRel where vpartid=i.vpartid and lower(vtypecat)='p') vCatPrimaryID,"&_
		" (select vname from trx_inveCPCatRel r, tlu_inveCPCategory c where r.vcatid=c.vcatid and r.vpartid=i.vpartid and lower(r.vtypecat)='p') vCatPrimaryName,"&_
		" isnull((select vpartid from trx_invecpbundle where vpartid=i.vpartid),'') vitembundlemain,"&_
		" isnull((select top 1 vbundlepartid from trx_invecpbundledetail where vbundlepartid=i.vpartid),'') vitembundledetail,"&_
		" i.vBrandID,ib.vname vbrandname,vPartNo,vSeri,vDesc,s.vLastUpdatePrc,"&_
		" (select a.vnickname from trx_hrdemployment a where a.vnoemp=COALESCE(s.vEditorNo,s.vCreatorNo)) vPriceEditorName,"&_
		" vBarCodeDesc,i.vstatusid,(select vname from tlu_invecpstatus where i.vstatusid=vid) vstsname,vStartPeriod,vEndPeriod,vShipWeight,"&_
		" (select vmarginpct from trx_invepricesetting where vpartid=i.vpartid) vmarginpct,"&_
		" (select vstsppn from trx_invepricesetting where vpartid=i.vpartid) vstsppn,"&_
		" (select vstsppncs from trx_invepricesetting where vpartid=i.vpartid) vstsppncs,"&_
		" vManufact,it.vname vActivationName,vMarketingInfo,vNote,i.vCreatorNo,i.vCreatorDateTime,i.vCreatorIP,i.vJaminanMurah,ISNULL(i.vAuthorisedWarranty, 'TRUE') AS vAuthorisedWarranty,ISNULL(i.vExtTxtWarranty,'') as vExtTxtWarranty, "&_
		" (select isNull(bun.vTrxID,'') from trx_inveCPBundle bun where bun.vPartID=i.vPartID) vTrxID,"&_
		" (select isNull(cominfo.bSN,'') from trx_inveComputerExtInfo cominfo where cominfo.vPartID=i.vPartID) bSN,"&_
		" (select isNull(cominfo.iLength,'') from trx_inveComputerExtInfo cominfo where cominfo.vPartID=i.vPartID) iLength,"&_
		" (select isNull(cominfo.iWidth,'') from trx_inveComputerExtInfo cominfo where cominfo.vPartID=i.vPartID) iWidth,"&_
		" (select isNull(cominfo.iHeight,'') from trx_inveComputerExtInfo cominfo where cominfo.vPartID=i.vPartID) iHeight,"&_
		" (select distinct isnull(buy_vendor.vVPLFromID,'') from trx_InveBuyingVendor buy_vendor where buy_vendor.vPartID=i.vPartID and buy_vendor.vVPLFromID is not null) vVPLFromID,"&_
		" (select distinct isnull(vendor_main.vName,'') from trx_vendorMain vendor_main,trx_InveBuyingVendor buy_vendor where vendor_main.vVndID=buy_vendor.vVPLFromID and buy_vendor.vPartID=i.vPartID) vVPLFromName"&_
		" from trx_invecomputer i inner join trx_inveselling s on i.vpartid=s.vpartid inner join trx_invebuying buy on buy.vpartid=i.vpartid and i.vpartid='"&currPartID&"'"&_
		" inner join tlu_invecpbrand ib on i.vbrandid=ib.vbrandid inner join tlu_invecpactivation it on it.vid=i.vactivation "
	
getRs.open xstring,conn

IF NOT getRs.eof THEN
	currPartID=getRs("vPartID")
	currBrandName=getRs("vbrandname")
	currSeri=getRs("vSeri")
	currDesc=getRs("vDesc")
	currPartNo=getRs("vPartNo")
	currBarCodeDesc=getRs("vBarCodeDesc")
	currStatusID=getRs("vstatusid")
	currStatusName=getRs("vstsname")
	currPeriodStart=getRs("vStartPeriod")
	currPeriodEnd=getRs("vEndPeriod")
	currShipWeight=getRs("vShipWeight")
	currManufacturer=getRs("vManufact")
	currActivationName=getRs("vActivationName")
	currMarketinginfo=getRs("vMarketingInfo")
	currNote=getRs("vNote")
	currCntnPrcCurrID=getRs("vcntnprccurrid")
	currCntnPrcCurrName=getRs("vcntnprccurrname")
	currCntnPrice=getRs("vcntnprc")
	currPrcCurrID=getRs("vprccurrid")
	currPrcCurrName=getRs("vprccurrname")
	currPrice=getRs("vprice")
	currSPrcCurrID=getRs("vsprccurrid")
	currSPrcCurrName=getRs("vsprccurrname")
	currSPrice=getRs("vsprice")
	currMinPrcCurrName=getRs("vMinPrcCurrName")
	currSPValidStart=getRs("vStartSPrice")
	currSPValidEnd=getRs("vEndSPrice")
	currSVndID=getRs("vSVndID")	
	currSVndName=getRs("vSVndName")
	currMPrcCurrID=getRs("vMPrcCurrID")
	currMPrcCurrName=getRs("vMPrcCurrName")
	currMinPrice=getRs("vminprice")
	currCatPrimaryID=getRs("vCatPrimaryID")
	currCatPrimaryName=getRs("vCatPrimaryName")
	curritembundlemain=getRs("vitembundlemain")
	curritembundledetail=getRs("vitembundledetail")
	currLastUpdatePrc=getRs("vLastUpdatePrc")
	currPriceEditorName=getRs("vPriceEditorName")
	currMarginPct=getRs("vMarginPct")
	currStsPPn=getRs("vStsPPn")
	crCreatorName=getRs("creatorname")
    crEditorName=getRs("editorname")
    crCreatorDateTime=getRs("vcreatordatetime")
    crEditorDateTime=getRs("veditordatetime") 
	crJaminanMurah = getRs("vJaminanMurah")
	crAuthorisedWarranty = getRs("vAuthorisedWarranty")
	crExtTxtWarranty = getRs("vExtTxtWarranty")
	crBundleID=getRs("vTrxID")
	crNeedSN = getRs("bSN")
	currLength = getRs("iLength")
	currWidth = getRs("iWidth")
	currHeight = getRs("iHeight")
	currVPLFromID = getRs("vVPLFromID")
	currSVndNameAdd = getRs("vVPLFromName")
END IF
getRs.close

IF crJaminanMurah THEN
	crJaminanMurah = "Yes"
ELSE
	crJaminanMurah = "No"
END IF

IF crAuthorisedWarranty THEN
	crAuthorisedWarranty = "Yes"
ELSE
	crAuthorisedWarranty = "No"
END IF
'--------------compare Web Price apakah ada koma atau tidak (khusus IDR aja)'-------------------
allow = false
WPInt = Int(currPrice)
WPDbl = cDbl(currPrice)
symbol = "*"

IF currPrcCurrID = "CUR01" THEN
	IF WPInt < WPDbl THEN
		allow = true
		currPrice=formatnumber(currPrice,2)
	ELSE
		allow = false
	END IF
END IF
	
IF currSPrice <> "" AND currSPrice>0 THEN
	currSpecialPrice = currSPrice * 1.1
	IF currSPrcCurrID = "CUR01" THEN
		currSpecialPrice = Round(currSpecialPrice + 0.5,0)
		currSPrice = formatnumber(currSPrice,2)
	ELSE
		currSpecialPrice = Round(currSpecialPrice + 0.5,0)
		currSPrice = formatnumber(currSPrice,2)
	END IF
ELSEIF currSPrice<0 THEN
	currSpecialPrice = currSPrice * 1.1
	IF currSPrcCurrID = "CUR01" THEN
		currSpecialPrice = Round(currSpecialPrice - 0.5,0)
		currSPrice = formatnumber(currSPrice,2)
	ELSE
		currSpecialPrice = Round(currSpecialPrice - 0.5,0)
		currSPrice = formatnumber(currSPrice,2)
	END IF
END IF
	
IF currCntnPrcCurrID <> "" THEN
	basicCurrency = currCntnPrcCurrID
	currencyRate = currPrcCurrID	
ELSE
	basicCurrency = currPrcCurrID
	currencyRate = currPrcCurrID
END IF
	
IF basicCurrency = "" THEN 
	basicCurrency = "CUR01"
	currencyRate = "CUR01"
END IF

IF basicCurrency <> "CUR01" THEN
	currencyRate = basicCurrency
	symbol = "/"
END IF
	
xCurrency = "select top 1 * from trx_Currency where vCurrID = '"&currencyRate&"' order by vCurrDate DESC"
getRs.open xCurrency,conn,3,1,0
IF NOT getRs.eof THEN
	currRateNow = getRs("vCurrValue")
END IF
getRs.close

IF symbol = "*" THEN
	IF currCntnPrcCurrID <> basicCurrency THEN
		currCntnPrcCurrID=basicCurrency
		currCntnPrice=currCntnPrice*currRateNow
	END IF
	IF currPrcCurrID<>basicCurrency THEN
		currPrcCurrID=basicCurrency
		IF currPrice > "0" THEN
			currPrice=currPrice*currRateNow
		END IF
	END IF
ELSEIF symbol = "/" THEN
	IF currCntnPrcCurrID <> basicCurrency THEN
		currCntnPrcCurrID=basicCurrency
		currCntnPrice=currCntnPrice/currRateNow
	END IF
	IF currPrcCurrID<>basicCurrency THEN
		currPrcCurrID=basicCurrency
		IF currPrice > "0" THEN
			currPrice=currPrice/currRateNow
		END IF
	END IF
END IF

IF basicCurrency = "CUR01" THEN
	currCntnPrcCurrName="IDR"
	currPrcCurrName="IDR"
ELSEIF basicCurrency = "CUR02" THEN
	currCntnPrcCurrName="USD"
	currPrcCurrName="USD"
ELSEIF basicCurrency = "CUR03" THEN
	currCntnPrcCurrName="JPY"
	currPrcCurrName="JPY"
END IF
	
set rsPromo = Server.CreateObject("ADODB.Recordset")
queryPromo = "select * from trx_InveCPBundle where vPartID = '" &currPartID& "'" 
rsPromo.open queryPromo, conn, 3, 1, 0

IF rsPromo.eof THEN
	adaPromo = 0
ELSE
	adaPromo = 1
	currBundleID = rsPromo("vTrxID")
END IF

rsPromo.close
set rsPromo = nothing

function ceknull(pString)
	str=trim(pString)
	IF isnull(str) or str="" or isEmpty(str) THEN
		ceknull="<font class='wordredregular'>n/a</font>"
	ELSE
		ceknull=Server.HTMLEncode(str)
	END IF
end function

function inveprodcatalog_menu(id)
	dim s,jmlmenu
	s=""
	dim mnu(8,3)

	mnu(0,0)="profile"
		mnu(0,1)= "Profile"
		mnu(0,2)="digoff_inve_prodcatalog_view.asp?crPartID="&currPartID&xLoadNext
	
		mnu(1,0)="detail"
		mnu(1,1)= "Detail"
		mnu(1,2)="digoff_inve_prodcatalogdet_view.asp?crPartID="&currPartID&xLoadNext
	
		mnu(2,0)="image"
		mnu(2,1)= "Image"
		mnu(2,2)="digoff_inve_prodcatalogimg_view.asp?crPartID="&currPartID&xLoadNext
		
		mnu(3,0)="overview"
		mnu(3,1)= "Overview"
		mnu(3,2)="digoff_inve_prodcatalogoverview_controller.asp?crAct=View&crPartID="&currPartID&xLoadNext
		
		mnu(4,0)="brochure"
		mnu(4,1)= "Brochure"
		mnu(4,2)="digoff_inve_prodcatalogbrochure_view.asp?crPartID="&currPartID&xLoadNext
		
		mnu(5,0)="offer"
		mnu(5,1)= "Offer"
		mnu(5,2)="digoff_inve_prodcatalogoffer_view.asp?crPartID="&currPartID&xLoadNext
		
		mnu(6,0)="vs"
		mnu(6,1)= "Voucher"
		mnu(6,2)="digoff_inve_vssku_view.asp?crPartID="&currPartID&xLoadNext
		
		mnu(7,0)="market"
		mnu(7,1)= "Marketing"
		mnu(7,2)="digoff_inve_prodcatalogmarket_view.asp?crPartID="&currPartID&xLoadNext
	
	if id<>"" then
		s=gen_menu(mnu,id)
	end if
	inveprodcatalog_menu=s
end function

function gen_menu(mnu,id)
	dim s,jmlmenu
	dim sts
	sts=false
	jmlmenu=ubound(mnu)
	for i=0 to jmlmenu - 1
			if mnu(i,0)=id then
				sts=true
				if i=0 then
					s=s&"<td><img src=""../image/tab/slc_tab_actleft_0" & warna & "0.gif"">"
				else
					s=s&"<td><img src=""../image/tab/slc_tab_noactrightactleft_0" & warna & "0.gif""></td>"
				end if
				s=s&"<td align=""center"" background=""../image/tab/slc_tab_bgact_0" & warna & "0.gif""><font class=wordactivetab><nobr>"&mnu(i,1)&"</nobr></font></td>"
				if i = jmlmenu-1 then
					s=s&"<td><img src=""../image/tab/slc_tab_actright_0" & warna & "0.gif""</td>"
				else
					s=s&"<td><img src=""../image/tab/slc_tab_actrightnoactleft_0" & warna & "0.gif""</td>"
				end if
            else
				if i=0 then
						s=s&"<td><img src=""../image/tab/slc_tab_noactleft_0" & warna & "0.gif""</td>"
					else
						if not sts then
						     s=s&"<td><img src=""../image/tab/slc_tab_noactrightnoactleft_0" & warna & "0.gif""></td>"
					    end if
                end if
					    if lcase(currAct)="edit" or lcase(currAct)="mutation" or lcase(currAct)="add" then
					     	s=s&"<td align=""center"" background=""../image/tab/slc_tab_bgnoact_0" & warna & "0.gif""><nobr><font class=wordnonactivetab>"&mnu(i,1)&"</font></nobr></td>"
					    else
					  	    s=s&"<td align=""center"" background=""../image/tab/slc_tab_bgnoact_0" & warna & "0.gif""><nobr><a class=""Tab"" href="&mnu(i,2)&" title=""" & mnu(i,1) & """><nobr>"&mnu(i,1)&"</a></nobr></td>"
					    end if
					    if i=jmlmenu-1 then
					  	    s=s&"<td><img src=""../image/tab/slc_tab_noactright_0" & warna & "0.gif""</td>"
					    else
					    end if
					    if (i <> jmlmenu - 1) then
				      	    sts = false
				      	else
				      	    sts = true
					    end if
		    end if
		    s=s&"</td>"&vbCrLf
	next
	gen_menu=s
end function

Function GenerateFormatDate(currTanggal, xFormat)
		on error resume next
		if not isdate(currTanggal) then
			GenerateFormatDate = ""
			exit function
		end if
				Dim Tanggalnya
				Dim Hasil
				Tanggalnya = currTanggal
				Dim ArrayBulan(12)
				Dim ArrayHari(7)
				Dim ArrayShortBulan(12)
				Dim ArrayShortHari(7)

				ArrayHari(0) = "Sunday"
				ArrayHari(1) = "Monday"
				ArrayHari(2) = "Tuesday"
				ArrayHari(3) = "Wednesday"
				ArrayHari(4) = "Thursday"
				ArrayHari(5) = "Friday"
				ArrayHari(6) = "Saturday"
				ArrayBulan(0) = "January"
				ArrayBulan(1) = "February"
				ArrayBulan(2) = "March"
				ArrayBulan(3) = "April"
				ArrayBulan(4) = "May"
				ArrayBulan(5) = "June"
				ArrayBulan(6) = "July"
				ArrayBulan(7) = "August"
				ArrayBulan(8) = "September"
				ArrayBulan(9) = "October"
				ArrayBulan(10) = "November"
				ArrayBulan(11) = "December"

				ArrayShortHari(0) = "Sun"
				ArrayShortHari(1) = "Mon"
				ArrayShortHari(2) = "Tues"
				ArrayShortHari(3) = "Wed"
				ArrayShortHari(4) = "Thurs"
				ArrayShortHari(5) = "Fri"
				ArrayShortHari(6) = "Sat"

				ArrayShortBulan(0) = "Jan"
				ArrayShortBulan(1) = "Feb"
				ArrayShortBulan(2) = "Mar"
				ArrayShortBulan(3) = "Apr"
				ArrayShortBulan(4) = "May"
				ArrayShortBulan(5) = "Jun"
				ArrayShortBulan(6) = "Jul"
				ArrayShortBulan(7) = "Aug"
				ArrayShortBulan(8) = "Sep"
				ArrayShortBulan(9) = "Oct"
				ArrayShortBulan(10) = "Nov"
				ArrayShortBulan(11) = "Dec"

				Hari = ArrayHari(Weekday(tanggalnya) - 1)
				HariShort = ArrayShortHari(Weekday(tanggalnya) - 1)
				Bulan = ArrayBulan(month(Tanggalnya) - 1)
				BulanShort = ArrayShortBulan(month(tanggalnya) - 1)
				Tahun = Year(tanggalnya)
                Tanggal = Day(tanggalnya)
                Waktu   = FormatDateTime(tanggalnya,vblongtime)
				Select Case xFormat
					Case "dd mmm yyyy hh:mn:ss"
						Hasil = Tanggal & " " & BulanShort & " " & Tahun & " " & Waktu
					case "DMY"
						Hasil = Tanggal & " " & BulanShort & " " & Tahun
					Case "day,dd mmm yyyy hh:mn:ss"
						Hasil = HariShort & ", " & Tanggal & " " & BulanShort & " " & Tahun & " " & waktu
					Case "daylong,dd mmmlong yyyy"
						Hasil = Hari & ", " & Tanggal & " " & Bulan & " " & Tahun
					Case "daylong,dd mmmlong yyyy hh:mn:ss"
						Hasil = Hari & ", " & Tanggal & " " & Bulan & " " & Tahun & " " & waktu
					Case "daylong,dd mmm yyyy"
						Hasil = Hari & ", " & Tanggal & " " & BulanShort & " " & Tahun
					Case "day,dd mmm yyyy"
						Hasil = HariShort & ", " & Tanggal & " " & BulanShort & " " & Tahun
					Case "hh:mn:ss"
						Hasil = Waktu
					case "kite"
						Hasil = BulanShort &" "& Tanggal&", "&Tahun
					case "kiteheader"
						Hasil = Hari & ", " & Tanggal & " " & BulanShort & ", " & Tahun
					case "mm yyyy"
						Hasil = Bulan & " "&Tahun
					case "mmm dd,yyyy hh:mn"
						Hasil = BulanShort & " " & Tanggal & ", " & Tahun & " " & waktu
					case "dd mm yyyy"
					    if len(Tanggal)=1 then
						Hasil = "0"&Tanggal & " " & BulanShort & " " & Tahun
						else
						Hasil = Tanggal & " " & BulanShort & " " & Tahun
						end if
					case else
						Hasil = Tanggal & " " & Bulan & " " & Tahun
				End Select

				GenerateFormatDate = Hasil
	End Function
	
	function formatcurr(amount, tipecurr)
		dim retval
		if not isnumeric(amount) then
			amount = 0
		end if
		if amount<0 then
			retval = "<font color='red'>("
			if (lcase(trim(tipecurr)) = "rp" or lcase(trim(tipecurr)) = "rp.") then
				retval = retval & formatnumber(amount,0)
			else
				retval = retval & formatnumber(amount,2)
			end if
			retval = retval & ")</font>"
		else
			if (lcase(trim(tipecurr)) = "rp" or lcase(trim(tipecurr)) = "rp.") then
				retval = formatnumber(amount,0)
			else
				retval = formatnumber(amount,2)
			end if
		end if
		formatcurr = retval
	end function
	
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
%>
<head>
    <title><%if currPartID <> "" then
				response.write "[vProfile]>"&Right(currPartID,8)
			end if
	%></title>   
    <LINK href="../css/gnr_all.css" rel=STYLESHEET type="text/css">
	<LINK href="../css/style.css" rel=STYLESHEET type="text/css">	
<script type="text/javascript" language="JavaScript1.2" src="../js/gjs_global.js"></script>
<script language="JavaScript">
	function fjs_viewbundle(){
		if(document.all.tr_viewbundle.style.display=="none"){
			document.all.tr_viewbundle.style.display="block"
		}
		else{
			document.all.tr_viewbundle.style.display="none"
		}
	}
</script>
</head>
<body leftmargin=0 topmargin=0 bgcolor="#25519A" style="text-align: center">
<table border="0" width="100%" id="table246" cellspacing="0" cellpadding="0">
	<tr><td>&nbsp;</td></tr>
 	<tr><td>&nbsp;</td></tr>
   <tr>
   		<td>
		    <!--#include file="../include/glob_tab_top.asp"-->
		    <%=inveprodcatalog_menu("profile") %>
		    <!--#include file="../include/glob_tab_mid.asp"-->
			<!-- Table disini -->
			<br>
				<table border=0 width="95%" align=center  border=0 cellpadding=2 cellspacing=2 border=0>
					<tr>
						<td style="border:solid 1 black" width=50%>
							&nbsp;<font face=verdana,arial size=2><strong><%=currpartid%></strong></font>
						</td>
						<td style="border:solid 1 black" width=50% align=right>
							&nbsp;<font face=verdana,arial size=2><strong><%=currBrandName&" "&currSeri%></strong></font>
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
						<td class="headerdata" height="20" align="center" colspan="3">
							Product Catalog
						</td>
					</tr>			
					<tr>
						<td rowspan=2 width=50% valign=top  style="border: 1px solid <%=xWarnaLine%>">
							<table width=100% topmargin=0 cellpadding="0" cellspacing="0" border="0">
								<tr height="20">
									<td class="field" width="20%">
										<nobr>Category Primary
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="77%"><nobr>
										<%=ceknull(currCatPrimaryName)%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Brand
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currBrandName)%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Seri
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currSeri)%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Description
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" align="justify">
                                   	 	<%response.write currDesc
										%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Jaminan Murah
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=crJaminanMurah%>&nbsp;
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Part Number
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currPartNo)%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Bar Code Desc
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currBarCodeDesc)%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										SN
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%		
										IF (crNeedSN=1 or crNeedSN=True) THEN 
										     response.write("Need Serial Number") 
										ELSE
										     response.write("No Serial Number")  
										END IF
										%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Length
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(crLength)%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Width
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(crWidth)%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Heigth
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(crHeight)%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Marketing Info
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<span style="color:#125;font-size:12;"><b>
											<%response.write currMarketinginfo%>
										</b></span>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								
								<tr height="20">
									<td class="field">
										Ship Weight
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currShipWeight)%> kg
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Bundle
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%
											IF curritembundlemain<>"" THEN 
												response.write "Has bundle <a onclick=""javascript:fjs_viewbundle()"" style=""cursor:hand""><font color=blue>>></font></a></a><br>"
											ELSEIF curritembundledetail<>"" THEN 
												xstring=" select b.vpartid,br.vname vbrandname,cp.vseri from trx_invecpbundle b "&_
														" inner join trx_invecpbundledetail bd on b.vtrxid=bd.vtrxid "&_
														" inner join trx_invecomputer cp on b.vpartid=cp.vpartid "&_
														" inner join tlu_invecpbrand br on br.vbrandid=cp.vbrandid "&_
														" and bd.vbundlepartid='"&currpartid&"'"
												getrs.open xstring,conn
												IF NOT getrs.eof THEN
													while not getrs.eof
														response.write "Bundle of <strong>"&getrs("vbrandname")&"-"&getrs("vseri")&"</strong><br>"
														getrs.movenext
													wend
												ELSE
													ceknull("")
												END IF
												getrs.close
											ELSE
												ceknull("")
											END IF
										%>
									</td>
								</tr>
								<%IF curritembundlemain<>"" THEN %>
								<tr height="20" id="tr_viewbundle" style="display:none">
									<td class="field" valign=top>&nbsp;
									</td>
									<td class="field" valign=top>&nbsp;
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%
											cntbundle=1
											xstring=" select bd.vbundlepartid,br.vname vbrandname,cp.vseri from trx_invecpbundle b "&_
													" inner join trx_invecpbundledetail bd on b.vtrxid=bd.vtrxid "&_
													" inner join trx_invecomputer cp on bd.vbundlepartid=cp.vpartid "&_
													" inner join tlu_invecpbrand br on br.vbrandid=cp.vbrandid "&_
													" and b.vpartid='"&currpartid&"'"
											getrs.open xstring,conn
											while not getrs.eof
												response.write cntbundle&". <a href=""digoff_inve_prodcatalog_view.asp?crpartid="&getrs("vbundlepartid")&""">"&getrs("vbundlepartid")&"</a>, "&getrs("vbrandname")&"&nbsp;"&getrs("vseri")&"<br>"
												cntbundle=cntbundle+1
												getrs.movenext
											wend
											getrs.close
										%>
									</td>
								</tr>
								<%END IF%>
								<tr height="25" style="display:none">
									<td class="field" valign=top>
										Vendor Warranty
									</td>
									<td class="field" width="1%" valign=top>
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" valign=top>
										<%
											'tanda : untuk mengerjakan beberapa statement di dalm satu line
											currVendorWarranty="":currCustWarranty=""
											currModelID="":currLongDefault="":currLongDefaultInv=""
											currtblbuka="<table border=0 >"
											currtbltutup="</table>"
											xstringWarVendor=" select m.vname, s.vlongwarranty from "&_
															 " tlu_invewarrantymodel m inner join tlu_invewarrantymodeldetail d "&_
															 " on m.vmodelid=d.vmodelid and vtype=0 "&_
															 " left join trx_invewarranty s on s.vmodelid=d.vmodelid and s.vtype=0 and s.vpartid='"&currpartid&"'"
											'response.write xstringWarVendor
											set getRsWar= conn.execute(xstringWarVendor)
											IF NOT getRsWar.eof THEN
												currVendorWarranty=currtblbuka
												while not getRsWar.eof 
													currlongwarranty=getRsWar("vlongwarranty")
													currVendorWarranty=currVendorWarranty&_
														"<tr>"&_
														"	<td valign=top><font face=verdana,arial size=1>"&getRsWar("vname")&"</font></td>"&_
														"	<td valign=top><font face=verdana,arial size=1>:</font></td>"&_
														"	<td valign=top><font face=verdana,arial size=1><nobr>"
															IF currlongwarranty<>"" THEN
																select case currlongwarranty
																	case "0":
																		currVendorWarranty=currVendorWarranty&"No Warranty"
																	case "-1"
																		currVendorWarranty=currVendorWarranty&"Life Time"
																	case else
																		currVendorWarranty=currVendorWarranty&ceknull(currlongwarranty)&" month"
																end select
															ELSE
																currVendorWarranty=currVendorWarranty&ceknull("")
															END IF
														currVendorWarranty=currVendorWarranty&"</font></td>"
														currVendorWarranty=currVendorWarranty&"</tr>"
													getRsWar.movenext
												wend
												currVendorWarranty=currVendorWarranty&currtbltutup
											END IF
											getRsWar.close
											set getRsWar=nothing
											IF currVendorWarranty<>"" THEN response.write currVendorWarranty END IF
										%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="25">
									<td class="field" valign=top>
										Customer Warranty
									</td>
									<td class="field" width="1%" valign=top>
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" valign=top>
									<%
										xstringWarCust=" select m.vname, s.vlongwarranty from "&_
														 " tlu_invewarrantymodel m inner join tlu_invewarrantymodeldetail d "&_
														 " on m.vmodelid=d.vmodelid and vtype=1 "&_
														 " left join trx_invewarranty s on s.vmodelid=d.vmodelid and s.vtype=1 and s.vpartid='"&currpartid&"'"
										'response.write xstringWarVendor
										set getRsWar= conn.execute(xstringWarCust)
										IF NOT getRsWar.eof THEN
											currCustomerWarranty=currtblbuka
											while not getRsWar.eof 
												currlongwarranty=getRsWar("vlongwarranty")
												currCustomerWarranty=currCustomerWarranty&_
													"<tr>"&_
													"	<td valign=top><font face=verdana,arial size=1>"&getRsWar("vname")&"</font></td>"&_
													"	<td valign=top><font face=verdana,arial size=1>:</font></td>"&_
													"	<td valign=top><font face=verdana,arial size=1><nobr>"
														IF currlongwarranty<>"" THEN
															select case currlongwarranty
																case "0":
																	currCustomerWarranty=currCustomerWarranty&"No Warranty"
																case "-1"
																	currCustomerWarranty=currCustomerWarranty&"Life Time"
																case else
																	currCustomerWarranty=currCustomerWarranty&ceknull(currlongwarranty)&" month"
															end select
														ELSE
															currCustomerWarranty=currCustomerWarranty&ceknull("")
														END IF
													currCustomerWarranty=currCustomerWarranty&"</font></td>"
													currCustomerWarranty=currCustomerWarranty&"</tr>"
												getRsWar.movenext
											wend
											currCustomerWarranty=currCustomerWarranty&currtbltutup
										END IF
										getRsWar.close
										set getRsWar=nothing
										IF currCustomerWarranty<>"" THEN response.write currCustomerWarranty END IF
									%>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Authorised Warranty
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=crAuthorisedWarranty%>&nbsp;
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Extended Text Warranty
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(crExtTxtWarranty)%>&nbsp;
									</td>
								</tr>
							</table>                
						</td>
						<td width=100% valign=top  style="border: 1px solid <%=xWarnaLine%>">
							<table width=100% topmargin=0 cellpadding="0" cellspacing="0" border="0">
								<tr height="20">
									<td class="field" width="20%">
										Activation
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="77%">
										<%=currActivationName%>&nbsp;
                                      	
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="20%">
										Auto Publish
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="77%">
                                        <a href="http://dos2:90/admin/Products/AutoPublishResult.aspx?searchfield=1&keyword=<%=currPartID%>" target="_blank">Status</a>
										||
                                        <a href="http://dos2:90/admin/Products/AutoPublishSKUUpdate.aspx?ProductTypeID=<%=currPartID%>" onClick="window.open('http://dos2:90/admin/Products/AutoPublishSKUUpdate.aspx?ProductTypeID=<%=currPartID%>', 'AutoPublish', 'left=' + (screen.width - 550) / 2 + ',top=' + (screen.height - 375) / 2 + ',width=550,height=375,scrollbars=yes,resizable=yes,location=no,status=yes,toolbar=no,menubar=no'); return false;">Update</a>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field" width="100">
										Status
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<%=currStatusName%>&nbsp;
									</td>
								</tr>
								<tr height="20" id="tr_periode" <%if currStatusID="" then response.write "style=""display:none"""%>>										
									<td class="field" width="100" valign=top>&nbsp;	
									</td>
									<td class="field" width="1%" valign=top>&nbsp;</td>
									<td width="1%"></td>
									<td class="fielddata">
										<table>
											<tr>
												<td class="fielddata">From</td>
												<td class="fielddata">:</td>
												<td class="fielddata"><%=GenerateFormatDate(currPeriodStart,"dd mmm yyyy")%>&nbsp;</td>
												<td class="fielddata">--&nbsp;&nbsp;To</td>
												<td class="fielddata">:</td>
												<td class="fielddata"><%=GenerateFormatDate(currPeriodEnd,"dd mmm yyyy")%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<%
									if currStsPPn = "1" then
										cntnPrc = currCntnPrice * 1.1
									else
										cntnPrc = currCntnPrice
									end if
								%>
								<tr height="20">
									<td class="field" width="100">
										VPL From
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<%=currSVndNameAdd%>&nbsp;
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="100"><p>VPL<br>
									  Before Tax
									</p></td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currCntnPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF currCntnPrcCurrName<>"" THEN response.write formatcurr(currCntnPrice,currCntnPrcCurrID) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field" width="100">
										VPL<br/>After Tax
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currCntnPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF currCntnPrcCurrName<>"" THEN response.write formatNumberRoundPrice(cntnPrc,currCntnPrcCurrID,2) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="100">
										PPn
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="100%"><span style="width:20%">
										<%
											if currStsPPn="1" then
												response.write "<img align=center src=""../image/sml_yes.gif"" width=10 height=10 title=""Jika kolom PPn dicentang, maka di SO akan muncul COG must add PPn"">"
											else
												response.write "<img align=center src=""../image/sml_no.gif"" width=10 height=10 title=""Jika kolom PPn dicentang, maka di SO akan muncul COG must add PPn"">"
											end if
										%></span>
									</td>
								</tr>
								<tr height="20">
									<td class="field" width="100">&nbsp;
									</td>
									<td class="field" width="1%">&nbsp;
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<font face='verdana' size='1'>Jika kolom <font color=red><strong>PPn</strong></font> dicentang(<img align=center src="../image/sml_yes.gif" valign=middle width=10 height=10>), maka di SO akan muncul <font color=red><strong>Harga Jual harus termasuk 10% PPn <%'=currPPNNilai%></strong>(pembelian item ini menggunakan PPn <%=currPPNNilai%>%)</font></font>&nbsp;
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>                                
								<%
									mrgPct = ""
									mrgVal = ""
									IF currPrice > "0" and currCntnPrice > "0" THEN
										mrgPct = ((currPrice-(currCntnPrice))/(currCntnPrice)) * 100													
										mrgVal = currPrice - currCntnPrice
									END IF
								%>
								<tr height="20">
									<td class="field" width="100">
										Margin (%)
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<%=formatnumber(mrgPct,2)%>&nbsp;
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="100">
										Margin (Value)
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<%=formatnumber(mrgVal,2)%>&nbsp;
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
                                <%
									IF currPrice > "0" THEN
										currWebPrice = currPrice * 1.1
									ELSE
										currWebPrice = currPrice
									END IF
									IF basicCurrency = "CUR01" THEN
										dim dblNumber
										dim intNumber
										dblNumber = CDbl(currWebPrice)
										intNumber = Int(currWebPrice)
										IF intNumber < dblNumber THEN
											Ceiling = intNumber + 1
										ELSE
											Ceiling = intNumber
										END IF
										currWebPrice=Ceiling
									END IF
								%>
								<tr height="20">
									<td class="field" width="100">
									Web Price<br>
									Before Tax</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF currPrcCurrName<>"" THEN response.write formatnumber(currPrice) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="100">
									Web Price<br>
									After Tax</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF currPrcCurrName<>"" THEN response.write formatcurr(currWebPrice,currPrcCurrID) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field" width="100">
										Special Price<br>Before Tax
									</td>
									<td class="field" width="1%" valign=top>
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currSPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF trim(currSPrcCurrName)<>"" THEN response.write formatnumber(currSPrice) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
                                <tr height="5"><td colspan="4"></td></tr>
                                <tr height="20">
									<td class="field" width="100">
										Special Price<br>After Tax
									</td>
									<td class="field" width="1%" valign=top>
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">
										<table width=100%>
											<tr>
												<td class="fielddata"><%=ceknull(currSPrcCurrName)%>&nbsp;</td>
												<td class="fielddata" align=right><%IF trim(currSPrcCurrName)<>"" THEN response.write formatnumber(currSpecialPrice,2) END IF%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="20">
									<td class="field" valign=top>&nbsp;
									</td>
									<td width="1%" class="field">&nbsp;
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<table>
											<tr>
												<td class="fielddata">From</td>
												<td class="fielddata">:</td>
												<td class="fielddata"><%=GenerateFormatDate(currSPValidStart,"dd mmm yyyy")%>&nbsp;</td>
												<td class="fielddata">--&nbsp;&nbsp;To</td>
												<td class="fielddata">:</td>
												<td class="fielddata"><%=GenerateFormatDate(currSPValidEnd,"dd mmm yyyy")%>&nbsp;</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field" width="100" valign=top>
										Last Update Price
									</td>
									<td class="field" width="1%" valign=top>
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="315" valign=top>
										<%=GenerateFormatDate(currLastUpdatePrc,"dd mmm yyyy hh:mn:ss")%>&nbsp;<%=currPriceEditorName%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field" width="100">
										Manufacturer
									</td>
									<td class="field" width="1%">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata" width="115">&nbsp;
										<%IF NOT isnull(currManufacturer) and currManufacturer<>"" THEN %>
											<nobr><a href="http://<%=ceknull(currManufacturer)%>">http://<%=ceknull(currManufacturer)%></a>
										<%END IF%>
									</td>
								</tr>
								<tr height="5"><td colspan="4"></td></tr>
								<tr height="20">
									<td class="field">
										Note
									</td>
									<td class="field">
										:
									</td>
									<td width="1%"></td>
									<td class="fielddata">
										<%=ceknull(currNote)%>
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
                <%response.write "<br>"%>
                <% 
				IF adaPromo = 1 THEN
					set rsPromoBP = Server.CreateObject("ADODB.Recordset")
					queryBP = "select * from trx_InveCPBundleDetail where vTrxID = '"&currBundleID&"' and vPromo = 0 order by vNoUrut ASC "
					rsPromoBP.open queryBP, conn, 3, 1, 0
					IF NOT rsPromoBP.eof THEN
				%>
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
                                            <tr style="border: 1px solid <%=xWarnaLine%>">
                                                <td <%=bgnoact%> width="10%" align="center" title="Type"><font class="wordblackmedium"><nobr><strong>Type</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="10%" align="center" title="SKU"><font class="wordblackmedium"><nobr><strong>SKU</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="35%" align="center" title="ItemDesc"><font class="wordblackmedium"><nobr><strong>Item Description</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="5%" align="center" title="Qty"><font class="wordblackmedium"><nobr><strong>Qty</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="15%" align="center" title="Start Date"><font class="wordblackmedium"><nobr><strong>Start Date</strong></nobr></font></td>
                                       			<td <%=bgnoact%> width="15%" align="center" title="End Date"><font class="wordblackmedium"><nobr><strong>End Date</strong></nobr></font></td>
												<td <%=bgnoact%> width="10%" align="center" title="Tag BP"><font class="wordblackmedium"><nobr><strong>Tag Date</strong></nobr></font></td>
											</tr>
                                            <%
											for bp = 0 to rsPromoBP.RecordCount - 1
											%>
                                            <tr>
                                                <td class="fielddata" valign=top align=center title="Promo" style="border: 1px solid">
                                                	<%IF rsPromoBP("vType") = 0 THEN Response.Write("Free") END IF%>
                                                    <%IF rsPromoBP("vType") = 1 THEN Response.Write("Bundle") END IF%>
                                                </td>
                                                <td class="fielddata" valign=top align=center title="SKU" style="border: 1px solid">
                                                    <%=rsPromoBP("vBundlePartID")%>
                                                </td>
                                                <td class="fielddata" align="justify" title="Product" valign="top" style="border: 1px solid">
                                                    <%
                                                        set rsBrandSeriDesc = server.CreateObject("ADODB.Recordset")
                                                        queryBrandSeriDesc = "select vBrandID, vSeri, vDesc from trx_InveComputer where vPartID = '"&rsPromoBP("vBundlePartID")&"' "
                                                        rsBrandSeriDesc.open queryBrandSeriDesc, conn, 3, 1, 0
                                                        IF NOT rsBrandSeriDesc.eof THEN
                                                            tempBrandID = rsBrandSeriDesc("vBrandID")
                                                            tempSeri = rsBrandSeriDesc("vSeri")
                                                            tempDesc = rsBrandSeriDesc("vDesc")
                                                            set rsBrandName = server.CreateObject("ADODB.Recordset")
                                                            queryBrandName = "select vName from tlu_InveCPBrand where vBrandID = '"&tempBrandID&"' "
                                                            rsBrandName.open queryBrandName, conn, 3, 1, 0
                                                            IF NOT rsBrandName.eof THEN
                                                                tempBrandName = rsBrandName("vName")
                                                            END IF
                                                        END IF
                                                    %>
                                                    <%=tempBrandName%><br><%=tempSeri%><br><%=tempDesc%>
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
                                                <td class="fielddata" valign=top align=center title="Quantity" style="border: 1px solid">
                                                    <%=rsPromoBP("vQty")%>
                                                </td>
                                                <td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <span>
                                                        <%=Day(rsPromoBP("vStartDateTime"))&" "&MonthName(Month(rsPromoBP("vStartDateTime")),1) &" "&Year(rsPromoBP("vStartDateTime"))%>
                                                    </span>                                              
                                                </td>
                                                <td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <span> 
                                                        <%=Day(rsPromoBP("vEndDateTime"))&" "&MonthName(Month(rsPromoBP("vEndDateTime")),1) &" "&Year(rsPromoBP("vEndDateTime"))%>
                                                    </span>
                                                </td>
												<td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <%if rsPromoBP("vTagDate")=true then response.write "<img src=""../image/sml_yes.gif"">" else response.write "-"%>
                                                </td>
                                            </tr>
											<%
												rsPromoBP.MoveNext
											next
											rsPromoBP.close
											set rsPromoBP = nothing
											%>
                                		</table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <%
					end if
				%>
                <%
					set rsPromoSD = Server.CreateObject("ADODB.Recordset")
					querySD = "select * from trx_InveCPBundleDetail where vTrxID = '"&currBundleID&"' and vPromo = 1 order by vNoUrut ASC "
					rsPromoSD.open querySD, conn, 3, 1, 0
					if not rsPromoSD.eof then
				%>
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
                                                <td <%=bgnoact%> width="35%" align="center" title="ItemDesc"><font class="wordblackmedium"><nobr><strong>Item Description</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="5%" align="center" title="Qty"><font class="wordblackmedium"><nobr><strong>Qty</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="15%" align="center" title="Start Date"><font class="wordblackmedium"><nobr><strong>Start Date</strong></nobr></font></td>
                                                <td <%=bgnoact%> width="15%" align="center" title="End Date"><font class="wordblackmedium"><nobr><strong>End Date</strong></nobr></font></td>
                                            	<td <%=bgnoact%> width="10%" align="center" title="Tag BP"><font class="wordblackmedium"><nobr><strong>Tag Date</strong></nobr></font></td>
                                            </tr>
                                            <%
											for sd = 0 to rsPromoSD.RecordCount - 1
											%>
                                            <tr id="trSD_<%=sd%>" name="trSD">
                                                <td class="fielddata" valign=top align=center title="Promo" style="border: 1px solid">
													<%IF rsPromoSD("vType") = 0 THEN Response.Write("Free") END IF%>
                                                    <%IF rsPromoSD("vType") = 1 THEN Response.Write("Bundle") END IF%>
                                                </td>
                                                <td class="fielddata" valign=top align=center title="SKU" style="border: 1px solid">
                                                    <%=rsPromoSD("vBundlePartID")%>
                                                </td>
                                                <td class="fielddata" align="left" title="Product" valign="top" style="border: 1px solid">
                                                	<%
														set rsBrandSeriDesc = server.CreateObject("ADODB.Recordset")
														queryBrandSeriDesc = "select vBrandID, vSeri, vDesc from trx_InveComputer where vPartID = '"&rsPromoSD("vBundlePartID")&"' "
														rsBrandSeriDesc.open queryBrandSeriDesc, conn, 3, 1, 0
														IF NOT rsBrandSeriDesc.eof THEN
															tempBrandID = rsBrandSeriDesc("vBrandID")
															tempSeri = rsBrandSeriDesc("vSeri")
															tempDesc = rsBrandSeriDesc("vDesc")
															set rsBrandName = server.CreateObject("ADODB.Recordset")
															queryBrandName = "select vName from tlu_InveCPBrand where vBrandID = '"&tempBrandID&"' "
															rsBrandName.open queryBrandName, conn, 3, 1, 0
															IF NOT rsBrandName.eof THEN
																tempBrandName = rsBrandName("vName")
															END IF
														END IF
													%>
                                                    <%=tempBrandName%><br><%=tempSeri%><br><%=tempDesc%>
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
                                                <td class="fielddata" valign=top align=center title="Quantity" style="border: 1px solid">
                                                    <%=rsPromoSD("vQty")%>
                                                </td>
                                                <td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <span>
                                                        <%=Day(rsPromoSD("vStartDateTime"))&" "&MonthName(Month(rsPromoSD("vStartDateTime")),1) &" "&Year(rsPromoSD("vStartDateTime"))%>
                                                    </span>                                               
                                                </td>
                                                <td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <span>
                                                        <%=Day(rsPromoSD("vEndDateTime"))&" "&MonthName(Month(rsPromoSD("vEndDateTime")),1) &" "&Year(rsPromoSD("vEndDateTime"))%>
                                                    </span>
                                                </td>
                                                <td class="fielddata" valign=top align=center style="border: 1px solid">
                                                    <%if rsPromoSD("vTagDate")=true then response.write "<img src=""../image/sml_yes.gif"">" else response.write "-"%>
                                                </td>
                                            </tr> 
											<%
												rsPromoSD.MoveNext
											next
											rsPromoSD.close
											set rsPromoSD = nothing
											%>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            	<%
					END IF
				%>
                <table height="5px">
                    <tr>
                        <td>&nbsp;</td>
                    </tr>
                </table>
                <%END IF%>                      
				<table border=0 width="95%" align=center  border=0 cellpadding=2 cellspacing=2 border=0>
					<tr>
						<td class="headerdata" height="20" align="center">
							Secondary Category
						</td>
						<td class="headerdata" height="20" align="center">
							Vendor
						</td>
					</tr>			
					<tr>
						<td width=50% valign=top  style="border: 1px solid <%=xWarnaLine%>">
							<table border="0" cellpadding=0 cellspacing=0 width=100%>
								<tr><td height=2></td></tr>
								<tr>
									<td width=4% height=17 <%=bgnoact%> align="center" title="No"><font class="wordblackmedium"><nobr><strong>No</strong></nobr></font></td>
									<td width=5% height=17 <%=bgnoact%> align="center" title="Category ID"><font class="wordblackmedium"><nobr><strong>ID</strong></nobr></font></td>
									<td width=90% height=17 <%=bgnoactright%> align="center" title="Category Name"><font class="wordblackmedium"><nobr><strong>Name</strong></nobr></font></td>
								</tr>
								<tr><td height=2></td></tr>
								<%
									xstring=" select r.vcatid,c.vname from trx_inveCPCatRel r, tlu_inveCPCategory c "&_
											" where r.vcatid=c.vcatid and r.vtypecat='S' and r.vpartid='"&currPartId&"'"
									getRs.open xstring,conn,1,3
									IF NOT getRs.eof THEN
										i=0
										while not getRs.eof
										currColor1 = xMouseOverGrid
										currColor2 = xEvenGrid
										currBackGround = ""
										IF (i mod 2 = 1) THEN
											currBackGround = currColor2
										ELSE
											currBackGround = "#ffffff"
										END IF
								%>
								<tr bgColor="<%=currBackGround%>" onMouseOver="bgColor='<%=currColor1%>'" onMouseOut="bgColor='<%=currBackGround%>'">
									<td class="wordfield" height=20><%=i+1%>&nbsp;</td>
									<td class="wordfield"><%=getRs("vcatid")%>&nbsp;</td>
									<td class="wordfield"><%=getRs("vname")%></td>
								</tr>
								<%
										i=i+1
										getRs.MoveNext
										wend
									ELSE
								%>
									<tr>
										<td colspan=3 class="wordfield"><font color=red>Data not Found</font></td>
									</tr>
								<%
									END IF
									getRs.close
								%>
							</table>
						</td>
						<td width=50% valign=top  style="border: 1px solid <%=xWarnaLine%>">
							<table border="0" cellpadding=0 cellspacing=0 width=100%>
								<tr><td height=2></td></tr>
								<tr>
									<td width=4% height=17 <%=bgnoact%> align="center" title="No"><font class="wordblackmedium"><nobr><strong>No</strong></nobr></font></td>
									<td width=31% height=17 <%=bgnoact%> align="center" title="Vendor Name"><font class="wordblackmedium"><nobr><strong>Vendor Name</strong></nobr></font></td>
									<td width=30% height=17 <%=bgnoact%> align="center" title="Phone"><font class="wordblackmedium"><nobr><strong>Phone</strong></nobr></font></td>
									<td width=20% height=17 <%=bgnoact%> align="center" title="Sales"><font class="wordblackmedium"><nobr><strong>Sales</strong></nobr></font></td>
									<td width=10% height=17 <%=bgnoact%> align="center" title="Primary Vendor"><font class="wordblackmedium"><nobr><strong>P</strong></nobr></font></td>
									<td width=5% height=17 <%=bgnoactright%> align="center" title="Status Updater"><font class="wordblackmedium"><nobr><strong>Stat</strong></nobr></font></td>
								</tr>
								<tr><td height=2></td></tr>
							    <%
									xstring=" select  m.vvndid,vname,b.vLPOID,isnull(b.vvndtype,'-')vvndtype,dbo.f_GetVndAddressAndPhone(m.vvndid,'phone','') phone, "&_
											" dbo.f_GetProductAndSales(m.vVndID,'sales') sales, b.vVPLFromID "&_
											" from trx_vendormain m,trx_invebuyingvendor b "&_
											" where b.vvndid=m.vvndid and b.vpartid='"&currPartId&"' order by vVPLFromID DESC, vVndType DESC"
									
									getRs.open xstring,conn,1,3
									IF NOT getRs.eof THEN
										i=0
										while not getRs.eof
										currColor1 = xMouseOverGrid
										currColor2 = xEvenGrid
										currBackGround = ""
										currupdater=getRs("vLPOID")
										IF NOT isnull(currupdater) THEN
											xcurrupdater = "<b>PO</b>"
										ELSE
											xcurrupdater = "<b>CG</b>"
										END IF
										IF (i mod 2 = 1) THEN
											currBackGround = currColor2
										ELSE
											currBackGround = "#ffffff"
										END IF
								%>
								<tr bgColor="<%=currBackGround%>" onMouseOver="bgColor='<%=currColor1%>'" onMouseOut="bgColor='<%=currBackGround%>'">
									<td class="wordfield" height=20 valign=top align="center"><%=i+1%>&nbsp;</td>
									<td class="wordfield" valign=top align="center"><a href="digoff_vendor_view.asp?crBhs=2&crVndID=<%=getRs("vvndid")%>"><%=getRs("vname")%></td>
									<td class="wordfield" valign=top align="center"><%=getRs("phone")%></td>
									<td class="wordfield" valign=top align="center"><%=getRs("sales")%></td>
									<td class="wordfield" valign=top align="center"><%if getRs("vvndid")= getRs("vVPLFromID") then response.write "<img src=""../image/sml_yes.gif"">" else response.write "-"%></td>
									<td class="wordfield" valign=top align="center"><%=xcurrupdater%></td>
								</tr>
								<%
										i=i+1
										getRs.MoveNext
										wend
									ELSE
								%>
									<tr>
										<td colspan=3 class="wordfield"><font color=red>Data not Found</font></td>
									</tr>
								<%
									END IF
									getRs.close
								%>
							</table>
						</td>
					</tr>
				</table>
                <br>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=center>
				<table width=95% border=0>
					<tr>
						<td><!-- #include file="../include/glob_footer_editor.asp" --></td>
					</tr>
				</table>
                <br>
			</td>
		</tr>
		<tr>
			<td colspan=2 align=right class=""wordfield"" >
				<font size=2><strong>What To do?</strong>
                [&nbsp;<a href="digoff_inve_prodcatalog.asp?crAct=Add&crCopy=Yes&crPartId=<%=currPartId%>&crLoad=1">Copy</a>&nbsp;]
                	&nbsp;&nbsp;&nbsp;||&nbsp;&nbsp;&nbsp;
                [&nbsp;<a href="digoff_inve_prodcatalog.asp?crAct=Edit&crPartId=<%=currPartId%>&crBhs=2&crLoad=1">Edit</a>&nbsp;]&nbsp;&nbsp;
				</font>
            </td>
		</tr>
	   <!--#include file="../include/glob_tab_bottom.asp"-->
	   <tr><td>&nbsp;</td></tr>
   	</table>
<%
set List = new CListPage
response.write List.GenerateCloseContent
response.write List.GenerateFooter
set List = nothing
set getRs = nothing
%>
</body>

<LINK href="../css/inve_prodcatalog.css" rel=STYLESHEET type="text/css">
<!--#include file="../include/glob_conn_close.asp" -->
</html>