************************************************************************
* ActiveVFP Version 6
*  - VFP Web development using a VFP MTDLL with ASP.NET
* http://activevfp.codeplex.com
* Author: CKF
* Last Modified: 8/22/2012
*********************************************************************************
#DEFINE crlf CHR(13)+CHR(10)
*********************************
Define Class ActiveVFP AS Session
*********************************
* Base objects
oRequest                = Null
oResponse               = Null
oSession                = Null
oContext                = Null
oApplication            = Null
oServer                 = Null
***************************
* ActiveVFP :: Init
Procedure Init() as Boolean
***************************
SET RESOURCE OFF
SET EXCLUSIVE OFF
SET CPDIALOG OFF
SET DELETED ON
SET EXACT OFF
SET SAFETY OFF
SET REPROCESS TO 2 SECONDS
*IF Application.StartMode == 5
	THIS.oRequest=NEWOBJECT('AVFPrequest')
	THIS.oSession=NEWOBJECT('AVFPsession')
	oMTS = CreateObject("MTxAS.AppServer.1")
	THIS.oContext = oMTS.GetObjectContext()
	THIS.oApplication=THIS.oContext.Item("Application")
	THIS.oServer=THIS.oContext.Item("Server")
	THIS.oResponse=THIS.oContext.Item("Response")
	THIS.oRequest.Load(THIS.oContext.Item("Request")) && modify request object
	If Type("THIS.oContext.Item('Session')") = 'O'    
		* Sessions might be disabled
		THIS.oSession.Load(THIS.oContext.Item("Session")) && modify session object
	ENDIF

DODEFAULT()
RETURN

************************************************************************
* ActiveVFP :: ERROR
************************************************************************
FUNCTION ERROR(nError, cMethod, nLine)
LOCAL lcDetails,i,x,lcErrMsg,lcAppStartPath,lcAppName

lcAppStartPath=JustPath(Application.ServerName)+"\"
lcAppName=JustStem(Application.ServerName)
lcDetails=[]

IF [MERGESCRIPT] $ UPPER(SYS(16,PROGRAM(-1)-2)) && if the previous was mergescript
	nline=nline -4 && compensate for script error line discrepancy
ENDIF
IF nError = 10
	lcDetails = " Compilation Error: Syntax error in the script (<B>"+UPPER(oProp.RunningPrg)+"</B>)that was called. "
ENDIF
AERROR(laErr)
FOR i=1 TO ALEN(laErr,2)
	FOR x=1 TO ALEN(laErr,1)
		lcDetails=lcDetails+TRANSFORM(laErr(x,i))+crlf
	ENDFOR
ENDFOR

lcErrMsg = '<B>'+UPPER(oProp.RunningPrg)+'</B>  ' +cMethod+'  err#='+STR(nError,5)+IIF(nError = 10,[],' line='+STR(nline,6))+;
	' '+MESSAGE()+lcDetails+' '+_VFP.SERVERNAME  
	
	COMreturnerror(lcErrMsg,_VFP.SERVERNAME)

ENDFUNC
ENDDEFINE
*********************************************************************************
*********************************************************************************
DEFINE CLASS AVFPhtml AS CUSTOM
*************************************************************
cAppStartPath = ""
cAppName = ""
cAction = ""
cSessID=""
*********************************************************************************
************************************************************************
* AVFPhtml :: htmltable
************************************************************************
* HTMLTABLE - eval html tables
*
* Last Modified: 6/26/09
*
***  1.) lcTableTag   -Name of the HTML table (from <htmltablename> in HTML)
***  2.) lcHTMLtable  -HTML table text to be evaluated
***  3.) lnTotPerPage -Total number of records per page
***  4.) lnpagenumbers-Include page numbers/# of page numbers to show(need to include lnStart and lcButton params)
***  5.) lcBar        -Alternating color for HTML table records (for example, #FFFFFF for white)
***  6.) llTableRec   -(fast write)<table>record</table>&do a response.write - for fastest output(usually a large list on 1 page) 
***  7.) lnStart      -Navigational - Page number to start list from, from URL
***  8.) lcButton     -Navigational - First, last, next, previous, from URL
************************************************************************
FUNCTION htmltable
LPARAMETERS lcTableTag,lcHTMLtable,lnTotPerPage,lnpagenumbers,lcBar,;
	llTableRec,lnStart,lcButton
LOCAL lnRowCount,lnTotPages,lnAtPos,lnAtPos1,lnAtPos2,lcHTML,;
	lnPageMax,lnPageBegin,lcPages,lnOccur,lcHTMLstr,lcStr1,lcStr2,;
	lnCount,lnX,lnZ,lcBarStr,lcDivStart,lcDivEnd, llpagenumbers
IF lnpagenumbers > 0 .AND. ISNULL(lnpagenumbers) != .T.
	llpagenumbers=.T.
ELSE
    llpagenumbers=.F.
ENDIF
lcHTML=''
lcHTMLstr=lcHTMLtable  && find the HTML table stuff
lnAtPos1 = ATC('&lt;'+ALLTRIM(lcTableTag)+'&gt;', lcHTMLstr)
IF lnAtPos1 = 0
	lnAtPos1 = ATC('<'+ALLTRIM(lcTableTag)+'>', lcHTMLstr)
	lcDivStart=THIS.GetString(SUBSTR(lcHTMLstr,lnAtPos1),;
		'<'+ALLTRIM(lcTableTag)+'>','<TABLE ')
ELSE
	lcDivStart=THIS.GetString(SUBSTR(lcHTMLstr,lnAtPos1),;
		'&lt;'+ALLTRIM(lcTableTag)+'&gt;','<TABLE ')
ENDIF
lnAtPos2 = ATC('&lt;/'+ALLTRIM(lcTableTag)+'&gt;', lcHTMLstr)
IF lnAtPos2 = 0
	lnAtPos2 = ATC('</'+ALLTRIM(lcTableTag)+'>', lcHTMLstr)
	lcDivEnd=THIS.GetString(SUBSTR(lcHTMLstr,lnAtPos1),;
		'</TABLE>','</'+ALLTRIM(lcTableTag)+'>')
ELSE
	lcDivEnd=THIS.GetString(SUBSTR(lcHTMLstr,lnAtPos1),;
		'</TABLE>','&lt;/'+ALLTRIM(lcTableTag)+'&gt;')
ENDIF
lnOccur=OCCURS('<TABLE ',UPPER(lcHTMLstr))
FOR lnCount = 1 TO lnOccur
	lnAtPos=ATC('<TABLE ',lcHTMLstr,lnCount)
	IF lnAtPos >lnAtPos1 .AND. lnAtPos< lnATPos2
		EXIT
	ENDIF
ENDFOR
IF lnAtPos > 0
	lcStr1     = LEFT(lcHTMLstr, lnAtPos - 1)
	lcHTMLstr = SUBSTR(lcHTMLstr, lnAtPos)
	lnAtPos    = ATC('</TABLE>', lcHTMLstr)
	lcStr2    = LEFT(lcHTMLstr, lnAtPos - 1)
	lcHTMLstr = SUBSTR(lcHTMLstr, lnAtPos)
	IF llTableRec
		lcTableStr="<Table "+THIS.GetString(lcStr2,"<TABLE ",">")+">"
	ENDIF
	lnAtPos = ATC('<TR', lcStr2)
	IF lnAtPos > 0
		lcStr1  = LEFT(lcStr2, lnAtPos - 1)
		lcStr2 = SUBSTR(lcStr2, lnAtPos)
		lcHTML = lcHTML + THIS.mergetext(lcStr1)
	ENDIF
	DO WHILE ATC('<TH', lcStr2) > 0
		lnAtPos = ATC('</TR>', lcStr2)
		lcStr1  = LEFT(lcStr2, lnAtPos + 5)
		lcStr2 = SUBSTR(lcStr2, lnAtPos + 5)
		lcHTML = lcHTML + IIF(llTableRec,lcTableStr+THIS.mergetext(lcStr1)+[</TABLE>],;
			THIS.mergetext(lcStr1))
	ENDDO
	lnRowCount = RECCOUNT()    && start page navigation stuff
	RELEASE START
	PUBLIC START
	IF llpagenumbers
		START = NVL(lnStart,1)
		IF START = 0 
			START= 1
		ENDIF
		lnPageMax = START * lnTotPerPage
		lnPageBegin = (lnPageMax - lnTotPerPage)+1
		IF lnRowCount < lnTotPerPage
			lnTotPages = 1
		ELSE
			IF MOD(lnRowCount, lnTotPerPage) > 0
				lnTotPages = INT(lnRowCount / lnTotPerPage) + 1
			ELSE
				lnTotPages = INT(lnRowCount / lnTotPerPage)
			ENDIF
		ENDIF		
	    oSession.value("totpages",lnTotPages)
		DO CASE
		CASE lcButton="First"
			START=1
			lnPageBegin=1
			lnPageMax=lnTotPerPage
		CASE lcButton="Prev"
			IF lnPageBegin < 1 .OR. START -1 = 0
				START=1
				lnPageBegin=1
				lnPageMax=lnTotPerPage
			ELSE
				START=START-1
				lnPageBegin=lnPageBegin-lnTotPerPage
				lnPageMax=lnPageMax-lnTotPerPage
			ENDIF
		CASE lcButton="Next"
			START=START+1
			IF START>lnTotPages
				START = lnTotPages
				lnPageMax =  START * lnTotPerPage
				lnPageBegin = (lnPageMax - lnTotPerPage)+1
				lnPageMax = lnRowCount
			ELSE
				lnPageBegin=lnPageBegin+lnTotPerPage
				lnPageMax=lnPageMax+lnTotPerPage
			ENDIF
		CASE lcButton="Last"
			START=lnTotPages
			lnPageMax =  START * lnTotPerPage
			lnPageBegin = (lnPageMax - lnTotPerPage)+1
			lnPageMax = lnRowCount
		OTHERWISE
			IF EMPTY(START)
				START=1
				lnPageBegin=1
				lnPageMax=lnTotPerPage
			ENDIF
		ENDCASE
	ELSE
		START=1
		lnPageBegin=1
		lnPageMax=lnTotPerPage
	ENDIF
	IF llTableRec  && writing out line by line
		lnTH=OCCURS('</th>',lcHTMLtable)
    	lnTHat=AT('<th',lcHTMLtable)
	    lnTHat2=AT('</th>',lcHTMLtable,lnTH)
	    oResponse.Write(ALLTRIM(This.mergetext(SUBSTR(lcHTMLtable,1,lnTHat-1))+[</tr>]))
	    oResponse.Write(lcTableStr+This.mergetext(SUBSTR(lcHTMLtable,lnTHat,(lnTHat2-lnTHat)+5)))
    ENDIF
	FOR lnX = lnPageBegin TO IIF(llpagenumbers,lnPageMax,lnRowCount) && write out those table recs	
		IF lnX <= lnRowCount
			GOTO lnX
			IF !EMPTY(lcBar) && alternating row colors
				IF MOD(lnX,2) = 0
					lcBarStr=STRTRAN(lcStr2,[<TD],[<TD bgcolor="]+lcBar+["])   && lcStr2 is row html with vfp variables 
					lcBarStr=STRTRAN(lcBarStr,[<td],[<td bgcolor="]+lcBar+["])
					lcBarStr=STRTRAN(lcBarStr,[<Td],[<td bgcolor="]+lcBar+["])
					lcBarStr=STRTRAN(lcBarStr,[<tD],[<td bgcolor="]+lcBar+["])
					IF llTableRec
					    lcHTML = lcTableStr+THIS.mergetext(lcBarStr)
					ELSE
					 	lcHTML = lcHTML +THIS.mergetext(lcBarStr)  && eval html row data
					ENDIF
				ELSE					
					IF llTableRec
					    lcHTML=lcTableStr+THIS.mergetext(lcStr2)
					ELSE
						lcHTML = lcHTML + THIS.mergetext(lcStr2)  && eval html row data
					ENDIF
				ENDIF
			ELSE
				IF llTableRec
					lcHTML = lcTableStr+THIS.mergetext(lcStr2)
				ELSE 
					lcHTML = lcHTML + THIS.mergetext(lcStr2)
				ENDIF
			ENDIF
			IF llTableRec
				oResponse.Write(lcHTML+[</TABLE>])
				oResponse.Flush
			ENDIF
		ENDIF
	ENDFOR
	IF llTableRec
		oResponse.Write([</TD></TR></TBODY></TABLE></BODY></HTML>])
	ELSE
		lcHTML=THIS.mergetext(lcDivStart)+lcHTML+"</TABLE>"+THIS.mergetext(lcDivEnd)
		lcHTMLtable  =STUFF(lcHTMLtable, lnAtPos1, lnAtPos2 - lnAtPos1+;
			LEN(lcTableTag)+9, lcHTML)
		IF llpagenumbers    && pages navigation
			lcPages=''   
			IF lnTotPages > 1
			    lngroupnumber = ceiling(START/lnpagenumbers)
				FOR lnZ = lngroupnumber*lnpagenumbers-(lnpagenumbers-1) TO IIF(lnTotpages<lngroupnumber*lnpagenumbers,lnTotPages,lngroupnumber*lnpagenumbers) &&lnTotPages
				      IF lnZ=START
				        lcPages=lcPages+[ <B>]+ALLTRIM(STR(lnZ))+[</B> ]+crlf
				      ELSE
						lcPages=lcPages+[ <a href="]+oProp.ScriptPath +[?action=];
						+oProp.Action+[&sid=]+oProp.SessID+[&page=]+ALLTRIM(STR(lnZ))+[&nav=]+[">];
						+ALLTRIM(STR(lnZ))+[</a> ]+crlf
					  ENDIF
				ENDFOR
			ENDIF
			lnAtPos = ATC("&lt;Pages&gt;",lcHTMLtable)
			IF lnAtPos = 0
				lnAtPos = ATC("<Pages>",lcHTMLtable)
			ENDIF
			IF ! EMPTY(lcPages) .AND. lnAtPos > 0
				lcHTMLtable= STUFF(lcHTMLtable,lnAtPos,13,lcPages)+crlf
			ELSE
				lcHTMLtable= STUFF(lcHTMLtable,lnAtPos,13,"")+crlf
			ENDIF
		ENDIF
	ENDIF
ENDIF
RETURN lcHTMLtable
************************************************************************
* AVFPhtml :: mergetext
* MERGETEXT - eval string containing vfp expressions
*
* Last Modified: 7/14/08
************************************************************************
FUNCTION mergetext
LPARAMETER lcStr
LOCAL lnAtP1,lnAtP2,lcValue

* used for error handling to return vfp expression in error
*!*	PRIVATE cEval
cEval=""

* translate needed characters
lcStr=STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(;
	lcStr,'&lt;',[<]),;
	'&gt;',[>]),'&quot;',["]),;
	'&lt;%','<%'),;
	'%&gt;','%>'),'&amp;',[&])

* eval expression in string
DO WHILE ATC('<%=', lcStr) > 0
	lnAtP1 = ATC('<%=', lcStr)
	lnAtP2 = ATC('%>', lcStr)
	cEval  = CHRTRAN(SUBSTR(lcStr,lnAtP1 + 3,lnAtP2 - lnAtP1 - 3),crlf,'')
	lcValue = EVAL(cEval)
    oProp.cEval = cEval
* always return character
	IF VARTYPE(lcValue) # "C"
		lcValue = TRANSFORM(lcValue)
	ENDIF

	lcStr  = STUFF(lcStr, lnAtP1, lnAtP2 - lnAtP1 + 2, lcValue)
ENDDO

RETURN lcStr
*
************************************************************************
* AVFPhtml :: mergescript
* MERGESCRIPT - execute embedded VFP code blocks with HTML
*
* Last Modified: 12/30/11
************************************************************************
FUNCTION mergescript
LPARAMETER lcStr
LOCAL lcExec,lcReplace
cEval=""
* translate needed characters
lcStr=STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(;
	lcStr,'&lt;',[<]),;
	'&gt;',[>]),'&quot;',["]),;
	'&lt;%','<%'),;
	'%&gt;','%>'),'&amp;',[&])

lcExec = "start"
DO WHILE !EMPTY(lcExec)
	lcExec = THIS.GetString(lcStr,"<%=","%>")
	IF !EMPTY(lcExec)
	   lcReplace = "<%="+lcExec+"%>"
	   lcStr = STRTRAN(lcStr,lcReplace,"<<"+ALLTRIM(lcExec)+">>")
	ENDIF
ENDDO
lcStr = STRTRAN(lcStr,"<%",CHR(13)+"ENDTEXT"+CHR(13))

cEval = "PRIVATE strVar"+CHR(13)+"TEXT TO strVar TEXTMERGE NOSHOW"+CHR(13);
 + STRTRAN(lcStr,"%>",CHR(13)+"TEXT TO strVar TEXTMERGE NOSHOW ADDITIVE"+CHR(13)) + ;
                CHR(13)+"ENDTEXT"+CHR(13)+"RETURN strVar" 
     
	   IF [INCLUDE] $ SYS(16,PROGRAM(-1)-1)   && prevent recursive error in TextMerge/ExecScript 
	        SET PROCEDURE TO codeblck ADDITIVE
            oCodeBlock = CREATEOBJECT( "cusCodeBlock") 
			oCodeBlock.SetCodeBlock(cEval) 
			oCodeBlock.lAccumulateMergedText = .T.
			SET TEXTMERGE ON
			oCodeBlock.Execute() 
			lcStr=oCodeBlock.GetMergedText()
	   ELSE
	        lcStr=EXECSCRIPT(cEval)  && Microsoft EXECSCRIPT seems to be reliable in VFP9
	   ENDIF	

RETURN lcStr

*
************************************************************************
* AVFPhtml :: GetString
* NOTE: This is for VFP 6 compatibility. For VFP 7 and above, can use STREXTRACT
* Last Modified: 3/19/08
*************************************************************************
FUNCTION GetString
LPARAMETERS lcStr,lcD1,lcD2
LOCAL lnStart
lnStart=ATC(lcD1,lcStr)
IF lnStart=0
   RETURN ""
ENDIF
lnStart=lnStart+len(lcD1)
lcRem=SUBSTR(lcStr,lnStart)
lnLast=ATC(lcD2,lcRem)
IF lnLast>0
   RETURN SUBSTR(lcRem,1,lnLast-1)
ENDIF   
ENDFUNC
*
************************************************************************
*
* AVFPhtml :: htmldropdown
*
* Last Modified: 7/15/08
*********************************
FUNCTION htmldropdown
LPARAMETERS lcKeyValue,lcDisplay,lcFormVar,lcFirstItem,lcSelectedVal,llComplete
LOCAL lcStr
IF EMPTY(lcDisplay)
	lcDisplay=lcKeyValue
ENDIF
IF EMPTY(llComplete)
	llComplete=.F.
ENDIF
lcStr=[<option ]+IIF(EMPTY(lcSelectedVal),[selected],;
	[])+[ value="">]+lcFirstItem+[</option>]+crlf
SCAN
	lcStr=lcStr+ [<option ]+IIF(PADR(lcSelectedVal,LEN(&lcKeyValue),' ')=;
		&lcKeyValue,[selected],[])+[ value="]+&lcKeyValue+[">]+&lcDisplay+[</option>]+crlf
ENDSCAN
IF llComplete = .T.
	lcStr=[<select name=]+lcFormVar+[>]+lcStr+[</select>]
ENDIF
RETURN lcStr
ENDFUNC
*
************************************************************************
*
* AVFPhtml :: htmldecode
*
*********************************
FUNCTION htmldecode
PARAMETERS lcString
RETURN this.htmlencode(lcString,.t.)
ENDFUNC
*
*
************************************************************************
*
* AVFPhtml :: htmlencode
*
*********************************
FUNCTION htmlencode
PARAMETERS lcString,decode
LOCAL ARRAY laEncoding[4,2]
laEncoding[1,1] = "&"
laEncoding[1,2] = "&amp;"
laEncoding[2,1] = ["]
laEncoding[2,2] = "&quot;"
laEncoding[3,1] = "<"
laEncoding[3,2] = "&lt;"
laEncoding[4,1] = ">"
laEncoding[4,2] = "&gt;"

FOR lnIndex = 1 TO ALEN(laEncoding,1)
	IF decode
		lcString = STRTRAN(lcString,laEncoding[lnIndex,2],laEncoding[lnIndex,1])
	ELSE
		lcString = STRTRAN(lcString,laEncoding[lnIndex,1],laEncoding[lnIndex,2])
	ENDIF
NEXT
RETURN lcString
ENDFUNC
*
ENDDEFINE

*********************************************************************************
DEFINE CLASS AVFPhttpupload AS CUSTOM 
Note: http upload works with VFP7 SP1 and above, not with VFP6
*************************************************************
DIMENSION aForm[1,2]
DIMENSION aFiles[1,7]
lnMaxSize = 999999999
lnChunkSize = 65000   &&1024
lnCpuSleep= 0
lnfCount = 0
progressid = ""
lcGifFile = ""

************************************************************************
*
* AVFPhttpupload :: INIT

*********************************
FUNCTION INIT
THIS.aForm = .NULL.
THIS.aFiles = .NULL.
THIS.lnMaxSize = 999999999 
ENDFUNC
************************************************************************
*
* AVFPhttpupload :: ERROR

*********************************
FUNCTION ERROR(nError, cMethod, nLine)
IF FILE(oProp.TempFile)
    DeleteFile(oProp.TempFile)
ENDIF
ENDFUNC
************************************************************************
*
* AVFPhttpupload :: httpupload
******************************
FUNCTION httpupload
LPARAMETER lcFilePath
LOCAL lcInput,lnAt,lnHeaderLen,lcHeaderID,lsRemain,lnNameBeginAt,;
	lnNameLength,lcFileName,lcOrig,lcVarName,lvVarValue,lcOutput,lcExtension,;
	lnStartFilePos,lnEndFilePos, N, bytesread, laBlob1, READSIZE, lncount, lncount2,;
	lnacount,starting_value,ending_value, lcTempName, llProgress
SYS(3050, 1, VAL(SYS (3050, 1, 0)) / 3)
lnAt=0
lnHeaderLen=0
lcHeaderID=""
lsRemain=""
lnNameBeginAt=0
lnNameLength=0
lcFileName=""
lcOrig=""
lcVarName=""
lvVarValue=""
lcOutput=""
lcExtension=""
lnStartFilePos=0
lnEndFilePos=0
lcInput=''
N=oRequest.oRequest.TotalBytes
bytesread=0
lcTempName = ""
llProgress= .F.
THIS.progressid = oRequest.QueryString("progressid")
IF ISNULL(THIS.progressid) .OR. EMPTY(THIS.progressid)
	lcTempName =SUBSTR(SYS(2015), 3, 10)
	llProgress = .F.
ELSE

	lcTempName = THIS.progressid
	llProgress = .T.
ENDIF
IF ! DIRECTORY(oProp.AppStartPath+'temp')
	MD oProp.AppStartPath+'temp'
ENDIF
oProp.TempFile=oProp.AppStartPath+'temp\'+lcTempName+".tmp"
IF VERSION(5) < 7    
	DIME laBlob1[THIS.lnchunksize] && load into array
ENDIF
DECLARE Sleep IN WIN32API INTEGER
DECLARE DeleteFile IN WIN32API STRING
m.starting_value=SECONDS()
FOR lncount = 1 TO (N/THIS.lnchunksize)
	READSIZE=THIS.lnchunksize
	laBlob1=oRequest.BinaryRead(READSIZE)
	IF llProgress
		THIS.updateprogress(oProp.TempName,READSIZE,N)
	ENDIF
	IF VERSION(5) < 7 
		FOR lncount2 = 1 TO THIS.lnchunksize
			lcInput=m.lcInput+ CHR(laBlob1(lncount2)) &&slows things down!!
			Sleep(THIS.lnCpuSleep)
		ENDFOR
		STRTOFILE(lcInput,oProp.TempFile,.T.)
		lcInput=""
		RELEASE laBlob1
		DIME laBlob1[THIS.lnchunksize]
	ELSE
		Sleep(THIS.lnCpuSleep)  && allow processor to do other things
		STRTOFILE(laBlob1,oProp.TempFile,.T.)
		lablob1=""
	ENDIF
	bytesread=bytesread+READSIZE
NEXT
READSIZE=N - bytesread
IF READSIZE <> 0
	laBlob1=oRequest.BinaryRead(READSIZE)
	IF llProgress
		THIS.updateprogress(lcTempName,READSIZE,N)
	ENDIF
	IF VERSION(5) < 7 
		FOR lncount = 1 TO READSIZE
			lcInput=m.lcInput+ CHR(laBlob1(lncount)) &&slows things down!!
			Sleep(THIS.lnCpuSleep)
		ENDFOR
		STRTOFILE(lcInput,oProp.TempFile,.T.)
		lcInput=""
		RELEASE laBlob1
	ELSE
		STRTOFILE(laBlob1,oProp.TempFile,.T.)
		laBlob1=""
	ENDIF
	bytesread=bytesread+READSIZE
ENDIF
IF FILE(oProp.TempFile)
	lcInput=FILETOSTR(oProp.TempFile)
	DeleteFile(oProp.TempFile)
ENDIF
lnaCount=0
THIS.lnfCount=0
lcHeaderID=LEFT(lcInput,AT(crlf,lcInput)-1)
lnHeaderLen=LEN(lcHeaderID)
lsRemain=lcInput
lnAt=AT(lcHeaderID,lcInput)
DO WHILE lnAt > 0
	lsRemain=SUBS(lsRemain,lnAt+lnHeaderLen)
	lnNameBeginAt=AT("name=",lsRemain)+5
	lnNameLength=AT(["],SUBS(lsRemain,lnNameBeginAt),2)-2
	lcVarName=SUBS(lsRemain,lnNameBeginAt+1,lnNameLength)
	IF SUBS(lsRemain,lnNamelength+lnNamebeginat+2,1)=';'
		lsRemain=SUBS(lsRemain,lnNameBeginAt+lnNameLength)
		lnFileNameAt=AT("filename=",lsRemain)+9
		lnNameLength=AT(["],SUBS(lsRemain,lnFileNameAt),2)-2
		lcFileName=SUBS(lsRemain,lnFileNameAt+1,lnNameLength)
		lcOrig=lcFileName
		IF EMPTY(RAT("\",lcFileName))
			IF EMPTY(RAT(":",lcFileName))
			ELSE
				lcFileName=SUBS(lcFileName,RAT(":",lcFileName)+1)
			ENDIF
		ELSE
			lcFileName=SUBS(lcFileName,RAT("\",lcFileName)+1)
		ENDIF
		lcExtension=SUBS(lcFileName,RAT(".",lcFileName))
		lnStartPos=AT("Content-Type:",lsRemain)
		lnStartPos=IIF(EMPTY(lnStartPos),lnFileNameAt,lnStartPos)
		lsRemain=SUBS(lsRemain,lnStartPos)
		lnStartFilePos=AT(CHR(13),lsRemain)+4
		lnEndFilePos=AT(lcHeaderID,lsRemain)-2
		lcOutput=SUBS(lsRemain,lnStartFilePos,lnEndFilePos-lnStartFilePos)
		IF ! EMPTY(lcFileName) .AND. ! EMPTY(lcOutPut) .AND. LEN(lcOutPut) < THIS.lnMaxSize
			STRTOFILE(lcOutput,lcFilePath+'\'+lcFileName)
			THIS.lnfCount= THIS.lnfCount+1
			m.ending_value=SECONDS()
			DIMENSION THIS.aFiles[THIS.lnfCount,7]
			THIS.aFiles[THIS.lnfCount,1] = lcVarName
			THIS.aFiles[THIS.lnfCount,2] = lcFilePath+'\'+lcFileName
			THIS.aFiles[THIS.lnfCount,3] = ALLTRIM(STR(LEN(lcOutPut)))
			THIS.aFiles[THIS.lnfCount,4] = lcOrig
			THIS.aFiles[THIS.lnfCount,5] = lcExtension
			THIS.aFiles[THIS.lnfCount,6] = lcFilePath+'\'+JUSTSTEM(lcFileName)
			THIS.aFiles[THIS.lnfCount,7] = STR(m.ending_value-m.starting_value)
			RELEASE lcOutput
		ENDIF
	ELSE
		IF !EMPTY(lcVarName)
			lnaCount= lnaCount+1
			lvVarValue=oHTML.GetString(lsRemain,crlf+crlf,crlf)
			*IIF(VERSION(5)>=700,STREXTRACT(lsRemain,crlf+crlf,crlf),;
				*oHTML.GetString(lsRemain,crlf+crlf,crlf))
			DIMENSION THIS.aForm[lnaCount,2]
			THIS.aForm[lnaCount,1] = lcVarName
			THIS.aForm[lnaCount,2] = lvVarValue
		ENDIF
	ENDIF
	lnAt=AT(lcHeaderID,lsRemain)
ENDDO

RELEASE lcInput
RETURN THIS.lnfCount

ENDFUNC
************************************************************************

* AVFPhttpupload :: getmform

*********************************
FUNCTION getmform
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1
DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aForm,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aForm,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aForm[ASUBSCRIPT(this.aForm,nLocation+1,1),ASUBSCRIPT(this.aForm,nLocation+1,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getfilename

*********************************
FUNCTION getfilename
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1


DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+1,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getfilesize

*********************************
FUNCTION getfilesize
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1


DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+2,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getfileorig

*********************************
FUNCTION getfileorig
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1


DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+3,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getfileex

*********************************
FUNCTION getfileex
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1


DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+4,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getfilestem

*********************************
FUNCTION getfilestem
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1


DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+5,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: getelapsedtime

*********************************
FUNCTION getelapsedtime
LPARAMETER cVar
LOCAL nLocation, llDone, nFrom
llDone=.F.
nFrom=1

DO WHILE llDone=.F.
	nLocation=ASCAN(THIS.aFiles,cVar,nFrom)
	IF nLocation=0
		llDone=.T.
	ELSE
		IF ASUBSCRIPT(THIS.aFiles,nLocation,2) = 1
			llDone=.T.
		ELSE
			nFrom=nLocation+1
		ENDIF
	ENDIF
ENDDO
IF nLocation > 0
	RETURN THIS.aFiles[ASUBSCRIPT(this.aFiles,nLocation+1,1),ASUBSCRIPT(this.aFiles,nLocation+6,2)]
ELSE
	RETURN ""
ENDIF

ENDFUNC
************************************************************************

* AVFPhttpupload :: setmaxsize

*********************************
FUNCTION setmaxsize
LPARAMETER MaxSize

THIS.lnMaxSize=MaxSize

ENDFUNC

*

************************************************************************

* AVFPhttpupload :: setchunksize

*********************************
FUNCTION setchunksize
LPARAMETER ChunkSize
IF ChunkSize <1024 .OR. ChunkSize>65000
	ChunkSize = 65000
ENDIF
THIS.lnchunksize=ChunkSize

ENDFUNC

*
************************************************************************

* AVFPhttpupload :: setcpusleep

*********************************
FUNCTION setcpusleep
LPARAMETER CpuSleep

THIS.lnCpuSleep=CpuSleep

ENDFUNC

*
************************************************************************

* AVFPhttpupload :: updateprogress

*********************************
FUNCTION updateprogress
LPARAMETER lcPID,lnTrans, lnTBytes

IF ! USED('progress')
	SELECT 0
	IF ! FILE(oProp.AppStartPath+'\temp\progress.DBF')
		CREATE TABLE (oProp.cAppStartPath+'\temp\progress') ;
			(PID c(10), TRANS N(12), tbytes N(12) , percent N(12))
		INDEX ON PID TAG PID

	ENDIF
	USE (oProp.cAppStartPath+'\temp\progress') SHARED
ENDIF
SELECT progress
SET ORDER TO PID
SEEK lcPID
IF FOUND()
	REPL TRANS WITH lnTrans+progress.TRANS
	REPL percent WITH (progress.TRANS/progress.tbytes)*100
ELSE
	APPEND BLANK
	REPL PID WITH lcPID
	REPL TRANS WITH lnTrans
	REPL tbytes WITH lnTbytes
	REPL percent WITH (lnTrans/lnTBytes)*100
ENDIF


ENDFUNC

*
************************************************************************

* AVFPhttpupload :: gettransfered

*********************************
FUNCTION gettransfered
LPARAMETER cVar
IF ! USED('progress')
	USE (oProp.cAppStartPath+'\temp\progress') IN 0 SHARED
ENDIF

SELECT progress
SET ORDER TO PID
SEEK cVar
IF FOUND()
	RETURN progress.TRANS
ELSE
	RETURN 0
ENDIF
ENDFUNC
*
************************************************************************

* AVFPhttpupload :: gettotal

*********************************
FUNCTION gettotal
LPARAMETER cVar
IF ! USED('progress')
	USE (oProp.cAppStartPath+'\temp\progress') IN 0 SHARED
ENDIF

SELECT progress
SET ORDER TO PID
SEEK cVar
IF FOUND()
	RETURN progress.tbytes
ELSE
	RETURN 0
ENDIF

ENDFUNC
*
************************************************************************

* AVFPhttpupload :: getpercent

*********************************
FUNCTION getpercent
LPARAMETER cVar
IF ! USED('progress')
	USE (oProp.cAppStartPath+'\temp\progress') IN 0 SHARED
ENDIF

SELECT progress
SET ORDER TO PID
SEEK cVar
IF FOUND()
	RETURN progress.percent
ELSE
	RETURN 0
ENDIF
ENDFUNC
*
************************************************************************

* AVFPhttpupload :: getpid

*********************************
FUNCTION getpid
RETURN SUBSTR(SYS(2015), 3, 10)
ENDFUNC
*********************************
ENDDEFINE

********************************************************
DEFINE CLASS AVFPsql AS CUSTOM
* Note: SQL/MSDE (as well as networked VFP DBFs) will need to be set up in COM+
*     and impersonate a user with proper network rights
************************************************************************
* AVFPsql :: LOGIN
*********************************
PROCEDURE LOGIN
LPARAMETER lcLogIn
SQLSetProp(0,"DispLogin",3)
SQLSetProp(0,"DispWarnings",.F.)
THIS.cconnectstring = lcLogIn

THIS.nsqlhandle = SQLSTRINGCONNECT(&lcLogIn)
IF THIS.nsqlhandle < 1
	THIS.lerror = .T.
	THIS.GetErrors()
	RETURN .F.
ENDIF
RETURN .T.
ENDPROC
************************************************************************
* AVFPsql :: ERROR
*********************************
PROCEDURE ERROR
LPARAMETERS nError, cMethod, nLine
#DEFINE SQL_SERVER '[SQL Server]'
DO CASE
CASE nError = 1466 AND THIS.nerror <> 1466
	THIS.nerror = 1466
	THIS.cerrormsg = MESSAGE()
	IF THIS.LOGIN(THIS.cconnectstring)
		THIS.lerror = .F.
		THIS.nerror = 1466     
		THIS.cerrormsg = ''
		RETRY
	ENDIF
ENDCASE
lnCount = AERROR(laError)
IF lnCount > 0
	ACOPY(laError,THIS.aerrors)
	lcErrorMsg = laError[2]
	lnLoc = ATC(SQL_SERVER,lcErrorMsg)
	IF lnLoc > 0
		lcErrorMsg = SUBSTR(lcErrorMsg,lnLoc + LEN(SQL_SERVER))
	ENDIF
	THIS.cerrormsg = lcErrorMsg
	THIS.nerror = laError[1]
ELSE
	THIS.cerrormsg = MESSAGE()
	THIS.nerror = nError
ENDIF
ENDPROC
************************************************************************
* AVFPsql :: DESTROY
*********************************
PROCEDURE DESTROY
IF THIS.nsqlres = 0
	=SQLCancel(THIS.nsqlhandle)
ENDIF

IF THIS.nsqlhandle > 0
	=SQLDisconnect(THIS.nsqlhandle)
ENDIF
ENDPROC
************************************************************************
* AVFPsql :: Execute
*********************************
PROCEDURE execute
LPARAMETER lcSQL,lcCursorName
lcSQL=IIF(TYPE('lcSQL')='C',lcSQL,THIS.csql)
lcCursorName=IIF(TYPE("lcCursorName")="C",lcCursorName,THIS.csqlcursor)
IF THIS.nsqlres = 0
	THIS.SQLCancel()
ENDIF
THIS.nsqlres = SQLExec(THIS.nsqlhandle,lcSQL,lcCursorName)
IF THIS.nsqlres = -1
	THIS.lerror = .T.
	THIS.GetErrors()
	IF (THIS.nodbcerror=14) AND ;
		THIS.LOGIN(THIS.cconnectstring)
		THIS.nsqlres = SQLExec(THIS.nsqlhandle,lcSQL,lcCursorName)
		IF THIS.nsqlres = -1
			THIS.lerror = .T.
			THIS.GetErrors()
		ELSE
			THIS.GetErrors('START')
		ENDIF
	ENDIF
ELSE
	THIS.lerror = .F.
	IF SQLCOMMIT(THIS.nsqlhandle) = -1
		THIS.lerror = .T.
		THIS.GetErrors()
	ENDIF
ENDIF
RETURN THIS.nsqlres
ENDPROC
************************************************************************
* AVFPsql :: GetErrors
*********************************
PROCEDURE GetErrors
LPARAMETER lcSTART
LOCAL lnLoc, lcErrorMsg
lcSTART=IIF(TYPE('lcSTART')='C',UPPER(lcSTART),'')
IF lcSTART = 'START'
	THIS.lerror=.F.
	THIS.nerror = 0
	THIS.nodbcerror = 0
	THIS.cerrormsg = ''
	RETURN
ENDIF
lnCount = AERROR(laError)
IF lnCount > 0
	ACOPY(laError,THIS.aerrors)
	lcDetails=""
	FOR i=1 TO ALEN(laError,2)
		FOR x=1 TO ALEN(laError,1)
			lcDetails=lcDetails+TRANSFORM(laError(x,i))+crlf
		ENDFOR
	ENDFOR
	
	lcErrorMsg = laError[2]
	lnLoc = ATC(SQL_SERVER,lcErrorMsg)
	IF lnLoc > 0
		lcErrorMsg = SUBSTR(lcErrorMsg,lnLoc + LEN(SQL_SERVER))
	ENDIF
	THIS.nerror = laError[1]
	THIS.nodbcerror = laError[5]
	THIS.cerrormsg = lcErrorMsg + ' ['+LTRIM(STR(THIS.nerror))+':'+LTRIM(STR(THIS.nodbcerror))+']'
ENDIF
ENDPROC
************************************************************************
* AVFPsql :: SQLCancel
*********************************
PROCEDURE SQLCancel
RETURN SQLCancel(THIS.nsqlhandle)
ENDPROC
*********************************
cconnectstring = ''
nsqlhandle = 0
csql = ''
csqlcursor = 'tCursor'
cerrormsg = ''
nerror = 0
nsqlres = -5
lerror = .F.
nodbcerror = .F.
DIMENSION aerrors[1,1]

ENDDEFINE
*

*********************************************************************************
DEFINE CLASS AVFPcookie AS CUSTOM 
*************************************************************

* AVFPcookie :: value

*********************************

*** Function: get value from cookie

************************************************************************
FUNCTION value
LPARAMETERS lcName
RETURN oRequest.Cookies(lcName)
ENDFUNC
************************************************************************

* AVFPcookie :: WriteCookie

*********************************

*** Function: Create a cookie to store some info. 

************************************************************************
FUNCTION Write
LPARAMETERS lcCookie,lValue,lcDate
oResponse.Cookies(lcCookie)=lValue
oResponse.Cookies(lcCookie).Path ="/"
IF !EMPTY(lcDate)
	oResponse.Cookies(lcCookie).Expires =lcDate && if date, save to disk("January 1, 2035")
ENDIF
RETURN
ENDFUNC
************************************************************************

* AVFPcookie :: DeleteCookie

*********************************

*** Function: Delete specified cookie. 

************************************************************************
FUNCTION Delete
LPARAMETERS lcCookie
oResponse.cookies(lcCookie)='expired'
oResponse.cookies(lcCookie).PATH ="/"
oResponse.cookies(lcCookie).Expires = "July 31, 1995"
RETURN
ENDFUNC
************************************************************************
ENDDEFINE

*********************************************************************************
DEFINE CLASS AVFPsessiontable AS CUSTOM 
*************************************************************
*!*	cAppStartPath=""
*!*	cSessID=""
cSessPref = "Cookie"
nSessPersist = 120 &&5400  
cSessID = ""
*********************************
*  AVFPsessiontable::value

*********************************

*** Function: emulate asp session with table
******************************************************************************
FUNCTION value
LPARAMETERS lcName,lcValue
LOCAL lcCurAlias,lnParams
lnParams=PARAMETERS()
lcCurAlias=ALIAS()
IF THIS.cSessPref = "URL"
	THIS.cSessID=oRequest.querystring("sid") && use URL
ELSE
	THIS.cSessID=oCookie.Value("sid")  && default is cookie
ENDIF

IF lnParams=1 .AND. EMPTY(THIS.cSessID)
		RETURN .NULL.
ENDIF

IF .NOT. FILE('session.DBF')
		CREATE TABLE ('session') ;
			(id c(10),timestamp T,vars M)
		INDEX ON id	TAG id
		INDEX ON timestamp TAG timestamp
ENDIF
IF .NOT. USED('session')
	USE ('session') IN 0 SHARED
ENDIF
SELECT session
SET ORDER TO id
IF EMPTY(THIS.cSessID) 
	THIS.cSessID=SUBSTR(SYS(2015), 3, 10)
	IF THIS.cSessPref != "URL"
		oCookie.Write("sid",THIS.cSessID)  && default is cookie
	ENDIF
ELSE
	IF lnParams=1
	    SEEK THIS.cSessID
		IF FOUND()
			RESTORE FROM MEMO vars ADDITIVE
		ENDIF
		IF TYPE('m.ts__&lcName')='U'
			IF !EMPTY(lcCurAlias)
				SELECT &lcCurAlias
			ENDIF
			RETURN .NULL.
		ELSE
			IF !EMPTY(lcCurAlias)
				SELECT &lcCurAlias
			ENDIF
			
	    	RETURN m.ts__&lcName
	    ENDIF
	 ENDIF
ENDIF
* recycling records
SCAN FOR (DATETIME()-session.timestamp)>;
   IIF(EMPTY(THIS.nSessPersist),120,THIS.nSessPersist)  &&86400 1 day &&5400 && 90 minutes    
   DO WHILE .NOT. RLOCK()
   ENDDO
   REPL ID WITH ''
   REPL vars WITH ''
   UNLOCK
ENDSCAN 
SEEK THIS.cSessID
IF EOF()
  DO WHILE .NOT. FLOCK()
  ENDDO
  LOCATE FOR EMPTY(id)  
  IF EOF()
	APPEND BLANK
	REPL id	WITH THIS.cSessID
	REPL timestamp WITH DATETIME()
  ELSE
    REPL id	WITH THIS.cSessID
	REPL timestamp WITH DATETIME()
  ENDIF	
ELSE
	DO WHILE .NOT. RLOCK()
	ENDDO
ENDIF
SEEK THIS.cSessID
IF FOUND()
   IF !EMPTY(vars)   
	RESTORE FROM MEMO vars ADDITIVE
   ENDIF
ENDIF
ts__&lcName=IIF(TYPE('lcValue')='C',THIS.trimMemo(lcValue),lcValue)
SAVE TO MEMO vars ALL LIKE ts__*
UNLOCK
IF !EMPTY(lcCurAlias)
	SELECT &lcCurAlias
ENDIF
RETURN .NULL.
ENDFUNC
************************************************************************

* AVFPsessiontable :: setSessPersist

*********************************
FUNCTION setSessPersist
LPARAMETER lnlength
THIS.nSessPersist=lnlength
ENDFUNC
************************************************************************

* AVFPsessiontable :: setSessPref

*********************************
FUNCTION setSessPref
LPARAMETER lcPref
THIS.cSessPref=lcPref
ENDFUNC
************************************************************************
* AVFPsessiontable :: trimMemo
*********************************
FUNCTION trimMemo
PARAMETER mem
PRIVATE SpStr, Loc
* Leading
SpStr = ''
Loc = 1
DO WHILE .T.
	DO CASE
	CASE SUBSTR(mem,Loc,1) = ' '
		SpStr = SpStr+' '
	CASE SUBSTR(mem,Loc,1) = CHR(10) OR SUBSTR(mem,Loc,1) = CHR(13)
* first non-space is CR or LF
* get rid of everything that came before
		mem = RIGHT(mem,LEN(mem)-Loc)
		Loc = 0
	OTHERWISE
* keep the leading spaces
		EXIT
	ENDCASE
	Loc = Loc+1
ENDDO
* Trailing
DO WHILE .T.
	IF RIGHT(mem,1) = ' ' OR RIGHT(mem,1) = CHR(10) OR ;
		RIGHT(mem,1) = CHR(13)
		mem = LEFT(mem,LEN(mem)-1)
	ELSE
		EXIT
	ENDIF
ENDDO
RETURN mem
ENDFUNC
*
ENDDEFINE
*
*

*********************************************************************************
DEFINE CLASS AVFPRequest AS CUSTOM 
*************************************************************
oRequest = .NULL.
************************************************************************
* AVFPRequest :: Init
*********************************
*** 
FUNCTION Init

ENDFUNC
************************************************************************
* AVFPRequest :: Load
************************************************************************
FUNCTION Load
LPARAMETERS loRequest
THIS.oRequest=loRequest
ENDFUNC
************************************************************************
* AVFPRequest :: ServerVariables
*********************************
*** 
FUNCTION ServerVariables
LPARAMETERS lcName
RETURN THIS.oRequest.ServerVariables(lcName).Item

ENDFUNC
************************************************************************
* AVFPRequest :: form 
*********************************
*** 
FUNCTION form
LPARAMETERS lcName
RETURN THIS.oRequest.Form(lcName).Item
ENDFUNC
************************************************************************
* AVFPRequest :: querystring 
*********************************
*** 
FUNCTION querystring
LPARAMETERS lcName
RETURN THIS.oRequest.QueryString(lcName).Item
ENDFUNC
************************************************************************
* AVFPRequest :: BinaryRead 
*********************************
*** 
FUNCTION BinaryRead
LPARAMETERS lnBytes
RETURN THIS.oRequest.BinaryRead(lnBytes)
ENDFUNC
************************************************************************
* AVFPRequest :: Cookies 
*********************************
*** 
FUNCTION Cookies
LPARAMETERS lcName
RETURN THIS.oRequest.Cookies(lcName).Item
ENDFUNC
************************************************************************

* AVFPRequest :: dumpvars

*********************************
FUNCTION dumpvars
LOCAL lcSTR
*[<html>]+crlf
lcSTR=[<CENTER><table Border=1 BGCOLOR="#FFFFFF" STYLE="Font:normal normal 9pt 'verdana'" width="90%" cellspacing="5" cellpadding="5">]+crlf
lcSTR=lcSTR+[<TR BGCOLOR="#E5E5E5"><TH align="left" COLSPAN=2>Query String</th></TR>]+crlf;
	+[<TR><TD COLSPAN=2><b></b>]+crlf;
	+THIS.oRequest.ServerVariables("QUERY_STRING").ITEM() +[</TD></TR>]+crlf
FOR EACH lcFormVar IN THIS.oRequest.QueryString
	lcVar =   THIS.oRequest.querystring(lcFormVar).ITEM()
	lcSTR=lcSTR+[<TR><TD><b>]+ lcFormVar +[</b></TD><TD>]+crlf;
		+lcVar +[</TD></TR>]+crlf
NEXT
lcSTR=lcSTR+[<TR BGCOLOR="#E5E5E5"><TH align="left">Form Variable</TH><TH align="left">Value</TH></TR>]+crlf
FOR EACH lcFormVar IN THIS.oRequest.FORM
	lcVar =   THIS.oRequest.FORM(lcFormVar)
	lcSTR=lcSTR+[    <TR><TD><b> ]+ lcFormVar +[ </b></TD><TD> ]+crlf;
		+lcVar +[</TD></TR>]+crlf
NEXT
lcSTR=lcSTR+[<TR BGCOLOR="#E5E5E5"><TH align="left">Cookie</TH><TH align="left">Value</TH></TR>]+crlf
FOR EACH lcCookie IN THIS.oRequest.Cookies
	lcValue =   THIS.oRequest.cookies(lcCookie).ITEM()
	lcSTR=lcSTR+[<TR><TD><b>]+ lcCookie +[</b></TD><TD>]+crlf;
		+lcValue +[</TD></TR>]+crlf
NEXT
lcSTR=lcSTR+[<TR BGCOLOR="#E5E5E5"><TH align="left">ServerVariable</TH><TH align="left">Value</TH></TR>]+crlf
FOR EACH lcServerVar IN THIS.oRequest.ServerVariables
	lcServer =   THIS.oRequest.servervariables(lcServerVar).ITEM()
	IF LEN(lcServer) = 0
		lcServer = "n/a"
	ENDIF
	lcSTR=lcSTR+[<TR><TD><b>]+ lcServerVar +[</b></TD><TD>]+crlf;
		+lcServer +[</TD></TR>]+crlf
NEXT
lcSTR=lcSTR+[</table></CENTER>]+crlf
	*+[</html>]+crlf

RETURN lcSTR
ENDFUNC
ENDDEFINE
*
*

*********************************************************************************
DEFINE CLASS AVFPSession AS CUSTOM 
*************************************************************
oSession = .NULL.
************************************************************************
* AVFPSession :: Init
*********************************
*** 
FUNCTION Init
ENDFUNC
************************************************************************
* AVFPSession :: Load
************************************************************************
FUNCTION Load
LPARAMETERS loSession
THIS.oSession=loSession
ENDFUNC
************************************************************************
* AVFPSession :: Value
************************************************************************
FUNCTION Value
LPARAMETERS lcName,lcValue
IF PARAMETERS()=1
	RETURN THIS.oSession.Value(lcName)
ELSE
	THIS.oSession.Value(lcName) = lcValue
	RETURN
ENDIF
ENDFUNC
************************************************************************
* AVFPSession :: timeout 
*********************************
*** 
Procedure Timeout
LParameters lnValue
This.oSession.Timeout = lnValue
ENDPROC

ENDDEFINE

*********************************************************************************
Define Class AVFPproperties As custom
*************************************************************
AppStartPath=""
ScriptPath=""
AppName=""
HtmlPath=""
DataPath=""
Action=""
SessID=""
TempFile=""
Ext=""
RunningPrg=""
************************************************************************
* AVFPproperties :: AppStartPath_Assign
************************************************************************
Procedure AppStartPath_Assign()
LParameters cPath
This.AppStartPath = cPath
ENDPROC
************************************************************************
* AVFPproperties :: AppStartPath_Access
************************************************************************
Procedure AppStartPath_Access
Return This.AppStartPath
EndProc
************************************************************************
* AVFPproperties :: ScriptPath_Assign
************************************************************************
Procedure ScriptPath_Assign()
LParameters cPath
This.ScriptPath = cPath
ENDPROC
************************************************************************
* AVFPproperties :: ScriptPath_Access
************************************************************************
Procedure ScriptPath_Access
Return This.ScriptPath
ENDPROC
************************************************************************
* AVFPproperties :: AppName_Assign
************************************************************************
Procedure AppName_Assign()
LParameters cName
This.AppName = cName
ENDPROC
************************************************************************
* AVFPproperties :: AppName_Access
************************************************************************
Procedure AppName_Access
Return This.AppName
ENDPROC
************************************************************************
* AVFPproperties :: HtmlPath_Assign
************************************************************************
Procedure HtmlPath_Assign()
LParameters cPath
This.HtmlPath = cPath
ENDPROC
************************************************************************
* AVFPproperties :: HtmlPath_Access
************************************************************************
Procedure HtmlPath_Access
Return This.HtmlPath
ENDPROC
************************************************************************
* AVFPproperties :: Action_Assign
************************************************************************
Procedure Action_Assign()
LParameters cAction
This.Action = cAction
ENDPROC
************************************************************************
* AVFPproperties :: Action_Access
************************************************************************
Procedure Action_Access
Return This.Action
ENDPROC
************************************************************************
* AVFPproperties :: SessID_Assign
************************************************************************
Procedure SessID_Assign()
LParameters cSess
This.SessID = cSess
ENDPROC
************************************************************************
* AVFPproperties :: SessID_Access
************************************************************************
Procedure SessID_Access
Return This.SessID
ENDPROC
************************************************************************
* AVFPproperties :: TempFile_Assign
************************************************************************
Procedure TempFile_Assign()
LParameters cFile
This.TempFile = cFile
ENDPROC
************************************************************************
* AVFPproperties :: TempFile_Access
************************************************************************
Procedure TempFile_Access
Return This.TempFile
ENDPROC
************************************************************************
* AVFPproperties :: DataPath_Assign
************************************************************************
Procedure DataPath_Assign()
LParameters cPath
This.DataPath = cPath
ENDPROC
************************************************************************
* AVFPproperties :: DataPath_Access
************************************************************************
Procedure DataPath_Access
Return This.DataPath
ENDPROC
************************************************************************
* AVFPproperties :: cEval_Assign
************************************************************************
Procedure Ext_Assign()
LParameters cLine
This.Ext = cLine
ENDPROC
************************************************************************
* AVFPproperties :: cEval_Access
************************************************************************
Procedure Ext_Access
Return This.Ext
ENDPROC
************************************************************************
* AVFPproperties :: RunningPrg_Assign
************************************************************************
Procedure RunningPrg_Assign()
LParameters cPrg
This.RunningPrg = cPrg
ENDPROC
************************************************************************
* AVFPproperties :: RunningPrg_Access
************************************************************************
Procedure RunningPrg_Access
Return This.RunningPrg
ENDPROC
*
ENDDEFINE

*************************************************************
DEFINE CLASS AVFPGraph AS Relation
*************************************************************
*: Adapted from wwWebGraphs by Rick Strahl
*************************************************************
*** Graph type:  0 -  Bar   
***              3 -  Block (horizontal bar)  
***              6 -  Line
***              7 -  Line Dots
***              18 - Pie  
***              19 - Exploded Pie
nGraphType = 0
cCaption  = ""

*** Office Web Components Chart object
oOWC = .NULL.

cOWCProgId = "OWC11.ChartSpace"  && Latest version
*cOWCProgId = "OWC.Chart"    && Office 2000/Office XP installs

*** Physical disk Path  the image is written to
cPhysicalPath = ""

*** Web Path that is mapped to the physical path
*** used to create an <img> tag for the graph
cLogicalPath = ""

*** Timeout for graphs before they are deleted
*** Images are checked for timeouts and deleted 
*** whenever a new image is created
nImageTimeout = 300  && 5 minutes - 300 seconds 

*** Resulting filename of the image - filename only
cImageName = ""

nImageWidth = 750
nImageHeight = 480

*** Default Colors used
cBackcolor = "lightyellow"
cSeries1Color = "darkred"
cSeries2Color = "darkblue"
cSeries3Color = "darkgreen"
cSeries4Color = "orange"
cSeries5Color = "purple"
cSeries6Color = "pink"

nShowLegend = 1   && Show with column name 2 - ???


FUNCTION Init
*!*	LPARAMETERS llNoExistCheck

*!*	IF !llNoExistCheck &&AND !ISCOMOBJECT(THIS.cOWCProgid)
*!*	   RETURN .F.
*!*	ENDIF
EXTERNAL ARRAY LAGRAPHS,LALEGEND

THIS.oOWC = CREATEOBJECT(THIS.cOWCProgId)
IF VARTYPE(THIS.oOWC) # "O"
   RETURN .F.
ENDIF

RETURN

************************************************************************
* AVFPgraph :: GraphSetup
****************************************
***  Function: Configures the basic Graph Layout
***    Assume: Basic layout used for all graph renderers
************************************************************************
FUNCTION GraphSetup()
LOCAL loGraph

loGraph = THIS.oOWC.Charts.Add()
loGraph.Type = THIS.nGraphType

IF THIS.nShowLegend > 0
  loGraph.HasLegend = .T.
ENDIF

loGraph.PlotArea.Interior.Color = THIS.cBackColor

IF !EMPTY(THIS.cCaption)
   THIS.oOWC.HasChartSpaceTitle = .T.
   THIS.oOWC.ChartSpaceTitle.Caption = THIS.cCaption
   THIS.oOWC.ChartSpaceTitle.Font.Bold = .T.
ENDIF

RETURN loGraph
ENDFUNC
*  AVFPgraph :: GraphSetup

************************************************************************
* AVFPgraph :: ShowGraphFromCursor
****************************************
***  Function: Graphs a chart from a cursor.
***    Assume: Col 1   - Label column
***            Col 2-n - Data columns
***    Return: nothing
************************************************************************
FUNCTION ShowGraphFromCursor()
LOCAL x, y, loGraph, lnFields, lnRows, lnSeries, lcElement, lcArray, oConst

lnFields = AFIELDS(laFields,ALIAS())
lnRows = RECCOUNT()
lnSeries = lnFields - 1

*** Create labal array from the first column of the table
DIMENSION laLabels[lnRows]

*** Create the value arrays and presize them each
FOR x = 1 TO lnSeries
  lcArray = "DIMENSION laValues" + TRANSFORM(x) + "[" + TRANSFORM(lnRows) + "]"
  &lcArray
ENDFOR

x = 0
SCAN
   x = x + 1
   laLabels[x] = TRIM(EVALUATE(FIELD(1)))  && First field value
   
   FOR y = 1 TO lnSeries 
      lcElement = "laValues" + TRANSFORM(y)+ "[" + TRANSFORM(x) + "]"
      &lcElement = EVALUATE(FIELD(y+1))
   ENDFOR  
ENDSCAN

oConst = THIS.oOWC.Constants
loGraph = THIS.GraphSetup()

FOR x = 1 TO lnSeries
  loGraph.SeriesCollection.Add()
  loSeries = loGraph.SeriesCollection(x-1)
  
  IF !INLIST(THIS.nGraphType,18,19,58,59) 
     *** Pies can't use interior colors well
     loSeries.Interior.Color = EVALUATE("THIS.cSeries" + TRANSFORM(x) +"Color" )
  ELSE
     IF x > 1
        EXIT   && PIes can't render more than a single series
     ENDIF
  ENDIF
  
  lcArray = "@laValues" + TRANSFORM(x) 
  loSeries.Caption = PROPER(STRTRAN(TRIM(laFields[x+1,1]),"_"," "))
  loSeries.SetData(oConst.chDimCategories, oConst.chDataLiteral, @laLabels)
  loSeries.SetData(oConst.chDimValues, oConst.chDataLiteral, &lcArray)
ENDFOR

ENDFUNC
*  AVFPgraph :: ShowGraphFromCursor

************************************************************************
* AVFPgraph :: ShowGraphFromArray
****************************************
***  Function: Graphs a chart from a cursor.
***    Assume: Col 1    - Label array
***            Col 2    - Data array
***            lcSeries - Legend Text for Series
***            to n columns
***    Return: nothing
************************************************************************
FUNCTION ShowGraphFromArray(laLabels,laSeries1,lcSeries1,laSeries2,lcSeries2,laSeries3,lcSeries3,;
                            laSeries4,lcSeries4,laSeries5,lcSeries5,laSeries6,lcSeries6)
LOCAL x, loGraph, lnSeries, lcArray, lcLabel

lnSeries = (PCOUNT() - 1) / 2

loGraph = THIS.GraphSetup()

oConst = THIS.oOWC.Constants

FOR x = 1 TO lnSeries
  loGraph.SeriesCollection.Add()
  
  IF !INLIST(THIS.nGraphType,18,19)  && Pies can't use interior colors well
     loGraph.SeriesCollection(x-1).Interior.Color = EVALUATE("THIS.cSeries" + TRANSFORM(x) +"Color" )
  ENDIF
  
  lcArray = "@laSeries" + TRANSFORM(x)
  lcLabel = EVALUATE("lcSeries" + TRANSFORM(x)) 
  loGraph.SeriesCollection(x-1).Caption = lcLabel
  loGraph.SeriesCollection(x-1).SetData(oConst.chDimCategories, oConst.chDataLiteral, @laLabels)
  loGraph.SeriesCollection(x-1).SetData(oConst.chDimValues, oConst.chDataLiteral, &lcArray)
ENDFOR

ENDFUNC
*  AVFPgraph :: ShowGraphFromArray


************************************************************************
* AVFPgraph :: ShowGraphFromMultiDimensionalArray
*****************************************************
***  Function: Graphs a chart from a cursor.
***    Assume: Col 1   - Label col
***            Col 2-n - Data cols
***    Return: nothing
************************************************************************
FUNCTION ShowGraphFromMultiDimensionalArray(laGraphs,laLegend)
LOCAL loGraph, lnSeries, lnRows, llLegend, x, y

lnSeries = ALEN(laGraphs,2) - 1
lnRows = ALEN(laGraphs,1)

loGraph = THIS.GraphSetup()

*** Check if a legend array was passed
IF TYPE([ALEN(laLegend)]) = "N"
  llLegend = .T.
ELSE
  llLegend = .F.  
  loGraph.HasLegend = .F.
ENDIF  

oConst = THIS.oOWC.Constants

DIMENSION laLabels[lnRows]
FOR y = 1 TO lnRows
   laLabels[y] = laGraphs[y,1]
ENDFOR

FOR x = 2 TO lnSeries + 1
  loGraph.SeriesCollection.Add()
  
  IF !INLIST(THIS.nGraphType,18,19)  && Pies can't use interior colors well
     loGraph.SeriesCollection(x-2).Interior.Color = EVALUATE("THIS.cSeries" + TRANSFORM(x) +"Color" )
  ENDIF
  
  *** Copy the array column into a 1D Array
  DIMENSION laArray[lnRows]
  FOR y = 1 TO lnRows
     laArray[y] = laGraphs[y,x]
  ENDFOR
  
  IF llLegend
     loGraph.SeriesCollection(x-2).Caption = laLegend[x-1]
  ENDIF
  loGraph.SeriesCollection(x-2).SetData(oConst.chDimCategories, oConst.chDataLiteral, @laLabels)
  loGraph.SeriesCollection(x-2).SetData(oConst.chDimValues, oConst.chDataLiteral, @laArray)
ENDFOR

ENDFUNC
*  AVFPgraph :: ShowGraphFromArray


************************************************************************
* AVFPgraph :: GetOutput
****************************************
***  Function: This method actually creates the graph on disk
***            and 
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetOutput()

*** Post Config Event for Graphic Layout
THIS.SetGraphicsOptions(THIS.oOWC.Charts(0))

*** Delete graph files that have timed out
*oUtil.cDeleteDir=SUBSTR(oProp.AppStartPath,1,ATC(oProp.AppName,oProp.AppStartPath)+8)+'r\temp\'
*otil.cDeleteExt='gif'
*oUtil.nDeleteAge=1200 && files older than 20 Minutes, erase 
*oUtil.DeleteFiles()
*(THIS.cPhysicalPath + "IMG*.gif",THIS.nImageTimeOut)

THIS.cImageName = "IMG" + SYS(2015) + ;
                  TRANSFORM(Application.ProcessId) + ".gif"

*** Now actually create the image on disk                  
THIS.oOWC.ExportPicture(THIS.cPhysicalPath + THIS.cImageName, "gif",;
                        THIS.nImageWidth, THIS.nImageHeight)

RETURN [<img src="] + THIS.cLogicalPath + THIS.cImageName + [">]
ENDFUNC
*  AVFPgraph :: GetOutput

************************************************************************
* AVFPgraph :: Clear
****************************************
***  Function: Clears the graphics area
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION Clear()

THIS.oOWC.Clear()

ENDFUNC
*  AVFPgraph :: Clear

************************************************************************
* AVFPgraph :: ShowGraphInForm
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION ShowGraphInForm(lcFormCaption)
LOCAL loForm as Form
loForm = CREATEOBJECT("GraphForm")

loForm.AddObject("oPicture","image")
loForm.Height = THIS.nImageHeight
loForm.Width = THIS.nImageWidth
loForm.AutoCenter = .T.

IF !EMPTY(lcFormCaption)
  loForm.Caption = lcFormCaption
ENDIF  

loForm.oPicture.Left = 0
loForm.oPicture.Top = 0
loForm.oPicture.Height = THIS.nImageHeight
loForm.oPicture.Width = THIS.nImageWidth

loForm.oPicture.Picture = THIS.cPhysicalPath + THIS.cImageName
loForm.oPicture.visible = .T.

loForm.Show()

RETURN loForm

ENDFUNC
*  AVFPgraph :: ShowGraphInForm

************************************************************************
* AVFPgraph :: GetGraphTypes
****************************************
***  Function:
***    Assume:
***      Pass:
***    Return:
************************************************************************
FUNCTION GetGraphTypes(lcCursorName)

IF EMPTY(lcCursorName)
  lcCursorName = "avfpWebGraphTypes"
ENDIF  

CREATE CURSOR  (lcCursorName) (name c(80),id i)

*** Using this format since this comes from a GetConstants 
*** export to header file which can be updated as needed
*** by re-running the constant export
TEXT TO lcVar NOSHOW
#define chChartTypeColumnClustered                        0
#define chChartTypeColumnStacked                          1
#define chChartTypeColumnStacked100                       2
#define chChartTypeBarClustered                           3
#define chChartTypeBarStacked                             4
#define chChartTypeBarStacked100                          5
#define chChartTypeLine                                   6
#define chChartTypeLineStacked                            8
#define chChartTypeLineStacked100                         10
#define chChartTypeLineMarkers                            7
#define chChartTypeLineStackedMarkers                     9
#define chChartTypeLineStacked100Markers                  11
#define chChartTypeSmoothLine                             12
#define chChartTypeSmoothLineStacked                      14
#define chChartTypeSmoothLineStacked100                   16
#define chChartTypeSmoothLineMarkers                      13
#define chChartTypeSmoothLineStackedMarkers               15
#define chChartTypeSmoothLineStacked100Markers            17
#define chChartTypePie                                    18
#define chChartTypePieExploded                            19
#define chChartTypePieStacked                             20
#define chChartTypeScatterMarkers                         21
#define chChartTypeScatterLine                            25
#define chChartTypeScatterLineMarkers                     24
#define chChartTypeScatterLineFilled                      26
#define chChartTypeScatterSmoothLine                      23
#define chChartTypeScatterSmoothLineMarkers               22
#define chChartTypeBubble                                 27
#define chChartTypeBubbleLine                             28
#define chChartTypeArea                                   29
#define chChartTypeAreaStacked                            30
#define chChartTypeAreaStacked100                         31
#define chChartTypeDoughnut                               32
#define chChartTypeDoughnutExploded                       33
#define chChartTypeRadarLine                              34
#define chChartTypeRadarLineMarkers                       35
#define chChartTypeRadarLineFilled                        36
#define chChartTypeRadarSmoothLine                        37
#define chChartTypeRadarSmoothLineMarkers                 38
#define chChartTypeStockHLC                               39
#define chChartTypeStockOHLC                              40
#define chChartTypePolarMarkers                           41
#define chChartTypePolarLine                              42
#define chChartTypePolarLineMarkers                       43
#define chChartTypePolarSmoothLine                        44
#define chChartTypePolarSmoothLineMarkers                 45
#define chChartTypeColumn3D                               46
#define chChartTypeColumnClustered3D                      47
#define chChartTypeColumnStacked3D                        48
#define chChartTypeColumnStacked1003D                     49
#define chChartTypeBar3D                                  50
#define chChartTypeBarClustered3D                         51
#define chChartTypeBarStacked3D                           52
#define chChartTypeBarStacked1003D                        53
#define chChartTypeLine3D                                 54
#define chChartTypeLineOverlapped3D                       55
#define chChartTypeLineStacked3D                          56
#define chChartTypeLineStacked1003D                       57
#define chChartTypePie3D                                  58
#define chChartTypePieExploded3D                          59
#define chChartTypeArea3D                                 60
#define chChartTypeAreaOverlapped3D                       61
#define chChartTypeAreaStacked3D                          62
#define chChartTypeAreaStacked1003D                       63
ENDTEXT

lnLines = ALINES(laLines,lcVar)

FOR x = 1 TO lnLines
   lcValue = GetString(laLines[x],"   ",CHR(13))
   INSERT INTO (lcCursorName) (name,id) VALUES ;
         ( GetString(laLines[x],"#define "," "), VAL(lcValue))
ENDFOR

REPLACE ALL name WITH STRTRAN(name,"chChartType","")

ENDFUNC
*  AVFPgraph :: GetGraphTypes



************************************************************************
* AVFPgraph :: SetGraphicsOptions
****************************************
***  Function: Override hook method that allows updating the look
***            of the graph.
***    Assume: virtual
***      Pass:
***    Return:
************************************************************************
FUNCTION SetGraphicsOptions(loChart)
RETURN
ENDFUNC
*  AVFPgraph :: SetGraphicsOptions

FUNCTION nGraphType_Assign(lvGraphType)

IF VARTYPE(lvGraphType) = "N"
   This.nGraphType = lvGraphType
   RETURN
ENDIF

IF VARTYPE(lvGraphType) = "C"
  lvGraphType = UPPER(lvGraphType)
  DO CASE
     CASE lvGraphType = "BAR3D"
        THIS.nGraphType = 50
     CASE lvGraphType = "BAR"
        THIS.nGraphType = 0
     CASE lvGraphType = "PIE3D"
        THIS.nGraphType = 58
     CASE lvGraphType = "PIE"
        THIS.nGraphType = 18
     CASE lvGraphType = "PIEEXPLODED3D"
        THIS.nGraphType = 59
     CASE lvGraphType = "PIEEXPLODED"
        THIS.nGraphType = 19
     CASE lvGraphType = "COLUMN3D"
        THIS.nGraphType = 51
     CASE lvGraphType = "COLUMN"
        THIS.nGraphType = 3
     CASE lvGraphType = "LINE3D"
        THIS.nGraphType = 54
     CASE lvGraphType = "LINEPOINTS"
        THIS.nGraphType = 7
     CASE lvGraphType = "LINE"
        THIS.nGraphType = 6
     
  ENDCASE
ENDIF


ENDDEFINE
*EOC AVFPgraph
 
*********************************************************************************
Define Class AVFPutilities As custom
*************************************************************

************************************************************************
* AVFPutilities :: DeleteFiles
****************************************
***  Function:  ERASE OLD Files
************************************************************************
FUNCTION DeleteFiles()
LPARAMETERS lcDeleteExt,lnDeleteAge,lcDeleteDir
*
lnNumber= ADIR(laFiles,lcDeleteDir+'*.'+lcDeleteExt)  && Create array
FOR i = 1 TO lnNumber  && Loop for number of files
IF DATETIME() - CTOT(DTOC(laFiles(i,3))+' '+laFiles(i,4)) > lnDeleteAge 
	ERASE (lcDeleteDir+laFiles(i,1))
ENDIF
ENDFOR

************************************************************************
* AVFPutilities :: urldecode
****************************************
***  Function:  URLDecodes a text string
***     Notes:  Replaces %hh tokens with ascii characters
***             Assuming fewer than 65000 '%'s
***     Input: tcInput - Text string to decode
***     Return: Decoded string
***     Author: Albert Ballinger
************************************************************************
FUNCTION urldecode()
Lparameter tcInput
Local lnPcntPos, lcOutput, lnPcntCount, i
tcInput = Strtran(tcInput, "+", " ")
Local Array laPercents[1]
lnPcntCount = Alines(laPercents, Strtran(tcInput, "%", Chr(13)))
lcOutput = laPercents[1]
For i = 2 To lnPcntCount
    lcOutput = m.lcOutput + Chr(Evaluate('0x' + Left(laPercents[i], 2))) + ;
        substr(laPercents[i], 3)
ENDFOR
DO CASE
CASE "</" $ lcOutput
CASE "SCRIPT" $ UPPER(lcOutput)
OTHERWISE
    Return lcOutput
ENDCASE
RETURN ""
ENDFUNC

************************************************************************
* AVFPutilities :: urlencode
****************************************
***  Function:  URLEncodes a text string
************************************************************************
FUNCTION urlencode()
PARAMETERS tcValue, llNoPlus
LOCAL lcResult, lcChar, lnSize, lnX
*** Do it in VFP Code
lcResult=""
 FOR lnX=1 to len(tcValue)
   lcChar = SUBSTR(tcValue,lnX,1)
   IF ATC(lcChar,"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") > 0
      lcResult=lcResult + lcChar
      LOOP
   ENDIF
   IF lcChar=" " AND !llNoPlus
      lcResult = lcResult + "+"
      LOOP
   ENDIF
   *** Convert others to Hex equivalents
   lcResult = lcResult + "%" + RIGHT(transform(ASC(lcChar),"@0"),2)
ENDFOR
 
RETURN lcResult
ENDFUNC
****************************************************************************
ENDDEFINE
****************************************************************************

****************************************************************************
*** Form class used to optionally display the graph
DEFINE CLASS GraphForm as Form
****************************************************************************
ShowWindow = 2 && Top Level Form
Caption = "Graph Results"
MaxButton = .F.
BorderStyle = 2
ENDDEFINE

*********************************************************************************
Define Class AVFPsmtp As custom
*************************************************************
cSMTPServer=''
*The IP Address or domain name of the SMTP server that is responsible for sending the message. The server must support 'pure' SMTP mail services and must not require a login prior to sending messages.
cSendername=''
*Display name of the person sending the message. This option is often overridden by the mail server with the actual user account, but some servers leave the name as set here.
cSenderEmail=''
*The sender's Email address.
cRecipient=''
*Either a single recipient email address or a list of comma delimited addresses. Limited to 1024 characters. To specify a name and email address use the following syntax (works for all Recipient and CC properties):
*Joe Doe <jdoe@yourserver.com>
cCCList=''
*Either a single CC recipient email address of a list of comma delimited addresses. Limited to 1024 characters.
cBCCList=''
*Blind CCs are sent messages but do not show on the recipient or CC listing of the receiver. Limited to 1024 characters.
cSubject=''
*The message subject or title.
cMessage=''
*The actual message body text. 
cAttachment=''
*The name of the file that is to be attached to the message. This property must be filled with a valid path pointing at the file to attach. Multiple files can be attached by separating each file with a comma.
cContentType=''
*Content type of the message allows you to send HTML messages for example. For HTML set the content type to text/html. Content types only work with text content types at this time since special encoding of the message text is not supported. All binary types should be sent as attachments.
cUsername=''
cPassword=''
*If you are sending mail to a mail server that requires username and password authentication, set these properties. Only unencrypted, plain text communication is supported using the AUTH LOGIN command. 
cImportance='normal' &&default = normal importance
*Importance levels
*0=low
*1=normal
*2=high
cErrorMsg=''
nErrorNum=0
nSendUsingPort= 2
nsmtpserverport= 25

*************************************************************
* AVFPsmtp :: CDOsend
*
********************************************
*** Function: Send SMTP email using CDO(2000/XP). 
************************************************************************
FUNCTION CDOsend
LOCAL llCDO
llCDO=.T.
DO CASE
* Tests for Windows 2000
	CASE "5.00" $ OS()
		llCDO = .T.
	CASE "NT 4.00" $ OS()  && must use CDONTS
		llCDO = .F.
* Tests for Windows 95
	CASE "4.00" $ OS()
		llCDO = .F.
* Tests for Windows 98
	CASE "4.10" $ OS()
		llCDO = .F.
* Tests for Windows Me
	CASE "4.90" $ OS()
		llCDO = .F.
	OTHERWISE  && XP?
		llCDO = .T.
ENDCASE
IF ! llCDO
	this.cErrorMsg='OS does not support CDO'
	RETURN .F.
ELSE
	LOCAL loMsg
	loMsg = NEWOBJECT("CDO.Message")

	WITH loMsg
		IF !EMPTY(this.cSendername)
	   		.FROM = this.cSendername
	   	ENDIF
	   	IF !EMPTY(this.cRecipient)
	   		.TO = this.cRecipient
	   	ENDIF
	   	IF !EMPTY(this.cBCCList)
	   		.BCC = this.cBCCList
	   	ENDIF
	   	IF !EMPTY(this.cSubject)
	   		.Subject = this.cSubject
	  	ENDIF
	   	IF !EMPTY(this.cMessage)
	   		.TextBody = this.cMessage
	   	ENDIF
	   .Configuration.FIELDS("http://schemas.microsoft.com/cdo/configuration/sendusing").VALUE = this.nSendUsingPort
	   .Configuration.FIELDS("http://schemas.microsoft.com/cdo/configuration/smtpserver").VALUE = this.csmtpserver
	   .Configuration.FIELDS("http://schemas.microsoft.com/cdo/configuration/smtpserverport").VALUE = this.nsmtpserverport
	   .Configuration.FIELDS.UPDATE
	   IF !EMPTY(this.cAttachment)
	  	  .AddAttachment(this.cAttachment)
       ENDIF 
	   .SEND
	ENDWITH
*!*		SET STEP ON 
*!*		this.nErrorNum = Err.Number
*!*	    this.cErrorMsg = Err.Description
*!*	    Err.Clear
	loMsg = NULL
	IF this.nErrorNum <> 0
		RETURN .F.
	ELSE
	    RETURN .T.
	ENDIF
ENDIF
*************************************************************
* AVFPsmtp :: CDONTSsend
*
********************************************
*** Function: Send SMTP email using CDONTS(NT). 
*** use CDONTS,IIS SMTP(must be started)
************************************************************************
LOCAL llCDONTS
#DEFINE CdoBodyFormatHTML  0 && Body property is HTML
#DEFINE CdoBodyFormatText  1 && Body property is plain text (default)
#DEFINE CdoMailFormatMime  0 && NewMail object is in MIME format
#DEFINE CdoMailFormatText  1 && NewMail object is plain text (default)
#DEFINE CdoLow     0         && Low importance
#DEFINE CdoNormal  1         && Normal importance (default)
#DEFINE CdoHigh    2         && High importance
#DEFINE CdoEncodingUUencode  0     && The attachment is to be in UUEncode format (default)  
#DEFINE CdoEncodingBase64    1     && The attachment is to be in base 64 format  
llCDONTS = .T.
DO CASE
* Tests for Windows 2000
	CASE "5.00" $ OS()
		llCDONTS = .T.
	CASE "NT 4.00" $ OS()  && must use CDONTS
		llCDONTS = .T.
* Tests for Windows 95
	CASE "4.00" $ OS()
		llCDONTS = .F.
* Tests for Windows 98
	CASE "4.10" $ OS()
		llCDONTS = .F.
* Tests for Windows Me
	CASE "4.90" $ OS()
		llCDONTS = .F.
	OTHERWISE  && XP?
		llCDONTS = .F.
ENDCASE
IF !llCDONTS
	this.cErrorMsg='OS does not support CDONTS'
	RETURN .F.
ELSE
	lcFileName=oRequest.FORM("txtMailAttachment")
	lcFileDesc=oRequest.FORM("txtMailFileDesc")
	oMail = CreateObject("CDONTS.Newmail")
	oMail.MailFormat = cdoMailFormatMIME
	IF !EMPTY(oRequest.FORM("txtMailFrom"))
		oMail.From = oRequest.FORM("txtMailFrom")
	ENDIF
	IF !EMPTY(oRequest.FORM("txtMailTo"))
		oMail.To = oRequest.FORM("txtMailTo")
	ENDIF
	IF !EMPTY(oRequest.FORM("txtMailBcc"))
		oMail.bcc = oRequest.FORM("txtMailBcc")
	ENDIF
	IF !EMPTY(oRequest.FORM("txtMailReplyTo"))
		oMail.Value("Reply-To") = oRequest.FORM("txtMailReplyTo")
	ENDIF
	IF !EMPTY(oRequest.FORM("D1"))
		oMail.Importance = VAL(oRequest.FORM("D1"))
	ENDIF
	IF !EMPTY(oRequest.FORM("txtMailSubject"))
		oMail.Subject = oRequest.FORM("txtMailSubject")
	ENDIF
	IF !EMPTY(oRequest.FORM("S1"))
		oMail.Body = oRequest.FORM("S1")
	ENDIF
	IF !EMPTY(lcFileName) .AND. FILE(lcFileName)
		oMail.AttachFile(lcFileName,lcFileDesc)
	ENDIF
	oMail.Send
ENDIF
RETURN .T.  && no error checking in cdonts

ENDFUNC
ENDDEFINE


DEFINE CLASS c1 as session olepublic
*!*	      DataSession = 2 && you should always run each object in their own datasession
*!*	      CallInfo = .NULL. && "magic" property
      proc MyDoCmd(cCmd as string,p2 as Variant,p3 as Variant,p4 as Variant,p5 as Variant) helpstring 'Execute a command'
                  
            &cCmd
            CLEAR PROGRAM 
			CLEAR RESOURCES

      proc MyEval(cExpr as string,p2 as Variant,p3 as Variant,p4 as Variant,p5 as Variant) helpstring 'Evaluate an expression'

            RETURN &cExpr
      FUNCTION ERROR(nError, cMethod, nLine)

            lcErrMsg = cMethod+'  err#='+STR(nError,5)+' line='+STR(nline,6)+;
	          ' '+MESSAGE()+_VFP.SERVERNAME
            STRTOFILE(lcErrMsg,oProp.AppStartPath+[temp\threaderror.txt])
	        COMreturnerror(lcErrMsg,_VFP.SERVERNAME)
      ENDFUNC
ENDDEFINE
