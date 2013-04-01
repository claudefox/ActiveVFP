#DEFINE crlf CHR(13)+CHR(10)
SET STEP ON
oProp=NEWOBJECT('AVFPproperties','activevfp.prg')  && create properties shared among classes
oHTML=NEWOBJECT('AVFPhtml','activevfp.prg')
oCookie=NEWOBJECT('AVFPcookie','activevfp.prg')
oProp.ScriptPath=oRequest.servervariables("SCRIPT_NAME")
oProp.SessID=NVL(oRequest.querystring("sid"),"")
oProp.AppStartPath=JUSTPATH(oRequest.ServerVariables("PATH_TRANSLATED"))+[\] &&JUSTPATH(APPLICATION.SERVERNAME)+"\"
oProp.AppName=STREXTRACT(oRequest.ServerVariables("PATH_INFO"),[/],[/])  &&blank if installed to root &&JUSTSTEM(APPLICATION.SERVERNAME)
oProp.RunningPrg=[main.prg]

lcScript=FILETOSTR(oProp.AppStartPath+'\prg\main.prg')
lcHTMLout= EXECSCRIPT(lcScript) &&main()  && run the application

RETURN lcHTMLout                              


*************************************************************
DEFINE CLASS server AS ActiveVFP OLEPUBLIC
*************************************************************
* server :: process
* Last Modified: 12/30/11
* ActiveVFP Version 5.61
*************************************************************
FUNCTION Process
LOCAL lcHTMLout,lcHTMLfile

* Base ActiveVFP Objects
oRequest = THIS.oRequest
oSession = THIS.oSession
oResponse = THIS.oResponse
oServer= this.oServer
oApplication=this.oApplication

oProp=NEWOBJECT('AVFPproperties')  && create properties shared among classes
oHTML=NEWOBJECT('AVFPhtml')
oCookie=NEWOBJECT('AVFPcookie')

* init properties
oProp.ScriptPath=oRequest.servervariables("SCRIPT_NAME")
oProp.SessID=NVL(oRequest.querystring("sid"),"")
oProp.AppStartPath=JUSTPATH(oRequest.ServerVariables("PATH_TRANSLATED"))+[\] &&JUSTPATH(APPLICATION.SERVERNAME)+"\"
oProp.AppName=STREXTRACT(oRequest.ServerVariables("PATH_INFO"),[/],[/])  &&blank if installed to root &&JUSTSTEM(APPLICATION.SERVERNAME)
oProp.RunningPrg=[main.prg]


**************************************************************
* REST-like APIs implementation
*
* Author: Victor Espina
* Date: April 2012
*
**************************************************************
* Create an instancia of RESTHelper class
LOCAL oRESTHelper, cRootPath
cRootPath = oRequest.serverVariables("APPL_PHYSICAL_PATH")
oRESTHelper = NEWOBJECT("restHelper","PRG\RESTHelper.PRG")
oRESTHelper.setFolder("ROOT", cRootPath)
oRESTHelper.setFolder("CONTROLLERS", ADDBS(cRootPath) + "prg/rest/controllers")
oRESTHelper.loadControllers()
 
* Use RESTHelper object to check if the request conforms a REST-like API call. If so, pass
* the request object to RESTHelper in order to let the right resource controller handle it.
IF oRESTHelper.isREST(oRequest)
 oProp.appStartPath = oRequest.serverVariables("APPL_PHYSICAL_PATH")
 RETURN oRESTHelper.handleRequest(oRequest, oProp)
ENDIF

* If the request was not a REST-like API call, check if MAIN.PRG exists in
* oProp.appStartPath. If it doesn't, we change oProp.appStartPath to point
* to app's root folder. This will avoid an innecessary "file not found" 
* error.
IF NOT FILE(oProp.appStartPath + "\PRG\MAIN.PRG")
 oProp.appStartPath = oRequest.serverVariables("APPL_PHYSICAL_PATH")
ENDIF

* If for some reason we can't find a valid processor, return
* an error message
IF NOT FILE(oProp.AppStartPath+'\prg\main.prg')
 RETURN "oRESTHelper: cannot find " + oProp.appStartPath+"\prg\main.prg"
ENDIF
**************************************************************

lcScript=FILETOSTR(oProp.AppStartPath+'\prg\main.prg')
lcHTMLout= EXECSCRIPT(lcScript) &&main()  && run the application

RETURN lcHTMLout

ENDFUNC

*************************************************************
* server :: debugger
* Last Modified: 12/30/11
* ActiveVFP Version 5.61
*************************************************************
FUNCTION Debugger()
LOCAL loFox, lcReturn
lcReturn = ""

loFox = GETOBJECT(,"visualfoxpro.application.9")

loFox.SetVar("oRequest",THIS.oRequest)
loFox.SetVar("oSession",THIS.oSession)
loFox.SetVar("oResponse",THIS.oResponse)
loFox.SetVar("oServer",THIS.oServer)
loFox.SetVar("oApplication",THIS.oApplication)

lcReturn = loFox.eval("proxystub()")


RELEASE loFox
RETURN lcReturn
ENDFUNC



ENDDEFINE




* VictorEspina Feb 6, 2012
* 
* VELoadPRG
* Dinamically loads a given PRG. This routines try to find the given
* PRG name in the request's PRG folder or the App's main PRG folder.
*
PROCEDURE VELoadPRG(pcName, poProp, poRequest)
 *
 LOCAL lcScript
 lcScript = ADDBS(poProp.AppStartPath) + "prg\" + pcName
 IF NOT FILE(lcScript)
  lcScript = ADDBS(poRequest.serverVariables("APPL_PHYSICAL_PATH")) + "prg\" + pcName
 ENDIF
 IF FILE(lcScript)
  lcScript = FILETOSTR(lcScript)
 ELSE
  lcScript = "RETURN 'File " + pcName + ".prg could not be found.'"
 ENDIF
 
 RETURN lcScript
 *
ENDPROC