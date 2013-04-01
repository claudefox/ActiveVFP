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


