* restHelper.prg
* Helper class that allow an easy implementation of REST-like calls
* within ActiveVFP
*
* Author : Victor Espina
* Version: 1.0
*
* VERSION HISTORY
*
* DATE		VERSION		DESCRIPTION
* =========	===========	================================
* Apr 2012	1.0			Initial version
* Aug 2012  1.1 CKF     Custom methods for MVC

* Test code
LOCAL oRESTHelper
oRESTHelper = CREATEOBJECT("restHElper")
oRESTHelper.setFolder("ROOT","c:\vfp5.61demo\")
oRESTHelper.setFolder("CONTROLLERS","c:\avfp5.61demo\prg\rest\controllers\")
oRESTHelper.loadControllers()
?oRESTHelper.handleRequest("/avfp5.61demo/users/")


RETURN


DEFINE CLASS restHelper AS Custom
 *
 Folders = NULL
 Controllers = NULL
 
 * Constructor
 *
 PROCEDURE Init
  *
  THIS.Folders = CREATEOBJECT("Dictionary")
    THIS.Folders.Add("ROOT","")
    THIS.Folders.Add("CONTROLLERS","")
    
  THIS.Controllers = CREATEOBJECT("Dictionary")
  * 
 ENDPROC
 
 
 * setFolder
 * Sets a folder location
 *
 PROCEDURE setFolder(pcFolderID, pcFolderPath)
  THIS.Folders.Values[pcFolderId] = ADDBS(CHRTRAN(pcFolderPath,"/","\"))
 ENDPROC
 
 
 * loadControllers
 * Look for controllers in the controller's folder
 * 
 PROCEDURE loadControllers
  *
  LOCAL ARRAY aControllers[1]
  LOCAL cPath,cFile,cController,nCount,i
  cPath = THIS.Folders.Values["CONTROLLERS"]
  nCount = ADIR(aControllers, cPath + "*.PRG")
  THIS.Controllers.Clear()
  FOR i = 1 TO nCount
   cFile = ADDBS(cPath) + aControllers[i,1]
   cController = LOWER(JUSTSTEM(cFile))
   THIS.Controllers.Add(cController, cFile)
  ENDFOR
  *
 ENDPROC


 * isREST
 * Check if the given request is a RESTful request. A request is considered
 * a REST-compliant request if:
 *
 * a) Does not contains a file extension
 * b) Does not contains a ? sign
 * c) The URL conforms to the form http://server/resource/
 *
 PROCEDURE isREST(poRequest)
  *
  LOCAL cUri
  IF VARTYPE(poRequest)<>"C"
   cUri = poRequest.serverVariables("SCRIPT_NAME")
  ELSE 
   cUri = poREquest
  ENDIF
    
  * Basic syntax check
  IF !EMPTY(JUSTEXT(cUri)) OR AT("?",cUri)<>0
   RETURN .F.
  ENDIF
  
  * Check if URI contains a reference to an available resource controller
  LOCAL lHasController,i
  lHasController = .F.
  FOR i = 1 TO THIS.Controllers.Count
   IF ATC("/" + THIS.Controllers.Keys[i], cUri) > 0  &&IF ATC("/" + THIS.Controllers.Keys[i] + "/", cUri) > 0
    lHasController = .T.
    EXIT
   ENDIF
  ENDFOR
  IF NOT lHasController
   RETURN .F.
  ENDIF
  
  * If we get through here, then the uri MAY be a REST api call
  RETURN .T.
  *
 ENDPROC
 
  
 * handleRequest
 * Handle a given HTTP request
 *
 PROCEDURE handleRequest(poRequest, poProps)
  *
  * Analyze the requested URL to obtain:
  * a) VERB
  * b) Resource controller
  * c) Resource's parameters
  *
  LOCAL cUri,cRESTUri,cVerb,cResource,cParams,i,j,cController,cBaseURL
  IF VARTYPE(poRequest)<>"C"
   cUri = poRequest.serverVariables("SCRIPT_NAME")
   cVerb = poRequest.serverVariables("REQUEST_METHOD")
  ELSE
   cUri = poRequest
   cVerb = "GET"
  ENDIF
  
  j = 0
  cResource = ""
  cController = ""
  FOR i = 1 TO THIS.Controllers.Count
   cResource = THIS.Controllers.Keys[i]
   cController = THIS.Controllers.Values[i]
   IF ATC(cResource, cUri) > 0
    j = ATC("/" + cResource, cUri)    &&ATC("/" + cResource + "/", cUri)
    IF j > 0
     EXIT
    ENDIF
   ENDIF
  ENDFOR
  IF EMPTY(cResource)   && There is no controller for the specified resource
   RETURN "restHelper: unhandled resource on REST call (" + cUri + ")"
  ENDIF 
  IF j = 0   && Invalid REST call
   RETURN "restHelper: unrecognized REST call (" + cUri + ")"
  ENDIF
  
  cBaseURL = LEFT(cUri, j - 1)
  cRESTUri = SUBSTR(cUri,j)  && /server/resource/params --> /resource/params
  cParams = SUBSTR(cRESTUri,LEN(cResource) + 3)

  * Convert the parameter list to a collection
  LOCAL ARRAY aParams[1]
  LOCAL nCount, cPAram
  LOCAL oParams AS Collection
  nCount = ALINES(aParams, STRT(cParams,"/",CHR(13)+CHR(10)))
  oParams = CREATEOBJECT("Collection")
  FOR EACH cParam IN aParams
   IF !EMPTY(cParam)
    oParams.Add(cParam)
   ENDIF
  ENDFOR
  
  
  * Use the verb + parameter count to deduce the actual action
  * to be invoked on the controller
  *
  LOCAL cAction
  cAction = ""
  DO CASE
     CASE cVerb == "GET" AND oParams.Count = 1 AND LOWER(oParams.Item[1]) == "info"
          cAction = "infoAction"
          
     CASE cVerb == "GET" AND oParams.Count = 1 AND LOWER(oParams.Item[1]) == "help"
          cAction = "helpAction"
  
     CASE cVerb == "GET" AND oParams.Count = 0
          cAction = "listAction"
          
     CASE cVerb == "GET" AND oParams.Count > 0 AND !(ISALPHA(oParams.Item[1]))
          cAction = "getAction"
          
     CASE cVerb == "POST" AND oParams.Count = 0
          cAction = "addAction"
          
     CASE cVerb == "POST" AND oParams.Count > 0
          cAction = "updAction"
          
     CASE cVerb == "DELETE" AND oParams.Count = 0
          cAction = "zapAction"
          
     CASE cVerb == "DELETE" AND oPArams.Count > 0
          cAction = "dropAction"
                    
     OTHERWISE
          cAction = oParams.Item[1]    &&"customAction"
  ENDCASE
  IF EMPTY(cAction)
   RETURN "restHelper: unrecognized REST call (" + cRESTUri + ")"
  ENDIF
  
  * Create a instance of the resource's controller
  LOCAL oDCH,oController,cHTMLOut
  IF NOT FILE(cController)
   RETURN "restHelper: controller file <b>" + cController + "</b> could not be found"
  ENDIF
  
  oDCH = NEWOBJECT("dch","prg\dch.prg")
  oDCH.initWithClass(cResource + "Controller", cController)
  IF NOT oDCH.createInstance()
   RETURN "restHelper: controller <b>" + cController + "</b> has some errors:<hr>" + FILETOSTR(FORCEEXT(cController,"TXT"))
  ENDIF

  oController = oDCH.Instance
  cHTMLOut = "" 
  TRY
   oController.Request = poRequest
   oController.Props = poProps
   oController.Params = oParams
   oController.homeFolder = ADDBS(JUSTPATH(cController))
   oController.baseURL = cBaseURL

   IF PEMSTATUS(oController, cAction, 5)
    cHTMLOut = EVALUATE("oController." + cAction + "()")
   ELSE
    cHTMLOut = "restHelper: the <b>" + PROPER(cResource) + "</b>'s controller does not implement the requested action (<b>" + cAction + "</b>)"
   ENDIF

  CATCH TO ex
   cHTMLOut = "restHelper: error while processing the request <b>" + cRESTUri + "</b>:<hr>" + ;
              ex.Message + "<hr>" + cController

  FINALLY
   oController = NULL
   oDCH.releaseInstance()
   
  ENDTRY

  RETURN cHTMLOut  
  *
 ENDPROC
 
 
 * compileController
 * Recompile a controller PRG (only if needed)
 *
 PROCEDURE compileController(pcPRG)
  *
  LOCAL cFXP, lResult
  cFXP = FORCEEXT(pcPRG, "FXP")
  lResult = .T.
  IF !FILE(cFXP) OR FDATE(pcPRG,1) > FDATE(cFXP,1)
   TRY
   	COMPILE (pcPRG)
   	
   CATCH TO ex
   	STRTOFILE(ex.Message, FORCEEXT(pcPRG, "TXT"))
   	lResult = .F.
   	
   ENDTRY
  ENDIF
  
  RETURN lResult
  *
 ENDPROC
 *
ENDDEFINE



* Dictionary (Class)
* Implementacion de un array asociativo
*
* Autor: Victor Espina
* Fecha: Abril 2012
*
* Uso:
* LOCAL oDict
* oDict = CREATE("Dictionary")
* oDict.Add("nombre","VICTOR")
* oDict.Add("apellido","ESPINA")
* ?oDict.Values["nombre"] --> "VICTOR"
*
* IF oDict.containsKey("apellido")
*  oDict.Values["apellido"] = "SARCOS"
* ENDIF
*
* FOR i = 1 TO oDict.Count
*  ?oDict.Keys[i], oDict.Values[i]
* ENDFOR 
* 
* oCopy = oDict.Clone()
* ?oCopy.Values["apellido"] --> "SARCOS"
*
DEFINE CLASS Dictionary AS Custom

 DIMEN Values[1]
 DIMEN Keys[1]
 Count = 0
 
 PROCEDURE Init(pnCapacity)
  IF PCOUNT() = 0
   pnCapacity = 0
  ENDIF
  DIMEN THIS.Values[MAX(1,pnCapacity)]
  DIMEN THIS.Keys[MAX(1,pnCapacity)]
  THIS.Count = pnCapacity
 ENDPROC
  
 PROCEDURE Values_Access(nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   RETURN THIS.Values[nIndex1]
  ENDIF
  LOCAL i
  FOR i = 1 TO THIS.Count
   IF THIS.Keys[i] == nIndex1
    RETURN THIS.Values[i]
   ENDIF
  ENDFOR
 ENDPROC

 PROCEDURE Values_Assign(cNewVal,nIndex1,nIndex2)
  IF VARTYPE(nIndex1) = "N"
   THIS.Values[nIndex1] = m.cNewVal
  ELSE
   LOCAL i
   FOR i = 1 TO THIS.Count
    IF THIS.Keys[i] == nIndex1
     THIS.Values[i] = m.cNewVal
     EXIT
    ENDIF
   ENDFOR
  ENDIF 
 ENDPROC


 * Clear
 * Elimina el contenido de la clase
 *
 PROCEDURE Clear
  DIMEN THIS.Values[1]
  DIMEN THIS.Keys[1]
  THIS.Count = 0
 ENDPROC
 
 * Add
 * Incluye un nuevo item en el diccionario
 *
 PROCEDURE Add(pcKey, puValue)
  IF THIS.ContainsKey(pcKey)
   RETURN .F.
  ENDIF
  THIS.Count = THIS.Count + 1
  DIMEN THIS.Values[THIS.Count]
  DIMEN THIS.Keys[THIS.Count]
  THIS.Values[THIS.Count] = puValue
  THIS.Keys[THIS.Count] = pcKey
 ENDPROC

 * containsKey
 * Permite determinar si existe un item registrado
 * con la clase indicada
 *
 PROCEDURE ContainsKey(pcKey)
  LOCAL i,lResult
  lResult = .F.
  FOR i = 1 TO THIS.Count
   IF THIS.Keys[i] == pcKey
    lResult = .T.
    EXIT
   ENDIF
  ENDFOR
  RETURN lResult  
 ENDPROC
 
 * getKeys
 * Copia en un array la lista de claves registradas
 *
 PROCEDURE getKeys(paTarget)
  IF THIS.Count = 0
   RETURN .F.
  ENDIF
  DIMEN paTarget[THIS.Count]
  ACOPY(THIS.Keys, paTarget)
  RETURN THIS.Count
 ENDPROC
 
 * Clone
 * Devuelve una copia del diccionario con todo su contenido
 *
 PROCEDURE Clone()
  LOCAL oClone,i
  oClone = CREATE(THIS.Class)
  FOR i = 1 TO THIS.Count
   oClone.Add(THIS.Keys[i], THIS.Values[i])
  ENDFOR
  RETURN oClone
 ENDPROC
 
ENDDEFINE


* restController (Class)
* Abstract class for REST resource's controllers
*
DEFINE CLASS restController AS Custom
 Request = NULL
 Props = NULL
 Params = NULL
 homeFolder = ""
 baseUrl = ""
 
 PROCEDURE infoAction
  RETURN "infoAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE helpAction
  RETURN "helpAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE getAction
  RETURN "getAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE listAction
  RETURN "listAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE addAction
  RETURN "addAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE updAction
  RETURN "updAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE dropAction
  RETURN "dropAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE zapAction
  RETURN "zapAction: <b>not implemented</b>"
 ENDPROC 
 PROCEDURE listParameters
  LOCAL cHTML,i
  cHTML = THIS.Class + " - Received Parameters:<ul>"
  FOR i = 1 TO THIS.PArams.Count
   cHTML = cHTML + "<li>" + THIS.Params.Item[i] + "</li>"
  ENDFOR
  RETURN cHTML
 ENDPROC
ENDDEFINE
 