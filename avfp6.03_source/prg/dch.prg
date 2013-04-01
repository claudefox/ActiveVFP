* DCH (Class)
* Dynamic Class Handler
*
* Author: Victor Espina
* Date: May 2012
*
* This class allows to create an instance of a given
* class stored in an external PRG file and then release 
* that PRG from memory after being used.
*
* This will allow the PRG to be modified without quiting AVFP.
*
DEFINE CLASS dch AS Custom
 *
 className = ""
 classPRG = ""
 Instance = NULL
 
 PROCEDURE initWithClass(pcClass, pcClassPRG)
  THIS.className = pcClass
  THIS.classPRG = pcClassPRG
 ENDPROC
 
 PROCEDURE createInstance
  IF NOT THIS.Recompile()
   RETURN NULL
  ENDIF
  THIS.Instance = NEWOBJECT(THIS.className, THIS.classPrg)
 ENDPROC
 
 PROCEDURE Recompile
  LOCAL cPRG, cFXP, lResult
  cPRG = THIS.classPRG
  cFXP = FORCEEXT(cPRG, "FXP")
  lResult = .T.
  IF !FILE(cFXP) OR FDATE(cPRG,1) > FDATE(cFXP,1)
   TRY
   	COMPILE (cPRG)
   CATCH TO ex
   	STRTOFILE(ex.Message, FORCEEXT(cPRG, "TXT"))
   	lResult = .F.
   ENDTRY
  ENDIF
  RETURN lResult
 ENDPROC
 
 PROCEDURE releaseInstance
  THIS.Instance = NULL
  CLEAR CLASS (THIS.className)
  CLEAR PROGRAM (THIS.classPRG)
  CLEAR RESOURCES
 ENDPROC
 *
ENDDEFINE