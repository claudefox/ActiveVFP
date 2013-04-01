**************************************************
* WEBTHREADS Version 2 for AVFP 
* Modified 1/26/13 CKF
* * Runs any vfp code on a background thread
* Uses CreateThreadObject in vfp2c32.fll
*  by Christian Ehlscheid
**************************************************

DEFINE CLASS ThreadManager AS Session OLEPUBLIC 
      
      CallInfo = .NULL. && "magic" property
      oEvent = NULL
      lSilent = .f. && just run the thread with no update page, defaults to having an update page
      nRecExpire = 5400  && how long to keep Event records
      
      PROCEDURE CreateThread(ThreadProc as String, ThreadProcParam as String)
            LOCAL lcID,ThreadProcParam
            PUBLIC loObj, loCallback, xj, lnCallId
			#INCLUDE vfp2c.h
			SET LIBRARY TO oProp.AppStartPath+[\vfp2c32.fll] ADDITIVE 
*!*				IF !InitVFP2C32(VFP2C_INIT_ALL) && you can vary the initialization parameter depending on what functions of the FLL you intend to use
*!*				  LOCAL laError[1], lnCount, xj, lcError
*!*				  lnCount = AERROREX('laError') 
*!*				  lcError = 'VFP2C32 Library initialization failed:' + CHR(13)
*!*				  FOR xj = 1 TO lnCount
*!*				    lcError = lcError + ;
*!*				    'Error No : ' + TRANSFORM(laError[1]) + CHR(13) + ;
*!*				    'Function : ' + laError[2] + CHR(13) + ;
*!*				    'Message : "' + laError[3] + '"'
*!*				  ENDFOR
*!*				  RETURN lcError   && show/log error and abort program initialization ..
*!*				ENDIF
            ThreadProcParam = IIF(EMPTY(ThreadProcParam),SUBSTR(SYS(2015),3,10),SUBSTR(SYS(2015),3,10)+[,]+ThreadProcParam)
            ALINES(laArr,ThreadProcParam,.F.,",")
            lcID = laArr[1]
            lcEventFile=oProp.AppStartPath+'temp\events.DBF'
            IF !FILE(lcEventFile)
					CREATE TABLE (lcEventFile) ;
						(id c(10), begin l, finish l , returnval c(50), count n(5,0), cancel l, action m, timestamp t, err l, err_txt m)
					INDEX ON id TAG id
					INDEX ON timestamp TAG timestamp

			ENDIF
            IF ! USED('events')
				USE (lcEventFile) IN 0 SHARED
			ENDIF
			SELECT events
			* recycling old events records
			*SCAN FOR (DATETIME()-events.timestamp)>5400 &&86400 1 day &&5400 90 minutes    
			SCAN FOR (DATETIME()-events.timestamp)>;
               IIF(EMPTY(THIS.nRecExpire),5400,THIS.nRecExpire)  &&86400 1 day &&5400 && 90 minutes    
				DO WHILE .NOT. RLOCK()
				ENDDO
			   	REPL ID WITH ''
			   	UNLOCK
			ENDSCAN 
			DO WHILE .NOT. FLOCK()
			ENDDO
			LOCATE FOR EMPTY(id)  
			IF EOF()
				APPEND BLANK
				REPLACE id WITH lcID  
				REPLACE finish WITH .f.
				REPLACE count WITH 0
				REPLACE begin WITH .t.
				REPLACE cancel WITH .f.
				REPLACE action WITH ''
				REPLACE timestamp WITH DATETIME()
				REPLACE err WITH .f.
				REPLACE err_txt WITH ''
			ELSE
			    REPLACE id WITH lcID  
				REPLACE finish WITH .f.
				REPLACE count WITH 0
				REPLACE begin WITH .t.
				REPLACE cancel WITH .f.
				REPLACE action WITH ''
				REPLACE timestamp WITH DATETIME()
				REPLACE err WITH .f.
				REPLACE err_txt WITH ''
			ENDIF	
			UNLOCK
			USE IN events
            
            IF VARTYPE(ThreadProc)='C'  
              
              *must be compiled object and compile is source date newer         
              IF !FILE(oProp.AppStartPath+"prg\"+ThreadProc+".fxp") OR ;
               FDATE(oProp.AppStartPath+"prg\"+ThreadProc+".prg",1) > FDATE(oProp.AppStartPath+"prg\"+ThreadProc+".fxp",1) 
              		COMPILE oProp.AppStartPath+"prg\"+ThreadProc+".prg"
              ENDIF

              m.loCallback = null   &&CREATEOBJECT('ExampleCallback')

				TRY 
				m.loObj = CreateThreadObject("activevfp.c1", m.loCallback) 
				
				CATCH 
				  AERROREX('laError') 
				  && the array "laError" now contains exact information 
				   &&laError
				ENDTRY 
				
				m.loObj.MyDoCmd("do '"+oProp.AppStartPath+"prg\"+ThreadProc+"' WITH '"+ThreadProcParam+"'")
                
                IF !this.lSilent
                    this.Start(ThreadProcParam)  && show update page
                ENDIF
                
                RETURN lcID  
            ENDIF 

      PROCEDURE Check
            LPARAMETERS lcID
            lcEventFile=oProp.AppStartPath+'temp\events.DBF'
            IF !USED('events')		
						USE (lcEventFile) IN 0 SHARED
			ENDIF
			SELECT events
				    
			SET ORDER TO 1 
			SEEK lcID
			IF FOUND()
					
			        replace count WITH count+1	
			       	SCATTER NAME THIS.oEvent BLANK MEMO
			       	this.oEvent.count=events.count	
			       	this.oEvent.action=events.action
			       	this.oEvent.id=events.id
			       	this.oEvent.begin=events.begin
			       	this.oEvent.finish=events.finish
			       	this.oEvent.cancel=events.cancel
			       	this.oEvent.timestamp=events.timestamp
			       	this.oEvent.err=events.err
			       	this.oEvent.err_txt=events.err_txt		
		         	IF events.finish .or. events.err
			         		RETURN .T.
		         	ELSE
			         	    RETURN .F.
		         	ENDIF
		    ENDIF
		    USE IN events		         		
      
      
      PROCEDURE Cancel
    
      		LPARAMETERS lcID
      		this.SetPath()  && set to data directory wherever activevfp.dll is
            UPDATE events SET cancel=.T.,finish = .f. WHERE id==lcID
	        RELEASE ALL
            
      PROCEDURE GetError
            LPARAMETERS lcID
            this.SetPath()
            *lcEventFile=oProp.AppStartPath+'temp\events.DBF'
            IF !USED('events')		
						USE events IN 0 SHARED
			ENDIF
			SELECT events
				    
			SET ORDER TO 1 
			SEEK lcID
			IF FOUND()
					
			        
		         	IF events.err 
			         		RETURN .T.
		         	ELSE
			         	    RETURN .F.
		         	ENDIF
		    ENDIF
		    USE IN events		         	
            
      PROCEDURE Canceled
            LPARAMETERS lcID
            this.SetPath()
            *lcEventFile=oProp.AppStartPath+'temp\events.DBF'
            IF !USED('events')		
						USE events IN 0 SHARED
			ENDIF
			SELECT events
				    
			SET ORDER TO 1 
			SEEK lcID
			IF FOUND()
					
			        
		         	IF events.cancel 
			         		RETURN .T.
		         	ELSE
			         	    RETURN .F.
		         	ENDIF
		    ENDIF
		    USE IN events		         		

       
	  FUNCTION Start(params)
	      LOCAL cEventID
	      ALINES(arr,params,.F.,",")
		  cEventID = arr[1]
		  oResponse.clear
		  oResponse.Expires=0
		  oResponse.Redirect(oProp.ScriptPath+[?action=]+oProp.Action+[&step=AsyncCheck&asyncID=]+cEventID)
          RETURN 
	  ENDFUNC

  	  PROCEDURE RecordError(cEventID,cErrorTxt)
	  	this.SetPath()
	  	UPDATE events SET err=.t.,finish = .f.,err_txt=ALLTRIM(cErrorTxt),begin =.f.,action=[Error] WHERE id=cEventID  &&,count = 0
	  
	  PROCEDURE StartWebEvent(cEventID,cStartActionTxt)
	  	this.SetPath()
	  	UPDATE events SET finish = .f.,begin =.t.,action=ALLTRIM(cStartActionTxt) WHERE id=cEventID  &&,count = 0
	  
	  PROCEDURE CompleteWebEvent(cEventID,cFinishActionTxt,cfilename)
	    IF EMPTY(cfilename)
	       cfilename=[ ]
	    ENDIF
	    IF EMPTY(cFinishActionTxt)
	        cFinishActionTxt=[ ]
	    ENDIF
	    &&,returnval=ALLTRIM(cfilename),action=ALLTRIM(LEFT(events.action,254))+CHR(13)+CHR(10)+ALLTRIM(cFinishActionTxt)
	    this.SetPath()
 	  	UPDATE events SET finish = .t.,begin =.f.,returnval=ALLTRIM(cfilename),action=ALLTRIM(LEFT(events.action,254))+CHR(13)+CHR(10)+ALLTRIM(cFinishActionTxt) WHERE id=cEventID  &&fix  &&,action=ALLTRIM(events.action)+CHR(13)+CHR(10)+ALLTRIM(cFinishActionTxt)
 	    RELEASE ALL
 	    CLEAR ALL
	    RETURN
	    
	  
	  
	  PROCEDURE StatusWebEvent(cEventID,cTxt,lAdditive)
	  	this.SetPath()
	  	IF lAdditive
		  	UPDATE events SET begin =.f.,action=ALLTRIM(LEFT(events.action,254))+CHR(13)+CHR(10)+ALLTRIM(cTxt) WHERE id=cEventID  &&
		ELSE
			UPDATE events SET begin =.f.,action=ALLTRIM(cTxt) WHERE id=cEventID  &&,count = 0
		ENDIF
		  	
	  PROCEDURE SetPath()
	
	    ASTACKINFO(myarray)
      	lcAppStartPath=JUSTPATH(myarray(1,4))   &&+"\"
      	lcTempPath= STRTRAN(lcAppStartPath,'prg','temp')+'\'  &&temp needs to be writable
	    SET PATH TO &lcTempPath
        RELEASE myarray
        RETURN
        
      FUNCTION ERROR(nError, cMethod, nLine)

            lcErrMsg = cMethod+'  err#='+STR(nError,5)+' line='+STR(nline,6)+;
	          ' '+MESSAGE()+_VFP.SERVERNAME
            STRTOFILE(lcErrMsg,oProp.AppStartPath+[temp\threaderror.txt])
	        COMreturnerror(lcErrMsg,_VFP.SERVERNAME)
	        
      ENDFUNC
      
      FUNCTION INIT
         SYS(3050, 1, VAL(SYS(3050, 1, 0)) / 3)
      ENDFUNC
      
      
          

ENDDEFINE


				    
*!*					    
				 	  
				 

DEFINE CLASS ExampleCallback AS Custom

	FUNCTION OnCallComplete(callid AS Long, result AS Variant, callcontext AS Variant) AS VOID
		 ? callid, result, callcontext
	ENDFUNC
	
	FUNCTION OnError(callid AS Long, callcontext AS Variant, errornumber AS Long, errorsource AS String, errordescription AS String) AS VOID
		 ? callid, callcontext, errornumber, errorsource, errordescription
	ENDFUNC

ENDDEFINE
