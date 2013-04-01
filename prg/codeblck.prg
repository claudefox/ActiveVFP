* -------------------------------------------------------- *
* CodeBlck.PRG : Code Block "Wrapper"
*
* Visual FoxPro Code Block Interpreter
*
* Created by: J. Randy Pearson, CYCLA Corporation
* Revision 3.2(a), October 5, 2001
* Record of Revision at End Of File.

* Public Domain
*
* Calling "wrapper" program for class cusCodeBlock in 
* Accepts same parameter calls
* as FP 2.6 version (same author).

* See notes at end of file for theory about how "runtime
* interpreter" works.

* To use JUST THE INTERPRETER in another framework, do one of these:
*
* 1) Cut everything between CUT HERE .. CUT HERE (search below) and
*    Paste into your application.
* - or -
*
* 2) SET PROCEDURE TO CodeBlck ADDITIVE, and then you can either use
*    this wrapper, or make direct calls to CLASS CusCodeBlock.
*

* You can also test out code interactively in a FORM by calling this
* PRG with no parameters:  = CodeBlck()

* (IMPORTANT: See notes at end of file on optional TEXTMERGE strategies!)

* ================================================================= *
* Wrapper code for Class CusCodeBlock below:

LPARAMETERS _0_qcCode, _0_qlFile, _0_qlEdit
* _0_qcCode : Text of code to run OR File name with code.
*              If blank, user gets screen to type code.
* _0_qlFile : .T. if 1st parameter is a file name. Internally
*              passed as -1 when recursive call made.
* _0_qlEdit : .T. if user gets to edit code before running. 

* Calling Examples:
* 1) Allow direct typing of code to run:
*      DO CodeBlck
* 2) Run the code contained in memo field "TheCode":
*      DO CodeBlck WITH TheCode
* 3) Same as 2, but allow review/edit first:
*      DO CodeBlck WITH TheCode, .f., .T.
* 4) Run the code found in file "TESTRUN.PRG":
*      DO CodeBlck WITH "TESTRUN.PRG", .T.
* 5) Same as 4, but allow user to review/edit:
*      DO CodeBlck WITH "TESTRUN.PRG", .T., .T.
*    [NOTE: The file doesn't get changed.]

LOCAL _0_qoCb
* object reference to code block or editor

* --- Deal with different calling methods:
IF m._0_qlFile
	* File name as 1st parameter.
	DO CASE
	CASE EMPTY( m._0_qcCode) OR NOT TYPE("m._0_qcCode") == 'C'
		* Process a file, but no file name passed.
		_0_qcCode = GETFILE( 'PRG|TXT', 'Select File', 'Execute')

	CASE '*' $ m._0_qcCode OR '?' $ m._0_qcCode
		_0_qcCode = GETFILE( m._0_qcCode, 'Select File', 'Execute')

	OTHERWISE
		* Explicit file name sent.
	ENDCASE
	
	IF EMPTY( m._0_qcCode)
		* File not found/selected.
		RETURN .F.
	ELSE
		* Store file contents to memvar.
		_0_qcCode = _0_qFile( m._0_qcCode)
	ENDIF
ENDIF

IF NOT TYPE( "m._0_qcCode") == 'C'
	* No code passed - see if any stored from last run.
	IF PROGRAM(1) == "CODEBLCK" AND TYPE( "m._0_qcPrev") == "C"
		_0_qcCode = m._0_qcPrev
	ELSE
		_0_qcCode = SPACE(0)
	ENDIF
	
	_0_qlEdit = .T.
ENDIF  [no code passed as parameter]

_0_qcCode = ALLTRIM( m._0_qcCode)

IF LEFT( m._0_qcCode, 1) == ";"
	* Assume dBW-style code block.  Translate each ;
	* to Cr-Lf so that this routine will run it.
	IF LEN( m._0_qcCode) > 1
		_0_qcCode = STRTRAN( SUBSTR( m._0_qcCode, 2), ";", CHR(13) + CHR(10) )
	ELSE
		_0_qcCode = ""
	ENDIF
ENDIF
	
IF m._0_qlEdit
	* Allow user to enter/edit code:
	_0_qoCb = CREATEOBJECT( "frmCodeBlockEditor")
	_0_qoCb.edtCodeBlock.Value = m._0_qcCode
	_0_qoCb.Show( 1)
	**READ EVENTS
ELSE
	_0_qoCb = CREATEOBJECT( "cusCodeBlock")
	_0_qoCb.SetCodeBlock( m._0_qcCode) 
	_0_qoCb.Execute()
	IF _0_qoCB.lError
		= MessageBox( "CAUTION: Error " + LTRIM( STR( _0_qoCB.nError)) + ;
			" occurred with message " + _0_qoCB.cErrorMessage + ;
			", while processing line " + _0_qoCB.cErrorCode + ".")
	ENDIF
	RETURN _0_qoCb.Result
ENDIF

RETURN
* -------------------------------------------------------- *

FUNCTION _0_qFile
*
* Get file contents.
*
PARAMETER pcFile

IF NOT FILE( m.pcFile)
	RETURN SPACE(0)
ENDIF
LOCAL lnSelect, lcCode
lnSelect = SELECT()
SELECT 0
CREATE CURSOR _0_qFile (Contents M)
APPEND BLANK
APPEND MEMO Contents FROM ( m.pcFile)
lcCode = Contents
USE
SELECT (m.lnSelect)

RETURN m.lcCode

ENDFUNC

**** <<<<< --- CUT HERE (START) to copy the main class into your system.

*** ------------------------------------------- ***
* Allow insertion into prg where CR is already defined:
#IFNDEF CR
	#DEFINE CR   CHR(13) + CHR(10)
#ENDIF

* Debugging/error control settings:
#DEFINE CODEBLOCK_WAIT_ON_ERROR  .F.       && displays wait window on error
#DEFINE CODEBLOCK_WAIT_TIMEOUT   5         && timeout on above, 0 = forever
#DEFINE CODEBLOCK_DEBUG          .F.       && set to FALSE unless debugging
#DEFINE CODEBLOCK_DEBUG_LEVEL    0         && set to 0 except in severe debugging
* 0 = normal, 1 = WAIT WINDOW, 2 = SUSPEND before each line
* applies only if CODEBLOCK_DEBUG = .T.

#DEFINE CODEBLOCK_SHOW_TEXT_FLAG .T.
* Compensates for shortcoming in VFP where you cannot check, via SET(), for
* the current setting of SHOW|NOSHOW in SET TEXTMERGE.
* If constant is TRUE, a property "lShowTextmerge" is included, which controls 
* whether output from textmerge commands is also sent via ? or ??. 
* IF FALSE, all output is sent via ?/?? and you need to use SET CONSOLE OFF 
* in your code to prevent output from going to screen.
*
* Recommended setting:
* .F. to avoid any output to screen ever.
* .T. to make environment most like VFP's current settings.

*** ------------------------------------------- ***
DEFINE CLASS CusCodeBlock AS RELATION
*** ------------------------------------------- ***

DIMENSION aRawCode[ 1]  && Raw original code, to allow trace back.
DIMENSION aCode[ 1, 3]  && Array to store pre-processed code.
DIMENSION aDiag[ 1, 4]  && Array to "diagram" structure of code block.

nRawLines       = 0
nDiagCount      = 0     && Counter for aDiag
nCodeLines      = 0     && Counter for aCode
nRawStart       = 0     && used internally
nRawEnd         = 0     && used internally

oCaller         = NULL   && calling object reference (see INIT)

cCodeBlock      = ""
nRecursionLevel = 0
lPreProcessed   = .F.
lScript         = .F.    && flag for whether code is ASP-style script

* TWO special flags to handle differences in variable
* declaration in non-compiled environment. It is 
* strongly recommended that these flags be set to TRUE:
* 
lInitializePrivates = .T. 
* Takes any PRIVATE statement and adds a "STORE .F. TO <var_list>"
* right afterward, thus avoiding errors when initialization
* occurs in nested code, such as:
*
* PRIVATE i
* IF <condition>
*   i = <value_1>
* ELSE
*   i = <value_2>
* ENDIF
* ? i  && error, without this switch turned on

lLocalToPrivate     = .T.  
* Interpret any LOCAL statement as a PRIVATE statement,
* to avoid numerous errors that occur otherwise. Always
* inserts a STORE .F. TO as well for maximum compatability
* with how the same compiled code would run.

#IF CODEBLOCK_DEBUG
lDontExecute    = .F.
#ENDIF

nCodeCheckingLevel = 2         && recommended setting: 2
* 0 = no optional checking (maximum performance)
* 1 = check and ignore (strip out) illegal/unsupported commands
* 2 = check and raise error if any illegal/unsupported commands
* NOTES:
* a) This replaces #DEFINE CODEBLOCK_TRUST_CODE in previous version.
* b) It's now possible to set this to 1 or 2 and have minimal
*    performance impact, since all checks are done only once
*    in the PreProcess() method.
* c) You must set this to at least 1 to enable AddIllegalCommand support.

DIMENSION aIllegal[ 1, 2]
aIllegal[ 1, 2] = .F.    && array of optional illegal commands to cause code abort

Result          = .T.    && used internally to establish RETURN value
xErrorReturn    = .F.    && value to RETURN if an error occurs (default .F.)
cReturnType     = "X"    && Type-check of return value. (X = any, can be multiple, i.e. "DT")

cExitCode       = ""     && [internal control variable]

lError          = .F.    && did error occur flag
nError          = 0      && VFP error (1089, user-defined, when program detects error)
cErrorMethod    = ""     && method in which VFP error occurred
nErrorLine      = 0      && line # of in pre-processed code in which VFP error occurred
nRawErrorLine   = 0      && line # in original code in which VFP error occurred
cErrorMessage   = ""     && error message
cErrorCode      = ""     && errant line of code, or similar message
lMacroError     = .F.    && did error occur when macro expanding user code?
nErrorRecursionLevel = 0 && level of recursion at which error occurred

* Properties for extended TEXTMERGE capabilities:
*
* (IMPORTANT: See notes at end of file on optional TEXTMERGE strategies!)
*
lAccumulateMergedText = .F.   && flag to accumulate all textmerge output
cAltMergeFunction     = ""    && alternative merge function name
cMergedText           = ""    && accumulated merged text
cTextmergeVariable    = ""    && variable to store accumulated text at end
lWriteAccumulatedText = .T.   && flag to write accumulated text at end
lClassicTextmerge     = .F.   && when not accumulating, use classic (slower) line-
*                             && by-line textmerge (vs. THIS.InternalMergeText function)
#IF CODEBLOCK_SHOW_TEXT_FLAG
lShowTextmerge        = .F.   && don't send textmerge via ?/??
#ENDIF

* ------------------------------------------------ *
FUNCTION Init( tcCode, toCaller )

THIS.AddIllegalCommand( "QUIT")
THIS.AddIllegalCommand( "CANC")
**THIS.AddIllegalCommand( "FOR EACH")

THIS.ResetProperties()

* 1) establish code block
IF NOT EMPTY( m.tcCode) AND TYPE( "m.tcCode") == "C"
	THIS.SetCodeBlock( m.tcCode)
ENDIF
* 2) caller object reference, if passed
IF PCOUNT() >= 2
	THIS.SetCallingObject( m.toCaller)
ENDIF

ENDFUNC
* ------------------------------------------------ *

FUNCTION ResetProperties
*
* Allow object reuse by resetting all properties
* needed to allow a second block to be executed.
*
THIS.lError          = .F.
THIS.nError          = 0 
THIS.cErrorMethod    = ""     
THIS.nErrorLine      = 0      
THIS.nRawErrorLine   = 0
THIS.cErrorMessage   = ""     
THIS.cErrorCode      = ""     
THIS.lMacroError     = .F.
THIS.nErrorRecursionLevel = 0 

THIS.cCodeBlock      = ""
THIS.cMergedText     = ""
THIS.cExitCode       = ""
THIS.nRecursionLevel = 0
THIS.lPreProcessed   = .F.
THIS.lScript         = .F.

DIMENSION THIS.aCode[ 1, 3]
DIMENSION THIS.aDiag[ 1, 4]
DIMENSION THIS.aRawCode[ 1]
THIS.nDiagCount = 0
THIS.nCodeLines = 0
THIS.nRawLines  = 0

THIS.nRawStart   = 0     
THIS.nRawEnd     = 0     

ENDFUNC  && ResetProperties

* ------------------------------------------------ *
FUNCTION BeforePreProcessCode  && <<--- Hook!
RETURN .T.  && .F. to kill processing
ENDFUNC  && BeforePreProcessCode
* ------------------------------------------------ *

FUNCTION PreProcessCode
*
* Analyze raw code block to facilitate Execute
* method. The following are accomplished:
*
* a) All control structure information is extracted
*    into an array to help control loops in Execute().
* b) Errors in control structures and presence of any
*    illegal/unsupported commands are identified up front.
* c) Cleaned up code is placed in an array.
* d) All text between TEXT..ENDTEXT is accumulated into
*    single strings to improve textmerge performance in loops.
*

IF NOT THIS.BeforePreProcessCode()  && Hook
	IF NOT THIS.lError
		THIS.SetError( "BeforePreprocessCode hook returned FALSE." )
	ENDIF
	RETURN .F.
ENDIF

LOCAL ARRAY laPoundIf[ 1, 2]
LOCAL ARRAY laWith[ 1]
LOCAL ARRAY laDefines[ 1, 2]
LOCAL ARRAY laIncludeCode[ 1]

LOCAL lnIncludeLines, lnIncludePointer

LOCAL lnWith, lnDefines, lnPoundIf
STORE 0 TO lnWith, lnDefines, lnPoundIf

LOCAL lnOldMemo, lnOldMline, lcOldExact, lcRawLine, lcUpper, lcCleanLine, ;
	jj, lnDepth, lnCleanLines, ;
	llInText, llContinued, llComment, lnAtPos, llMatch, llProblem, ;
	lcTextBlock, llNewNest, lcStr, lcStr2, ;
	lnAt, lnAt2, lnAtStart, lcConstant, lcValue
	
lnOldMemo = SET( "MEMOWIDTH" )
SET MEMOWIDTH TO 8192  
*!*	lnOldMline = _MLINE
*!*	_MLINE = 0
lcOldExact = SET( 'EXACT')
SET EXACT OFF

IF THIS.lScript OR THIS.cCodeBlock = "<HTML>"
	THIS.InvertScript()
ENDIF

lnDepth = 0

* Read the code in:
THIS.nRawLines = ALINES( THIS.aRawCode, THIS.cCodeBlock )

* Dimension the processed code array:
DIMENSION THIS.aCode[ THIS.nRawLines, 3]
* Max possible, so we don't re-DIM all the time.
* New: 2nd column is for tracking original code line for error messages.

lcCleanLine = ""
lnCleanLines = 0

SET MEMOWIDTH TO m.lnOldMemo

THIS.nRawStart = 1
THIS.nRawEnd = 0

DO WHILE .T.
	* Cannot use FOR..ENDFOR, since nRawLines can grow,
	* if a #INCLUDE file is inserted.
	THIS.nRawEnd = THIS.nRawEnd + 1
	IF THIS.nRawEnd > THIS.nRawLines
		EXIT
	ENDIF

	* Read next line of raw code:
	*!* lcRawLine = MLINE( m.lcRawCode, 1, _MLINE )
	lcRawLine = THIS.aRawCode[ THIS.nRawEnd]
	
	* This section accumulates any multi-line statements,
	* gathers any text between TEXT..ENDTEXT and otherwise
	* cleans up the code for further processing:
	IF m.llInText  
		* Within TEXT..ENDTEXT: don't touch,
		* unless "ENDTEXT" is found.
		* lnCleanLines = m.lnCleanLines + 1
		IF LTRIM( STRTRAN( UPPER( m.lcRawLine ), CHR( 9))) = "ENDT"
			* First, the accumulated block of text:
			lnCleanLines = m.lnCleanLines + 1
			THIS.SetCodeLine( m.lnCleanLines, m.lcTextBlock)
			
			* Then, the ENDTEXT:
			lnCleanLines = m.lnCleanLines + 1
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDTEXT")
			lnDepth = m.lnDepth - 1
			llInText = .F.
			
			IF m.lnCleanLines > THIS.nRawLines  && unlikely
				DIMENSION THIS.aCode[ m.lnCleanLines, 3]
			ENDIF
			
			THIS.SetCodeLine( m.lnCleanLines, "ENDTEXT")
		ELSE  && simply accumulate the text block
			lcTextBlock = m.lcTextBlock + m.lcRawLine + CR
		ENDIF
		LOOP
	ENDIF
	* Strip out any TAB's in the code (simplifies later processing):
	IF m.llContinued
		* 2nd or later line in multi-line
		* statement, attach but don't LTRIM(),
		* since we could be in middle of delimited string.
		lcCleanLine = m.lcCleanLine + TRIM( STRTRAN( ;
			m.lcRawLine, CHR(9), SPACE(1) ) )
	ELSE  && new line of code
		THIS.nRawStart = THIS.nRawEnd
		lcCleanLine = ALLTRIM( STRTRAN( ;
			m.lcRawLine, CHR(9), SPACE(1) ) )
	ENDIF
	IF EMPTY( m.lcCleanLine) && blank line
		LOOP
	ENDIF

	*!*		IF LTRIM( m.lcCleanLine) = "*"  && comment line
	*!*			LOOP
	*!*		ENDIF

	IF LTRIM( m.lcCleanLine) = "&" + "&"
		* Comment line (note: can't type 2 &'s together in VFP)
		LOOP
	ENDIF
	
	* Look for "&&" in line:
	lnAtPos = AT( "&" + "&", m.lcCleanLine )
	* (Note gymnastics used to avoid compile error.)
	IF m.lnAtPos > 0
		lcCleanLine = TRIM( LEFT( m.lcCleanLine, m.lnAtPos - 1))
		IF EMPTY( m.lcCleanLine)
			LOOP
		ENDIF
		llComment = .T.
	ELSE
		llComment = .F.
	ENDIF

	* Check for semi-colon at end of line--denoting line continuation:
	IF RIGHT( m.lcCleanLine, 1) = CHR( 59) AND THIS.nRawEnd < THIS.nRawLines
		IF m.llComment
			* Not allowed on same line!
			THIS.SetError( "Syntax Error: Semi-Colon and double-& on same line.", ;
				m.lcCleanLine )
			lcCleanLine = SPACE(0)
			EXIT
		ELSE
			llContinued = .T.
			lcCleanLine = LEFT( m.lcCleanLine, LEN( m.lcCleanLine) - 1)
			LOOP
		ENDIF
	ELSE
		llContinued = .F.
	ENDIF

	* Check/ignore for comment lines. [09/18/2001: This check was moved down to 
	* here so we captured continuation comments correctly.]
	IF LTRIM( m.lcCleanLine) = "*"  && comment line
		LOOP
	ENDIF

	* NEW SECTION -- PLEASE TEST AND USE WITH CAUTION!
	*
	* Substitute for any #DEFINE's. For first round, this handles constant
	* replacement only. No support for #IF yet--use IF if possible.
	FOR jj = 1 TO m.lnDefines
		lnAtStart = 1
		lcConstant = laDefines[ m.jj, 1]
		DO WHILE m.lnAtStart <= LEN( m.lcCleanLine)
			lnAt = ATC( m.lcConstant, SUBSTR( m.lcCleanLine, m.lnAtStart)) + ;
				m.lnAtStart - 1
			*lnAtStart = m.lnAt + LEN( m.lcConstant)
			DO CASE
			CASE m.lnAt < m.lnAtStart
				EXIT
			CASE NOT ( ;
				m.lnAt = 1 OR ;
				SUBSTR( m.lcCleanLine, m.lnAt -1, 1) $ "[ (*-+/!^%$#<>&.=" )
				*
				lnAtStart = m.lnAtStart + 1
				LOOP
			CASE NOT ( ;
				m.lnAt + LEN( m.lcConstant) - 1 = LEN( m.lcCleanLine) OR ;
				SUBSTR( m.lcCleanLine, m.lnAt + LEN( m.lcConstant), 1) $ ;
					"] )*-+/!^%$#<>.=" )
				*
				lnAtStart = m.lnAtStart + 1
				LOOP
			OTHERWISE
				lcCleanLine = STUFF( m.lcCleanLine, m.lnAt, ;
					LEN( m.lcConstant), laDefines[ m.jj, 2] )
				lnAtStart = m.lnAt + LEN( laDefines[ m.jj, 2] )
			ENDCASE
		ENDDO
	ENDFOR
	
	* Deal with #IF structures:
	*
	DO CASE
	CASE NOT LEFT( m.lcCleanLine, 1) == "#"
		= .F.
	CASE PADR( UPPER( m.lcCleanLine), 4) == "#IF "
		lnPoundIf = m.lnPoundIf + 1
		DIMENSION laPoundIf[ m.lnPoundIf, 2]
		lcStr = SUBSTR( m.lcCleanLine, 4)
		lcType = TYPE( m.lcStr )
		DO CASE
		CASE m.lcType = "L"
			IF EVALUATE( m.lcStr) = .T.
				laPoundIf[ m.lnPoundIf, 1 ] = .T.
				laPoundIf[ m.lnPoundIf, 2 ] = .T.
			ENDIF
		OTHERWISE
			THIS.SetError( "Invalid #IF directive: " + m.lcStr )
			EXIT
		ENDCASE
		LOOP
	CASE PADR( UPPER( m.lcCleanLine), 6) == "#ELIF "
		IF m.lnPoundIf <= 0
			THIS.SetError( "Mismatched #ELIF directive." )
			EXIT
		ENDIF
		IF laPoundIf[ m.lnPoundIf, 2 ] = .T. && previous TRUE case
			laPoundIf[ m.lnPoundIf, 1 ] = .F.
			LOOP
		ENDIF
		lcStr = SUBSTR( m.lcCleanLine, 6)
		lcType = TYPE( m.lcStr )
		DO CASE
		CASE m.lcType = "L"
			IF EVALUATE( m.lcStr) = .T.
				laPoundIf[ m.lnPoundIf, 1 ] = .T.
				laPoundIf[ m.lnPoundIf, 2 ] = .T.
			ENDIF
		OTHERWISE
			THIS.SetError( "Invalid #ELIF directive: " + m.lcStr )
			EXIT
		ENDCASE
		LOOP
	CASE PADR( UPPER( m.lcCleanLine), 5) == "#ELSE"
		IF m.lnPoundIf <= 0
			THIS.SetError( "Mismatched #ELSE directive." )
			EXIT
		ENDIF
		IF laPoundIf[ m.lnPoundIf, 1 ] = .F.
			IF laPoundIf[ m.lnPoundIf, 2] = .F.
				laPoundIf[ m.lnPoundIf, 1] = .T.
			ENDIF
		ELSE
			laPoundIf[ m.lnPoundIf, 1 ] = .F.
		ENDIF
		LOOP
	CASE PADR( UPPER( m.lcCleanLine), 6) == "#ENDIF"
		lnPoundIf = m.lnPoundIf - 1
		IF m.lnPoundIf < 0
			THIS.SetError( "Mismatched #ENDIF directive." )
			EXIT
		ENDIF
		LOOP
	ENDCASE  && #IF stuff

	IF m.lnPoundIf > 0 AND laPoundIf[ m.lnPoundIf, 1 ] = .F.
		* We're in a FALSE part of a #IF -- ignore line.
		LOOP
	ENDIF
		
	* Insert any #INCLUDE files:
	IF PADR( UPPER( m.lcCleanLine), 9) == "#INCLUDE "
		lcStr = ALLTRIM( SUBSTR( m.lcCleanLine, 9))
		IF FILE( m.lcStr)
			lcStr = THIS.FileToStr( m.lcStr )

			* Parse the INCLUDE file to an array:
			SET MEMOWIDTH TO 8192
			lnIncludeLines = ALINES( laIncludeLines, m.lcStr)
			SET MEMOWIDTH TO m.lnOldMemo
			
			* Insert the included lines back into the *raw* code:
			THIS.nRawLines = THIS.nRawLines + m.lnIncludeLines
			DIMENSION THIS.aRawCode[ THIS.nRawLines]
			FOR lnIncludePointer = 1 TO m.lnIncludeLines
				AINS( THIS.aRawCode, THIS.nRawEnd + m.lnIncludePointer) 
				THIS.aRawCode[ THIS.nRawEnd + m.lnIncludePointer] = laIncludeLines[ m.lnIncludePointer]
			ENDFOR
			* Re-dim the processed code if needed for new size:
			IF ALEN( THIS.aCode, 1) < THIS.nRawLines
				DIMENSION THIS.aCode[ THIS.nRawLines, 3 ]
			ENDIF
			
			* Loop back so we process the next line (from include file):
			LOOP
		ELSE
			THIS.SetError( "#INCLUDE file not found: " + m.lcStr )
			EXIT
		ENDIF
	ENDIF
	
	* Process any #DEFINE's:
	IF PADR( UPPER( m.lcCleanLine), 8) == "#DEFINE "
		lnAt = AT( SPACE(1), m.lcCleanLine )
		IF LEN( m.lcCleanLine) = m.lnAt
			THIS.SetError( "Illegal #DEFINE format.")
			EXIT
		ENDIF
		lcConstant = LTRIM( SUBSTR( m.lcCleanLine, m.lnAt + 1))
		lnAt = AT( SPACE(1), m.lcConstant)
		IF m.lnAt = 0 OR m.lnAt = LEN( m.lcConstant)
			THIS.SetError( "Illegal #DEFINE format.")
			EXIT
		ENDIF
		lcValue = ALLTRIM( SUBSTR( m.lcConstant, m.lnAt + 1))
		IF EMPTY( m.lcValue) 
			THIS.SetError( "Illegal #DEFINE format.")
			EXIT
		ENDIF
		lcConstant = UPPER( LEFT( m.lcConstant, m.lnAt - 1))
		llConstantExists = .F.
		FOR jj = 1 TO m.lnDefines
			IF laDefines[ m.jj, 1] == m.lcConstant
				llConstantExists = .T.
				EXIT
			ENDIF
		ENDFOR
		IF m.llConstantExists
			* Emulate VFP's annoying behavior:
			THIS.SetError( "Constant " + m.lcConstant + ;
				" was already created with a #DEFINE.")
			EXIT
		ENDIF
		lnDefines = m.lnDefines + 1
		DIMENSION laDefines[ m.lnDefines, 2]
		laDefines[ m.lnDefines, 1] = m.lcConstant
		laDefines[ m.lnDefines, 2] = m.lcValue
		LOOP
	ENDIF
	
	* --- Create an upper-case version of the line:
	lcUpper = UPPER( m.lcCleanLine)

	* --- Process any WITH or ENDWITH statements separately:
	IF PADR( m.lcUpper, 5) == "WITH "
		* We don't "do" WITH's, we store argument in stack and
		* apply them to subsequent lines of code.
		lnWith = lnWith + 1
		DIMENSION laWith[ m.lnWith]
		laWith[ m.lnWith] = ALLTRIM( SUBSTR( m.lcCleanLine, 5))
		IF m.lnWith > 1  && WITH's are cascading
			laWith[ m.lnWith] = laWith[ m.lnWith - 1] + ;
				laWith[ m.lnWith]
		ENDIF
		
		LOOP
	ENDIF
	IF PADR( m.lcUpper, 4) == "ENDW"
		* ENDWITH - see note above on WITH
		lnWith = MAX( 0, m.lnWith - 1)
		IF m.lnWith = 0
			DIMENSION laWith[ 1]
			laWith[ 1] = ""
		ELSE
			DIMENSION laWith[ m.lnWith]
		ENDIF
		LOOP
	ENDIF
	IF m.lnWith > 0
		* We're within WITH..ENDWITH, so STRTRAN() as needed first.
		lnAtStart = 1
		lcCleanLine = SPACE(1) + m.lcCleanLine
		* The SPACE(1) is prepended to handle lines starting with a "."
		DO WHILE m.lnAtStart <= LEN( m.lcCleanLine)
			lnAt = AT( SPACE(1) + ".", SUBSTR( m.lcCleanLine, m.lnAtStart))
			IF m.lnAt = 0
				EXIT
			ENDIF
			lcStr = LTRIM( UPPER( SUBSTR( m.lcCleanLine, m.lnAt + m.lnAtStart - 1)))
			IF NOT INLIST( m.lcStr, ".T.", ".F.", ".AND.", ".OR.", ".NOT.")
				lcCleanLine = STUFF( m.lcCleanLine, ;
					m.lnAt + m.lnAtStart - 1, 2, ;
					SPACE(1) + laWith[ m.lnWith] + ".")
			ENDIF
			lnAtStart = m.lnAt + m.lnAtStart + 1
		ENDDO
		lcCleanLine = LTRIM( m.lcCleanLine)
		lcUpper = UPPER( m.lcCleanLine )
	ENDIF
	
	* --- Optional check of code for non-supported commands:
	IF THIS.nCodeCheckingLevel > 0  
		llProblem = .F.
		DO CASE			
		CASE THIS.IllegalCommandFound( m.lcUpper )
			llProblem = .T.
			* This method can also set THIS.lError for certain commands.
		
		CASE PADR( m.lcUpper, 9) == "CLEAR ALL" OR ;
			PADR( m.lcUpper, 8) == "CLEA ALL" OR ;
			PADR( m.lcUpper, 10) == "CLEAR MEMO" OR ;
			PADR( m.lcUpper, 9) == "CLEA MEMO" OR ;
			PADR( m.lcUpper, 7) == "RETU TO" OR ;
			PADR( m.lcUpper, 8) == "RETUR TO" OR ;
			PADR( m.lcUpper, 9) == "RETURN TO" OR ;
			PADR( m.lcUpper, 8) == "RELE ALL" OR ;
			PADR( m.lcUpper, 9) == "RELEA ALL" OR ;
			PADR( m.lcUpper, 10) == "RELEAS ALL" OR ;
			PADR( m.lcUpper, 11) == "RELEASE ALL" 
			*
			* These are known to break the system.
			llProblem = .T.
			
		CASE PADR( m.lcUpper, 4) == "REST" AND ;
			"FROM " $ m.lcUpper AND ;
			NOT "ADDI" $ m.lcUpper
			*
			* THIS.SetError( "Can't have RESTORE FROM w/o ADDITIVE.", m.lcUpper)
			* EXIT
			llProblem = .T.
			
		CASE INLIST( PADR( m.lcUpper, 4), "PROC", "FUNC")
			llProblem = .T.

		ENDCASE
		
		IF m.llProblem
			IF THIS.nCodeCheckingLevel >= 2  && raise error
				THIS.SetError( "Command not supported.", m.lcUpper )
				EXIT
			ELSE  && ignore the line
				LOOP	
			ENDIF
		ENDIF
	ENDIF  && Optional checking.

	lnCleanLines = m.lnCleanLines + 1
	llMatch = .T.
	llNewNest = .F.
	
	* --- Main part of pre-processing. Create a "code diagram":
	DO CASE
	CASE m.lcUpper = "TEXT"
		llInText = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "TEXT")
		lcTextBlock = ""
	CASE m.lcUpper = "DO WHIL"
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "DO WHILE")
		
		* Convert any abbreviated (DO WHIL) so Execute method does not
		* have to worry about this:
		lnAT = ATC( "WHIL", m.lcCleanLine)
		lcCleanLine = "DO WHILE " + ;
			IIF( SUBSTR( m.lcCleanLine, m.lnAt + 4, 1) $ "eE", ;
				LTRIM( SUBSTR( m.lcCleanLine, m.lnAt + 5)), ;
				LTRIM( SUBSTR( m.lcCleanLine, m.lnAt + 4)) )
			
	CASE m.lcUpper = "FOR EACH"
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "FOR EACH")
	CASE m.lcUpper = "FOR "
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "FOR")
	CASE m.lcUpper = "SCAN"
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "SCAN")
	CASE m.lcUpper = "IF " OR m.lcUpper = "IF("
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "IF")
	CASE m.lcUpper = "DO CASE"
		llNewNest = .T.
		lnDepth = m.lnDepth + 1
		THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "DO CASE")
	CASE m.lcUpper = "CASE"
		llNewNest = .T.
		llMatch = THIS.FindMatch( m.lnDepth, "CASE", "DO CASE")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "CASE")
		ENDIF
	CASE m.lcUpper = "OTHE"
		lcCleanLine = "OTHERWISE"
		llNewNest = .T.
		llMatch = THIS.FindMatch( m.lnDepth, "CASE")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "OTHERWISE")
		ENDIF
	CASE m.lcUpper = "ELSE"
		llNewNest = .T.
		llMatch = THIS.FindMatch( m.lnDepth, "IF" )
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ELSE")
		ENDIF
	CASE m.lcUpper = "ENDC"
		lcCleanLine = "ENDCASE"
		llMatch = THIS.FindMatch( m.lnDepth, "CASE", "OTHERWISE")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDCASE")
			lnDepth = m.lnDepth - 1
		ENDIF
	CASE m.lcUpper = "ENDD"
		lcCleanLine = "ENDDO"
		llMatch = THIS.FindMatch( m.lnDepth, "DO WHILE")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDDO")
			lnDepth = m.lnDepth - 1
		ENDIF
	CASE m.lcUpper = "ENDI"
		lcCleanLine = "ENDIF"
		llMatch = THIS.FindMatch( m.lnDepth, "IF", "ELSE")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDIF")
			lnDepth = m.lnDepth - 1
		ENDIF
	CASE m.lcUpper = "ENDS"
		lcCleanLine = "ENDSCAN"
		llMatch = THIS.FindMatch( m.lnDepth, "SCAN" )
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDSCAN")
			lnDepth = m.lnDepth - 1
		ENDIF
	CASE m.lcUpper = "ENDF" OR m.lcUpper = "NEXT"
		lcCleanLine = "ENDFOR"
		llMatch = THIS.FindMatch( m.lnDepth, "FOR")
		IF m.llMatch
			THIS.AddDiagramItem( m.lnCleanLines, m.lnDepth, "ENDFOR")
			lnDepth = m.lnDepth - 1
		ENDIF

	OTHERWISE
		* normal code--no action needed now
	ENDCASE
	
	IF m.llMatch = .F.
		THIS.SetError( "Nesting Error", m.lcCleanLine )
		EXIT
	ENDIF
	
	DO CASE
	CASE PADR( m.lcUpper, 7) == "PRIVATE" AND ;
		THIS.lInitializePrivates
		*
		IF m.lnCleanLines + 1 > THIS.nRawLines && need 2 lines
			DIMENSION THIS.aCode[ m.lnCleanLines + 1, 3]
		ENDIF
		* First the PRIVATE itself:
		THIS.SetCodeLine( m.lnCleanLines, m.lcCleanLine)
		lnCleanLines = m.lnCleanLines + 1
		lnAtPos = AT( SPACE(1), m.lcCleanLine )
		* Create a matching initialization statement:
		lcCleanLine = "STORE .F. TO " + ;
			ALLTRIM( SUBSTR( m.lcCleanLine, m.lnAtPos))
		THIS.SetCodeLine(m.lnCleanLines, m.lcCleanLine)
	
	CASE PADR( m.lcUpper, 5) == "LOCAL" AND ;
		( ")" $ m.lcUpper OR "]" $ m.lcUpper ) AND ;
		THIS.lLocalToPrivate
		*
		* LOCAL [ARRAY]. This must be parsed and handled
		* separately.
		IF m.lnCleanLines + 1 > THIS.nRawLines && need 2 lines
			DIMENSION THIS.aCode[ m.lnCleanLines + 1, 3]
		ENDIF
		* Extract the variable names:
		lnAtPos = AT( SPACE(1), m.lcCleanLine )
		lcStr = ALLTRIM( SUBSTR( m.lcCleanLine, m.lnAtPos))
		* Get rid of optional "ARRAY" word:
		IF UPPER( m.lcStr) = "ARRAY" + SPACE(1)
			lcStr = ALLTRIM( SUBSTR( m.lcStr, 6))
		ENDIF
		* Now create a list of variables *without*
		* the dimensions included:
		lcStr2 = STRTRAN( STRTRAN( m.lcStr, "(", "["), ")", "]" )
		DO WHILE "[" $ m.lcStr2
			lnAt = AT( "[", m.lcStr2 )
			lnAt2 = AT( "]", m.lcStr2 )
			IF m.lnAt2 = 0 OR m.lnAt2 < m.lnAt
				llProblem = .T.
				EXIT
			ELSE
				* Take out the dimensions:
				lcStr2 = STUFF( m.lcStr2, m.lnAt, 1 + m.lnAt2 - m.lnAt, "")
				LOOP
			ENDIF
		ENDDO
		IF m.llProblem
			THIS.SetError( "Invalid dimension syntax for local arrays." )
			EXIT
		ENDIF
		* Create a PRIVATE statement:
		THIS.SetCodeLine( m.lnCleanLines, "PRIVATE " + m.lcStr2)
		lnCleanLines = m.lnCleanLines + 1
		* Also insert a DIMENSION statement, since LOCAL
		* arrays are dimensioned right in their declaration:
		THIS.SetCodeLine( m.lnCleanLines, "DIMENSION " + m.lcStr)
		
	CASE PADR( m.lcUpper, 5) == "LOCAL" AND ;
		THIS.lLocalToPrivate
		*
		* (See above CASE for LOCAL ARRAY.)
		IF m.lnCleanLines + 1 > THIS.nRawLines && need 2 lines
			DIMENSION THIS.aCode[ m.lnCleanLines + 1, 3]
		ENDIF
		* Extract the variable names:
		lnAtPos = AT( SPACE(1), m.lcCleanLine )
		* Create a PRIVATE statement:
		THIS.SetCodeLine( m.lnCleanLines, "PRIVATE " + ;
			ALLTRIM( SUBSTR( m.lcCleanLine, m.lnAtPos)))
		lnCleanLines = m.lnCleanLines + 1
		* Also create an initialization statement, since LOCAL
		* variables all start with .F. value by default:
		THIS.SetCodeLine( m.lnCleanLines, "STORE .F. TO " + ;
			ALLTRIM( SUBSTR( m.lcCleanLine, m.lnAtPos)))
		
	OTHERWISE
		IF m.lnCleanLines > THIS.nRawLines  && unlikely
			DIMENSION THIS.aCode[ m.lnCleanLines, 3]
		ENDIF
		THIS.SetCodeLine( m.lnCleanLines, m.lcCleanLine)
		
	ENDCASE
ENDDO

IF m.lcOldExact = "ON"
	SET EXACT ON
ENDIF

*!*	_MLINE = m.lnOldMline
SET MEMOWIDTH TO m.lnOldMemo

THIS.nCodeLines = m.lnCleanLines
THIS.lPreProcessed = .T.

IF NOT THIS.lError
	IF NOT THIS.AfterPreProcessCode()
		IF NOT THIS.lError
			THIS.SetError( "AfterPreprocessCode hook returned FALSE." )
		ENDIF
		RETURN .F.
	ENDIF
ENDIF

RETURN NOT THIS.lError

ENDFUNC  && PreProcessCode

* ------------------------------------------------ *
FUNCTION AfterPreProcessCode  && <<-- Hook!
RETURN .T.  && .F. to kill processing
ENDFUNC  && AfterPreProcessCode

* ------------------------------------------------ *
FUNCTION SetCodeLine( lnLine, lcCode)
THIS.aCode[ m.lnLine, 1] = m.lcCode
THIS.aCode[ m.lnLine, 2] = THIS.nRawStart
THIS.aCode[ m.lnLine, 3] = THIS.nRawEnd
ENDFUNC  && SetCodeLine

* ------------------------------------------------ *

FUNCTION InvertScript
*
* Inverts ASP-style script in cCodeBlock and 
* changes to VFP code with TEXT..ENDTEXT surrounding
* all previous script text.
*
LPARAMETERS tcCode, tlReturnCode

IF TYPE( "m.tcCode") <> "C"
	tcCode = THIS.cCodeBlock
ENDIF

DO WHILE .T.
	lnLen = LEN( m.tcCode )
	lnAt1 = AT( "<%=", m.tcCode )
	IF m.lnAt1 = 0 OR m.lnAt1 >= m.lnLen - 3
		EXIT
	ENDIF
	lnAt2 = AT( "%>", SUBSTR( m.tcCode, m.lnAt1 + 3) )
	IF m.lnAt2 = 0
		EXIT
	ENDIF
	IF m.lnAt1 = 1  && <%=%> ???
		tcCode = STUFF( m.tcCode, m.lnAt1, 5, "")
		LOOP
	ENDIF
	tcCode = STUFF( m.tcCode, m.lnAt1, m.lnAt2 + 4, ;
		"<<" + ALLTRIM( SUBSTR( m.tcCode, m.lnAt1 + 3, m.lnAt2 - 1 )) + ">>")
ENDDO

tcCode = STRTRAN( m.tcCode, "<%", CR + "ENDTEXT" + CR )
tcCode = STRTRAN( m.tcCode, "%>", CR + "TEXT" + CR )
tcCode = "TEXT" + CR + m.tcCode + CR + "ENDTEXT" + CR

IF m.tlReturnCode
	RETURN m.tcCode
ELSE
	THIS.cCodeBlock = m.tcCode
	THIS.lScript = .F.
ENDIF

ENDFUNC && InvertScript

* ------------------------------------------------ *
FUNCTION GetPreProcessedCodeBlock( tnStartLine, tnEndLine, ;
	lIncludeOriginalLineNumbers )
*
* Return pre-processed code as a string. Allows you to see the
* results of the preprocessing step.
*
tnStartLine = IIF( TYPE( "m.tnStartLine") = "N", m.tnStartLine, 1)
tnStartLine = MAX( 1, m.tnStartLine)

tnEndLine = IIF( TYPE( "m.tnEndLine") = "N", m.tnEndLine, THIS.nCodeLines)
tnEndLine = MIN( m.tnEndLine, THIS.nCodeLines)

LOCAL ii, lcCode, llTextNext, lcLine
lcCode = ""
FOR ii = m.tnStartLine TO m.tnEndLine
	lcLine = THIS.aCode[ m.ii, 1]
	IF m.lIncludeOriginalLineNumbers
		lcLine = m.lcLine + ' [' + TRANS( THIS.aCode[ m.ii, 2]) + ;
			IIF( THIS.aCode[ m.ii, 2] = THIS.aCode[ m.ii, 3], [], ;
				', ' + TRANS( THIS.aCode[ m.ii, 3])) + ']'
	ENDIF
	DO CASE
	CASE m.llTextNext
		llTextNext = .F.
		lcCode = m.lcCode + m.lcLine
	CASE m.lcCode == "TEXT"
		llTextNext = .T.
		lcCode = m.lcCode + m.lcLine + CR
	OTHERWISE
		lcCode = m.lcCode + m.lcLine + CR
	ENDCASE
ENDFOR

RETURN m.lcCode

ENDFUNC  && GetPreProcessedCodeBlock

* ------------------------------------------------ *
FUNCTION GetErrorContext( tnLines, tlNoHtml )
*
* Fetches context information for an error. If the error occurred
* in a macro expansion of user code, returns the errant line PLUS
* some surrounding lines of code.
*
tnLines = IIF( TYPE( "m.tnLines") = "N", m.tnLines, 2)
* # of lines on each side of error line to show (default = 2)
LOCAL lcLine, lcCode, ii
lcCode = ""
DO CASE
CASE THIS.lError = .F.
	= .F. && no error
	
CASE THIS.nError = 1089  
	* user-defined (code structure error)
	* no surrounding code available, as pre-processing 
	* wasn't able to be completed
	lcCode = "Structural error in code: " + THIS.cErrorCode
	
CASE THIS.lMacroError = .T.  && typical case
	FOR ii = THIS.nErrorLine - m.tnLines ;
		TO THIS.nErrorLine + m.tnLines
		*
		lcLine = THIS.GetPreProcessedCodeBlock( m.ii, m.ii, .T. )
		IF NOT EMPTY( m.lcLine )
			IF m.tlNoHtml
				lcCode = m.lcCode + m.lcLine + CR
			ELSE
				lcCode = m.lcCode + ;
					IIF( m.ii = THIS.nErrorLine, [<B>], []) + ;
					m.lcLine + ;
					IIF( m.ii = THIS.nErrorLine, [</B>], []) + ;
					[<BR>] + CR
			ENDIF
		ENDIF
	ENDFOR
	IF EMPTY( m.lcCode)
		lcCode = THIS.cErrorCode
	ENDIF
	
OTHERWISE  && Error in CodeBlock itself.
	lcCode = "Error in CodeBlock processing program: " + THIS.cErrorCode
	
ENDCASE

RETURN m.lcCode

ENDFUNC  && GetErrorContext

* ------------------------------------------------ *

PROTECTED FUNCTION FindMatch( tnDepth, tcMatch1, tcMatch2)
*
* Find matching ending line (ENDFOR, ENDIF, etc.).
*
LOCAL ll2Items, llMatch
ll2Items = TYPE( "m.tcMatch2") = "C"

LOCAL lnCount
FOR lnCount = THIS.nDiagCount TO 1 STEP -1
	IF THIS.aDiag[ m.lnCount, 2] = m.tnDepth AND ;
		( THIS.aDiag[ m.lnCount, 3] = m.tcMatch1 OR ;
		m.ll2Items AND THIS.aDiag[ m.lnCount, 3] = m.tcMatch2 )
		*
		llMatch = .T.
		EXIT
	ENDIF
	IF THIS.aDiag[ m.lnCount, 2] < m.tnDepth OR ;
		THIS.aDiag[ m.lnCount, 2] = m.tnDepth AND ;
		THIS.aDiag[ m.lnCount, 3] <> m.tcMatch1 AND ;
		( m.ll2Items = .F. OR THIS.aDiag[ m.lnCount, 3] <> m.tcMatch2 )
		* Nesting error.
		EXIT
	ENDIF
ENDFOR

RETURN m.llMatch

ENDFUNC  && FindMatch
* ------------------------------------------------ *

PROTECTED FUNCTION AddDiagramItem( tnLine, tnDepth, tcType)
*
* Adds a code diagram item to the array at a
* depth count specified in the 2nd parameter. Called
* from PreProcessCode().
*
THIS.nDiagCount = THIS.nDiagCount + 1
DIMENSION THIS.aDiag[ THIS.nDiagCount, 4 ]
THIS.aDiag[ THIS.nDiagCount, 1 ] = m.tnLine
THIS.aDiag[ THIS.nDiagCount, 2 ] = m.tnDepth
THIS.aDiag[ THIS.nDiagCount, 3 ] = m.tcType

ENDFUNC  && AddDiagramItem
* ------------------------------------------------ *

FUNCTION Execute( _0_qnStart, _0_qnEnd )
*
* Execute the codeblock.
*
IF THIS.nRecursionLevel = 0 
	IF TYPE( "m._0_qnStart" ) = "C"  
		* Code passed to method directly -- 
		* supported for backward compatability.
		THIS.ResetProperties()
		THIS.SetCodeBlock( m._0_qnStart )
	ENDIF
	IF NOT THIS.lPreProcessed
		THIS.PreProcessCode()
	ENDIF
	_0_qnStart = 1
	_0_qnEnd = THIS.nCodeLines
ENDIF

PRIVATE _0_qcLine, _0_qcMacroString, _0_qnPointer 
* PRIVATE so Error() can see them!
LOCAL _0_qcUpper, _0_qnSelect, _0_qnChildStart, _0_qnChildEnd, ;
	_0_qnDepth, _0_qlInTrueCase, _0_qlFoundTrueCase, _0_qnDiagItem, ;
	_0_qcString, _0_qnAtPos, _0_qcMerged
	
LOCAL _0_qcForEach, _0_qcForIn, _0_qnForPtr, _0_qnForLinePtr, ;
	_0_qcForDoCmd, _0_qcForObj

* Make variable names as obscure as possible, to minimize possibility
* of conflict with variable name in code to be executed.

FOR _0_qnPointer = m._0_qnStart TO m._0_qnEnd

	_0_qcLine = LTRIM( THIS.aCode[ m._0_qnPointer, 1] )
	_0_qcUpper = UPPER( m._0_qcLine )

	DO CASE
	CASE PADR( m._0_qcUpper, 4) == "TEXT"
		
		* Find matching ENDTEXT and fetch block:
		
		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "TEXT"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF

		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		* Look for ENDTEXT at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDTEXT"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* Pointer moved to ENDTEXT.

		IF THIS.lError = .T.
			* Syntax error found in child block--can't proceed.
			* Should ONLY occur if no ENDTEXT was found.
			EXIT
		ENDIF
		
		* All of TEXT block is stored in 1 record/array element:
		_0_qcString = THIS.aCode[ m._0_qnChildEnd, 1 ]
		
		IF THIS.lAccumulateMergedText
			THIS.ProcessText( m._0_qcString )
		ELSE
			* NOTE: You cannot just use "\" to send entire
			* block, since that command does not handle CR-LFs.
		
			IF THIS.lClassicTextmerge
				THIS.ClassicMerge( m._0_qcString )
			ELSE
				_0_qcString = THIS.InternalMergeText( m._0_qcString )
				* {10/08/1998}:
				IF THIS.lError  && Error merging text, don't write. 
					EXIT
				ENDIF
				IF _TEXT > 0  && Output file open.
					= FWRITE( _TEXT, m._0_qcString )
				ENDIF
				IF CODEBLOCK_SHOW_TEXT_FLAG = .F. OR ;
					THIS.lShowTextmerge = .T.
					* Send string, in case output is going to printer or display.
					* This is most like classic merge behavior.
					IF RIGHT( m._0_qcString, 2) = CR
						* Eliminate final CR, since ? puts one first.
						? LEFT( m._0_qcString, ;
							LEN( m._0_qcString) - 2 )
					ELSE
						? m._0_qcString
					ENDIF
				ENDIF
			ENDIF
		ENDIF
		
		* End CASE: "TEXT"
		
	CASE m._0_qcUpper = "\\"
		* Single textmerge line, no new line.
		
		* Get text that follows \\
		_0_qcString = SUBSTR( m._0_qcLine, 3)
		
		IF THIS.lAccumulateMergedText
			THIS.ProcessText( m._0_qcString )
		ELSE
			IF THIS.lClassicTextmerge
				* Don't call ClassicMerge(), since that is line oriented
				* and would insert a spurious CR-LF.
				_0_qcMacroString = m._0_qcLine
				
				&_0_qcMacroString
			ELSE
				_0_qcString = THIS.InternalMergeText( m._0_qcString)
				* {10/29/1998}:
				IF THIS.lError  && Error merging text, don't write. 
					EXIT
				ENDIF
				IF _TEXT > 0
					= FWRITE( _TEXT, m._0_qcString )
				ENDIF
				IF CODEBLOCK_SHOW_TEXT_FLAG = .F. OR ;
					THIS.lShowTextmerge = .T.
					*
					?? m._0_qcString
				ENDIF
			ENDIF
		ENDIF

	CASE m._0_qcUpper = "\"
		* Single textmerge line, new line.
		
		* Get text that follows \
		_0_qcString = SUBSTR( m._0_qcLine, 2)
		
		IF THIS.lAccumulateMergedText
			THIS.ProcessText( CR + m._0_qcString )
		ELSE
			IF THIS.lClassicTextmerge
				THIS.ClassicMerge( m._0_qcString )
			ELSE
				_0_qcString = THIS.InternalMergeText( m._0_qcString )
				* {10/29/1998}:
				IF THIS.lError  && Error merging text, don't write. 
					EXIT
				ENDIF
				IF _TEXT > 0
					= FWRITE( _TEXT, ;
						CR + m._0_qcString )
				ENDIF
				IF CODEBLOCK_SHOW_TEXT_FLAG = .F. OR ;
					THIS.lShowTextmerge = .T.
					*
					? m._0_qcString
				ENDIF
			ENDIF
		ENDIF

	CASE PADR( m._0_qcUpper, 1) = "="
		* EVAL() is faster than macro sub.
		= EVALUATE( SUBSTR( m._0_qcLine, 2) )

	CASE PADR( m._0_qcUpper, 8) == "DO WHILE"
	
		_0_qcMacroString = SUBSTR( m._0_qcLine, 9)
		* portion of line after DO WHILE
		
		* Find matching ENDDO and make recursive call:
		
		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "DO WHILE"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		* Look for ENDDO at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDDO"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* where to start after child--right after ENDDO
		
		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		* Simulate original DO WHILE block by re-constructing it
		* here with a recursive call to this method for the code
		* inside the nested structure:
		*
		DO WHILE &_0_qcMacroString
			IF _0_qnChildStart <= _0_qnChildEnd
				* Call child code recursively:
				THIS.nRecursionLevel = THIS.nRecursionLevel + 1
				THIS.Result = THIS.Execute( m._0_qnChildStart, m._0_qnChildEnd)
				THIS.nRecursionLevel = THIS.nRecursionLevel - 1
			ENDIF

			IF NOT EMPTY( THIS.cExitCode)
				IF THIS.cExitCode = 'LOOP'
					THIS.cExitCode = SPACE(0)
					LOOP
				ENDIF
				IF THIS.cExitCode = 'EXIT'
					THIS.cExitCode = SPACE(0)
				ENDIF
				EXIT
			ENDIF
		ENDDO
		* End CASE: "DO WHILE"
		
	CASE PADR( m._0_qcUpper, 4) == "SCAN"
	
		_0_qcMacroString = IIF( LEN( m._0_qcLine) = 4, "", SUBSTR( m._0_qcLine, 5))
		* portion of line after SCAN
		
		* Find matching ENDSCAN and make recursive call:
		
		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "SCAN"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		* Look for ENDSCAN at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDSCAN"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* where to start after child--right after ENDSCAN
		
		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		* Simulate original SCAN block by re-constructing it
		* here with a recursive call to this method for the code
		* inside the nested structure:
		*
		SCAN &_0_qcMacroString
			IF _0_qnChildStart <= _0_qnChildEnd
				* Call childcode recursively:
				THIS.nRecursionLevel = THIS.nRecursionLevel + 1
				THIS.Result = THIS.Execute( m._0_qnChildStart, m._0_qnChildEnd)
				THIS.nRecursionLevel = THIS.nRecursionLevel - 1
			ENDIF

			IF NOT EMPTY( THIS.cExitCode)
				IF THIS.cExitCode = 'LOOP'
					THIS.cExitCode = SPACE(0)
					LOOP
				ENDIF
				IF THIS.cExitCode = 'EXIT'
					THIS.cExitCode = SPACE(0)
				ENDIF
				EXIT
			ENDIF
		ENDSCAN
		* End CASE: "SCAN"

	CASE PADR( m._0_qcUpper, 9) == "FOR EACH "
	
		* NOTE: CANNOT handle like other control structures, since
		* VFP has bug that doesn't let you macro expand a FOR EACH command!
		
		_0_qcForEach = ALLTRIM( SUBSTR( m._0_qcUpper, 9))
		* portion of line after FOR EACH
		_0_qnAtPos = AT( SPACE(1), m._0_qcForEach )

		IF m._0_qnAtPos = 0
			THIS.SetError( "FOR EACH syntax error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qcForIn = LTRIM( SUBSTR( m._0_qcForEach, m._0_qnAtPos))
		_0_qcForEach = LEFT( m._0_qcForEach, m._0_qnAtPos - 1)
		IF EMPTY( m._0_qcForIn) OR m._0_qcForIn <> "IN " OR ;
			LEN( m._0_qcForIn) = 3 OR EMPTY( SUBSTR( m._0_qcForIn, 4)) 
			*
			THIS.SetError( "FOR EACH syntax error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qcForIn = LTRIM( SUBSTR( _0_qcForIn, 4))
		
		* Find matching ENDFOR and make recursive call:
		
		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "FOR EACH"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		* Look for ENDFOR at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDFOR"
			*
			* Don't worry about "NEXT"--PreProcess() changes to "ENDFOR".
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* where to start after child--right after ENDFOR
		
		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		* Simulate FOR EACH block by determining whether it is
		* handling (1) an array or (2) an OLE collection, and then
		* translating the inner code:
		*
		* WARNING: Does not support inner nested code at this time!
		* (ie, you cannot put IF's, etc inside FOR EACH..ENDFOR).
		*
		* Is it an array?:
		_0_qlForArray = TYPE( "ALEN(" + _0_qcForIn + ")" ) = "N"
		IF m._0_qlForArray
			_0_qnForLen = ALEN( &_0_qcForIn )
			FOR _0_qnForPtr = 1 TO m._0_qnForLen
				_0_qcForObj = m._0_qcForIn + "[" + LTRIM( STR( _0_qnForPtr)) + "]"
				FOR _0_qnForLinePtr = m._0_qnChildStart TO m._0_qnChildEnd
					* Translate command:
					_0_qcForDoCmd = ForEachTran( THIS.aCode[ m._0_qnForLinePtr, 1 ], ;
						m._0_qcForEach, m._0_qcForObj)
					* Run it:
					&_0_qcForDoCmd
					IF THIS.lError
						EXIT
					ENDIF
				ENDFOR
				IF THIS.lError
					EXIT
				ENDIF
			ENDFOR
		ELSE  && OLE collection
			_0_qnForPtr = 1
			DO WHILE NOT THIS.lError
				_0_qcForObj = m._0_qcForIn + "[" + LTRIM( STR( _0_qnForPtr)) + "]"
				IF TYPE( m._0_qcForObj) <> "O"  && past end of objects
					EXIT
				ENDIF
				FOR _0_qnForLinePtr = m._0_qnChildStart TO m._0_qnChildEnd
					* Translate command:
					_0_qcForDoCmd = THIS.ForEachTran( THIS.aCode[ m._0_qnForLinePtr, 1], ;
						m._0_qcForEach, m._0_qcForObj)
					* Run it:
					&_0_qcForDoCmd
					IF THIS.lError = .T. 
						EXIT
					ENDIF
				ENDFOR
				_0_qnForPtr = m._0_qnForPtr + 1
			ENDDO
		ENDIF

		IF THIS.lError = .T. && Error processing child block--can't proceed.
			EXIT
		ENDIF
		
		* End CASE: "FOR EACH"

	CASE PADR( m._0_qcUpper, 4) == "FOR "
	
		_0_qcMacroString = SUBSTR( m._0_qcLine, 5)
		* portion of line after FOR
		
		* Find matching ENDFOR and make recursive call:
		
		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "FOR"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		* Look for ENDFOR at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDFOR"
			*
			* Don't worry about "NEXT"--PreProcess() changes to "ENDFOR".
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* where to start after child--right after ENDFOR
		
		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		* Simulate original FOR block by re-constructing it
		* here with a recursive call to this method for the code
		* inside the nested structure:
		*
		FOR &_0_qcMacroString
			IF _0_qnChildStart <= _0_qnChildEnd
				* Call childcode recursively:
				THIS.nRecursionLevel = THIS.nRecursionLevel + 1
				THIS.Result = THIS.Execute( m._0_qnChildStart, m._0_qnChildEnd)
				THIS.nRecursionLevel = THIS.nRecursionLevel - 1
			ENDIF

			IF NOT EMPTY( THIS.cExitCode)
				IF THIS.cExitCode = 'LOOP'
					THIS.cExitCode = SPACE(0)
					LOOP
				ENDIF
				IF THIS.cExitCode = 'EXIT'
					THIS.cExitCode = SPACE(0)
				ENDIF
				EXIT
			ENDIF
		ENDFOR
		* End CASE: "FOR"

	CASE PADR( m._0_qcUpper, 3) == "IF " OR PADR( m._0_qcUpper, 3) == "IF("
	
		_0_qlInTrueCase = EVAL( SUBSTR( m._0_qcLine, 3) )
		_0_qlTrueCaseFound = m._0_qlInTrueCase
		
		* Find matching ELSE/ENDIF:

		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "IF"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF

		IF m._0_qlInTrueCase  && IF expression was TRUE
			_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
		ELSE
			_0_qnChildStart = THIS.nCodeLines + 1
		ENDIF
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]

		* Look for ELSE/ENDIF at same depth:
		FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
				EXIT  && Desired item found.
			ENDIF
		ENDFOR

		DO CASE
		CASE m._0_qnDiagItem > THIS.nDiagCount
			THIS.SetError( "Code error: missing ENDIF.", m._0_qcLine, m._0_qnPointer )
			EXIT

		CASE THIS.aDiag[ m._0_qnDiagItem, 3] == "ELSE"

			IF m._0_qlInTrueCase  && IF expression was TRUE
				_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
				_0_qlInTrueCase = .F.
			ELSE
				_0_qlInTrueCase = .T.
				_0_qlTrueCaseFound = .T.
				_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
			ENDIF

			* Now look for ENDIF as well.
			FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
				IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
					EXIT  && Desired item found.
				ENDIF
			ENDFOR
	
			IF m._0_qnDiagItem > THIS.nDiagCount OR ;
				THIS.aDiag[ m._0_qnDiagItem, 3] <> "ENDIF"
				*
				THIS.SetError( "Code error: missing ENDIF.", m._0_qcLine, m._0_qnPointer )
				EXIT
			ENDIF
					
			IF m._0_qlInTrueCase
				_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
			ENDIF
			
		CASE THIS.aDiag[ m._0_qnDiagItem, 3] == "ENDIF"
			* ENDIF, but no ELSE
			_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
			* Note: If IF expr was FALSE, ChildStart will be
			* higher than ChildEnd, so no code will run.
			
		OTHERWISE
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDCASE

		* Point to ENDIF as starting point after recursive call:
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]

		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		IF m._0_qlTrueCaseFound AND ;
			m._0_qnChildStart <= m._0_qnChildEnd
			*
			* Call child code recursively:
			THIS.nRecursionLevel = THIS.nRecursionLevel + 1
			THIS.Result = THIS.Execute( m._0_qnChildStart, m._0_qnChildEnd)
			THIS.nRecursionLevel = THIS.nRecursionLevel - 1
		ENDIF
		
		* End CASE: "IF"

	CASE PADR( m._0_qcUpper, 7) == "DO CASE"
	
		_0_qlInTrueCase = .F.
		_0_qlFoundTrueCase = .F.
		
		* Find matching ENDCASE and make recursive
		* call to this method:

		FOR _0_qnDiagItem = 1 TO THIS.nDiagCount
			IF THIS.aDiag[ m._0_qnDiagItem, 1] = m._0_qnPointer
				EXIT  && Desired item found.
			ENDIF
		ENDFOR
		IF _0_qnDiagItem > THIS.nDiagCount OR ;
			THIS.aDiag[ m._0_qnDiagItem, 3] <> "DO CASE"
			*
			THIS.SetError( "Code diagram error.", m._0_qcLine, m._0_qnPointer )
			EXIT
		ENDIF
		_0_qnDepth = THIS.aDiag[ m._0_qnDiagItem, 2]
		_0_qnChildStart = THIS.nCodeLines + 1

		DO WHILE .T.

			FOR _0_qnDiagItem = m._0_qnDiagItem + 1 TO THIS.nDiagCount
				IF THIS.aDiag[ m._0_qnDiagItem, 2] = m._0_qnDepth
					EXIT  && Desired item found.
				ENDIF
			ENDFOR
	
			DO CASE
			CASE m._0_qnDiagItem > THIS.nDiagCount  && not found
				THIS.SetError( "Code diagram error.", ;
					m._0_qcLine, m._0_qnPointer )
				EXIT
			CASE THIS.aDiag[ m._0_qnDiagItem, 3] == "CASE"
				IF m._0_qlInTrueCase
					_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
					_0_qlInTrueCase = .F.
				ELSE
					IF NOT m._0_qlFoundTrueCase  && no true case yet
						IF EVAL( SUBSTR( LTRIM( ;
							THIS.aCode[ THIS.aDiag[ m._0_qnDiagItem, 1], 1 ] ), 5))  && TRUE case
							_0_qlInTrueCase = .T.
							_0_qlFoundTrueCase = .T.
							_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
						ENDIF
					ENDIF
				ENDIF
			CASE THIS.aDiag[ m._0_qnDiagItem, 3] == "OTHERWISE"
				IF m._0_qlInTrueCase
					_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
					_0_qlInTrueCase = .F.
				ELSE
					IF NOT m._0_qlFoundTrueCase  && no true case--use OTHERWISE
						_0_qlInTrueCase = .T.
						_0_qlFoundTrueCase = .T.
						_0_qnChildStart = THIS.aDiag[ m._0_qnDiagItem, 1] + 1
					ENDIF
				ENDIF
			CASE THIS.aDiag[ m._0_qnDiagItem, 3] == "ENDCASE"
				IF m._0_qlInTrueCase
					_0_qnChildEnd = THIS.aDiag[ m._0_qnDiagItem, 1] - 1
				ENDIF
				EXIT
			OTHERWISE
				THIS.SetError( "Code diagram error.", ;
					m._0_qcLine, m._0_qnPointer )
				EXIT
			ENDCASE
		ENDDO
		
		* Unless error, should be pointing at ENDCASE.
		_0_qnPointer = THIS.aDiag[ m._0_qnDiagItem, 1]
		* where to start after child--right after ENDCASE

		IF THIS.lError = .T. && Error getting child block--can't proceed.
			EXIT
		ENDIF
		
		IF m._0_qlFoundTrueCase = .T. AND ;
			m._0_qnChildStart <= m._0_qnChildEnd

			* Call child code recursively:
			THIS.nRecursionLevel = THIS.nRecursionLevel + 1
			THIS.Result = THIS.Execute( m._0_qnChildStart, m._0_qnChildEnd)
			THIS.nRecursionLevel = THIS.nRecursionLevel - 1
		ENDIF
		
		* End CASE: "CASE"

	CASE PADR( m._0_qcUpper, 4) == "LOOP"
		* This should only occur in recursive call, since only applies
		* within a control structure (DO WHILE, SCAN, or FOR).  Further,
		* cannot macro substitute these lines since child block is
		* called recursively. Instead we create a code indicating how the
		* block was terminated, and exit back to the simulated control
		* structure. Those structures (above) check for cExitCode's of
		* "LOOP" or "EXIT" and act accordingly.
		
		IF THIS.nRecursionLevel = 0
			THIS.SetError( "LOOP statement not within control structure.", ;
				m._0_qcLine )
			EXIT
		ENDIF
		
		THIS.cExitCode = "LOOP"
		EXIT
		
	CASE PADR( m._0_qcUpper, 4) == "EXIT"
		* See comment above on "LOOP"--applies exactly to EXIT, too.
		
		IF THIS.nRecursionLevel = 0
			THIS.SetError( "EXIT statement not within control structure.", ;
				m._0_qcLine )
			EXIT
		ENDIF

		THIS.cExitCode = "EXIT"
		EXIT

	CASE PADR( m._0_qcUpper, 4) == "RETU"
		THIS.cExitCode = "RETURN"
		THIS.Result = .T.  && default
		_0_qnAtPos = AT( SPACE(1), m._0_qcLine )
		
		IF m._0_qnAtPos > 0
			_0_qcMacroString = ALLTRIM( SUBSTR( m._0_qcLine, m._0_qnAtPos ))
			IF NOT EMPTY( m._0_qcMacroString)
				* RETURN <something>
				THIS.Result = EVAL( m._0_qcMacroString )
			ENDIF
		ENDIF


	OTHERWISE
		IF EMPTY( THIS.cExitCode)

			* Normal line of code, just do it:
			_0_qcMacroString = m._0_qcLine
			
			* Special controls when compiled in debugging mode:
			#IF CODEBLOCK_DEBUG
				IF NOT THIS.lDontExecute
					#IF CODEBLOCK_DEBUG_LEVEL = 1
						WAIT WINDOW TIMEOUT CODEBLOCK_WAIT_TIMEOUT m._0_qcMacroString
					#ELSE
						#IF CODEBLOCK_DEBUG_LEVEL = 2
							SUSPEND
						#ENDIF
					#ENDIF
					* Execute the command:
					&_0_qcMacroString
				ENDIF
			#ELSE

			* Execute the command:
			&_0_qcMacroString
			
			#ENDIF
		ENDIF

	ENDCASE  && main CASE for command types

	IF THIS.lError OR NOT EMPTY( THIS.cExitCode)
		* Some error or exit code encountered.
		EXIT
	ENDIF

ENDFOR  && each lone of code

IF THIS.cExitCode == "ERROR"
	THIS.lError = .T.
ENDIF

IF THIS.nRecursionLevel = 0 AND NOT THIS.lError
	* We're about to return from entire block,
	* perform type-check, if specified:
	IF NOT EMPTY( THIS.cReturnType) AND ;
		NOT "X" $ THIS.cReturnType
		*
		* a specific type is required
		IF NOT TYPE( "THIS.Result") $ UPPER( THIS.cReturnType)
			* incompatible type
			THIS.SetError( "RETURN type did not match requirement. + ", ;
				TYPE( "THIS.Result") + " was attempted, when " + ;
				THIS.cReturnType + " was specified." )
				
			* THIS.Result will be set below, due to THIS.lError flag.
		ENDIF
	ENDIF

	* Handle any accumulated merged text:
	IF THIS.lAccumulateMergedText
		THIS.ProcessAccumulatedText()
	ENDIF
ENDIF

IF THIS.lError = .T.
	THIS.Result = THIS.xErrorReturn
ENDIF

RETURN THIS.Result

ENDFUNC  && Execute

* ------------------------------------------------ *

PROTECTED FUNCTION ForEachTran( tcCode, tcEach, tcObj)
*
* Special translation of FOR EACH lines.
*
LOCAL ii, lcRet, llOkChar, lnEachLen, lnLen
lcret = ""
ii = 1
llOkChar = .T.
lnLen = LEN( m.tcCode)
lnEachLen = LEN( m.tcEach)
DO WHILE ii <= m.lnLen
	IF m.llOkChar AND UPPER( SUBSTR( m.tcCode, m.ii)) = m.tcEach
		IF m.ii + m.lnEachLen >= m.lnLen OR ;
			SUBSTR( m.tcCode, m.ii + m.lnEachLen, 1) $ " )]*-+!@#$%^&/<>=."
			*
			lcRet = m.lcRet + m.tcObj
		ELSE
			lcRet = m.lcRet + m.tcEach
		ENDIF
		ii = m.ii + m.lnEachLen
		llOkChar = .F.
	ELSE
		lcRet = m.lcRet + SUBSTR( m.tcCode, m.ii, 1)
		llOkChar = SUBSTR( m.tcCode, m.ii, 1) $ " ([*-+!@#$%^&/<>=."
		ii = m.ii + 1
	ENDIF
ENDDO

RETURN lcRet

ENDFUNC  && ForEachTran
* ------------------------------------------------ *

FUNCTION ProcessText( tcText )
*
* Centralized method to process and TEXT, \, or \\ text. Used
* only when THIS.lAccumulateMergedText is TRUE.
*
* NOTE: Override this method in your subclass this to handle 
* text your own way. You should also override the method
* ProcessAccumulatedText if you do this.
*
LOCAL lcMerged
IF NOT EMPTY( THIS.cAltMergeFunction)
	* User-specified alternative merge function specified.
	lcMerged =EVAL( THIS.cAltMergeFunction + "( m.tcText )" )
ELSE  && We do the merge here.
	lcMerged = THIS.InternalMergeText( m.tcText )
ENDIF

THIS.cMergedText = THIS.cMergedText + m.lcMerged

ENDFUNC  && ProcessText( tcText )

* ------------------------------------------------ *

FUNCTION ProcessAccumulatedText
*
* Override in your subclass to preclude this behavior.
* If you send as-you-go, simply create an empty method.
*
IF TYPE( "THIS.cTextmergeVariable") = "C" AND ;
	NOT EMPTY( THIS.cTextmergeVariable )
	*
	STORE THIS.cMergedText TO ( THIS.cTextmergeVariable )
ENDIF
IF THIS.lWriteAccumulatedText AND ;
	NOT EMPTY( THIS.cMergedText )
	*
	IF _TEXT > 0
		= FWRITE( _TEXT, THIS.cMergedText )
	ENDIF
ENDIF

ENDFUNC  && ProcessAccumulatedText

* ------------------------------------------------ *

FUNCTION SetError( tcMessage, tcCode, tnLine )
*
* Indicate that an error has occurred. This is used when
* preprocessing detects invalid code prior to a VFP error.
* Error number is set to 1089, which can be checked by
* calling routine.
*
THIS.lError = .T.
THIS.cExitCode = "ERROR"  && need this to exit recursive calls

THIS.nError = 1089 && user-defined
THIS.nErrorRecursionLevel = THIS.nRecursionLevel

IF TYPE( "m.tcMessage") == "C"
	THIS.cErrorMessage = m.tcMessage
ENDIF

IF TYPE( "m.tcCode") == "C"
	THIS.cErrorCode = m.tcCode
ENDIF

IF TYPE( "m.tnLine") == "N"
	THIS.nErrorLine = m.tnLine
	IF m.tnLine > 0 AND TYPE( "THIS.aCode[ m.tnLine, 2]") = "N"
		THIS.nRawErrorLine = THIS.aCode[ m.tnLine, 2]
	ENDIF
ENDIF

#IF CODEBLOCK_WAIT_ON_ERROR
	WAIT WINDOW TIMEOUT CODEBLOCK_WAIT_TIMEOUT THIS.cErrorMessage
#ENDIF

ENDFUNC  && SetError
* ------------------------------------------------------------- *

#IF NOT CODEBLOCK_DEBUG
* If debugging, let VFP error handler take over.

FUNCTION Error( tnError, tcMethod, tnLine)
*
* Codeblock error handler.
*
THIS.lError = .T.
THIS.cExitCode = "ERROR"  && need this to exit recursive calls

THIS.nError = m.tnError
THIS.nErrorRecursionLevel = THIS.nRecursionLevel
THIS.cErrorMessage = MESSAGE()

THIS.cErrorCode = MESSAGE( 1)
THIS.cErrorMethod = m.tcMethod
THIS.nErrorLine = m.tnLine
IF m.tnLine > 0 AND TYPE( "THIS.aCode[ m.tnLine, 2]") = "N"
	THIS.nRawErrorLine = THIS.aCode[ m.tnLine, 2]
ENDIF

*!*	IF "&_0_QCMACRO" $ UPPER( THIS.cErrorCode) 
*!*		THIS.lMacroError = .T.
*!*		IF TYPE( "m._0_qcLine") = "C"
*!*			* Show command that was attempted instead of & command:
*!*			THIS.cErrorCode = _0_qcLine
*!*		ENDIF
*!*		IF TYPE( "m._0_qnPointer") = "N"
*!*			* Show command that was attempted instead of & command:
*!*			THIS.nErrorLine = _0_qnPointer
*!*		ENDIF
*!*	ELSE
*!*		THIS.lMacroError = .F.
*!*	ENDIF

* {10/08/1998} 
* Changed above IF..ENDIF to the CASE code below to handle additional
* errors found when EVAL'ing CASE and IF expressions
* in user's code.
DO CASE
CASE '&_0_QCMACRO' $ UPPER( THIS.cErrorCode) 
	THIS.lMacroError = .T.
	IF TYPE( 'm._0_qcLine') = 'C'
		* Show command that was attempted instead of & command:
		THIS.cErrorCode = _0_qcLine
	ENDIF
	IF TYPE( 'm._0_qnPointer') = 'N'
		* Show command that was attempted instead of & command:
		THIS.nErrorLine = _0_qnPointer
		IF TYPE( "THIS.aCode[ m._0_qnPointer, 2]") = "N"
			THIS.nRawErrorLine = THIS.aCode[ m._0_qnPointer, 2]
		ENDIF
	ENDIF
CASE '_0_QLINTRUECASE =' $ UPPER( THIS.cErrorCode)
	THIS.lMacroError = .T.
	IF TYPE( 'm._0_qcLine') = 'C'
		* Show command that was attempted instead of & command:
		THIS.cErrorCode = _0_qcLine
	ENDIF
	IF TYPE( 'm._0_qnPointer') = 'N'
		* Show command that was attempted instead of & command:
		THIS.nErrorLine = _0_qnPointer
		IF TYPE( "THIS.aCode[ m._0_qnPointer, 2]") = "N"
			THIS.nRawErrorLine = THIS.aCode[ m._0_qnPointer, 2]
		ENDIF
	ENDIF
OTHERWISE
	THIS.lMacroError = .F.
ENDCASE

ENDFUNC  && Error

#ENDIF  && not debugging

* ------------------------------------------------------------- *

PROTECTED FUNCTION IllegalCommandFound( tcCommand)
*
* Scans array THIS.aIllegal to see if current command is disallowed.
*
* Array is built via methods AddIllegalCommand() and 
* DropIllegalCommand(). In the base class for cusCodeBlock, QUIT and
* CANCEL are disallowed. It is easy to subclass this behaviour.

LOCAL ii, llFound
llFound = .F.

FOR ii = 1 TO ALEN( THIS.aIllegal, 1)
	IF NOT EMPTY( THIS.aIllegal[ m.ii, 1])
		IF PADR( m.tcCommand, LEN( THIS.aIllegal[ m.ii, 1])) ;
			== THIS.aIllegal[ m.ii, 1]
			*
			llFound = .T.
			EXIT
		ENDIF
	ENDIF
ENDFOR

RETURN m.llFound  && TRUE == BAD!

ENDFUNC  && IllegalCommandFound
* ------------------------------------------------------------- *
FUNCTION AddIllegalCommand( tcCommand)

LOCAL ii
FOR ii = 1 TO ALEN( THIS.aIllegal, 1)
	IF EMPTY( THIS.aIllegal[ m.ii, 1])
		EXIT
	ENDIF
ENDFOR
IF m.ii > ALEN( THIS.aIllegal, 1)
	DIMENSION THIS.aIllegal[ m.ii, 1]
ENDIF
THIS.aIllegal[ m.ii, 1] = UPPER( m.tcCommand)

ENDFUNC  && AddIllegalCommand
* ------------------------------------------------------------- *
FUNCTION DropIllegalCommand( tcCommand )

LOCAL ii
FOR ii = 1 TO ALEN( THIS.aIllegal, 1)
	IF NOT EMPTY( THIS.aIllegal[ m.ii, 1] ) AND ;
		THIS.aIllegal[ m.ii, 1] == UPPER( m.tcCommand)
		*
		THIS.aIllegal[ m.ii, 1] = .F.
		EXIT
	ENDIF
ENDFOR

ENDFUNC  && DropIllegalCommand
* ------------------------------------------------------------- *

FUNCTION InternalMergeText( tcText )
*
* Basic merge function to be used if no
* user-specified function is given, and if
* "classic merge" is not on.
*
IF NOT SET( 'TEXTMERGE') = "ON"
	* We're not merging, so just return text.
	RETURN m.tcText
ENDIF

LOCAL lcText, lcOut, lcDelim_1, lcDelim_2, lnLen, lnAt1, lnAt2, ;
	lcType, lxValue, lcSubStr
	
lcText = m.tcText
lcOut = ""

* Determine current delimiters:
lcDelim_1 = SET( 'TEXTMERGE', 1)
lnLen = INT( LEN( m.lcDelim_1) / 2)
* Good luck if you have unequal length delimiters!
lcDelim_2 = SUBSTR( m.lcDelim_1, m.lnLen + 1)
lcDelim_1 = LEFT( m.lcDelim_1, m.lnLen )

*DO WHILE .T.
DO WHILE NOT THIS.lError
	* In case an error gets set while handling a line within a 
	* larger text block.

	lnAt1 = AT( m.lcDelim_1, m.lcText)  && Position of open delimiter.

	IF m.lnAt1 > 0 AND ;
		LEN( m.lcDelim_1) + m.lnAt1 < LEN( m.lcText )
		* 
		* Open delimiter found and there's room for close delimiter. Look for one,
		* starting after open delimiter:
		lnAt2 = AT( m.lcDelim_2, SUBSTR( m.lcText, m.lnAt1 + LEN( m.lcDelim_1)))
		
		IF m.lnAt2 > 0
			* Close delimiter found.
			IF m.lnAt1 > 1		
				* If open delimiter is not at beginning of text, 
				* output everything that comes before it.
				lcOut = m.lcOut + LEFT( m.lcText, m.lnAt1 - 1)
			ENDIF
			* Now EVALUATE the text between delimiters, and apply
			* TRANSFORM as a simple type conversion, if necessary. For more
			* precise conversion, code block should contain specific character
			* expressions. [One problem here is difference in how this handles
			* DECIMALS, compared to classic textmerge.]
			*
			lcSubStr = SUBSTR( m.lcText, m.lnAt1 + LEN( m.lcDelim_1), m.lnAt2 - 1)
			lxValue = EVAL( m.lcSubStr )
			IF THIS.lError
				* Error occurred during EVAL(). Provide a different
				* set of error stats:
				THIS.cErrorCode = m.lcSubStr
				THIS.nErrorLine = 0  && can't really tell w/ TEXT, so don't confuse user
				THIS.nRawErrorLine = 0
				THIS.lMacroError = .T.
				EXIT
			ENDIF
			lcType = TYPE( "m.lxValue" )
			* Note: We didn't check TYPE() before EVAL(), since that could 
			* cause a method or function referenced by the expression to be
			* called twice!
			IF m.lcType = "U"
				lcOut = m.lcOut + "[Script Error: " + m.lcSubStr + "]"
			ELSE
				DO CASE
				CASE m.lcType = "C"
					lcOut = m.lcOut + m.lxValue
				OTHERWISE
					lcOut = m.lcOut + TRANSFORM( m.lxValue, "")
				ENDCASE
			ENDIF
			
			IF m.lnAt1 + m.lnAt2 + LEN( m.lcDelim_1 + m.lcDelim_2) - 2 >= LEN( m.lcText)
				* Nothing more after the close delimiter.
				EXIT
			ELSE
				* Strip off everything we just processed in preparation
				* for re-entrance to DO loop.
				lcText = SUBSTR( m.lcText, m.lnAt1 + m.lnAt2 + ;
					LEN( m.lcDelim_1 + m.lcDelim_2) - 1 )
			ENDIF
		ELSE  && No matching close delimiter found.
			* Just send on everything and exit.
			lcOut = m.lcOut + m.lcText
			EXIT
		ENDIF
	ELSE  && No candidate delimiter pair found.
		* Just send on everything and exit.
		lcOut = m.lcOut + m.lcText
		EXIT
	ENDIF
ENDDO

RETURN m.lcOut

ENDFUNC  && InternalMergeText

* ------------------------------------------------------------- *

FUNCTION ClassicMerge( _0_qcText )
*
* Attempt to mimic VFP regular textmerge behavior by just
* using the "\" command on each line in the block. This will
* produce most consistent formatting at the cost of performance. 
*
LOCAL _0_qcMacroString, _0_qnMline, _0_qnMemo, _0_qnCounter

_0_qnMemo = SET("MEMOWIDTH")
SET MEMOWIDTH TO 254

_0_qnMline = _MLINE
_MLINE = 0

FOR _0_qnCounter = 1 TO MEMLINES( m._0_qcText)
	* Use TEXTMERGE output command "\" to output each
	* line that existed between TEXT and ENDTEXT:
	_0_qcMacroString = "\" + MLINE( THIS.cChildBlock, 1, _MLINE )
	&_0_qcMacroString
ENDFOR

_MLINE = m._0_qnMline
SET MEMOWIDTH TO m._0_qnMemo

ENDFUNC  && ClassicMerge

* ------------------------------------------------------------- *

FUNCTION GetMergedText( )
*
* Gets corresponding property.
*
RETURN THIS.cMergedText

ENDFUNC  && GetMergedText
* ------------------------------------------------------------- *

FUNCTION SetTextmergeVariable( tcVarName)
*
* Sets property "cTextmergeVariable".
*
IF TYPE( "m.tcVarname") == "C" AND NOT EMPTY( m.tcVarName )
	THIS.cTextmergeVariable = m.tcVarName
ELSE
	THIS.cTextmergeVariable = NULL
ENDIF

ENDFUNC  && SetTextmergeVariable
* ------------------------------------------------------------- *

FUNCTION SetAltMergeFunction( tcFuncName)
*
* Set alternative (external) text merge function.
*
IF TYPE( "m.tcFuncName") = "C" AND NOT EMPTY( m.tcFuncName)
	THIS.cAltMergeFunction = ALLTRIM( ;
		STRTRAN( STRTRAN( STRTRAN( STRTRAN( m.tcFuncName, ;
			"(", ""), ")", ""), "[", ""), "]", "") )
ELSE
	THIS.cAltMergeFunction = NULL
ENDIF

ENDFUNC  && SetAltMergeFunction
* ------------------------------------------------------------- *

FUNCTION SetCodeBlock( tcCodeBlock)

IF TYPE( "m.tcCodeBlock") $ "CM"
	THIS.cCodeBlock = m.tcCodeBlock
ELSE
	THIS.lError = .T.
ENDIF

ENDFUNC  && SetCodeBlock

* ------------------------------------------------------------- *

FUNCTION SetScript( tcScript)
THIS.SetCodeBlock( m.tcScript)
THIS.lScript = .T.
ENDFUNC

* ------------------------------------------------------------- *

FUNCTION SetCodeBlockFromFile( tcFileSpec)
*
* Get file contents.
*
IF TYPE( "m.tcFileSpec") <> "C"
	THIS.cCodeBlock = ""
	THIS.SetError( "No file name passed to SetCodeBlockFromFile() method." )
	RETURN
ENDIF
	
IF NOT FILE( m.tcFileSpec)
	THIS.cCodeBlock = ""
	THIS.SetError( "File " + m.tcFileSpec + " not found." )
	RETURN
ENDIF

THIS.cCodeBlock = THIS.FileToStr( m.tcFileSpec )

ENDFUNC  && SetCodeBlockFromFile

* ------------------------------------------------------------- *

FUNCTION FileToStr( tcFileSpec )

LOCAL lnSelect, lcReturn
lnSelect = SELECT()
SELECT 0
CREATE CURSOR _0_qFile (Contents M)
APPEND BLANK
APPEND MEMO Contents FROM ( m.tcFileSpec)
lcReturn = Contents
USE
SELECT ( m.lnSelect)

RETURN m.lcReturn

ENDFUNC  && FileToStr

* ------------------------------------------------------------- *

FUNCTION SetScriptFromFile( tcFileSpec)
THIS.SetCodeBlockFromFile( m.tcFileSpec)
THIS.lScript = .T.
ENDFUNC

* ------------------------------------------------------------- *

FUNCTION SetReturnType( tcType)
*
* Sets property "cReturnType".
*
IF TYPE( "m.tcType") == "C"
	THIS.cReturnType = UPPER( ALLTRIM( m.tcType))
ENDIF

ENDFUNC  && SetReturnType
* ------------------------------------------------------------- *

FUNCTION SetErrorReturnValue( txValue)
*
* Sets property "xErrorReturn".
*
IF TYPE( "m.txValue") == "U"
	THIS.xErrorReturn = .F. && default
ELSE
	THIS.xErrorReturn = m.txValue
ENDIF

ENDFUNC  && SetErrorReturnValue
* ------------------------------------------------------------- *

FUNCTION SetCallingObject( toCaller)
*
* Sets property "oCaller".
*
IF TYPE( "m.toCaller") == "O" AND NOT ISNULL( m.toCaller )
	THIS.oCaller = m.toCaller
ELSE
	THIS.oCaller = NULL
ENDIF

ENDFUNC  && SetCallingObject
* ------------------------------------------------ *

*** ------------------------------------------- ***
ENDDEFINE  && CLASS CusCodeBlock
*** ------------------------------------------- ***

*** ------------------------------------------- ***
DEFINE CLASS RunVFPScript AS CusCodeBlock
*** ------------------------------------------- ***
* Sub-class of CodeBlock setup for WCS script maps.
* REQUIRES: Existence of an external PRIVATE variable
* named RESPONSE, so to support Response.Write() syntax.
*
lAccumulateMergedText = .T.
lWriteAccumulatedText = .F.

*!*	* Override native CodeBlock behavior:
FUNCTION ProcessText( tcText )
SET TEXTMERGE ON
LOCAL lcMerged
lcMerged = THIS.InternalMergeText( m.tcText )
* CRITICAL: Send the text right away with Response.Write,
*  so it will merge with any calls the user makes to same fn.
Response.Write( m.lcMerged )
ENDFUNC

FUNCTION ProcessAccumulatedText
* override this method, since we've been writing as we go
* we don't have any task to perform at the end
ENDFUNC

*** ------------------------------------------- ***
ENDDEFINE  && CLASS RunVFPScript
*** ------------------------------------------- ***

**** <<<<< --- CUT HERE (END)

*** ------------------------------------------- ***
DEFINE CLASS FrmCodeBlockEditor AS FORM
*** ------------------------------------------- ***
* Codeblock Editor Form
*
* Handy for testing code blocks. How to call:
*
* 1) DO CodeBlck  (with no arguments)
*
* 2) DO CodeBlck WITH <parm1>, <parm2>, .T.
*
* 3) Call class directly.

	nStartTime = 0
	nEndTime = 0
	
	* Why don't these 2 produce a size box???
	BorderStyle = 3
	SizeBox = .T.
	
	Height = 247
	Width = 428
	DoCreate = .T.
	ShowTips = .T.
	AutoCenter = .T.
	Caption = "Code Block Editor"
	MinHeight = 160
	MinWidth = 300
	WhatsThisButton = .F.
	*-- Width (in pixels) of command buttons (default = 84).
	commandbuttonwidth = 84
	*-- Height (in pixels) of command buttons (default = 25).
	commandbuttonheight = 25
	*-- Pixels between objects (default 8).
	objectspacing = 8
	*-- Object reference to codeblock execution module.
	codeblockobject = .NULL.
	Name = "frmcodeblockeditor"


	ADD OBJECT edtcodeblock AS editbox WITH ;
		FontName = "Courier New", ;
		FontSize = 11, ;
		AllowTabs = .T., ;
		Height = 132, ;
		Left = 12, ;
		TabIndex = 1, ;
		ToolTipText = "Enter Visual FoxPro code and press execute.", ;
		Top = 24, ;
		Width = 384, ;
		Name = "edtCodeBlock"


	ADD OBJECT cmdcancel AS commandbutton WITH ;
		Top = 204, ;
		Left = 216, ;
		Height = 25, ;
		Width = 84, ;
		Cancel = .T., ;
		Caption = "\<Cancel", ;
		TabIndex = 4, ;
		ToolTipText = "Cancel without executing any code.", ;
		Name = "cmdCancel"


	ADD OBJECT cmdfont AS commandbutton WITH ;
		Top = 204, ;
		Left = 324, ;
		Height = 25, ;
		Width = 84, ;
		Caption = "\<Font...", ;
		TabIndex = 5, ;
		ToolTipText = "Change the font for the code block.", ;
		Name = "cmdFont"


	ADD OBJECT cmdverify AS commandbutton WITH ;
		Top = 204, ;
		Left = 120, ;
		Height = 25, ;
		Width = 84, ;
		Caption = "\<Verify Code", ;
		TabIndex = 3, ;
		ToolTipText = "Attempt to compile code block to check for syntax errors.", ;
		Name = "cmdVerify"


	ADD OBJECT cmdexecute AS commandbutton WITH ;
		Top = 204, ;
		Left = 24, ;
		Height = 25, ;
		Width = 84, ;
		Caption = "\<Execute", ;
		Default = .T., ;
		TabIndex = 2, ;
		ToolTipText = "Run the above code using the CodeBlock() runtime interpreter.", ;
		Name = "cmdExecute"


	ADD OBJECT lblcodeblock AS label WITH ;
		Caption = "Visual FoxPro Code Block:", ;
		Height = 17, ;
		Left = 5, ;
		Top = 3, ;
		Width = 151, ;
		Name = "lblCodeBlock"

	ADD OBJECT lblTiming AS label WITH ;
		Caption = "Time (sec): 0.0", ;
		Height = 17, ;
		Left = 190, ;
		Top = 3, ;
		Width = 100, ;
		Name = "lblTiming"

	PROCEDURE Show
		LPARAMETERS nStyle

		LOCAL lnSpace

		* Disable screen refresh until we figure out everything:
		THIS.LockScreen = .T.

		* Grow/shrink edit region to use as much space as possible:
		THIS.edtCodeBlock.Width = THIS.Width - 2 * THIS.ObjectSpacing

		THIS.edtCodeBlock.Height = ;
			THIS.Height - ;
			THIS.edtCodeBlock.Top - ;
			THIS.CommandButtonHeight - ;
			2 * THIS.ObjectSpacing

		* Position command buttons vertically:
		THIS.cmdExecute.Top = THIS.Height - THIS.CommandButtonHeight - THIS.ObjectSpacing
		THIS.cmdVerify.Top = THIS.Height - THIS.CommandButtonHeight - THIS.ObjectSpacing
		THIS.cmdCancel.Top = THIS.Height - THIS.CommandButtonHeight - THIS.ObjectSpacing
		THIS.cmdFont.Top = THIS.Height - THIS.CommandButtonHeight - THIS.ObjectSpacing

		* Position command buttons horizontally:
		lnSpace = INT( (THIS.Width - 4 * THIS.CommandButtonWidth) / 5)
		THIS.cmdExecute.Left = m.lnSpace
		THIS.cmdVerify.Left = m.lnSpace * 2 + THIS.CommandButtonWidth * 1
		THIS.cmdCancel.Left = m.lnSpace * 3 + THIS.CommandButtonWidth * 2
		THIS.cmdFont.Left = m.lnSpace * 4 + THIS.CommandButtonWidth * 3

		* Show re-painted screen:
		THIS.LockScreen = .F.
	ENDPROC


	PROCEDURE Resize
		THIS.Show()
	ENDPROC


	PROCEDURE Init
		THIS.CommandButtonHeight = MAX( 20, THIS.CommandButtonHeight)
		THIS.CommandButtonWidth = MAX( 64, THIS.CommandButtonWidth)
		THIS.ObjectSpacing = MAX( THIS.ObjectSpacing, 2)

		THIS.MinWidth = 4 * THIS.CommandButtonWidth + 5 * THIS.ObjectSpacing

		THIS.cmdExecute.Height = THIS.CommandButtonHeight
		THIS.cmdExecute.Width = THIS.CommandButtonWidth

		THIS.cmdVerify.Height = THIS.CommandButtonHeight
		THIS.cmdVerify.Width = THIS.CommandButtonWidth

		THIS.cmdCancel.Height = THIS.CommandButtonHeight
		THIS.cmdCancel.Width = THIS.CommandButtonWidth

		THIS.cmdFont.Height = THIS.CommandButtonHeight
		THIS.cmdFont.Width = THIS.CommandButtonWidth

		THIS.edtCodeBlock.Left = THIS.ObjectSpacing
		THIS.edtCodeBlock.Top = ;
			THIS.lblCodeBlock.Top + THIS.lblCodeBlock.Height + 4
	ENDPROC


	PROCEDURE lblTiming.Refresh
		THIS.Caption = "Time (sec): " + STR( THISFORM.nEndTime - THISFORM.nStartTime, 10, 4)
	ENDPROC
	  
	PROCEDURE cmdcancel.Click
		THISFORM.Release()
	ENDPROC


	PROCEDURE cmdfont.Click
		LOCAL lcGetFont, lnAt1, lnAt2
		lcGetFont = GETFONT()
		IF NOT EMPTY( m.lcGetFont)
			lnAt1 = AT( ",", m.lcGetFont)
			lnAt2 = RAT( ",", m.lcGetFont)
			THISFORM.LockScreen = .T.
			THISFORM.edtCodeBlock.FontName = ;
				LEFT( m.lcGetFont, m.lnAt1 - 1)
			THISFORM.edtCodeBlock.FontSize = ;
				VAL( SUBSTR( m.lcGetFont, m.lnAt1 + 1, ;
				m.lnAt2 - m.lnAt1 - 1))
			THISFORM.edtCodeBlock.FontBold = ;
				"B" $ SUBSTR( m.lcGetFont, m.lnAt2 + 1)
			THISFORM.edtCodeBlock.FontItalic = ;
				"I" $ SUBSTR( m.lcGetFont, m.lnAt2 + 1)
			THISFORM.LockScreen = .F.
		ENDIF
	ENDPROC


	PROCEDURE cmdverify.Click
		IF EMPTY( THISFORM.edtCodeBlock.Value)
			?? CHR(7)
			= MessageBox( "Nothing to verify!")
		ELSE
			LOCAL lcFile, lnSelect, lcDir
			lcDir = ADDBS( SYS(2023))
			lcFile = "C" + RIGHT(SYS(3), 7)
			lnSelect = SELECT()
			CREATE CURSOR (m.lcFile) (Code M)
			SELECT (m.lcFile)
			INSERT INTO (m.lcFile) VALUES ( ;
				IIF( EMPTY( THISFORM.edtCodeBlock.SelText), ;
					THISFORM.edtCodeBlock.Value, ;
					THISFORM.edtCodeBlock.SelText) )
			COPY MEMO Code TO (m.lcDir + m.lcFile + ".PRG")
			USE IN (m.lcFile)
			COMPILE (m.lcDir + m.lcFile + ".PRG")
			IF FILE(m.lcDir + m.lcFile + ".ERR")
				MODIFY COMMAND (m.lcDir + m.lcFile + ".ERR")
				ERASE (m.lcDir + m.lcFile + ".ERR")
			ELSE
				= MessageBox( "No compilation errors.")
			ENDIF
			IF FILE(m.lcDir + m.lcFile + ".PRG")
				ERASE (m.lcDir + m.lcFile + ".PRG")
			ENDIF
			IF FILE(m.lcDir + m.lcFile + ".FXP")
				ERASE (m.lcDir + m.lcFile + ".FXP")
			ENDIF
			SELECT (m.lnSelect)
		ENDIF
	ENDPROC


	PROCEDURE cmdexecute.Click
		LOCAL lcRetVal, loCb, lcTmpResult, llToScreen

		IF EMPTY( THISFORM.edtCodeBlock.Value)
			?? CHR(7)
			= MessageBox( "Nothing to execute!")
		ELSE
			THISFORM.CodeBlockObject = CREATE( "cusCodeBlock", , THISFORM )

			llToScreen = UPPER( THISFORM.Name) == WOUTPUT()
			IF m.llToScreen
				ACTIVATE SCREEN
			ENDIF
			
			loCb = THISFORM.CodeBlockObject
			loCb.SetCodeBlock( ;
				IIF( EMPTY( THISFORM.edtCodeBlock.SelText), ;
					THISFORM.edtCodeBlock.Value, ;
					THISFORM.edtCodeBlock.SelText) )

			THISFORM.nStartTime = SECONDS()
			lcTmpResult = loCb.Execute()
			THISFORM.nEndTime = SECONDS()
			
			IF loCb.lError
				= MessageBox( "CAUTION: Error " + LTRIM( STR( loCb.nError)) + ;
					" occurred with message " + loCb.cErrorMessage + ;
					", while processing line " + loCb.cErrorCode + ".")
			ELSE
				DO CASE
				CASE TYPE( "m.lcTmpResult") = "L"
					lcRetVal = IIF( m.lcTmpResult, ".T.", ".F.")
				CASE TYPE( "m.lcTmpResult") = "C"
					lcRetVal = m.lcTmpResult
				CASE TYPE( "m.lcTmpResult") = "D"
					lcRetVal = DTOC( m.lcTmpResult)
				OTHERWISE
					lcRetVal = [TYPE "] + TYPE( "m.lcTmpResult") + ["]
				ENDCASE

				= MessageBox( "Code block executed successfully!" + ;
					" RETURN value was: " + m.lcRetVal )
			ENDIF

			RELEASE loCb
			THISFORM.CodeBlockObject = .NULL.
			
			IF m.llToScreen
				ACTIVATE WINDOW ( THISFORM.Name )
			ENDIF
			THISFORM.lblTiming.Refresh
			THISFORM.edtCodeBlock.SetFocus
		ENDIF
	ENDPROC


ENDDEFINE
*
*-- EndDefine: FrmCodeBlockEditor
**************************************************

*** >>>>> "Theory, Record of Revision, and Special Notes" <<<<<
***        =============================================

* CODEBLCK.PRG : Code Block Class
*
* Author: Randy Pearson ( RandyP@cycla.com )
* Public Domain
* Purpose: Executes un-compiled, structured FoxPro code 
*   at runtime.
*
* General Strategy:
*   If code were all simple "in-line" code (no SCAN, DO, etc.),
*   it is easy to run uncompiled by simply storing each line to
*   a memory variable and then macro substituting each line.
*
*   Thus, we adopt that approach, but when we encounter any control
*   structure, we capture the actual code block within the structure,
*   create an artificial simulation of the structure, and pass the
*   internal code block recursively to this same routine.  Nesting
*   is handled automatically by this approach.
*
* Limitations: 
*   - Does not support embedded subroutines (PROC or FUNC)
*     within code passed.
*     Performs implied RETURN .T. if subroutine found.
*     To use UDF's, capture each PROC/FUNC as its own code
*     block and call CODEBLCK repeatedly as needed.
*   - Doesn't accept TEXT w/ no ENDTEXT (although FP doc's
*     suggest that this is acceptable to FP.
*   - Limited support FOR EACH..ENDFOR. No nested code between
*     FOR EACH..ENDFOR will work.
*

* ------------------------------------------------------------- *
* KEY NOTES about RETURN VALUES <-- IMPORTANT, PLEASE READ!!
*
*!*	The class can *return* a value if you call the Execute() method
*!*	as a function. The actual value returned is dependent on 3 factors:

*!*	1. did an error occur running the code?
*!*	2. did you specify a special return value to use in error conditions?
*!*	3. did the code itself include a RETURN statement?

*!*	IF an error occurred
*!*		IF you specified an error return value (property xErrorReturn)
*!*			RETURN THIS.xErrorReturn
*!*		ELSE
*!*			RETURN .F.
*!*		ENDIF
*!*	ELSE no error occurred
*!*		IF code included a RETURN statement that was executed
*!*			RETURN that value (or .T. if a plain "RETURN")
*!*		ELSE 
*!*			RETURN .T.
*!*		ENDIF
*!*	ENDIF

* ------------------------------------------------------------- *

* I - MINIMUM CALLING SYNTAX EXAMPLE
* ==================================
** SETUP: Ensure this class is available in memory. If using this file
**   natively you must:
**
**   SET PROCEDURE TO CodeBlck.PRG ADDITIVE
**
** If using the wwEval class with West-Wind Web Connection, all you
** need is for WWEVAL.PRG to be included (which it is by default).

*!*	LOCAL oCode
*!*	oCode = CREATEOBJECT( "cusCodeBlock")
*!*	* Assumes this class is already loaded into memory via
*!*	* SET PROCEDURE or equivalent.

*!*	oCode.SetCodeBlock( MyTable.MyMemoField )
*!*	* Whatever you pass has structured FoxPro code.
*!*   * This could also be a memory variable.

*!*	oCode.Execute()
*!*   RELEASE oCode

* II - DEALING WITH ERRORS IN THE CODE
* ====================================
* To pre-establish the RETURN value should an error arise,
* use the "xErrorReturn" Property:

*!*	LOCAL oCode, nResult
*!*	oCode = CREATEOBJECT( "cusCodeBlock")
*!*	oCode.xErrorReturn = -1  && can be any type
*!*	oCode.SetCodeBlock( MyTable.MyMemoField )
*!*	nResult = oCode.Execute()
*!*	RELEASE oCode
*!*	IF nResult = -1
*!*		= MessageBox( ...
*!*	ENDIF

* Instead, you could also just check to see if an error occurred by 
* looking at a few properties afterward:

*!*	LOCAL oCode
*!*	oCode = CREATEOBJECT( "cusCodeBlock")
*!*	oCode.SetCodeBlock( MyTable.MyMemoField )
*!*	oCode.Execute()
*!*	IF oCode.lError
*!*		= MessageBox( oCode.GetErrorContext( 2, .T. ))
*!*	ENDIF
*!*	RELEASE oCode

* Of course, combinations of the above can be used.

* III - PREVENTING ACCESS TO DANGEROUS COMMANDS 
* =============================================

* There are some commands I won't even try to let the program 
* execute. These include CLEAR ALL and CLOSE ALL. There are also
* some commands that you normally won't want to provide access to,
* except in certain circumstances. The *default* ones that are
* precluded are:
*
*  QUIT
*  CANCEL
*
* To ADD another command to the disallowed list, use the 
* AddIllegalCommand() Method, e.g.:
*
*!*	* Before calling the Execute() method:
*!*	oCode.AddIllegalCommand( "ZAP")

* To instead DROP one of the restrictions that I have added (or that you
* have added to the Init() method in a sub-class), use the
* DropIllegalCommand() Method, e.g.:

*!*	* Before calling the Execute() method:
*!*	oCode.DropIllegalCommand( "QUIT")

* IV - CODE CONTAINED IN FILES
* ============================
* If your code is in a file (perhaps a non-compiled PRG), you 
* can either extract it yourself to a memo field or variable;
* or you can use the Method: SetCodeBlockFromFile( pcFileSpec).

* V - SPECIAL "TESTING" MODE
* ==========================
* NOTE: Applies only if compiled with CODEBLOCK_DEBUG = .T.  !!
*
* You can test the structure of your code without executing any
* non-structured code (i.e, all the stuff between DO..ENDDO,
* FOR..ENDFOR, etc., by setting the following property before
* calling the Execute() method:
*
*!*	oCode.lDontExecute = .T.

* VI - OPTIONS FOR TEXTMERGE PROCESSING
* =====================================
* Applications that involve significant TEXTMERGE operations can be very slow
* if we simply output one line at a time using macro substitution. In addition,
* some applications would prefer that we simply "accumulate" all merged output
* to be handled in bulk after the code block is processed. Further, some
* applications would rather use their own merge function. These needs are 
* addressed by offering several options as discussed below:
*
*    (a) Property "lAccumulateMergedText", is the key flag, which causes ALL 
*        merged text to be accumulated for final disposition at end of code 
*        block execution. Set to TRUE for any of the advanced operations
*        listed above. (Default = .T., accumulate.)
*
*        Example:  oCode.lAccumulateMergedText = .T.
*
*    (b) Property "cAltMergeFunction" allows you to specify your
*        own function to process textmerge strings. Your function is 
*        responsible for anything within textmerge delimiters, if applicable.
*        Set this with SetAltMergeFunction(). Applies only when accumulate
*        flag is set.
*
*        Example:  oCode.SetAltMergeFunction( "MyMerge" )
*
*        NOTE: You can also sub-class and override InternalMergeText()
*
*    (c) Property "cMergedText" is used internally to accumulate the text. Use
*        METHOD GetMergedText() to obtain this text from your calling routine after
*        the code block is processed. Always returns "" unless lAccumulateMergeText 
*        is TRUE.
*
*    (d) Property "cTextmergeVariable" allows you to specify a variable name (or
*        external object property name)
*        to which the accumulated text should be stored
*        after the code block is processed. Only applies when the accumulation
*        flag is set. Your calling routine is responsible for declaring the 
*        variable as PRIVATE or PUBLIC before executing the code block. Set 
*        this with SetTextmergeVariable().
*
*        Example:  PRIVATE lcTextOut
*                  oCode.SetTextmergeVariable( "lcTextOut" )
*
*    (e) Alternatively, property "lWriteAccumulatedText" is a flag that, if set,
*        causes any accumulated text to be written to the proper destination, 
*        as determined by _TEXT at the completion of the block. (Presumably
*        not used in conjuction with cTextMergeVariable.) Use this when you want
*        the performance gains of accumulation, but still want the output to go
*        to the file specified by SET TEXTMERGE TO.
*
* Even when _not_ accumulating, CodeBlock has characteristics that affect text
* processing. For example, all text between any TEXT and ENDTEXT pair is accumulated 
* prior to any output. This text is then processed through internal METHOD "InternalMergeText",
* which approximates the results of line-by-line merging by stripping out and
* EVAL'uating and expressions between textmerge delimiters. To attain strings as
* the result, TRANSFORM() is used, which is not perfect (differences include treatment
* of DECIMALS).
*
* If the above is not acceptable, either make certain your merged expressions produce
* character results, or set the property "lClassicTextmerge" to TRUE, which will result
* in a (slower) line-by-line merge using macro substitution and "\".
*
* CAUTION: Be sure you have SET TEXTMERGE ON, if you want this class
*   to evaluate delimited strings. This routine respects that setting. Also, if
*   you want this routine to write to a file, be sure _TEXT is properly set,
*   typically via SET TEXTMERGE TO <file>.

*!*	VII - EXTERNAL REFERENCES AND "THIS"
*!*	====================================
*!*	In general, you cannot use "THIS" in your code, since it would refer to
*!*	the CodeBlock object rather than your own. You can, however, pass in a reference
*!*	from your application or set the "oCaller" property directly as follows:

*!*		LOCAL loCB
*!*		loCB = CREATE( "cusCodeBlock" )
*!*		loCB.SetCodeBlock( MyCode.MemoFieldWithCode )
*!*		loCB.SetCallingObject( THIS ) && or maybe THIS.oApp
*!*		loCB.Execute()

*!*	Now you could make external references in your code block like this:

*!*		WAIT WINDOW "Codeblock called by " + THIS.oCaller.Name


*!*	VIII - OBJECT REUSE
*!*	===================
*!*	To execute more than one block of code without instantiating a new object,
*!*	you MUST use the ResetProperties() method. Example:

*!*		LOCAL loCB
*!*		loCB = CREATE( "cusCodeBlock" )
*!*		SCAN
*!*			loCB.SetCodeBlock( WeeklyMaintenance.TaskCode )
*!*			loCB.Execute()
*!*			loCB.ResetProperties()
*!*		ENDSCAN

*!*	IX - Why "FOR EACH" has limited Support
*!*	=======================================
*!*			* BUG in VFP prevents using macro substitution on this command.
*!*			* Major difficulties were encountered to support this. Here
*!*			* are some points:
*!*			*  - must be supported by simulating a regular FOR..ENDFOR (for arrays)
*!*			*  - needed different approach to support OLE collections
*!*			*  - performance and complexity would suffer appreciably if we
*!*			*    supported nested code
*!*			*  - must parse out all existence of the first element and array
*!*			*    name in child block and translate to THIS.<new_properties>
*!*			*  - must support *nested* FOR EACH (yuck)
*!*			* This seems much too complex and not worth it, for now.
*!*			*
*!*			* As such, you CAN use FOR EACH for arrays or OLE collections, as long as 
*!*			* you have no control structures between FOR EACH and ENDFOR. You can still
*!*			* use functions like IIF() to good effect. Also, remember you can often
*!*			* use SETALL() to eliminate some FOR EACH structures completely.

*!*	X - Web Connection Script Maps
*!*	==============================
*!*	Codeblock can be used to execute West-Wind Web Connection style script 
*!*	map files (.WCS files).  With these, an HTML file can contain nested command
*!*	structures in ASP style, like:

*!*	<HTML>
*!*	<% SCAN %>
*!*	<P><%= CustNo %>
*!*	<% ENDSCAN %>
*!*	</HTML>

*!*	If you execute this code, it is first "inverted" into VFP code using textmerge
*!*	and then executed. The inverted code would look like:

*!*	TEXT
*!*	<HTML>
*!*	ENDTEXT

*!*	SCAN 
*!*	TEXT
*!*	<P><<CustNo>>
*!*	ENDTEXT
*!*	ENDSCAN 

*!*	TEXT
*!*	</HTML>
*!*	ENDTEXT

*!*	This in conjunction with the advanced textmerge features discussed above
*!*	will produce a dynamic web page.

*!*	RECORD OF REVISIONS
*!*	===================
*!*	Still To Do
*!*	-----------
*!*	 - Support for PARAMETERS and LPARAMETERS, possibly using SetParameters() 
*!*	   and aParameters.
*!*	   

*!*	10/05/2001 [Revision 3.2(a)]
*!*	----------
*!*	 - Revised pre-processor to use ALINES() vs. _MLINE. (Thanks to Paul Mrozowski
*!*	   for suggested code changes.) Note that this may exclude pre-VFP 6.0.
*!*  - Support for determining *original* line number in error reporting.
*!*	 - Addition of new property, nRawErrorLine that gives line number in 
*!*	   original (non pre-processed) code.
*!*	 - GetPreProcessedCodeBlock() method now accepts a 3rd parameter
*!*	   requesting that original code line numbers be appended to each line.
*!*	 - GetErrorContext() now returns original line numbers after each line.
*!*	 - Original code now parsed to aRawCode array property.
*!*	 - Pre-processed code array aCode now has 3 columns, the 2nd and 3rd
*!*	   being used to track original code line number(s).

*!*	09/18/2001 [Revision 3.1(f)]
*!*	----------
*!*	 - Revised pre-processor and code executor routines to accept IF syntax like "IF(.T.)"
*!*	 - Modified pre-processor code for comment line recognition, so that continuation
*!*	   marks on comment lines cause the next line to be treated as part of the comment.

*!*	09/07/2001 [Revision 3.1(e)]
*!*	----------
*!* [Thanks to Andrus Moor for suggesting following 3 changes/fixes:]
*!*	 - Fixed cmdVerify.Click to ERASE any .ERR file created during code Verify.
*!*	 - Wrapper code now has RETURN value of _0_qoCB.Result (vs. simple RETURN).
*!*	 - Use SYS(2023) folder for cmdVerify operations.

*!*	10/28/1998 [Revision 3.1(d)]
*!*	----------
*!*	 - Applied TEXT..ENDTEXT fixes from Rev 3.1(c) to \ and \\ statements.
*!*	 - Set flag lMacroError when errors occur during text evaluation.
 
*!*	10/08/1998 [Revision 3.1(c)]
*!*	----------
*!*	 - Added support for LOCAL ARRAY statements.
*!*	 - Fixed problem where user errors ocurring in TEXT..ENDTEXT blocks
*!*	   did not cause code to stop executing until after end of block.
*!*	 - Added support for determining nature of errors occurring in
*!*	   TEXT blocks.
*!*	 - Fixed error handling of IF and CASE statements where the evaluation
*!*	   of expressions following either command produced a runtime error.
   
*!*	07/19/1998 [Revision 3.1(b)]
*!*	----------
*!*	 - Execute(), added call to ResetProperties() when code block is passed directly in
*!*	   to allow for multiple calls to Execute using older calling syntax.
 
 
*!*	06/02/1998 [Revision 3.1(a)]
*!*	----------
*!*	 - Added support for interpreter-time compiler directives. Support is now available
*!*	   for codeblocks that include #DEFINE's, #INCLUDE's and 
*!*	   #IF--#ELIF--#ELSE--#ENDIF constructs. Note that these constants are evaluated just
*!*	   before codeblock execution, and NOT at compile time. Also, any #INCLUDE's file must
*!*	   be available in the VFP path at runtime. Despite these limits, this support can help
*!*	   creating flexible scripts that employ framework constants and local settings.
*!*	 
*!*	 - Added error-handling property (flag) "lMacroError" which signifies whether any error
*!*	   resulted during a macro expansion of a user's line of code. 
*!*	  
*!*	 - Improved GetErrorContext() function so it can return either HTML or plain text, and
*!*	   uses the lMacroError flag to customize the context message.
*!*	 
*!*	 - Added limited support for FOR EACH construct. Works for arrays and OLE collections,
*!*	   but DOES NOT support nested control structures. (See separate notes on this.)
*!*	   
*!*	 - Greatly improved the WITH..ENDWITH support. Nested definitions work. Proper
*!*	   handling of strings like .AND. (actually works better than compiled VFP here).
*!*	   Everything done in pre-process phase for better performance.
*!*	 
*!*	 - Handle all local-to-private translation in pre-process phase. This is now controlled
*!*	   by a new property "lLocalToPrivate" rather than a #DEFINE.
*!*	   
*!*	 - Handle all private initialization in pre-process phase. This is now controlled
*!*	   by a new property "lInitializePrivates" rather than a #DEFINE.
*!*	   
*!*	 - Added SubClass, "RunVFPScript", which is set up to work in conjunction with 
*!*	   an existing RESPONSE object. To call:
*!*	   
*!*	     oCB = CREATE( "RunVfpScript")
*!*	     oCB.SetScriptFromFile( "TEST.WCS")
*!*	     oCB.Execute()
*!*	     IF oCB.lError
*!*	       Response.Write( oCB.GetErrorContext() )
*!*	     ENDIF
*!*	     
*!*	 - Eliminated virtually all #DEFINE constants. Many steps that were previously
*!*	   optional have been moved into the Pre-Process stage, and are no longer optional,
*!*	   since they have much less performance impact there.
*!*	   
*!*	05/15/1998 [Revision 3.0(d)]
*!*	----------
*!*	 - Corrected problem where running InvertScript() by itself failed to
*!*	   clear the lScript property thus leading to double inversion.
*!*	   
*!*	05/13/1998 [Revision 3.0(c)]
*!*	----------
*!*	 - Fixed SCAN bug where no SCAN clauses in VFP 5 produced error.
*!*	 - Added hooks at beginning and end of code pre-processing to allow
*!*	   developers to add transformations from their code to VFP code. See
*!*	   BeforePreProcessCode() and AfterPreProcessCode().
*!*	 - Verify array size during pre-processing to avoid error in case
*!*	   pre-processed code actually gets larger than original code.
*!*	 - Added a line to clear cMergedText in ResetProperties() method.
*!*	 - Removed #CODEBLOCK_STORAGE_DBF flag and all associated code that
*!*	   allowed CURSOR storage vs. array storage.
*!*	 - Added method GetPreProcessedCodeBlock() to allow returning pre-processed
*!*	   code block as a string.
*!*	 - Fixed problem where an error in user's code did not result in the actual
*!*	   line of code being stored to the cErrorCode property.
*!*	 - Added method GetErrorContext(), which for normal errors, returns the
*!*	   offending line of the user's code plus the 2 surrounding lines before 
*!*	   and after this line.
*!*	   
*!*	05/04/1998 [Revision 3.0(a)]
*!*	----------
*!*	*** MAJOR REWRITE -- COMPLETE NEW ARCHITECTURE

*!*	Changed architecture from where the recursive call created another *object* to
*!*	where the Execute() method simply calls itself within the current, single object.

*!*	Changed strategy to include a PreProcess() step, wherein the entire code
*!*	is analyzed once at the beginning. This allows for enhanced perfromance and 
*!*	simplifications to the balance of the code because:

*!*	- Can check/clean up control structures once and then count on a given syntax
*!*	  within nested loops, etc.
*!*	- Can opt to checkfor illegal/non-supported commands without taking anywhere
*!*	  near as big a performance hit (see notes above on CODEBLOCK_CODE_CHECKING constant).
*!*	- Any multi-line text between TEXT..ENDTEXT can be gathered into a single string
*!*	  before code processing. For TEXTMERGE in a SCAN loop, this can significantly
*!*	  enhance performance.
*!*	- Many errors are flagged before the code starts to run, thus preventing your
*!*	  application from getting into an inconsistent state.

*!*	Added special handling of PRIVATE and LOCAL statements to help bridge the gap to
*!*	a compile environment and to eliminate the most common problems people have had:

*!*	a) Automatically translate any LOCAL command to a PRIVATE command to avoid problems
*!*	   where LOCAL variables went out-of-scope within nested structures. A second
*!*	   command STORE .F. TO <var_list> is inserted to simulate how LOCAL's work.
*!*	   (I strongly recommend using this translation, but the beahvior *can* be 
*!*	   eliminated by setting a constant CODEBLOCK_LOCAL_TO_PRIVATE to .F.)

*!*	b) Added an optional beahior to initialize PRIVATE variables to .F., to avoid 
*!*	   scope issue where a variable isn't set until inside say an IF.ENDIF and then 
*!*	   a reference is attempted outside the structure.
*!*	   (I strongly recommend using this initialization, but the beahvior *can* be 
*!*	   eliminated by setting a constant CODEBLOCK_INITIALIZE_PRIVATE to .F.)

*!*	Added an "oCaller" property to allow you to pass in an external object reference
*!*	to your application. This allows you to make external references without PUBLIC
*!*	or PRIVATE commands. See notes above on "External References and THIS", but here 
*!*	is a quick example of how to reference once property is set:

*!*		WAIT WINDOW "Codeblock called by " + THIS.oCaller.Name

*!*	INIT() method now accepts only two (optional) parameters: (1) the code block itself;
*!*	and (2) an external object reference. I recommend NOT using these, but rather
*!*	calling properties and methods directly.

*!*	Re-packaged auxiliary material into single PRG file, including wrapper function and
*!*	CodeBlock editing FORM. See top of file for how to use these, if wanted.

*!*	Introduced several optional approaches to handling TEXTMERGE situations. You can
*!*	opt to accumulate all merged text, process at the end, call an external merge 
*!*	routine, use various internal merge functions, save merged text to a variable
*!*	that you create and use after the code has run, etc.  See the section above
*!*	entitled "OPTIONS FOR TEXTMERGE PROCESSING".

*!*	Fixed a problem with WITH..ENDWITH if you had 2 or more of these nested. (NOTE: You 
*!*	must have constant CODEBLOCK_WITH_SUPPORT set to .T. for WITH's to be honored!)

*!*	Added a new method ResetProperties() to allow the object to be reused without
*!*	having to reinstantiate. This could be helpful for performance in certain server
*!*	applications that do lots of script/template processing. See notes under topic
*!*	"Object Reuse" above.
