Attribute VB_Name = "modErrHandler"
Option Compare Database
Option Explicit

Public Function ErrHandler( _
ByVal pvlngErrNum As Long, _
ByVal pvstrErrDesc As String, _
Optional ByVal pvstrErrCustom As String, _
Optional ByVal pvstrErrProcedure As String, _
Optional ByVal pvstrErrModule As String) As Long
Dim strErrorMessage As String
On Error GoTo ErrHandler_Err
'****************************************************************************************
'*  Name            :       ErrHandler                                                  *
'*  Author          :       Joseph Solomon, CPA                                         *
'*  Purpose         :       This procedure is a general use Error Handling function. The*
'*                              function will log the Error to tblError.  If tblError   *
'*                              does not exist, this function will create it.           *
'*  Return Value    :       ErrHandler will return a value derived from the VBA MsgBox  *
'*                              Function's vbAbortRetryIgnore enumeration based on the  *
'*                              User's Selection.  If this function Errors out, it will *
'*                              not log the Error, but will prompt the user to Abort,   *
'*                              Retry, or Ignore.                                       *
'*  Last Update     :       2012/12/13                                                  *
'*                                                                                      *
'*  Parameters/Variables:   Description:                                                *
'*  ~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~               *
'*  pvlngErrNum             Error number passed into the Error handler.                 *
'*  pvstrErrDesc            VBA's Error Description                                     *
'*  pvstrErrCustom          Optional custom Error Description                           *
'*  pvstrErrProcedure       Optional name of procedure calling the Error.               *
'*  pvstrErrModule          Optional name of module containing the procedure.           *
'*  strErrorMessage         Prompt that shows in the Error Message Box.                 *
'*                                                                                      *
'*  Usage Example:                                                                      *
'*  ~~~~~~~~~~~~~~~                                                                     *
'*  On Error GoTo MyProcedure_Err                                                       *
'*  <...code...>                                                                        *
'*  MyProcedure_Err:                                                                    *
'*                                                                                      *
'*  Dim strErrorMessage as String                                                       *
'*  Select Case Err.Number                                                              *
'*      Case is = 11 'Division by zero                                                  *
'*          strErrorMessage = "The Application tried to divide by zero.                 *
'*      Case Is = 75, 76 'Path/File Access Error                                        *
'*          strErrorMessage = "There was a problem accessing the network.  Check " & _  *
'*              "your connection and try again."                                        *
'*  End Select                                                                          *
'*                                                                                      *
'*  Select Case ErrHandler(Err.Number, Err.Description, strErrorMessage, "DataImport", _*
'*      "modImport")                                                                    *
'*      Case Is = vbIgnore: Resume Next                                                 *
'*      Case Is = vbAbort: Exit Sub                                                     *
'*      Case Is = vbRetry: Resume                                                       *
'*  End Select                                                                          *
'****************************************************************************************
    
    'Create Error Message
    If Not IsMissing(pvstrErrCustom) Then 'Optional Error Message parameter was passed
         strErrorMessage = _
            pvstrErrCustom & vbCrLf & vbCrLf & _
            "Error # : " & pvlngErrNum & vbCrLf & _
            "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
            "Please make a selection:"
    Else 'No Optional Error Message Passed.  Use default.
        strErrorMessage = _
            "Error # : " & pvlngErrNum & vbCrLf & _
            "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
            "Please make a selection:"
    End If

    'Create Error table if it does not exist
    DoCmd.SetWarnings False
    If DCount("[Name]", "MSysObjects", "Type = 1 AND [Name] = 'tblErrorLog'") < 1 Then
        DoCmd.RunSQL _
            "CREATE TABLE tblErrorLog " & _
            "(ErrorID AUTOINCREMENT PRIMARY KEY, ErrorNumber INTEGER, ErrorDescription LONGTEXT, " & _
            "ProcedureName CHAR, ModuleName CHAR, ErrorTimeStamp DATETIME, User CHAR, " & _
            "ComputerName CHAR);"
    End If
    
    'Log the Error
    DoCmd.RunSQL _
        "INSERT INTO tblErrorLog " & _
        "(ErrorNumber, ErrorDescription, ProcedureName, ModuleName, " & _
        "ErrorTimeStamp, User, ComputerName) " & _
        "VALUES (" & _
        pvlngErrNum & ", '" & pvstrErrDesc & "', '" & pvstrErrProcedure & "', '" & _
        pvstrErrModule & "', #" & Now & "#, '" & Environ("USERNAME") & "', '" & _
        Environ("COMPUTERNAME") & "');"
    DoCmd.SetWarnings True
    
    ErrHandler = MsgBox( _
        Prompt:=strErrorMessage, _
        Buttons:=vbAbortRetryIgnore Or vbExclamation, _
        Title:="Error " & pvlngErrNum)
        
ErrHandler_Exit:
    Exit Function
    
ErrHandler_Err:
    On Error Resume Next
    Select Case MsgBox( _
            Prompt:="There was an Error handling the Error...Ironic, huh?" & vbCrLf & _
                "This Error will not be logged." & _
                vbCrLf & vbCrLf & _
                "Error # : " & pvlngErrNum & vbCrLf & _
                "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
                "Please make a selection:", _
            Buttons:=vbAbortRetryIgnore Or vbExclamation, _
            Title:="Error Handling Error")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo ErrHandler_Exit

End Function

