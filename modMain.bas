Attribute VB_Name = "modMain"
Option Explicit
Option Compare Text
DefInt A-Z

'API Types
Public Type SHITEMID
   cb             As Long
   abID           As Byte
End Type

Public Type ITEMIDLIST
   mkid           As SHITEMID
End Type

Public Type BROWSEINFO
   hOwner         As Long
   pidlRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfn           As Long
   lParam         As Long
   iImage         As Long
End Type

'API Constants
Public Const BIF_RETURNONLYFSDIRS = &H1

'API Functions
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
                "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
                "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

'Public variables
Public gobjServer As New SQLDMO.SQLServer
Public gsServerNameSW As String
Public gsLoginSW As String
Public gsPasswordSW As String
Public gvDatabasesSW As Variant
Public gsDestDirSW As String
Public gbDelOldFilesSW As Boolean
Public giDelWeeksSW As Integer
Public gbSaveDTSPackagesSW As Boolean
Public gbScriptAlertsSW As Boolean
Public gbScriptServerLoginsSW As Boolean
Public gbScriptAgentJobsSW As Boolean
Public gbScriptBackupDevicesSW As Boolean
Public gbRunUnAttendedSW As Boolean

Public gbNoGUI As Boolean
Public gsUnattendedLog As String
Public gsLogFile As String

' new command line options
Public gbScriptDB       As Boolean
Public gbScriptTabs     As Boolean
Public gbScriptViews    As Boolean
Public gbScriptSPs      As Boolean
Public gbScriptRules    As Boolean
Public gbScriptDefs     As Boolean
Public gbScriptRoles    As Boolean
Public gbScriptUsers    As Boolean
Public gbScriptUDTs     As Boolean
Public gbScriptFTCs     As Boolean
Public gbObjsInSepFiles As Boolean
Public gbDatePrefix     As Boolean
Public gsSelTables      As String
Public gbScriptDrops1st As Boolean

Public Function BrowseForFolder(szPrompt As String) As String
   Dim biInfo As BROWSEINFO
   Dim pidl As Long
   Dim szPath As String
   
   szPath = Space$(512)
   
   biInfo.hOwner = 0&
   biInfo.pidlRoot = 0&
   biInfo.lpszTitle = szPrompt
   biInfo.ulFlags = BIF_RETURNONLYFSDIRS
   
   pidl = SHBrowseForFolder(biInfo)
   SHGetPathFromIDList ByVal pidl, ByVal szPath
   
   BrowseForFolder = Trim$(szPath)
End Function

Public Sub SaveLog(ByVal gsLogFile As String)
    Dim iFileNumber As Integer
    
    On Error Resume Next
    iFileNumber = FreeFile
    Open gsLogFile For Output As #iFileNumber
    Print #iFileNumber, gsUnattendedLog
    Close #iFileNumber
    DoEvents
End Sub

Public Sub StatusMessage(ByVal strMessage As String)
    Const ciMaxVisualLogLength As Integer = 30000
    Const cstrLogDateTimeFormat As String = "yyyymmdd Hh:Mm:Ss"
    
    Dim strLogEntry As String
    Dim iPos As Integer
    
    On Error Resume Next
    strLogEntry = Format(Now, cstrLogDateTimeFormat) & vbTab & strMessage
    gsUnattendedLog = gsUnattendedLog & strLogEntry & vbNewLine
    If Not gbNoGUI Then
        With frmMain.lblStatus
            .Caption = " " & strMessage
            .Refresh
        End With
        If Len(gsUnattendedLog) > ciMaxVisualLogLength Then
            iPos = InStr(Right$(gsUnattendedLog, ciMaxVisualLogLength), vbNewLine)
            With frmMain.txtLog
                .Text = Right$(gsUnattendedLog, ciMaxVisualLogLength - iPos - 1)
                .SelStart = Len(.Text)
                .Refresh
            End With
        Else
            With frmMain.txtLog
                .Text = gsUnattendedLog
                .SelStart = Len(.Text)
                .Refresh
            End With
        End If
    End If
End Sub

Public Function CreateDataInsert(ByVal vsTableName As String, _
    ByVal vbIdentityPK As Boolean, ByVal vsServer As String, _
    ByVal vsDatabase As String, ByVal vsLogin As String, ByVal vsPwd As String, _
    ByVal vbTrustedConnection As Boolean) As String
    
    Dim k As Integer
    Dim strSql As String, strSql1 As String, strMainSql As String
    Dim sFieldVal As String, sOutput As String
    Dim RS As New ADODB.Recordset
    Dim sDSN As String
    
    If vbTrustedConnection Then
        sDSN = "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
            "Trusted Connection=True;Initial Catalog=" & vsDatabase & _
            ";Data Source=" & vsServer & ""
    Else
        sDSN = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & _
            vsLogin & ";Password=" & vsPwd & ";Initial Catalog=" & vsDatabase & _
            ";Data Source=" & vsServer & ""
    End If
    
    If vbIdentityPK Then
        sOutput = "SET IDENTITY_INSERT " & vsTableName & " ON" & vbNewLine
    Else
        sOutput = ""
    End If
    
    strSql = "insert into " & vsTableName & "("
    RS.Open vsTableName, sDSN, adOpenForwardOnly, adLockReadOnly

    For k = 0 To RS.Fields.Count - 1
        strSql = strSql & RS.Fields(k).Name & ","
    Next
    
    strSql = Left(strSql, Len(strSql) - 1) & ") values("
            
    Do While Not RS.EOF

        For k = 0 To RS.Fields.Count - 1
    
            If Not IsNull(RS.Fields(k).Value) Then
                sFieldVal = RS.Fields(k).Value
            Else
                sFieldVal = "Null"
            End If
            
            Select Case RS.Fields(k).Type
            
                Case adChar, adVarChar, adVarWChar, adLongVarWChar, adLongVarBinary
                    sFieldVal = Replace(sFieldVal, "'", "''")
                    strSql1 = strSql1 & "'" & sFieldVal & "',"
                
                Case adBoolean
                    If sFieldVal = "true" Then
                        strSql1 = strSql1 & "1,"
                    Else
                        strSql1 = strSql1 & "0,"
                    End If
                    
                Case adDBDate
                    If sFieldVal = "Null" Then
                        strSql1 = strSql1 & sFieldVal & ","
                    Else
                        strSql1 = strSql1 & "'" & Format(sFieldVal, "dd-mmm-yyyy") & "',"
                    End If
                
                Case adDBTimeStamp
                    If sFieldVal = "Null" Then
                        strSql1 = strSql1 & sFieldVal & ","
                    Else
                        strSql1 = strSql1 & "'" & Format(sFieldVal, "dd-mmm-yyyy") & " " & Format(sFieldVal, "Long Time") & "',"
                    End If
                    
                Case Else
                    strSql1 = strSql1 & sFieldVal & ","
            End Select
        
        Next
                
        strSql1 = Left(strSql1, Len(strSql1) - 1)
        strMainSql = strSql & strSql1 & ")"
        sOutput = sOutput & strMainSql & vbCrLf
        strSql1 = ""
        strMainSql = ""
        RS.MoveNext
    
    Loop
    
    RS.Close
    Set RS = Nothing
    
    If vbIdentityPK Then
        sOutput = sOutput & "SET IDENTITY_INSERT " & vsTableName & " OFF"
    End If
    
    CreateDataInsert = sOutput

End Function

Public Sub GenerateScripts(strServerName As String, strDir As String, _
    Optional vDatabases As Variant, _
    Optional bTrustedConnection As Boolean = True, _
    Optional strLogin As String, Optional strPassword As String, _
    Optional bDelOldFiles As Boolean = False, Optional iDelWeeks As Integer = 4, _
    Optional bScriptDatabase As Boolean = True, _
    Optional bScriptDBTables As Boolean = True, _
    Optional bScriptDBViews As Boolean = True, _
    Optional bScriptDBSPs As Boolean = True, _
    Optional bScriptDBRules As Boolean = True, _
    Optional bScriptDBDefaults As Boolean = True, _
    Optional bScriptDBRoles As Boolean = True, _
    Optional bScriptDBFullText As Boolean = True, _
    Optional bScriptDBUsers As Boolean = True, _
    Optional bScriptDBUDTs As Boolean = True, _
    Optional bSaveDTSPackages As Boolean = False, _
    Optional bScriptAlerts As Boolean = False, _
    Optional bScriptAgentJobs As Boolean = False, _
    Optional bScriptServerLogins As Boolean = False, _
    Optional bScriptBackupDevices As Boolean = False, _
    Optional bScriptToSeparateFiles As Boolean = False, _
    Optional bIncludeDSPrefix As Boolean = True, _
    Optional sInsertScriptTables As String = vbNullString, _
    Optional bScriptDropsFirst As Boolean = True)

    Dim lScriptOptions As Long
    Dim lTableScriptOptions As Long
    Dim objLogin As SQLDMO.Login
    Dim objDatabase As SQLDMO.Database
    Dim objCatalog As SQLDMO.FullTextCatalog
    Dim objTab As SQLDMO.Table
    Dim objView As SQLDMO.View
    Dim objProc As SQLDMO.StoredProcedure
    Dim objRule As SQLDMO.Rule
    Dim objDefault As SQLDMO.Default
    Dim objUser As SQLDMO.User
    Dim objRole As SQLDMO.DatabaseRole
    Dim objUDT As SQLDMO.UserDefinedDatatype
    Dim objBackupDevice As SQLDMO.BackupDevice
    Dim qryResults As SQLDMO.QueryResults
    Dim objPackage As DTS.Package
    Dim strScript As String, strScript2 As String
    Dim sFilePrefix As String, strDelCommand As String, strFile As String
    Dim strSQLquery As String, strPackageFile As String
    Dim i As Integer, i2 As Integer, j As Integer, iFileNumber As Integer
    Dim dtDelDate As Date
    Dim msgResult As VbMsgBoxResult
    Dim strCheckPoint As String
    Dim strPackageName As String
    Dim vInsertScriptTables As Variant
    Dim vInsertTable As Variant
    Dim sPKCol As String
    Dim bIdentityPK As Boolean
    
    lScriptOptions = SQLDMOScript_Default Or SQLDMOScript_Drops Or _
        SQLDMOScript_IncludeHeaders Or SQLDMOScript_Permissions Or _
        SQLDMOScript_OwnerQualify
    
    lTableScriptOptions = lScriptOptions Or SQLDMOScript_Indexes Or SQLDMOScript_Triggers Or _
        SQLDMOScript_DRI_All Or SQLDMOScript_Bindings
    
    If bIncludeDSPrefix Then
        sFilePrefix = Format(Date, "yyyymmdd") & "."
    End If
    
    ' cannot do both
    If bScriptToSeparateFiles And bScriptDropsFirst Then
        bScriptDropsFirst = False
    End If
    
    If bScriptDropsFirst And Not bScriptToSeparateFiles Then
        lScriptOptions = lScriptOptions Or SQLDMOScript_IncludeIfNotExists
        lScriptOptions = lScriptOptions Xor SQLDMOScript_Drops
        lTableScriptOptions = lTableScriptOptions Xor SQLDMOScript_Drops Or SQLDMOScript_IncludeIfNotExists
    End If
    
    'Deleting old files
    If bDelOldFiles Then
        strCheckPoint = "deleting"
        On Error Resume Next
        StatusMessage "Deleting old files in [" & strDir & "]:"
        dtDelDate = DateAdd("ww", -iDelWeeks, Date)
        If InStr(strDir, ":") > 0 Then ChDrive Left$(Trim$(strDir), 1)
        ChDir strDir
        strFile = Dir("*.*")
        i = 0
        While Len(strFile) <> 0
            If DateValue(FileDateTime(strFile)) <= dtDelDate Then
                i = i + 1
                Kill strFile
                StatusMessage "- Deleted file " & CStr(i) & ": " & strFile
                DoEvents
            End If
            strFile = Dir
        Wend
        StatusMessage "* " & CStr(i) & " old files deleted."
    End If
    
    'Save DTS Packages
    On Error GoTo ErrorHandler
    If gobjServer.Issysadmin And bSaveDTSPackages Then
        strCheckPoint = "dts"
        StatusMessage "Saving DTS packages as Structured Storage Files:"
        Set qryResults = gobjServer.ExecuteWithResults("USE msdb SELECT DISTINCT name FROM sysdtspackages")
        With qryResults
            For i = 1 To .Rows
                strPackageName = .GetColumnString(i, 1)
                StatusMessage "- Saving DTS package " & CStr(i) & ": " & strPackageName
                Set objPackage = New DTS.Package
                If gobjServer.LoginSecure Then
                    objPackage.LoadFromSQLServer ServerName:=strServerName, PackageName:=strPackageName, Flags:=DTSSQLStgFlag_UseTrustedConnection
                Else
                    objPackage.LoadFromSQLServer ServerName:=strServerName, ServerUserName:=gsLoginSW, ServerPassword:=gsPasswordSW, PackageName:=strPackageName
                End If
                DoEvents
                strPackageFile = strDir & sFilePrefix & strPackageName & ".dts"
                objPackage.SaveToStorageFile strPackageFile
                DoEvents
NextDTSPackage:
                objPackage.UnInitialize
                Set objPackage = Nothing
            Next i
        End With
        StatusMessage "* " & CStr(i - 1) & " DTS packages saved."
        DoEvents
    End If

    'Script Operators/Alerts
    If gobjServer.Issysadmin And bScriptAlerts Then
        strCheckPoint = "alerts"
        StatusMessage "Generating script for Operators and Alerts:"
        strScript = gobjServer.JobServer.Operators.Script(ScriptType:=lScriptOptions)
        DoEvents
        strScript2 = gobjServer.JobServer.Alerts.Script(ScriptType:=lScriptOptions, Script2Type:=SQLDMOScript2_AgentNotify)
        DoEvents
        iFileNumber = FreeFile
        strFile = strDir & sFilePrefix & strServerName & "_Alerts.sql"
        Open strFile For Output As #iFileNumber
        Print #iFileNumber, strScript
        DoEvents
        Print #iFileNumber, strScript2
        DoEvents
        Close #iFileNumber
        i = gobjServer.JobServer.Alerts.Count
        i2 = gobjServer.JobServer.Operators.Count
        StatusMessage "* " & CStr(i2) & " Operators and " & CStr(i) & " Alerts scripted."
        DoEvents
    End If
    
    'Script SQL Agent Jobs
    If gobjServer.Issysadmin And bScriptAgentJobs Then
        strCheckPoint = "jobs"
        StatusMessage "Generating script for SQL Agent Jobs:"
        strScript = gobjServer.JobServer.Jobs.Script(ScriptType:=lScriptOptions)
        DoEvents
        iFileNumber = FreeFile
        strFile = strDir & sFilePrefix & strServerName & "_Jobs.sql"
        Open strFile For Output As #iFileNumber
        Print #iFileNumber, strScript
        DoEvents
        Close #iFileNumber
        i = gobjServer.JobServer.Jobs.Count
        StatusMessage "* " & CStr(i) & " SQL Agent Jobs scripted."
        DoEvents
    End If
    
    'Script Server Logins
    If gobjServer.Issysadmin And bScriptServerLogins Then
        strCheckPoint = "logins"
        StatusMessage "Generating script for Server Logins:"
        iFileNumber = FreeFile
        strFile = strDir & sFilePrefix & strServerName & "_Logins.sql"
        Open strFile For Output As #iFileNumber
        i = 0
        For Each objLogin In gobjServer.Logins
            i = i + 1
            StatusMessage "- Generating T-SQL code for Server Login " & CStr(i) & ": " & objLogin.Name
            strScript = objLogin.Script(ScriptType:=lScriptOptions, Script2Type:=SQLDMOScript2_LoginSID)
            DoEvents
            Print #iFileNumber, strScript
            DoEvents
        Next objLogin
        Close #iFileNumber
        StatusMessage "* " & CStr(i) & " Server Logins scripted."
        DoEvents
    End If
    
    'Script Backup Devices
    If gobjServer.Issysadmin And bScriptBackupDevices Then
        strCheckPoint = "backupdevices"
        StatusMessage "Generating script for Backup Devices:"
        iFileNumber = FreeFile
        strFile = strDir & sFilePrefix & strServerName & "_BackupDevices.sql"
        Open strFile For Output As #iFileNumber
        i = 0
        For Each objBackupDevice In gobjServer.BackupDevices
            i = i + 1
            StatusMessage "- Generating T-SQL code for Backup Device " & CStr(i) & ": " & objBackupDevice.Name
            strScript = objBackupDevice.Script(ScriptType:=lScriptOptions)
            DoEvents
            Print #iFileNumber, strScript
            DoEvents
        Next objBackupDevice
        Close #iFileNumber
        StatusMessage "* " & CStr(i) & " Backup Devices scripted."
        DoEvents
    End If
    
    'Script Databases
    If Not IsMissing(vDatabases) Then
    
        strCheckPoint = "databases"
        i = 0
        StatusMessage "Generating scripts for Databases:"
        
        For j = 0 To UBound(vDatabases)
            
            ' Script the Database
            i = i + 1
            Set objDatabase = gobjServer.Databases(vDatabases(j))
            
            If bScriptDatabase Then
            
                StatusMessage "- Generating T-SQL code for Database " & CStr(i) & ": " & vDatabases(j)
                DoEvents
                strScript = objDatabase.Script(ScriptType:=lScriptOptions)
                DoEvents
                iFileNumber = FreeFile
                strFile = strDir & sFilePrefix & objDatabase.Name & "_Database.sql"
                Open strFile For Output As #iFileNumber
                Print #iFileNumber, strScript
                DoEvents
                
                If bScriptDBFullText Then
                    'Script the FullText catalogs
                    StatusMessage "  - Generating scripts for FullText Catalogs:"
                    i2 = 0
                    For Each objCatalog In objDatabase.FullTextCatalogs
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for FullText Catalog " & CStr(i2) & ": " & objCatalog.Name
                        strScript = objCatalog.Script(ScriptType:=lScriptOptions, Script2Type:=SQLDMOScript2_FullTextCat)
                        Print #iFileNumber, strScript
                        Print #iFileNumber, "GO"
                        DoEvents
                    Next objCatalog
                    StatusMessage "  * " & CStr(i2) & " FullText Catalogs scripted."
                    DoEvents
                End If
                
                Close #iFileNumber
            
            End If
            
            If bScriptDBTables Then
            
                'Script Tables, Indexes and Triggers
                StatusMessage "  - Generating scripts for Tables, Indexes and Triggers:"
                
                If Not bScriptToSeparateFiles Then
                    iFileNumber = FreeFile
                    strFile = strDir & sFilePrefix & objDatabase.Name & "_Tables.sql"
                    Open strFile For Output As #iFileNumber
                End If
                
                i2 = 0
                
                ' script drops first
                If bScriptDropsFirst Then
                    If Not bScriptToSeparateFiles Then
                        For Each objTab In objDatabase.Tables
                            If Not objTab.SystemObject Then
                                i2 = i2 + 1
                                StatusMessage "    - Generating T-SQL code for Table " & CStr(i2) & ": " & objTab.Name
                                strScript = objTab.Script(ScriptType:=SQLDMOScript_Drops)
                                DoEvents
                                Print #iFileNumber, strScript
                                DoEvents
                            End If
                        Next
                    End If
                End If
                
                i2 = 0
                
                For Each objTab In objDatabase.Tables
                
                    If Not objTab.SystemObject Then
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for Table " & CStr(i2) & ": " & objTab.Name
                        strScript = objTab.Script(ScriptType:=lTableScriptOptions, Script2Type:=IIf(bScriptDBFullText, SQLDMOScript2_FullTextIndex, SQLDMOScript2_Default))
                        DoEvents
                        
                        If bScriptToSeparateFiles Then
                            iFileNumber = FreeFile
                            strFile = strDir & sFilePrefix & objDatabase.Name & "_" & objTab.Name & ".sql"
                            Open strFile For Output As #iFileNumber
                            Print #iFileNumber, strScript
                            Close #iFileNumber
                        Else
                            Print #iFileNumber, strScript
                        End If
                        
                        DoEvents
                    End If
                    
                Next objTab
                
                If Not bScriptToSeparateFiles Then Close #iFileNumber
                
                StatusMessage "  * " & CStr(i2) & " Tables scripted."
                DoEvents
                
            End If
            
            If bScriptDBViews Then
                'Script Views
                StatusMessage "  - Generating scripts for Views:"
                
                If Not bScriptToSeparateFiles Then
                    iFileNumber = FreeFile
                    strFile = strDir & sFilePrefix & objDatabase.Name & "_Views.sql"
                    Open strFile For Output As #iFileNumber
                End If
                
                i2 = 0
                
                ' script drops first
                If bScriptDropsFirst Then
                    If Not bScriptToSeparateFiles Then
                        For Each objView In objDatabase.Views
                            If Not objView.SystemObject Then
                                i2 = i2 + 1
                                StatusMessage "    - Generating T-SQL code for View " & CStr(i2) & ": " & objView.Name
                                strScript = objView.Script(ScriptType:=SQLDMOScript_Drops)
                                DoEvents
                                Print #iFileNumber, strScript
                                DoEvents
                            End If
                        Next
                    End If
                End If
                
                i2 = 0
                
                For Each objView In objDatabase.Views
                
                    If Not objView.SystemObject Then
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for View " & CStr(i2) & ": " & objView.Name
                        strScript = objView.Script(ScriptType:=lScriptOptions)
                        DoEvents
                        
                        If bScriptToSeparateFiles Then
                            iFileNumber = FreeFile
                            strFile = strDir & sFilePrefix & objDatabase.Name & "_" & objView.Name & ".sql"
                            Open strFile For Output As #iFileNumber
                            Print #iFileNumber, strScript
                            Close #iFileNumber
                        Else
                            Print #iFileNumber, strScript
                        End If
                        
                        DoEvents
                    End If
                    
                Next objView
                
                If Not bScriptToSeparateFiles Then Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Views scripted."
                DoEvents
                
            End If
                
            If bScriptDBSPs Then
                'Script Stored Procedures
                StatusMessage "  - Generating scripts for Stored Procedures:"
                
                If Not bScriptToSeparateFiles Then
                    iFileNumber = FreeFile
                    strFile = strDir & sFilePrefix & objDatabase.Name & "_Procs.sql"
                    Open strFile For Output As #iFileNumber
                End If
                
                i2 = 0
                ' script drops first
                If bScriptDropsFirst Then
                    If Not bScriptToSeparateFiles Then
                        For Each objProc In objDatabase.StoredProcedures
                            If Not objProc.SystemObject Then
                                i2 = i2 + 1
                                StatusMessage "    - Generating T-SQL code for Procedure " & CStr(i2) & ": " & objProc.Name
                                strScript = objProc.Script(ScriptType:=SQLDMOScript_Drops)
                                DoEvents
                                Print #iFileNumber, strScript
                                DoEvents
                            End If
                        Next
                    End If
                End If
                
                i2 = 0
                For Each objProc In objDatabase.StoredProcedures
                
                    If Not objProc.SystemObject Then
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for Procedure " & CStr(i2) & ": " & objProc.Name
                        strScript = objProc.Script(ScriptType:=lScriptOptions)
                        DoEvents
                        
                        If bScriptToSeparateFiles Then
                            iFileNumber = FreeFile
                            strFile = strDir & sFilePrefix & objDatabase.Name & "_" & objProc.Name & ".sql"
                            Open strFile For Output As #iFileNumber
                            Print #iFileNumber, strScript
                            Close #iFileNumber
                        Else
                            Print #iFileNumber, strScript
                        End If
                        
                        DoEvents
                    End If
                    
                Next objProc
                
                If Not bScriptToSeparateFiles Then Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Stored Procedures scripted."
                DoEvents
            End If
            
            If bScriptDBRules Then
                'Script Rules
                StatusMessage "  - Generating scripts for Rules:"
                iFileNumber = FreeFile
                Open strDir & sFilePrefix & objDatabase.Name & "_Rules.sql" For Output As #iFileNumber
                i2 = 0
                For Each objRule In objDatabase.Rules
                    i2 = i2 + 1
                    StatusMessage "    - Generating T-SQL code for Rule " & CStr(i2) & ": " & objRule.Name
                    strScript = objRule.Script(ScriptType:=lScriptOptions)
                    DoEvents
                    Print #iFileNumber, strScript
                    DoEvents
                Next objRule
                Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Rules scripted."
                DoEvents
            End If
            
            If bScriptDBDefaults Then
                'Script Defaults
                StatusMessage "  - Generating scripts for Defaults:"
                iFileNumber = FreeFile
                Open strDir & sFilePrefix & objDatabase.Name & "_Defaults.sql" For Output As #iFileNumber
                i2 = 0
                For Each objDefault In objDatabase.Defaults
                    i2 = i2 + 1
                    StatusMessage "    - Generating T-SQL code for Default " & CStr(i2) & ": " & objDefault.Name
                    strScript = objDefault.Script(ScriptType:=lScriptOptions)
                    DoEvents
                    Print #iFileNumber, strScript
                    DoEvents
                Next objDefault
                Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Defaults scripted."
                DoEvents
            End If
            
            If bScriptDBUsers Then
                'Script Users
                StatusMessage "  - Generating scripts for Users:"
                iFileNumber = FreeFile
                Open strDir & sFilePrefix & objDatabase.Name & "_Users.sql" For Output As #iFileNumber
                i2 = 0
                For Each objUser In objDatabase.Users
                    If Not objUser.SystemObject Then
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for User " & CStr(i2) & ": " & objUser.Name
                        strScript = objUser.Script(ScriptType:=lScriptOptions, Script2Type:=SQLDMOScript2_LoginSID)
                        DoEvents
                        Print #iFileNumber, strScript
                        DoEvents
                    End If
                Next objUser
                Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Users scripted."
                DoEvents
            End If
            
            If bScriptDBRoles Then
                'Script Database Roles
                StatusMessage "  - Generating scripts for Database Roles:"
                iFileNumber = FreeFile
                Open strDir & sFilePrefix & objDatabase.Name & "_DBRoles.sql" For Output As #iFileNumber
                i2 = 0
                For Each objRole In objDatabase.DatabaseRoles
                    If Not objRole.IsFixedRole Then
                        i2 = i2 + 1
                        StatusMessage "    - Generating T-SQL code for Database Role " & CStr(i2) & ": " & objRole.Name
                        strScript = objRole.Script(ScriptType:=lScriptOptions)
                        DoEvents
                        Print #iFileNumber, strScript
                        DoEvents
                    End If
                Next objRole
                Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Database Roles scripted."
                DoEvents
            End If
            
            If bScriptDBUDTs Then
                'Script User Defined Datatypes
                StatusMessage "  - Generating scripts for User Defined Datatypes:"
                iFileNumber = FreeFile
                Open strDir & sFilePrefix & objDatabase.Name & "_UDTs.sql" For Output As #iFileNumber
                i2 = 0
                For Each objUDT In objDatabase.UserDefinedDatatypes
                    i2 = i2 + 1
                    StatusMessage "    - Generating T-SQL code for User Defined Datatype " & CStr(i2) & ": " & objUDT.Name
                    strScript = objUDT.Script(ScriptType:=lScriptOptions)
                    DoEvents
                    Print #iFileNumber, strScript
                    DoEvents
                Next objUDT
                Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " User Defined Datatypes scripted."
                DoEvents
            End If
            
            If sInsertScriptTables <> vbNullString Then
            
                StatusMessage "  - Generating insert scripts for selected Tables:"
                
                If Not bScriptToSeparateFiles Then
                    iFileNumber = FreeFile
                    strFile = strDir & sFilePrefix & objDatabase.Name & "_Insert.sql"
                    Open strFile For Output As #iFileNumber
                End If
                
                i2 = 0
                vInsertScriptTables = Split(sInsertScriptTables, ",")
                
                For i = LBound(vInsertScriptTables) To UBound(vInsertScriptTables)
                    
                    If vInsertScriptTables(i) <> "" Then
                    
                        Set objTab = objDatabase.Tables.Item(vInsertScriptTables(i))
                        
                        If Not objTab.SystemObject Then
                            i2 = i2 + 1
                            StatusMessage "    - Generating T-SQL code for Table " & CStr(i2) & ": " & objTab.Name
                            
                            If Not objTab.PrimaryKey Is Nothing Then
                                sPKCol = objTab.PrimaryKey.KeyColumns.Item(1)
                                bIdentityPK = objTab.Columns.Item(sPKCol).Identity
                            Else
                                bIdentityPK = False
                            End If
                            
                            strScript = CreateDataInsert(objTab.Name, _
                                bIdentityPK, _
                                strServerName, vDatabases(j), strLogin, strPassword, _
                                bTrustedConnection)
                            DoEvents
                            
                            If bScriptToSeparateFiles Then
                                iFileNumber = FreeFile
                                strFile = strDir & sFilePrefix & objDatabase.Name & "_" & objTab.Name & "_Insert.sql"
                                Open strFile For Output As #iFileNumber
                                Print #iFileNumber, strScript
                                Close #iFileNumber
                            Else
                                Print #iFileNumber, strScript
                            End If
                                
                            DoEvents
                        End If
                        
                    End If
                    
                Next
                
                If Not bScriptToSeparateFiles Then Close #iFileNumber
                StatusMessage "  * " & CStr(i2) & " Table Inserts scripted."
                DoEvents
            
            End If
            
            StatusMessage "* Database " & objDatabase.Name & " scripted."
NextDatabase:
        Next j
        StatusMessage CStr(i) & " Databases scripted."
    End If
    
    StatusMessage "Ready."
    Exit Sub
    
ErrorHandler:       ' Error-handling routine.
    StatusMessage "Error: 0x" & Hex$(Err.Number) & vbTab & Error(Err.Number)
    Select Case Err.Number
        Case &H35           'File does not exist
            StatusMessage "Error deleting file. Resuming with next one."
            Resume Next
        Case &H80030002     'DTS Package version number to high
            StatusMessage "DTS Package " & strPackageName & " has wrong version number. Package skipped."
            Resume NextDTSPackage
        Case &H80045510     'Nonexistant database
            StatusMessage "Database " & vDatabases(j) & " does not exist. Resuming with next one."
            i = i - 1
            Resume NextDatabase
        Case Else           'Unanticipated error
            If gbNoGUI Then
                If strCheckPoint = "databases" Then
                    StatusMessage "Error scripting database. Resuming with next one."
                    Resume NextDatabase
                Else
                    StatusMessage "Unanticipated error."
                    StatusMessage "Checkpoint: " & strCheckPoint
                    StatusMessage "Aborting program."
                    SaveLog gsLogFile
                End If
            Else
                msgResult = MsgBox("An SQL-DMO error occurred. The error has been written to the log. " & _
                                   "Do you want to abort the program?" & vbNewLine & vbNewLine & _
                                   "Choose 'Abort' to abort the program, 'Retry' to retry the action that caused " & _
                                   "the error, or 'Ignore' to stop scripting this database and continue with " & _
                                   "the next one (if any). ", vbExclamation Or vbAbortRetryIgnore, "Error!")
                Select Case msgResult
                    Case vbAbort
                        Close #iFileNumber
                        StatusMessage "Program aborted by user."
                        SaveLog gsLogFile
                        Unload frmMain
                    Case vbRetry
                        StatusMessage "Retrying..."
                        Resume
                    Case vbIgnore
                        Close #iFileNumber
                        If strCheckPoint = "databases" Then
                            StatusMessage "Ignoring error, resuming with next database."
                            Resume NextDatabase
                        Else
                            StatusMessage "Error cannot be ignored. Aborting program."
                            MsgBox "Sorry, unable to continue from this point. Program will be terminated instead.", _
                                   vbCritical, "Critical Error!"
                            SaveLog gsLogFile
                            Unload frmMain
                        End If
                End Select
            End If
    End Select
End Sub

Private Sub Main()
    Dim vaSwitches As Variant, vTemp As Variant
    Dim strSwitch As String, strValue As String
    Dim i As Integer
    
    'Parse out the command-line parameters (if any)
    vaSwitches = Split(Command$, "/")
    On Error Resume Next
    For i = 1 To UBound(vaSwitches)
        vTemp = Split(CStr(vaSwitches(i)), "=")
        strSwitch = UCase$(Trim$(CStr(vTemp(0))))
        strValue = Trim$(CStr(vTemp(1)))
        Select Case strSwitch
            Case "S", "SERVER"                  'Servername
                gsServerNameSW = strValue
            Case "U", "USER"                    'Login
                gsLoginSW = strValue
            Case "P", "PW", "PASSWORD"          'Password
                gsPasswordSW = strValue
            Case "DB", "DATABASE"               'Databases (comma-delimited list)
                gvDatabasesSW = Split(strValue, ",")
            Case "DBA", "DBAOPTIONS"            'DBA Scriptiong options. "ALL" or any combination of A, D, L and J
                If UCase$(strValue) = "ALL" Then
                    gbSaveDTSPackagesSW = True
                    gbScriptAlertsSW = True
                    gbScriptServerLoginsSW = True
                    gbScriptAgentJobsSW = True
                    gbScriptBackupDevicesSW = True
                Else
                    If InStr(UCase$(strValue), "D") > 0 Then gbSaveDTSPackagesSW = True
                    If InStr(UCase$(strValue), "A") > 0 Then gbScriptAlertsSW = True
                    If InStr(UCase$(strValue), "L") > 0 Then gbScriptServerLoginsSW = True
                    If InStr(UCase$(strValue), "J") > 0 Then gbScriptAgentJobsSW = True
                    If InStr(UCase$(strValue), "B") > 0 Then gbScriptBackupDevicesSW = True
                End If
            Case "DEL"                          'Number of weeks to keep old files
                giDelWeeksSW = CInt(strValue)
                gbDelOldFilesSW = True
            Case "DIR", "DEST", "DESTDIR"       'Destination directory for the scripts
                gsDestDirSW = strValue
                If Right$(gsDestDirSW, 1) <> "\" Then gsDestDirSW = gsDestDirSW & "\"
            Case "BG", "BACK", "BACKGROUND"
                gbRunUnAttendedSW = True
            
            Case "ScriptDB", "SDB"
                gbScriptDB = CBool(strValue)
            Case "ScriptTables", "STA"
                gbScriptTabs = CBool(strValue)
            Case "ScriptViews", "SVI"
                gbScriptViews = CBool(strValue)
            Case "ScriptSPs", "SSP"
                gbScriptSPs = CBool(strValue)
            Case "ScriptRules", "SRU"
                gbScriptRules = CBool(strValue)
            Case "ScriptDefs", "SDE"
                gbScriptDefs = CBool(strValue)
            Case "ScriptRoles", "SRO"
                gbScriptRoles = CBool(strValue)
            Case "ScriptUsers", "SUS"
                gbScriptUsers = CBool(strValue)
            Case "ScriptUDTs", "SUD"
                gbScriptUDTs = CBool(strValue)
            Case "ScriptFTCs", "SFT"
                gbScriptFTCs = CBool(strValue)
            Case "DBObjsInSepFiles", "SEP"
                gbObjsInSepFiles = CBool(strValue)
            Case "DatePrefixOnFiles", "PRE"
                gbDatePrefix = CBool(strValue)
            Case "InsertScriptTables", "INS"
                gsSelTables = strValue & ","
            Case "ScriptDropsFirst", "SDF"
                gbScriptDrops1st = CBool(strValue)
        End Select
    Next i
    
    DoEvents
    'Determine if the program will be run unattended
    If (Len(gsServerNameSW) = 0 Or Len(gsDestDirSW) = 0) Or _
       (UBound(gvDatabasesSW) = 0 And Not gbRunUnAttendedSW) Then       'GUI
        gbNoGUI = False
        frmMain.Show
    Else        'Unattended
        gbNoGUI = True
        With gobjServer
            .Name = gsServerNameSW
            .ApplicationName = App.Title
            DoEvents
            gsLogFile = gsDestDirSW & Format(Date, "yyyymmdd") & "." & gobjServer.Name & "_Log.txt"
            If Len(gsLoginSW) > 0 Then
                .LoginSecure = False
                .Login = gsLoginSW
                .Password = gsPasswordSW
                StatusMessage "Using Login: " & gsLoginSW
            Else
                .Login = ""
                .Password = ""
                .LoginSecure = True
                StatusMessage "Using Trusted Connection"
            End If
            On Error GoTo ErrHandler
            StatusMessage "Connecting to: " & gobjServer.Name
            .Connect
            DoEvents
            
            GenerateScripts gsServerNameSW, gsDestDirSW, gvDatabasesSW, .LoginSecure, _
                gsLoginSW, gsPasswordSW, gbDelOldFilesSW, giDelWeeksSW, _
                gbScriptDB, gbScriptTabs, gbScriptViews, gbScriptSPs, _
                gbScriptRules, gbScriptDefs, gbScriptRoles, gbScriptFTCs, _
                gbScriptUsers, gbScriptUDTs, _
                gbSaveDTSPackagesSW, gbScriptAlertsSW, gbScriptAgentJobsSW, _
                gbScriptServerLoginsSW, gbScriptBackupDevicesSW, _
                gbObjsInSepFiles, gbDatePrefix, gsSelTables, gbScriptDrops1st
                
            DoEvents
            SaveLog gsLogFile
            .Close
            DoEvents
        End With
        Set gobjServer = Nothing
    End If
    Exit Sub
    
ErrHandler:
    StatusMessage "Error: 0x" & Hex$(Err.Number) & vbTab & Error(Err.Number)
    StatusMessage "Failed to connect to the server. Program aborted."
    SaveLog gsLogFile
End Sub

