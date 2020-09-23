VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQLScripter"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7950
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "&Clear Log"
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      ToolTipText     =   "Clear the Log."
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveLog 
      Caption         =   "&Save Log"
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      ToolTipText     =   "Save the Log."
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      ToolTipText     =   "Start generating the scripts."
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      ToolTipText     =   "Exit the program (by special request :)"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Disconnect"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      ToolTipText     =   "Disconnect from the currently selected server."
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame fraServerOptions 
      Caption         =   "Server Scripting options (sysadmins only):"
      Height          =   1935
      Left            =   240
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkOptionBackup 
         Caption         =   "Script Backup Devices"
         Height          =   375
         Left            =   3960
         TabIndex        =   42
         ToolTipText     =   "Script the Backup Devices."
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkOptionDTS 
         Caption         =   "Save DTS Packages"
         Height          =   375
         Left            =   240
         TabIndex        =   36
         ToolTipText     =   "Save DTS Packages as Structured Storage File."
         Top             =   360
         Width           =   3135
      End
      Begin VB.CheckBox chkOptionJobs 
         Caption         =   "Script SQL Agent Jobs"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "Script the MS SQL Server Agent Jobs."
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkOptionLogins 
         Caption         =   "Script Server Logins"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Script the Server Logins."
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox chkOptionAlerts 
         Caption         =   "Script Operators/Alerts"
         Height          =   375
         Left            =   3960
         TabIndex        =   33
         ToolTipText     =   "Script the Operators and Alerts."
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraDBOptions 
      Caption         =   "Database Scripting options:"
      Height          =   3975
      Left            =   240
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkScriptDropsFirst 
         Caption         =   "Script database object drops first"
         Height          =   255
         Left            =   3240
         TabIndex        =   48
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CheckBox chkScriptDatabase 
         Caption         =   "Script the Database"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdSelectTables 
         Caption         =   "..."
         Height          =   315
         Left            =   6600
         TabIndex        =   46
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkInsertScript 
         Caption         =   "Create INSERT scripts for selected tables"
         Height          =   255
         Left            =   3240
         TabIndex        =   45
         Top             =   1920
         Width           =   3255
      End
      Begin VB.CheckBox chkDSFilePrefix 
         Caption         =   "Prefix files with date stamp"
         Height          =   315
         Left            =   3240
         TabIndex        =   44
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkSingleFile 
         Caption         =   "Script each Table, View and SP in separate files"
         Height          =   375
         Left            =   3240
         TabIndex        =   43
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox chkDBOptionRules 
         Caption         =   "Script the Rules"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox chkDBOptionDefaults 
         Caption         =   "Script the Defaults"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkDBOptionSPs 
         Caption         =   "Script the Stored Procedures"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox chkDBOptionViews 
         Caption         =   "Script the Views"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   6015
      End
      Begin VB.CheckBox chkDBOptionTables 
         Caption         =   "Script the Tables (including indexes, triggers and bindings)"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   6015
      End
      Begin VB.CheckBox chkDBOptionUDT 
         Caption         =   "Script the User Defined Datatypes"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   2895
      End
      Begin VB.CheckBox chkDBOptionUsers 
         Caption         =   "Script the Database Users"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CheckBox chkDBOptionFullText 
         Caption         =   "Include Full Text Catalogs and Indexes with DB and Table scripts"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3480
         Width           =   6015
      End
      Begin VB.CheckBox chkDBOptionRoles 
         Caption         =   "Script the Database Roles"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   2415
      End
   End
   Begin VB.Frame fraScripting 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   7455
      Begin VB.CommandButton cmdFetchDBs 
         Caption         =   "Connect && &Fetch Databases"
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         ToolTipText     =   "Connect to the server and fetch the list of available databases."
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ListBox lstDatabases 
         Height          =   3435
         Left            =   2640
         Style           =   1  'Checkbox
         TabIndex        =   13
         ToolTipText     =   "Select (check) the databases that you want to create scripts for."
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         ToolTipText     =   "Type the Login to use for SQL Server Authentication mode."
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboServers 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   2640
         List            =   "frmMain.frx":0444
         TabIndex        =   3
         ToolTipText     =   "Choose a server from the list or type the server-name in the box."
         Top             =   240
         Width           =   4695
      End
      Begin VB.Frame fraAuthentication 
         Caption         =   "Authentication mode:"
         Height          =   855
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Choose your authentication mode."
         Top             =   720
         Width           =   3255
         Begin VB.OptionButton optAuthMode 
            Caption         =   "NT (Trusted)"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "NT Authentication."
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optAuthMode 
            Caption         =   "SQL Server"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            ToolTipText     =   "MS SQL Server Authentication."
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5160
         PasswordChar    =   "*"
         TabIndex        =   10
         ToolTipText     =   "Type the password to use for SQL Server Authentication mode."
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtDirectory 
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         ToolTipText     =   "Destination directory for the script-files."
         Top             =   5520
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowseDir 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   19
         ToolTipText     =   "Browse for a destination directory for the script files."
         Top             =   5520
         Width           =   375
      End
      Begin VB.TextBox txtDelWeeks 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   630
         TabIndex        =   15
         Text            =   "4"
         ToolTipText     =   "Number of weeks to keep old script-files."
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox chkDelWeeks 
         Caption         =   "Delete files older then:"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         ToolTipText     =   "Check to delete old script-files."
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblDirectory 
         Caption         =   "Destination directory:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label lblServer 
         Caption         =   "Server:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblDatabases 
         Caption         =   "Databases:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblLogin 
         Caption         =   "Login:"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblDelWeeks 
         Caption         =   "weeks."
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   2805
         Width           =   975
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   5895
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   7455
   End
   Begin MSComctlLib.TabStrip tsPage 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11456
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Scripting"
            Key             =   "Scripting"
            Object.ToolTipText     =   "On this tab you select what you want to script."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            Object.ToolTipText     =   "On this tab you can select the scripting options."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Log"
            Key             =   "Log"
            Object.ToolTipText     =   "This tab contains the log of the scripting actions."
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   26
      ToolTipText     =   "Current activity of SQLScripter."
      Top             =   7200
      Width           =   7935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
DefInt A-Z

Private mobjApplication As New SQLDMO.Application

Private Sub cboServers_Click()
    cmdDisconnect_Click
    If cboServers.ListIndex <> -1 Then
        gobjServer.Name = cboServers.List(cboServers.ListIndex)
        If CBool(cboServers.ItemData(cboServers.ListIndex)) Then
            optAuthMode(0).Value = True
            optAuthMode_Click 0
        Else
            optAuthMode(1).Value = True
            optAuthMode_Click 1
        End If
    End If
End Sub

Private Sub chkDelWeeks_Click()
    If chkDelWeeks.Value = 0 Then
        With txtDelWeeks
            .BackColor = vbButtonFace
            .Enabled = False
        End With
    Else
        With txtDelWeeks
            .Enabled = True
            .BackColor = vbWindowBackground
        End With
    End If
End Sub

Private Sub chkInsertScript_Click()

    If chkInsertScript.Value = vbUnchecked Then
        gsSelTables = vbNullString
    End If
    
End Sub

Private Sub cmdBrowseDir_Click()
    txtDirectory.Text = BrowseForFolder("Select a directory")
End Sub

Private Sub cmdClearLog_Click()
    txtLog.Text = ""
    gsUnattendedLog = ""
End Sub

Private Sub cmdDisconnect_Click()
    On Error Resume Next
    With gobjServer
        If .VerifyConnection(SQLDMOConn_CurrentState) Then
            StatusMessage "Disconnecting from server " & gobjServer.Name
            .DisConnect
        End If
    End With
    With lstDatabases
        .Visible = False
        .Clear
    End With
    fraServerOptions.Visible = False
    cmdFetchDBs.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFetchDBs_Click()
    Dim objDatabase As SQLDMO.Database
    Dim i As Integer
    Dim qryResults As SQLDMO.QueryResults
    
    If Len(cboServers.Text) = 0 Then
        MsgBox "You'll have to choose or enter a servername first.", vbExclamation, "Error!"
        cboServers.SetFocus
    Else
        Me.MousePointer = vbArrowHourglass
        With gobjServer
            .Name = cboServers.Text
            DoEvents
            .ApplicationName = App.Title
            If optAuthMode(1).Value = True Then
                .LoginSecure = False
                .Login = txtLogin.Text
                .Password = txtPassword.Text
                StatusMessage "Using Login " & txtLogin.Text
            Else
                .Login = ""
                .Password = ""
                .LoginSecure = True
                StatusMessage "Using Trusted Connection"
            End If
            On Error GoTo ErrHandler
            StatusMessage "Connecting to " & gobjServer.Name
            .Connect
            DoEvents
            If .Issysadmin Then
                StatusMessage "User is sysadmin"
                fraServerOptions.Enabled = True
'                chkOptionDTS.Value = 1
'                chkOptionAlerts.Value = 1
'                chkOptionJobs.Value = 1
'                chkOptionLogins.Value = 1
'                chkOptionBackup.Value = 1
            Else
                StatusMessage "User is not a sysadmin"
                fraServerOptions.Enabled = False
                chkOptionDTS.Value = 2
                chkOptionAlerts.Value = 2
                chkOptionJobs.Value = 2
                chkOptionLogins.Value = 2
                chkOptionBackup.Value = 2
            End If
            StatusMessage "Fetching databases:"
            On Error GoTo 0
            For Each objDatabase In gobjServer.Databases
                'This freaky "On Error" hack needs to be used since the .IsUser method
                ' will raise an error (instead of returning False as it should) when called
                ' with an invalid login.
                On Error GoTo ErrNoUser
                If objDatabase.IsUser(.Login) Or .Issysadmin Then
ResumeSaUser:
                    If ((Not objDatabase.SystemObject) Or .Issysadmin) Then
                        i = i + 1
                        StatusMessage "- Found Database " & CStr(i) & ": " & objDatabase.Name
                        lstDatabases.AddItem objDatabase.Name
                    End If
                End If
ResumeNoUser:
            DoEvents
            Next objDatabase
            StatusMessage CStr(i) & " Databases found."
        End With
        Me.MousePointer = vbNormal
        cmdFetchDBs.Visible = False
        lstDatabases.Visible = True
        StatusMessage "Ready."
    End If
    Exit Sub
    
ErrHandler:
    StatusMessage "Error: 0x" & Hex$(Err.Number) & vbTab & Error(Err.Number)
    If Err.Number = &H80044818 Then
        MsgBox "Your login credentials are not recognised by the server.", vbExclamation, "Error!"
    Else
        MsgBox "Error: 0x" & Hex$(Err.Number) & vbNewLine & Err.Description, vbExclamation, "Error!"
    End If
    Me.MousePointer = vbNormal
    Exit Sub
    
ErrNoUser:
    If gobjServer.Issysadmin Then
        Resume ResumeSaUser
    Else
        Resume ResumeNoUser
    End If
End Sub

Private Sub cmdGenerate_Click()
    Dim strServerName As String
    Dim strDatabases As String
    Dim vDatabases As Variant
    Dim strDir As String
    Dim msgResult As VbMsgBoxResult
    Dim i As Integer
    
    If chkInsertScript.Value = vbUnchecked Then
        gsSelTables = vbNullString
    End If
    
    If Len(txtDirectory.Text) = 0 Then
        MsgBox "Choose an output directory for the scripts first.", vbExclamation, "Error!"
        On Error Resume Next
        txtDirectory.SetFocus
        Exit Sub
    End If
    
    If lstDatabases.SelCount = 0 And chkOptionDTS.Value <> 1 And chkOptionJobs.Value <> 1 And _
       chkOptionLogins.Value <> 1 And chkOptionAlerts.Value <> 1 And chkOptionBackup.Value <> 1 Then
        MsgBox "You haven't selected anything to script!", vbExclamation, "Error!"
        Exit Sub
    End If
    
    Me.MousePointer = vbArrowHourglass
    fraScripting.Enabled = False
    fraServerOptions.Enabled = False
    fraDBOptions.Enabled = False
    cmdExit.Enabled = False
    cmdDisconnect.Enabled = False
    cmdGenerate.Enabled = False
    tsPage.Tabs.Item("Log").Selected = True
    strDir = Trim$(txtDirectory.Text)
    If Right$(strDir, 1) <> "\" Then strDir = strDir & "\"
    txtDirectory.Text = strDir
    strServerName = gobjServer.Name
    gsLogFile = strDir & Format(Date, "yyyymmdd") & "." & strServerName & "_Log.txt"
    For i = 0 To lstDatabases.ListCount - 1
        If lstDatabases.Selected(i) Then strDatabases = strDatabases & "," & lstDatabases.List(i)
    Next i
    vDatabases = Split(Mid$(strDatabases, 2, Len(strDatabases)), ",")
    
    GenerateScripts strServerName, strDir, vDatabases, optAuthMode(0).Value, Trim$(txtLogin.Text), _
                    txtPassword.Text, chkDelWeeks.Value = 1, CInt(txtDelWeeks.Text), _
                    chkScriptDatabase.Value = 1, _
                    chkDBOptionTables.Value = 1, chkDBOptionViews.Value = 1, _
                    chkDBOptionSPs.Value = 1, chkDBOptionRules.Value = 1, _
                    chkDBOptionDefaults.Value = 1, chkDBOptionRoles.Value = 1, _
                    chkDBOptionFullText.Value = 1, chkDBOptionUsers.Value = 1, _
                    chkDBOptionUDT.Value = 1, chkOptionDTS.Value = 1, _
                    chkOptionAlerts.Value = 1, chkOptionJobs.Value = 1, _
                    chkOptionLogins.Value = 1, chkOptionBackup.Value = 1, _
                    chkSingleFile.Value = 1, chkDSFilePrefix.Value = 1, _
                    gsSelTables, chkScriptDropsFirst.Value = 1
                    
    DoEvents
    fraScripting.Enabled = True
    fraServerOptions.Enabled = True
    fraDBOptions.Enabled = True
    Me.MousePointer = vbNormal
    cmdExit.Enabled = True
    cmdDisconnect.Enabled = True
    cmdGenerate.Enabled = True
    MsgBox "Scripting done.", vbInformation, "Information"
End Sub

Private Sub cmdSaveLog_Click()
    SaveLog gsLogFile
    MsgBox "Log saved as:" & vbNewLine & gsLogFile, vbInformation, "Information."
End Sub

Private Sub cmdSelectTables_Click()
    Dim objTab  As SQLDMO.Table
    Dim sDBName As String
    Dim i       As Integer
    ' Dim sTemp   As String
    
    ' use the global gobjServer object
    ' get the database name from the first selected entry in the list: lstDatabases.Selected(0)
    If lstDatabases.SelCount > 0 Then
    
        For i = 0 To lstDatabases.ListCount
            If lstDatabases.Selected(i) Then
                sDBName = lstDatabases.List(i)
                Exit For
            End If
        Next
        
        If gobjServer.Databases(sDBName).Tables.Count > 0 Then
            
            Load frmSelectTables
            
            With frmSelectTables
                
                .lstTables.Clear
                
                For Each objTab In gobjServer.Databases(sDBName).Tables
                
                    ' add each of the tables in the database to lstTables on frmSelectTables
                    If Not objTab.SystemObject Then
                        .lstTables.AddItem objTab.Name
                        ' sTemp = sTemp & objTab.Name & vbNewLine
                        If InStr(1, gsSelTables, objTab.Name & ",") <> 0 Then
                            .lstTables.Selected(.lstTables.NewIndex) = True
                        End If
                        
                    End If
                    
                Next
                
                ' when the form is hidden grab the list of selected tables and store it in a mod level var
                .Show vbModal
                
                If .lstTables.ListCount > 0 Then
                    If .lstTables.SelCount > 0 Then
                        
                        gsSelTables = vbNullString
                        
                        For i = 0 To .lstTables.ListCount - 1
                            If .lstTables.Selected(i) Then
                                gsSelTables = gsSelTables & .lstTables.List(i) & ","
                            End If
                        Next
                        
                        If gsSelTables <> vbNullString Then
                            chkInsertScript.Value = vbChecked
                        End If
                       
                    Else
                        gsSelTables = vbNullString
                        chkInsertScript.Value = vbUnchecked
                    End If
                End If
                
            End With
            
        Else
            MsgBox "There are no user tables in the selected database", vbOKOnly + vbInformation
        End If
    
    Else
        MsgBox "Please select a database first", vbOKOnly + vbInformation
    End If
        
End Sub

Private Sub Form_Load()
    Dim objServerGroup As SQLDMO.ServerGroup
    Dim objRegisteredServer As SQLDMO.RegisteredServer
    Dim i As Integer, j As Integer
    
    For Each objServerGroup In mobjApplication.ServerGroups
        
        For Each objRegisteredServer In objServerGroup.RegisteredServers
            
            cboServers.AddItem objRegisteredServer.Name
            cboServers.ItemData(cboServers.NewIndex) = CStr(objRegisteredServer.UseTrustedConnection)
        
        Next objRegisteredServer
    
    Next objServerGroup
    
    ' set options specified on command line
    If Len(gsServerNameSW) > 0 Then
        cboServers.Text = gsServerNameSW
        If Len(gsLoginSW) > 0 Then
            optAuthMode(1).Value = True
            optAuthMode_Click 1
            txtLogin.Text = gsLoginSW
            txtPassword.Text = gsPasswordSW
        Else
            optAuthMode(0).Value = True
            optAuthMode_Click 0
        End If
        If gbDelOldFilesSW Then
            chkDelWeeks.Value = 1
            txtDelWeeks.Text = CStr(giDelWeeksSW)
        Else
            chkDelWeeks.Value = 0
        End If
        txtDirectory.Text = gsDestDirSW
        cmdFetchDBs_Click
        
        ' set main options
        chkScriptDatabase.Value = Abs(CInt(gbScriptDB))
        chkDBOptionTables.Value = Abs(CInt(gbScriptTabs))
        chkDBOptionViews.Value = Abs(CInt(gbScriptViews))
        chkDBOptionSPs.Value = Abs(CInt(gbScriptSPs))
        chkDBOptionRules.Value = Abs(CInt(gbScriptRules))
        chkDBOptionDefaults.Value = Abs(CInt(gbScriptDefs))
        chkDBOptionRoles.Value = Abs(CInt(gbScriptRoles))
        chkDBOptionUsers.Value = Abs(CInt(gbScriptUsers))
        chkDBOptionUDT.Value = Abs(CInt(gbScriptUDTs))
        chkDBOptionFullText.Value = Abs(CInt(gbScriptFTCs))
        chkSingleFile.Value = Abs(CInt(gbObjsInSepFiles))
        chkDSFilePrefix.Value = Abs(CInt(gbDatePrefix))
        chkScriptDropsFirst.Value = Abs(CInt(gbScriptDrops1st))
        
        If gsSelTables <> vbNullString Then
            chkInsertScript.Value = vbChecked
        Else
            chkInsertScript.Value = vbUnchecked
        End If
        
        On Error GoTo ErrHandler
        If gobjServer.Issysadmin Then
            chkOptionDTS.Value = Abs(CInt(gbSaveDTSPackagesSW))
            chkOptionAlerts.Value = Abs(CInt(gbScriptAlertsSW))
            chkOptionJobs.Value = Abs(CInt(gbScriptAgentJobsSW))
            chkOptionLogins.Value = Abs(CInt(gbScriptServerLoginsSW))
            chkOptionBackup.Value = Abs(CInt(gbScriptBackupDevicesSW))
            If Not IsEmpty(gvDatabasesSW) Then
                For i = 0 To lstDatabases.ListCount - 1
                    For j = 0 To UBound(gvDatabasesSW)
                        If lstDatabases.List(i) = gvDatabasesSW(j) Then
                            lstDatabases.Selected(i) = True
                            Exit For
                        End If
                    Next j
                Next i
            End If
        End If
    End If
    Exit Sub
    
ErrHandler:
    MsgBox "Cannot connect to server " & cboServers.Text, vbExclamation, "Error!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gobjServer.Close
    Set gobjServer = Nothing
    Set mobjApplication = Nothing
End Sub

Private Sub optAuthMode_Click(Index As Integer)
    If Index = 0 Then
        With txtLogin
            .BackColor = vbButtonFace
            .Enabled = False
        End With
        With txtPassword
            .BackColor = vbButtonFace
            .Enabled = False
        End With
    Else
        With txtLogin
            .Enabled = True
            .BackColor = vbWindowBackground
        End With
        With txtPassword
            .Enabled = True
            .BackColor = vbWindowBackground
        End With
    End If
End Sub

Private Sub tsPage_Click()
    Select Case tsPage.SelectedItem.Key
        Case "Scripting"
            txtLog.Visible = False
            fraServerOptions.Visible = False
            fraDBOptions.Visible = False
            fraScripting.Visible = True
        Case "Log"
            fraScripting.Visible = False
            fraServerOptions.Visible = False
            fraDBOptions.Visible = False
            txtLog.Visible = True
        Case "Options"
            fraScripting.Visible = False
            txtLog.Visible = False
            fraDBOptions.Visible = True
            fraServerOptions.Visible = True
    End Select
End Sub

