VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AddToPOPUPMENUS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize Windows POPUP Menus"
   ClientHeight    =   3900
   ClientLeft      =   3255
   ClientTop       =   1965
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5775
   Begin VB.TextBox txtExCV 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2340
      TabIndex        =   7
      ToolTipText     =   "String to display"
      Top             =   2115
      Width           =   2445
   End
   Begin VB.CheckBox chkEx 
      Caption         =   "Include me in ""EXPLORER"" POPUP Menu and ""START"" Button POPUP Menu"
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   765
      TabIndex        =   6
      Top             =   1575
      Width           =   3390
   End
   Begin VB.CheckBox chkMC 
      Caption         =   "Include me in ""My Computer"" POPUP Menu"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   765
      TabIndex        =   3
      Top             =   360
      Width           =   3480
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   3375
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   4365
      TabIndex        =   2
      Top             =   3375
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   360
      Left            =   2115
      TabIndex        =   0
      Top             =   3375
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Height          =   3030
      Left            =   405
      TabIndex        =   9
      Top             =   90
      Width           =   4965
      Begin VB.CommandButton cmdExCmdB 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   240
         Left            =   4080
         TabIndex        =   15
         Top             =   2500
         Width           =   280
      End
      Begin VB.CommandButton cmdMCCmdB 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   240
         Left            =   4080
         TabIndex        =   14
         Top             =   1060
         Width           =   280
      End
      Begin VB.TextBox txtExCmd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1935
         TabIndex        =   8
         ToolTipText     =   "Action/program to run"
         Top             =   2475
         Width           =   2445
      End
      Begin VB.TextBox txtMCCV 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   4
         ToolTipText     =   "String to display"
         Top             =   585
         Width           =   2400
      End
      Begin VB.TextBox txtMCCmd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1980
         TabIndex        =   5
         ToolTipText     =   "Action/program to run"
         Top             =   1035
         Width           =   2400
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Command"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   2520
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Command"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   720
         TabIndex        =   12
         Top             =   1035
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Current Value"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current Value"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   630
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   585
      Top             =   900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L1nd@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   0
      Left            =   405
      TabIndex        =   16
      Top             =   3240
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L1nd@"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Index           =   1
      Left            =   450
      TabIndex        =   17
      Top             =   3285
      Width           =   990
   End
End
Attribute VB_Name = "AddToPOPUPMENUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(c)L1nd@
'eMail:      linda.69@mailcity.com
'Purpose: Add program to Windows popup menus.
  
'Deleting of registry keys not implemented
'use Regedit instead or extend this program.

'You know the drill.

Option Explicit

Dim strTemp$

Const strKeyName = "THISIADDED"              'this is the main key
Const strDisplayed = "My &Program"             'and name
                                                                    'change these with ur own

Private Sub Form_Load()
  GetSettings
  Dialog.Filter = "Executable files (*.exe;*.com)|*.exe;*.com|All Files|*.*"
End Sub

Sub GetSettings()
  'This is where you will put your "String" to be include
  'in "My Computer" POPUP Menu
  strTemp = _
    GetStringValue("HKEY_CLASSES_ROOT" + _
                            "\CLSID" + _
                            "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                            "\SHELL\" + _
                            strKeyName, _
                            "")                              'Just ("") to get the default value which
                                                             'is the string that is shown in the menu
  
  'Check if we are already included.
  'Note: Values are NULL Terminated
  chkMC.Value = Abs(strTemp = strDisplayed + Chr(0))
  
  'Key not yet there
  If strTemp = "" Then strTemp = "Key Not Found"
  'Display gotten value
  txtMCCV.Text = strTemp

  
  'This is where you will put the "Action" to be done
  'or "Program" to be run when "String" is clicked.
  strTemp = _
      GetStringValue("HKEY_CLASSES_ROOT" + _
                              "\CLSID" + _
                              "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                              "\SHELL\" + _
                               strKeyName + _
                              "\Command", _
                              "")                  'Just ("") to get the default value.
  
  'Key not yet there
  If strTemp = "" Then strTemp = "Key Not Found"
  'Display gotten value
  txtMCCmd.Text = strTemp

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  'This is where you will put your "String" to be include in
  'the "EXPLORER" POPUP Menu AND "START" Button POPUP Menu
  strTemp = _
    GetStringValue("HKEY_CLASSES_ROOT" + _
                            "\DIRECTORY" + _
                            "\SHELL\" + _
                            strKeyName, _
                            "")                             'Just ("") to get the default value which
                                                            'is the string that is shown in the menu
  
  'Check if we are already included.
  'Note: Values are NULL Terminated
  chkEx.Value = Abs(strTemp = strDisplayed + Chr(0))
  
  'Key not yet there
  If strTemp = "" Then strTemp = "Key Not Found"
  'Display gotten value
  txtExCV.Text = strTemp

  'This is where you will put the "Action" to be done
  'or "Program" to be run when "String" is clicked.
  strTemp = _
      GetStringValue("HKEY_CLASSES_ROOT" + _
                              "\DIRECTORY" + _
                              "\SHELL\" + _
                               strKeyName + _
                              "\Command", _
                              "")                  'Just ("") to get the default value.
  
  'Key not yet there
  If strTemp = "" Then strTemp = "Key Not Found"
  'Display gotten value
  txtExCmd.Text = strTemp

End Sub

Sub SaveSettings()

'Sorry, not fully customized.

  If (chkMC.Value = Checked) Then
    If (txtMCCV.Text = "Key Not Found") Then
      CreateKey "HKEY_CLASSES_ROOT" + _
                       "\CLSID" + _
                       "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                       "\SHELL\" + _
                        strKeyName                  'I Hard-coded the key name
                                                            'you can change with your own
      CreateKey "HKEY_CLASSES_ROOT" + _
                       "\CLSID" + _
                       "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                       "\SHELL\" + _
                        strKeyName + _
                        "\Command"                   'Dont change THIS!!!
    End If
    
    If ((txtMCCmd.Text <> "Key Not Found") And _
        (txtMCCmd.Text <> "")) Then
       SetStringValue "HKEY_CLASSES_ROOT" + _
                              "\CLSID" + _
                              "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                              "\SHELL\" + _
                              strKeyName, _
                              "", _
                              strDisplayed                 'I hard-coded this one
                                                                  'you can change with your own
    
       SetStringValue "HKEY_CLASSES_ROOT" + _
                              "\CLSID" + _
                              "\{20D04FE0-3AEA-1069-A2D8-08002B30309D}" + _
                              "\SHELL\" + _
                              strKeyName + _
                              "\Command", _
                              "", _
                              txtMCCmd.Text
    Else
      MsgBox "Please fill up Command Box" + Chr(13) + _
                   "with the path of your program." + Chr(13) + _
                   "ie.    c:\windows\notepad.exe     ", vbCritical, "Error"
      txtMCCmd.SetFocus
      Exit Sub
    End If
  End If
  
  If (chkEx.Value = Checked) Then
    If (txtExCV.Text = "Key Not Found") Then
      CreateKey "HKEY_CLASSES_ROOT" + _
                      "\DIRECTORY" + _
                      "\SHELL\" + _
                       strKeyName                  'I Hard-coded the key name
                                                            'you can change with your own
                                                            
      CreateKey "HKEY_CLASSES_ROOT" + _
                      "\DIRECTORY" + _
                      "\SHELL\" + _
                      strKeyName + _
                      "\Command"                    'Dont change THIS!!!
    End If
    
    If ((txtExCmd.Text <> "Key Not Found") And _
        (txtExCmd.Text <> "")) Then
       SetStringValue "HKEY_CLASSES_ROOT" + _
                              "\DIRECTORY" + _
                              "\SHELL\" + _
                              strKeyName, _
                              "", _
                             strDisplayed            'I hard-coded this one
                                                            'you can change with your own
    
       SetStringValue "HKEY_CLASSES_ROOT" + _
                              "\DIRECTORY" + _
                              "\SHELL\" + _
                              strKeyName + _
                              "\Command", _
                              "", _
                              txtExCmd.Text
    Else
      MsgBox "Please fill up Command Box" + Chr(13) + _
                   "with the path of your program." + Chr(13) + _
                   "ie.    c:\windows\notepad.exe     ", vbCritical, "Error"
      txtExCmd.SetFocus
    End If
  End If
  GetSettings          'show the result
End Sub

Private Sub Command1_Click()
  SaveSettings
  End
End Sub

Private Sub Command2_Click()
  End
End Sub

Private Sub Command3_Click()
  SaveSettings
End Sub

Private Sub chkEx_Click()
'  txtExCV.Enabled = -chkEx.Value
  txtExCmd.Enabled = -chkEx.Value
  cmdExCmdB.Enabled = -chkEx.Value
End Sub

Private Sub chkMC_Click()
'  txtMCCV.Enabled = -chkMC.Value
  txtMCCmd.Enabled = -chkMC.Value
  cmdMCCmdB.Enabled = -chkMC.Value
End Sub

Private Sub txtExCmd_KeyDown(KeyCode As Integer, Shift As Integer)
  CmdEnabled
End Sub

Private Sub txtMCCmd_KeyDown(KeyCode As Integer, Shift As Integer)
  CmdEnabled
End Sub

Private Sub cmdExCmdB_Click()
  Dialog.Action = 1
  If Dialog.filename <> "" Then _
    txtExCmd.Text = Dialog.filename: _
    CmdEnabled
End Sub

Private Sub cmdMCCmdB_Click()
  Dialog.Action = 1
  If Dialog.filename <> "" Then _
    txtMCCmd.Text = Dialog.filename: _
    CmdEnabled
End Sub

Sub CmdEnabled()
  Command1.Enabled = True
  Command3.Enabled = True
End Sub
