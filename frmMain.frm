VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Generator Version 2.0"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   4500
      Left            =   4770
      TabIndex        =   9
      Top             =   -45
      Width           =   2415
      Begin VB.CheckBox chkCustom 
         Caption         =   "Create Custom Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cmbKeyType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         TabIndex        =   17
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox txtCustomChar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtNumPass 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Text            =   "100"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "Key Type"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   16
         Top             =   3480
         Width           =   810
      End
      Begin VB.Label lblCust 
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Mask Key"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   15
         Top             =   2640
         Width           =   1530
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   20
         X2              =   2400
         Y1              =   1940
         Y2              =   1940
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   20
         X2              =   2400
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passwords To Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2130
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicated Passwords"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2130
      End
      Begin VB.Label lblDupes 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   700
         TabIndex        =   11
         Top             =   1440
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1500
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   4695
      Begin VB.CommandButton cmdSet 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set"
         Height          =   285
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtCustom 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.ComboBox cmbMask 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   400
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Mask"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Mask"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   4695
      Begin VB.CommandButton cmdGen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generate"
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2480
         Width           =   1095
      End
      Begin VB.ListBox lstPass 
         Height          =   2595
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":0002
         TabIndex        =   1
         Top             =   260
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'NEW UPDATES:
'   NOW GENERATES "SPECIAL" CHARACTERS
'   NOW IT WILL TELL YOU HOW MANY DUPLICATE PASSWORDS WERE DELETED
'   YOU CAN CREATE YOUR OWN CUSTOM MASK KEYS (#,X,O)(Not the string itself)


'HOW TO USE A CUSTOM KEY MASK:
' CLICK "CREATE CUSTOM KEY"
' PUT THE KEY YOU WANT IN THE KEY FIELD. SAY YOU WANT "T" TO BE YOU KEY
' THEN SELECT THE TYPE. "T" COULD BE NUMERIC, ALPHABETIC, ETC.
' THEN JUST HIT GENERATE AND IT AUTOMATICALLY CHECKS THAT TYPE NOW.

Private Sub chkCustom_Click()
'enable/disable the controls to create the custom key
If chkCustom.Value Then
 lblCust.Enabled = True
 txtCustomChar.Enabled = True
 lblType.Enabled = True
 cmbKeyType.Enabled = True
Else
 txtCustomChar = ""
 lblCust.Enabled = False
 txtCustomChar.Enabled = False
 lblType.Enabled = False
 cmbKeyType.Enabled = False
End If
End Sub

Private Sub cmdGen_Click()
'Generate the passwords
GeneratePasswords lstPass, CLng(txtNumPass), cmbMask.Text, True, chkCustom.Value, txtCustomChar, cmbKeyType.ListIndex
lblDupes.Caption = GetNumberOfIdenticalUniques
End Sub

Private Sub cmdSet_Click()
'Add the custom mask to the list box and select it
If txtCustom <> "" Then
 cmbMask.AddItem txtCustom
 cmbMask.ListIndex = cmbMask.ListCount - 1
 txtCustom = ""
End If
End Sub

Private Sub Form_Load()

'For our password mask we are going to set it up like this:
'
'# = Numbers
'X = Letters
'
'     (Not Zero's, their capital o's)
'New: O = Special Characters (Not really recommended in a real
'                             app. since most users dont know
'                             AnsiCharacters like ¶,Æ,¥,†,‰, ect.)
'This can be whatever you want

cmbMask.AddItem "####-####-####-####", 0
cmbMask.AddItem "XXXX-XXXX-XXXX-XXXX", 1
cmbMask.AddItem "XXXX-XXXX-####-####", 2
cmbMask.AddItem "####-####-XXXX-XXXX", 3
cmbMask.AddItem "####-XXXX-####-XXXX", 4
cmbMask.AddItem "XXXX-####-XXXX-####", 5
cmbMask.AddItem "CDKEY-####-####-####-####", 6
cmbMask.AddItem "CDKEY-XXXX-XXXX-XXXX-XXXX", 7
cmbMask.AddItem "CDKEY-####-####-XXXX-XXXX", 8
cmbMask.AddItem "CDKEY-XXXX-XXXX-####-####", 9
cmbMask.AddItem "CDKEY-####-XXXX-####-XXXX", 10
cmbMask.AddItem "CDKEY-XXXX-####-XXXX-####", 11
cmbMask.AddItem "KEY-####-####-####-####", 12
cmbMask.AddItem "KEY-XXXX-XXXX-XXXX-XXXX", 13
cmbMask.AddItem "KEY-####-####-XXXX-XXXX", 14
cmbMask.AddItem "KEY-XXXX-XXXX-####-####", 15
cmbMask.AddItem "KEY-####-XXXX-####-XXXX", 16
cmbMask.AddItem "KEY-XXXX-####-XXXX-####", 17

'New Password uses(DO NOT USE FOR PASSWORDS, YOU'LL SEE WHY)
'Theres are just to show you how to add custom mask keys
cmbMask.AddItem "KEY-OOOO-####-OOOO-####", 18
cmbMask.AddItem "KEY-OOOO-XXXX-OOOO-XXXX", 19
cmbMask.AddItem "KEY-OOOO-OOOO-OOOO-OOOO", 20
cmbMask.AddItem "OOOOOOOOOOOOOOOOOOOOOOO", 21
cmbMask.AddItem "XOXO-#O#O-XOXO-#O#O", 22
cmbMask.AddItem "XXXX-OOOO-OOOO-####", 23
cmbMask.ListIndex = 0

'Setup our custom key type combo box
cmbKeyType.AddItem "Alphabetic", 0
cmbKeyType.AddItem "Numeric", 1
cmbKeyType.AddItem "Special Characters", 2
cmbKeyType.ListIndex = 0
End Sub
