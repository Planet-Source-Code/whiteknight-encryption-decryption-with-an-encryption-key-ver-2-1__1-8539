VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   569
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Encryption Steps"
      Height          =   3615
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   8415
      Begin VB.TextBox txt_stp4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "frmMain.frx":0000
         Top             =   3000
         Width           =   8055
      End
      Begin VB.TextBox txt_stp3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmMain.frx":0021
         Top             =   2160
         Width           =   8055
      End
      Begin VB.TextBox txt_stp2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmMain.frx":0042
         Top             =   1320
         Width           =   8055
      End
      Begin VB.TextBox txt_stp1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "frmMain.frx":0063
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label3 
         Caption         =   "Step 4 of encryption"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   7935
      End
      Begin VB.Label Label3 
         Caption         =   "Step 3 of encryption"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   7935
      End
      Begin VB.Label Label3 
         Caption         =   "Step 2 of encryption"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   7935
      End
      Begin VB.Label Label3 
         Caption         =   "Step 1 of encryption"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.TextBox txt_Temp 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Width           =   8295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encrypt / Decrypt"
      Height          =   2055
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton bttn_Decrypt 
         Caption         =   "Decrypt String"
         Height          =   255
         Left            =   6600
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton bttn_encrypt 
         Caption         =   "Encrypt String"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txt_Decr 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   4320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txt_norm 
         Height          =   675
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txt_Key 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label4 
         Caption         =   "Decrypted String:"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Unencrypted String:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Encryption Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu MNU_File 
      Caption         =   "&File"
      Begin VB.Menu MNU_About 
         Caption         =   "&About"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub bttn_Decrypt_Click()
        txt_stp1.Text = ""
        txt_stp2.Text = ""
        txt_stp3.Text = ""
        txt_stp4.Text = ""
txt_Decr.Text = Decrypt(txt_Decr.Text, txt_Key.Text)
End Sub

Private Sub bttn_encrypt_Click()
      txt_stp1.Text = ""
      txt_stp2.Text = ""
      txt_stp3.Text = ""
      txt_stp4.Text = ""
txt_Decr.Text = Encrypt(txt_norm.Text, txt_Key.Text)

End Sub

Private Sub Form_Load()
txt_Temp.Text = "This Example Encrypts a String by using an Encryption Key.  It converts each letter in the key to it's ASCII value then (depending on the value of the previous value) adds/subtracts the value to the previous " & _
"value.  It then returns a value (ex. '403 432 421 502 493 411'). It takes that value and breaks it down to 2 digit numbers (ex. '40 34 32 42 15 02 49 34 11').  It takes those numbers and adds 100(so they will be valid characters) " & _
"and converts them to valid characters.  It then returns the Encrypted String." & vbCrLf & "If the receiving user does not know the encryption key then they will not be able to decrypt the encrypted string." & vbCrLf & "Please e-mail all comments " & _
"and questions to witenite87@excite.com.  Check http://camalot.virtualave.net/CamalotVB/index.html (Camalot VB) For more examples Soon." & vbCrLf & vbCrLf & "Note: If the first encrypted string returns values >999 then it WILL Crash!!!"
Me.Caption = "Encrypt/Decrypt Using Encryption Key ver " & App.Major & "." & App.Minor & "." & App.Revision & " By WhiteKnight"
End Sub

Private Sub MNU_About_Click()
frmAbout.Show vbModal
End Sub

Private Sub MNU_Exit_Click()
End
End Sub
