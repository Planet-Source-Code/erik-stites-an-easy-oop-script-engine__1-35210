VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Quick Script"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHex 
      Caption         =   "Hex output"
      Height          =   255
      Left            =   420
      TabIndex        =   3
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute Instruction"
      Default         =   -1  'True
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   1635
   End
   Begin VB.TextBox txtScript 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   3675
   End
   Begin VB.Label lblResult 
      Caption         =   "[output]"
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   660
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Welcome to Object Oriented Programming...

'Similar to a CPU [cough...'emulation']... the values will remain
    'from the previous instruction
Dim Scripting As New clsScript

Private Sub cmdExecute_Click()
    On Error GoTo CallErr
    
    Dim MNEM As String
    
    'Get the name of the method to call
    MNEM = Scripting.Parse(txtScript.Text)
    
    'If it returned no mnemonic, don't call a method
    If Not MNEM = "" Then
        'This class has methods, methods can be
            'called by name from a string
        CallByName Scripting, MNEM, VbMethod
    End If
    
    'Display results of instruction
    If chkHex.Value = Checked Then
        lblResult.Caption = Hex(Scripting.X)
    Else
        lblResult.Caption = Scripting.X
    End If
    txtScript.SetFocus
    Exit Sub
    
    'This error will occur when a function is called that is not in the class
CallErr:
    MsgBox "Mnemonic not recognized.", vbExclamation, "Error"
    txtScript.SetFocus
End Sub
