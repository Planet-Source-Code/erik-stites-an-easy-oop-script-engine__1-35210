VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'The Best part of using this kind of scripting; to add instructions,
    'just add more functions to this class module

'Change to Double for more accurate results
Public X As Single 'Is public so external objects can read it
Private Y As Single 'This can only be affected by this module

'I believe that these functions are self explanatory
    'Well, the V on some of them are to keep errors from occurring
    'i.e. AND is a reserved word in VB... this could be avoided with string
    'manipulation such as when the user enters AND as a mnemonic
    'a 'V' is added to the string before the callbyname
Public Sub DIV()
    If Y = 0 Then
        Y = 1 'prevent Division by Zero
    End If
    X = X / Y
End Sub

Public Sub MULT()
    X = X * Y
End Sub

'You can enter negative numbers
Public Sub ADD()
Attribute ADD.VB_Description = "Add X to Y -> X"
    X = X + Y
End Sub

Public Sub SUBT()
    X = X - Y
End Sub

Public Sub MODX()
    X = X Mod Y
End Sub

Public Sub EOR()
Attribute EOR.VB_Description = "Exclusive Or X with Y -> X"
    X = X Xor Y
End Sub

Public Sub ORV()
    X = X Or Y
End Sub

Public Sub NOR()
    X = Not (X Or Y)
End Sub

Public Sub INVERT()
    X = Not X
End Sub

Public Sub ANDV()
    X = X And Y
End Sub

Public Sub NAND()
    X = Not (X And Y)
End Sub

Public Function Parse(strInput As String) As String
    On Error GoTo ParseErr
    
    Dim DataArr() As String
    
    DataArr() = Split(strInput, " ")
    
    Parse = DataArr(0) 'Returns the mnemonic name
    X = Val(DataArr(1)) 'Returns first value
    Y = Val(DataArr(2)) 'Second Value
    
ParseErr:
End Function

