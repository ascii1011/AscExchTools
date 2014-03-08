VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Exmerge_Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarAccount As Variant 'local copy
Friend Property Let Account(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Account = 5
    mvarAccount = vData
End Property


Friend Property Set Account(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Account = Form1
    Set mvarAccount = vData
End Property


Friend Property Get Account() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Account
    If IsObject(mvarAccount) Then
        Set Account = mvarAccount
    Else
        Account = mvarAccount
    End If
End Property



