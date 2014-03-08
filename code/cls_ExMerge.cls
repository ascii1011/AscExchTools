VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_ExMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Accounts"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarAccounts As Accounts


Public Property Get Accounts() As Accounts
    If mvarAccounts Is Nothing Then
        Set mvarAccounts = New Accounts
    End If


    Set Accounts = mvarAccounts
End Property


Public Property Set Accounts(vData As Accounts)
    Set mvarAccounts = vData
End Property
Private Sub Class_Terminate()
    Set mvarAccounts = Nothing
End Sub



