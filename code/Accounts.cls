VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Accounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"cls_Account"
Attribute VB_Ext_KEY = "Member0" ,"cls_Account"
'local variable to hold collection
Private mCol As Collection

Public Function Add(vName As String, vGroup As String, SessionName As String, SessionDesc As String, SessionID As Variant, SessionType As Variant, BackupType As Variant, LastBackup As Variant, LastResults As Variant, MailSize As Variant, MailMsgCount As Variant, AD_SID As Variant, Optional sKey As String) As cls_Account
    'create a new object
    Dim objNewMember As cls_Account
    Set objNewMember = New cls_Account


    'set the properties passed into the method
    objNewMember.vName = vName
    objNewMember.vGroup = vGroup
    objNewMember.SessionName = SessionName
    objNewMember.SessionDesc = SessionDesc
    If IsObject(SessionID) Then
        Set objNewMember.SessionID = SessionID
    Else
        objNewMember.SessionID = SessionID
    End If
    If IsObject(SessionType) Then
        Set objNewMember.SessionType = SessionType
    Else
        objNewMember.SessionType = SessionType
    End If
    If IsObject(BackupType) Then
        Set objNewMember.BackupType = BackupType
    Else
        objNewMember.BackupType = BackupType
    End If
    If IsObject(LastBackup) Then
        Set objNewMember.LastBackup = LastBackup
    Else
        objNewMember.LastBackup = LastBackup
    End If
    If IsObject(LastResults) Then
        Set objNewMember.LastResults = LastResults
    Else
        objNewMember.LastResults = LastResults
    End If
    If IsObject(MailSize) Then
        Set objNewMember.MailSize = MailSize
    Else
        objNewMember.MailSize = MailSize
    End If
    If IsObject(MailMsgCount) Then
        Set objNewMember.MailMsgCount = MailMsgCount
    Else
        objNewMember.MailMsgCount = MailMsgCount
    End If
    If IsObject(AD_SID) Then
        Set objNewMember.AD_SID = AD_SID
    Else
        objNewMember.AD_SID = AD_SID
    End If
    objNewMember.vName = vName
    objNewMember.vGroup = vGroup
    objNewMember.SessionName = SessionName
    objNewMember.SessionDesc = SessionDesc
    If IsObject(SessionID) Then
        Set objNewMember.SessionID = SessionID
    Else
        objNewMember.SessionID = SessionID
    End If
    If IsObject(SessionType) Then
        Set objNewMember.SessionType = SessionType
    Else
        objNewMember.SessionType = SessionType
    End If
    If IsObject(BackupType) Then
        Set objNewMember.BackupType = BackupType
    Else
        objNewMember.BackupType = BackupType
    End If
    If IsObject(LastBackup) Then
        Set objNewMember.LastBackup = LastBackup
    Else
        objNewMember.LastBackup = LastBackup
    End If
    If IsObject(LastResults) Then
        Set objNewMember.LastResults = LastResults
    Else
        objNewMember.LastResults = LastResults
    End If
    If IsObject(MailSize) Then
        Set objNewMember.MailSize = MailSize
    Else
        objNewMember.MailSize = MailSize
    End If
    If IsObject(MailMsgCount) Then
        Set objNewMember.MailMsgCount = MailMsgCount
    Else
        objNewMember.MailMsgCount = MailMsgCount
    End If
    If IsObject(AD_SID) Then
        Set objNewMember.AD_SID = AD_SID
    Else
        objNewMember.AD_SID = AD_SID
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cls_Account
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

