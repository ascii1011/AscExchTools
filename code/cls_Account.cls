VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSessionName As String 'local copy
Private mvarSessionDesc As String 'local copy
Private mvarSessionID As Variant 'local copy
Private mvarSessionType As Variant 'local copy
Private mvarBackupType As Variant 'local copy
Private mvarLastBackup As Variant 'local copy
Private mvarLastResults As Variant 'local copy
Private mvarMailSize As Variant 'local copy
Private mvarMailMsgCount As Variant 'local copy
Private mvarAD_SID As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarDisplayName As Variant 'local copy
Private mvarGroupName As Variant 'local copy
Friend Property Let GroupName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GroupName = 5
    mvarGroupName = vData
End Property


Friend Property Set GroupName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.GroupName = Form1
    Set mvarGroupName = vData
End Property


Friend Property Get GroupName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GroupName
    If IsObject(mvarGroupName) Then
        Set GroupName = mvarGroupName
    Else
        GroupName = mvarGroupName
    End If
End Property



Friend Property Let DisplayName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DisplayName = 5
    mvarDisplayName = vData
End Property


Friend Property Set DisplayName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DisplayName = Form1
    Set mvarDisplayName = vData
End Property


Friend Property Get DisplayName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DisplayName
    If IsObject(mvarDisplayName) Then
        Set DisplayName = mvarDisplayName
    Else
        DisplayName = mvarDisplayName
    End If
End Property



Friend Property Let AD_SID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AD_SID = 5
    mvarAD_SID = vData
End Property


Friend Property Set AD_SID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AD_SID = Form1
    Set mvarAD_SID = vData
End Property


Friend Property Get AD_SID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AD_SID
    If IsObject(mvarAD_SID) Then
        Set AD_SID = mvarAD_SID
    Else
        AD_SID = mvarAD_SID
    End If
End Property



Friend Property Let MailMsgCount(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MailMsgCount = 5
    mvarMailMsgCount = vData
End Property


Friend Property Set MailMsgCount(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MailMsgCount = Form1
    Set mvarMailMsgCount = vData
End Property


Friend Property Get MailMsgCount() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MailMsgCount
    If IsObject(mvarMailMsgCount) Then
        Set MailMsgCount = mvarMailMsgCount
    Else
        MailMsgCount = mvarMailMsgCount
    End If
End Property



Friend Property Let MailSize(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MailSize = 5
    mvarMailSize = vData
End Property


Friend Property Set MailSize(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.MailSize = Form1
    Set mvarMailSize = vData
End Property


Friend Property Get MailSize() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MailSize
    If IsObject(mvarMailSize) Then
        Set MailSize = mvarMailSize
    Else
        MailSize = mvarMailSize
    End If
End Property



Friend Property Let LastResults(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastResults = 5
    mvarLastResults = vData
End Property


Friend Property Set LastResults(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LastResults = Form1
    Set mvarLastResults = vData
End Property


Friend Property Get LastResults() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastResults
    If IsObject(mvarLastResults) Then
        Set LastResults = mvarLastResults
    Else
        LastResults = mvarLastResults
    End If
End Property



Friend Property Let LastBackup(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LastBackup = 5
    mvarLastBackup = vData
End Property


Friend Property Set LastBackup(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.LastBackup = Form1
    Set mvarLastBackup = vData
End Property


Friend Property Get LastBackup() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastBackup
    If IsObject(mvarLastBackup) Then
        Set LastBackup = mvarLastBackup
    Else
        LastBackup = mvarLastBackup
    End If
End Property



Friend Property Let BackupType(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackupType = 5
    mvarBackupType = vData
End Property


Friend Property Set BackupType(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BackupType = Form1
    Set mvarBackupType = vData
End Property


Friend Property Get BackupType() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BackupType
    If IsObject(mvarBackupType) Then
        Set BackupType = mvarBackupType
    Else
        BackupType = mvarBackupType
    End If
End Property



Friend Property Let SessionType(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionType = 5
    mvarSessionType = vData
End Property


Friend Property Set SessionType(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SessionType = Form1
    Set mvarSessionType = vData
End Property


Friend Property Get SessionType() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionType
    If IsObject(mvarSessionType) Then
        Set SessionType = mvarSessionType
    Else
        SessionType = mvarSessionType
    End If
End Property



Friend Property Let SessionID(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionID = 5
    mvarSessionID = vData
End Property


Friend Property Set SessionID(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SessionID = Form1
    Set mvarSessionID = vData
End Property


Friend Property Get SessionID() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionID
    If IsObject(mvarSessionID) Then
        Set SessionID = mvarSessionID
    Else
        SessionID = mvarSessionID
    End If
End Property



Friend Property Let SessionDesc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionDesc = 5
    mvarSessionDesc = vData
End Property


Friend Property Get SessionDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionDesc
    SessionDesc = mvarSessionDesc
End Property



Friend Property Let SessionName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionName = 5
    mvarSessionName = vData
End Property


Friend Property Get SessionName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionName
    SessionName = mvarSessionName
End Property


