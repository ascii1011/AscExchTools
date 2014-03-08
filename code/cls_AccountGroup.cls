VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_AccountGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarDisplayName As Variant 'local copy
Private mvarSessionName As Variant 'local copy
Private mvarSessionID As Variant 'local copy
Private mvarUserCount As Variant 'local copy
Private mvarLastBackup As Variant 'local copy
Private mvarLastResults As Variant 'local copy
Private mvarSelected As Variant 'local copy
Private mvarTotalMailMsgCount As Variant 'local copy
Private mvarTotalUserMailSize As Variant 'local copy
Friend Property Let TotalUserMailSize(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalUserMailSize = 5
    mvarTotalUserMailSize = vData
End Property


Friend Property Set TotalUserMailSize(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TotalUserMailSize = Form1
    Set mvarTotalUserMailSize = vData
End Property


Friend Property Get TotalUserMailSize() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalUserMailSize
    If IsObject(mvarTotalUserMailSize) Then
        Set TotalUserMailSize = mvarTotalUserMailSize
    Else
        TotalUserMailSize = mvarTotalUserMailSize
    End If
End Property



Friend Property Let TotalMailMsgCount(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalMailMsgCount = 5
    mvarTotalMailMsgCount = vData
End Property


Friend Property Set TotalMailMsgCount(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TotalMailMsgCount = Form1
    Set mvarTotalMailMsgCount = vData
End Property


Friend Property Get TotalMailMsgCount() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalMailMsgCount
    If IsObject(mvarTotalMailMsgCount) Then
        Set TotalMailMsgCount = mvarTotalMailMsgCount
    Else
        TotalMailMsgCount = mvarTotalMailMsgCount
    End If
End Property



Friend Property Let Selected(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Selected = 5
    mvarSelected = vData
End Property


Friend Property Set Selected(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Selected = Form1
    Set mvarSelected = vData
End Property


Friend Property Get Selected() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Selected
    If IsObject(mvarSelected) Then
        Set Selected = mvarSelected
    Else
        Selected = mvarSelected
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



Friend Property Let UserCount(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserCount = 5
    mvarUserCount = vData
End Property


Friend Property Set UserCount(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.UserCount = Form1
    Set mvarUserCount = vData
End Property


Friend Property Get UserCount() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserCount
    If IsObject(mvarUserCount) Then
        Set UserCount = mvarUserCount
    Else
        UserCount = mvarUserCount
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



Friend Property Let SessionName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SessionName = 5
    mvarSessionName = vData
End Property


Friend Property Set SessionName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.SessionName = Form1
    Set mvarSessionName = vData
End Property


Friend Property Get SessionName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SessionName
    If IsObject(mvarSessionName) Then
        Set SessionName = mvarSessionName
    Else
        SessionName = mvarSessionName
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



