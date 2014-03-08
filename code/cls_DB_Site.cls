VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_DB_Site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarSite As Variant 'local copy
Private mvarDisplayName As Variant 'local copy
Private mvarDescription As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarServer As Variant 'local copy
Private mvarDatabaseName As Variant 'local copy
Private mvarUserName As Variant 'local copy
Private mvarPassword As Variant 'local copy
Friend Property Let Password(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Password = 5
    mvarPassword = vData
End Property


Friend Property Set Password(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Password = Form1
    Set mvarPassword = vData
End Property


Friend Property Get Password() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Password
    If IsObject(mvarPassword) Then
        Set Password = mvarPassword
    Else
        Password = mvarPassword
    End If
End Property



Friend Property Let UserName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserName = 5
    mvarUserName = vData
End Property


Friend Property Set UserName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.UserName = Form1
    Set mvarUserName = vData
End Property


Friend Property Get UserName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UserName
    If IsObject(mvarUserName) Then
        Set UserName = mvarUserName
    Else
        UserName = mvarUserName
    End If
End Property



Friend Property Let DatabaseName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatabaseName = 5
    mvarDatabaseName = vData
End Property


Friend Property Set DatabaseName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DatabaseName = Form1
    Set mvarDatabaseName = vData
End Property


Friend Property Get DatabaseName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabaseName
    If IsObject(mvarDatabaseName) Then
        Set DatabaseName = mvarDatabaseName
    Else
        DatabaseName = mvarDatabaseName
    End If
End Property



Friend Property Let Server(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Server = 5
    mvarServer = vData
End Property


Friend Property Set Server(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Server = Form1
    Set mvarServer = vData
End Property


Friend Property Get Server() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Server
    If IsObject(mvarServer) Then
        Set Server = mvarServer
    Else
        Server = mvarServer
    End If
End Property



Friend Property Let Description(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Friend Property Set Description(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Description = Form1
    Set mvarDescription = vData
End Property


Friend Property Get Description() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    If IsObject(mvarDescription) Then
        Set Description = mvarDescription
    Else
        Description = mvarDescription
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



Friend Property Let Site(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Site = 5
    mvarSite = vData
End Property


Friend Property Set Site(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Site = Form1
    Set mvarSite = vData
End Property


Friend Property Get Site() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Site
    If IsObject(mvarSite) Then
        Set Site = mvarSite
    Else
        Site = mvarSite
    End If
End Property



