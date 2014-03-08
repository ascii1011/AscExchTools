VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'frmMain.Caption = "Backup Solutions - Version " & App.Major & "." & App.Minor & "." & App.Revision
'cls_Main Class
'
'
'

Option Explicit


Private str_Main_Caption As String
Private str_App_Version As String
            

Public Property Let MainCaption(ByVal vData As String)
    str_Main_Caption = vData
End Property
Public Property Get MainCaption() As String
    MainCaption = str_Main_Caption
End Property
            

Public Property Let AppVersion(ByVal vData As String)
    str_App_Version = vData
End Property
Public Property Get AppVersion() As String
    AppVersion = str_App_Version
End Property

          
            
Public Sub Init()
    MainCaption = "Ascii ToolSet"
    AppVersion = App.Major & "." & App.Minor & "." & App.Revision
End Sub


Public Sub Put_MainCaption(ByVal vData As String)
    frmMain.Caption = MainCaption & " - " & vData
End Sub
Public Sub clr_MainCaption()
    frmMain.Caption = MainCaption
End Sub
