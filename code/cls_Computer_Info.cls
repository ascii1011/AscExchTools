VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Computer_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'cls_Computer_Info Class
'
'
'

Option Explicit

Private str_AccountName As String
Private str_ComputerName As String

'''''''''''''''''''''''''''''''Grab current user name'''''''''''''''''''''''''''''''''
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                                        (ByVal lpBuffer As String, nSize As Long) As Long
            
'''''''''''''''''''''''''''''''Grab computer name'''''''''''''''''''''''''''''''''
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                                        (ByVal lpBuffer As String, nSize As Long) As Long
            

Public Property Let Account(ByVal vData As String)
    str_AccountName = vData
End Property
Public Property Get Account() As String
    Account = str_AccountName
End Property


Public Property Let ComputerName(ByVal vData As String)
    str_ComputerName = vData
End Property
Public Property Get ComputerName() As String
    ComputerName = str_ComputerName
End Property

          
Public Sub Init()
    
    ClrVars
    
    Get_AccountName
    Get_ComputerName

End Sub

Private Sub ClrVars()
    Account = ""
    ComputerName = ""
End Sub


Private Sub Get_AccountName()
    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    Account = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Sub
            

Private Function Get_ComputerName()
    Dim strBuffer As String
    Dim lngBufSize As Long
    Dim lngStatus As Long

    lngBufSize = 255
    strBuffer = String$(lngBufSize, " ")
    lngStatus = GetComputerName(strBuffer, lngBufSize)
    If lngStatus <> 0 Then
        ComputerName = Left(strBuffer, lngBufSize)
    End If
End Function
