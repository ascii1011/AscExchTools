VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_MSSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"cls_DB_Sites"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Private Type SQL_Con_Struct
    Con As New ADODB.Connection
    Cmd As New ADODB.Command
    Error As String
End Type

Private Type Sql_Results_Struct
    Query As String
    Data As New ADODB.Recordset
End Type

Private SQLcn As SQL_Con_Struct
'Public pubSQLcn As SQL_Con_Struct
'Public rs As Sql_Results_Struct



''''''''''''''''abstraction to cls_db_sites'''''''''''''

Private mvarcls_DB_Sites As cls_DB_Sites


Public Property Get cls_DB_Sites() As cls_DB_Sites
    If mvarcls_DB_Sites Is Nothing Then
        Set mvarcls_DB_Sites = New cls_DB_Sites
    End If


    Set cls_DB_Sites = mvarcls_DB_Sites
End Property


Public Property Set cls_DB_Sites(vData As cls_DB_Sites)
    Set mvarcls_DB_Sites = vData
End Property
Private Sub Class_Terminate()
    Set mvarcls_DB_Sites = Nothing
End Sub




'''''''''''''''''''''functionality'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'here is an example of a query
'
'   Dim msgs As Sql_Results_Struct              'init structure
'
'   msgs.Query = " select * from tablename "    'set query statement
'
'   SQL_Query_auto msgs.Query, msgs.Data        'Process query request
'                                               'SQL_Query_auto -> SQL_Connect -> SQL_Query -> SQL_Init
'
'   SQL_Close_Clear msgs.Data                   'clean up
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SQL_Query_auto(sQuery As String, rsData As ADODB.Recordset)
    Set rsData = Nothing
    If SQL_Connect = True Then SQL_Query sQuery, rsData
End Sub

Public Function SQL_Connect() As Boolean
    SQL_Connect = False
    
On Error GoTo SQLError:

    SQL_Connect = SQL_Status
    If SQL_Connect = True Then Exit Function

    SQLcn.Con.Provider = "sqloledb"
    SQLcn.Con.CursorLocation = adUseClient
    SQLcn.Con.Properties("Data Source").Value = "127.0.0.1"
    SQLcn.Con.Properties("Initial Catalog").Value = "siboney"
    SQLcn.Con.Properties("User ID").Value = "sa"
    SQLcn.Con.Properties("Password").Value = "gggg"
    SQLcn.Con.Open
    
    If SQLcn.Con.State = 1 Then SQL_Connect = True
    Exit Function
    
SQLError:
    'SQLErrorAction Err
End Function
Public Sub SQL_Query(sQuery As String, rsData As ADODB.Recordset)
    SQL_Init sQuery
    
On Error GoTo SQLError:
    Set rsData = SQLcn.Cmd.Execute
    Exit Sub
    
SQLError:
    'SQLErrorAction Err
End Sub


Private Sub SQL_Init(sQuery As String)
    
On Error GoTo SQLError:
    Set SQLcn.Cmd.ActiveConnection = SQLcn.Con
    SQLcn.Cmd.CommandType = adCmdText
    SQLcn.Cmd.CommandText = sQuery
    Exit Sub
    
SQLError:
    'SQLErrorAction Err
End Sub


Private Sub SQL_Init(sQuery As String)
    
On Error GoTo SQLError:
    Set SQLcn.Cmd.ActiveConnection = SQLcn.Con
    SQLcn.Cmd.CommandType = adCmdText
    SQLcn.Cmd.CommandText = sQuery
    Exit Sub
    
SQLError:
    'Err
End Sub
