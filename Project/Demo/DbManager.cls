VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Top database API class. Abstract factory for DbConnection."
'@Folder "Demo"
'@ModuleDescription "Top database API class. Abstract factory for DbConnection."
'@PredeclaredId
Option Explicit

Private Type TDbManager
    Connections As Scripting.Dictionary '''' Children collection
    TimeStamp As Double
    DefaultInstance As Boolean
End Type
Private this As TDbManager


Public Function Create() As DbManager
    Dim Instance As DbManager
    Set Instance = New DbManager
    Instance.Init
    Set Create = Instance
End Function

Friend Sub Init()
    Set this.Connections = New Scripting.Dictionary
    this.Connections.CompareMode = TextCompare
    this.DefaultInstance = False
    this.TimeStamp = GetEpoch()
End Sub

Private Sub Class_Initialize()
    this.DefaultInstance = True
    Dim InstanceType As String
    InstanceType = IIf(Me Is DbManager, "Default", "Regular")
    Debug.Print EpochToString(this.TimeStamp) & ": DbManager/" & _
                InstanceType & " - Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Dim InstanceType As String
    InstanceType = IIf(this.DefaultInstance, "Default", "Regular")
    Debug.Print GetTimeStampMs & ": DbManager/" & _
                InstanceType & " - Class_Terminate"
    If ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_CLEANUP_CASCADE Then
        CleanUp
    ElseIf ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE Then
        Set ObjectStore = Nothing
    End If
End Sub

Friend Sub CleanUp()
    Dim DbConn As DbConnection
    Dim DbConnHandle As Variant
    For Each DbConnHandle In this.Connections.Keys
        Set DbConn = this.Connections(DbConnHandle)
        DbConn.CleanUp
    Next DbConnHandle
    Set DbConn = Nothing
    this.Connections.RemoveAll
    Set this.Connections = Nothing
End Sub

Public Function CreateConnection(ByVal DbConnStr As String) As DbConnection
    Dim DbConn As DbConnection
    Set DbConn = DbConnection.Create(DbConnStr)
    Set this.Connections(DbConnStr) = DbConn
    Set CreateConnection = DbConn
End Function
