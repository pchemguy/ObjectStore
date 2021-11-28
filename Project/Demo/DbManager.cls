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
    InstanceCount As Long    '''' Number of DbManager objects; only use on the default instance
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
    If Me Is DbManager Then
        this.InstanceCount = 0
    Else
        DbManager.InstanceAdd
    End If
    this.DefaultInstance = True
    Dim InstanceType As String
    InstanceType = IIf(Me Is DbManager, "Default", "Regular")
    Debug.Print EpochToString(this.TimeStamp) & ": DbManager   /" & _
                InstanceType & " - Class_Initialize"
End Sub

Private Sub Class_Terminate()
    If ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_CLEANUP_CASCADE Then
        CleanUp
    End If
    DbManager.InstanceDel
    Dim InstanceType As String
    InstanceType = IIf(this.DefaultInstance, "Default", "Regular")
    Debug.Print GetTimeStampMs & ": DbManager   /" & _
                InstanceType & " - Class_Terminate"
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

Public Property Get InstanceCount() As Long
    If Me Is DbManager Then
        InstanceCount = this.InstanceCount
    Else
        InstanceCount = DbManager.InstanceCount
    End If
End Property

Public Property Let InstanceCount(ByVal Value As Long)
    If Me Is DbManager Then
        this.InstanceCount = Value
    Else
        DbManager.InstanceCount = Value
    End If
End Property

Public Sub InstanceAdd()
    If Me Is DbManager Then
        this.InstanceCount = this.InstanceCount + 1
    Else
        DbManager.InstanceCount = DbManager.InstanceCount + 1
    End If
End Sub

Public Sub InstanceDel()
    If Me Is DbManager Then
        this.InstanceCount = this.InstanceCount - 1
        If this.InstanceCount = 0 Then
            If ObjectStoreGlobals.ReferenceLoopManagementMode = _
               REF_LOOP_OBJECT_STORE Then
                Set ObjectStore = Nothing
            End If
        End If
    Else
        DbManager.InstanceCount = DbManager.InstanceCount - 1
    End If
End Sub

Public Function CreateConnection(ByVal DbConnStr As String) As DbConnection
    Dim DbConn As DbConnection
    Set DbConn = DbConnection.Create(DbConnStr)
    Set this.Connections(DbConnStr) = DbConn
    Set CreateConnection = DbConn
End Function
