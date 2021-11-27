VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Prepares database queries."
'@Folder "Demo"
'@ModuleDescription "Prepares database queries."
'@PredeclaredId
Option Explicit

Private Type TDbStatement
    '''' Parent reference. For the purpose of this demo, parent reference is
    '''' defined as variant. Normally, it should be defined as a specific type.
    DbConn As Variant
    DbStmtID As String
    TimeStamp As Double
    DefaultInstance As Boolean
End Type
Private this As TDbStatement


Public Function Create(ByVal DbConn As DbConnection, ByVal DbStmtID As String) As DbStatement
    Dim Instance As DbStatement
    Set Instance = New DbStatement
    Instance.Init DbConn, DbStmtID
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbConn As DbConnection, ByVal DbStmtID As String)
    this.DbStmtID = DbStmtID
    '''' In practice, replace conditional with appropriate
    '''' branch based on selected approach.
    If ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE Then
        '''' Using parent object handle
        Dim DbConnHandle As String
        DbConnHandle = "CONNECTION:" & DbConn.DbConnStr
        '''' Assume that handles are created in a way that uniquely identify
        '''' objects, meaning if the handle already exists, it can be used.
        If ObjectStore.AllocateHandle(DbConnHandle) = -1 Then
            If ObjectStore.RefSet(DbConnHandle, DbConn) <> DbConnHandle Then
                Err.Raise 17, "DbStatement/Constructor", _
                          "Failed to save object refererence"
            End If
        End If
        this.DbConn = DbConnHandle
    Else
        '''' Using parent object reference
        Set this.DbConn = DbConn
    End If
    this.DefaultInstance = False
    this.TimeStamp = GetEpoch()
End Sub

Private Sub Class_Initialize()
    this.DefaultInstance = True
    Dim InstanceType As String
    InstanceType = IIf(Me Is DbStatement, "Default", "Regular")
    Debug.Print EpochToString(this.TimeStamp) & ": DbStatement /" & _
                InstanceType & " - Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Dim InstanceType As String
    InstanceType = IIf(this.DefaultInstance, "Default", "Regular")
    Debug.Print GetTimeStampMs & ": DbStatement /" & _
                InstanceType & " - Class_Terminate"
End Sub

Friend Sub CleanUp()
    '''' If DbConn is defined as Parent class or specific type, amend appropriately
    this.DbConn = Empty
End Sub

Public Property Get DbConn() As DbConnection
    '''' In practice, replace conditional with appropriate
    '''' branch based on selected approach.
    If ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE Then
        '''' Using parent object handle
        Set DbConn = ObjectStore.RefGet(this.DbConn)
    Else
        '''' Using parent object reference
        Set DbConn = this.DbConn
    End If
End Property
