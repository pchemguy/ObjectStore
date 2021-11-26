VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Handles database connections. Abstract factory for DbStatement."
'@Folder "Demo"
'@ModuleDescription "Handles database connections. Abstract factory for DbStatement."
'@PredeclaredId
Option Explicit

Private Type TDbConnection
    DbConnStr As String
    Statements As Scripting.Dictionary  '''' Children collection
    TimeStamp As Double
    DefaultInstance As Boolean
End Type
Private this As TDbConnection


Public Function Create(ByVal DbConnStr As String) As DbConnection
    Dim Instance As DbConnection
    Set Instance = New DbConnection
    Instance.Init DbConnStr
    Set Create = Instance
End Function

Friend Sub Init(ByVal DbConnStr As String)
    this.DbConnStr = DbConnStr
    Set this.Statements = New Scripting.Dictionary
    this.Statements.CompareMode = TextCompare
    this.DefaultInstance = False
    this.TimeStamp = GetEpoch()
End Sub

Private Sub Class_Initialize()
    this.DefaultInstance = True
    Dim InstanceType As String
    InstanceType = IIf(Me Is DbConnection, "Default", "Regular")
    Debug.Print EpochToString(this.TimeStamp) & ": DbConnection/" & _
                InstanceType & " - Class_Initialize"
End Sub

Private Sub Class_Terminate()
    Dim InstanceType As String
    InstanceType = IIf(this.DefaultInstance, "Default", "Regular")
    Debug.Print GetTimeStampMs & ": DbConnection/" & _
                InstanceType & " - Class_Terminate"
End Sub

Friend Sub CleanUp()
    Dim DbStmt As DbStatement
    Dim DbStmtHandle As Variant
    For Each DbStmtHandle In this.Statements.Keys
        Set DbStmt = this.Statements(DbStmtHandle)
        DbStmt.CleanUp
    Next DbStmtHandle
    Set DbStmt = Nothing
    this.Statements.RemoveAll
    Set this.Statements = Nothing
End Sub

Public Property Get DbConnStr() As String
    DbConnStr = this.DbConnStr
End Property

Public Function CreateStatement(ByVal DbStmtID As String) As DbStatement
    Dim DbStmt As DbStatement
    Set DbStmt = DbStatement.Create(Me, DbStmtID)
    Set this.Statements(DbStmtID) = DbStmt
    Set CreateStatement = DbStmt
End Function
