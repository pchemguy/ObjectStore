Attribute VB_Name = "Main"
'@Folder "Demo"
Option Explicit


Private Sub Main()
    ObjectStoreGlobals.ReferenceLoopManagementMode = REF_LOOP_NO_MANAGEMENT
    
    Dim dbm As DbManager
    Set dbm = DbManager.Create()
    
    Dim DbConnStr As String
    DbConnStr = "SomeConnectionString"
    Dim dbc As DbConnection
    Set dbc = dbm.CreateConnection(DbConnStr)
    
    Dim DbStmtID As String
    DbStmtID = "SomeStatementID"
    Dim dbs As DbStatement
    Set dbs = dbc.CreateStatement(DbStmtID)
End Sub
