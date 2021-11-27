---
layout: default
title: CleanUp cascade
nav_order: 3
permalink: /managed-refs
---

Because DbManager is not involved in circular references, its destructor can initiate traversal of the object hierarchy unraveling reference loops. For this purpose, the DbManager class and its descendants contain CleanUp routines:

**DbManager**

```vb
Private Sub Class_Terminate()
    CleanUp
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
```

---

**DbConnection**

```vb
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
```

---

**DbStatement**

```vb
Friend Sub CleanUp()
    Set this.DbConn = Nothing
End Sub
```

---

Setting *ReferenceLoopManagementMode = REF_LOOP_CLEANUP_CASCADE* and running the Main.Main() again should produce immediate pane output similar to this:

```
2021-11-27 02:51:53.172: DbManager   /Regular - Class_Initialize
2021-11-27 02:51:53.173: DbConnection/Regular - Class_Initialize
2021-11-27 02:51:53.174: DbStatement /Regular - Class_Initialize
2021-11-27 02:51:53.170: DbManager   /Regular - Class_Terminate
2021-11-27 02:51:53.170: DbStatement /Regular - Class_Terminate
2021-11-27 02:51:53.170: DbConnection/Regular - Class_Terminate
```

The message from the initiating DbManager class comes first, followed by the affected classes in the reverse order from bottom to top as expected. Each CleanUp routine should go through all affected children and call their CleanUp routines first, and then complete the clean-up process once the control is returned from the subtree underneath. In particular, parent object references should be set to Nothing, as well as the children collections. The clean-up process essentially severs the links to related objects making subsequent traversing of the hierarchy impossible. However, because the CleanUp routines are called from their parents, the control traverses the call stack back up, cleaning up all the visited objects.
