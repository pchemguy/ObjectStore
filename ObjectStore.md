---
layout: default
title: ObjectStore
nav_order: 4
permalink: /objectstore
---

The other alternative prevents the formation of circular references by introducing an additional ObjectStore class (located in the *Object Store* folder). This class wraps a dictionary and acts as a global object store, with object handles being the dictionary keys. The amended class diagram is shown in [Fig. 1](#CircularReferenceResolved).

<a name="CircularReferenceResolved"></a>  
<div align="center"><img src="https://github.com/pchemguy/ObjectStore/raw/develop/Assets/Diagrams/CircularReferenceResolved.svg" alt="Circular References Resolved" width="75%" /></div>
<p align="center"><b>Fig. 1. Simplified database library class diagram with ObjectStore</b></p>  

The important part is how ObjectStore is accessed. Accessing an ObjectStore instance via a regular reference would result in a three-node loop. Instead, a public ObjectStore variable, such as the predeclared instance, should be used. At the same time, the ObjectStore collection should be destroyed during termination to free the stored objects. The simplest way to achieve this goal is to destroy the ObjectStore variable itself. Unfortunately, setting the predeclared instance variable to Nothing is not supported and would crash the application. On the other hand, the similarly behaving auto-assigned variables can be destroyed by setting them to Nothing. Thus, a public auto-assigned variable ObjectStore named after the class (mimicking predeclared instances) is declared in the _ObjectStoreGlobals_ regular module located in the same folder.

The child object in a circular reference relationship often takes its parent reference via the factory and saves it via the constructor, for example:

**DbStatement**

```vb
Private Type TDbStatement
    DbConn As DbConnection
    DbStmtID As String
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
    Set this.DbConn = DbConn
End Sub
```

With ObjectStore, the factory's signature and the calling code remain the same, and the child needs four changes. The private parent attribute should be changed from parent class to the handle type, such as String or Currency:

```vb
Private Type TDbStatement
    DbConn As String         '''' <--- instead of DbConnection
    DbStmtID As String
End Type
Private this As TDbStatement
```

The constructor should submit the parent object reference to ObjectStore and save the handle in its private field. This code also assumes that if the object handle already exists in the ObjectStore, that saved object reference can be used (e.g., an older DbStatement sibling has already submitted its parent reference to ObjectStore):

```vb
Friend Sub Init(ByVal DbConn As DbConnection, ByVal DbStmtID As String)
    this.DbStmtID = DbStmtID
    Dim DbConnHandle As String
    DbConnHandle = "CONNECTION:" & DbConn.DbConnStr
    If ObjectStore.AllocateHandle(DbConnHandle) = -1 Then
        If ObjectStore.RefSet(DbConnHandle, DbConn) <> DbConnHandle Then
            Err.Raise 17, "DbStatement/Constructor", _
                      "Failed to save object refererence"
        End If
    End If
    this.DbConn = DbConnHandle
End Sub
```

The parent getter needs to be defined or amended (any visibility will work for local use):

```vb
Public Property Get DbConn() As DbConnection
    Set DbConn = ObjectStore.RefGet(this.DbConn)
End Property
```

Finally, any references to the `this.DbConn` parent class attribute must be replaced with the getter `DbConn`, possibly caching its value within the local procedural scope as necessary.

---

Running Main.Main() with `ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE` setting should produce output similar to this:

```
2021-11-27 03:19:39.682: DbManager   /Regular - Class_Initialize
2021-11-27 03:19:39.683: DbConnection/Regular - Class_Initialize
2021-11-27 03:19:39.683: DbStatement /Regular - Class_Initialize
2021-11-27 03:19:39.680: ObjectStore - Class_Initialize
2021-11-27 03:19:39.680: ObjectStore - Class_Terminate
2021-11-27 03:19:39.680: DbManager   /Regular - Class_Terminate
2021-11-27 03:19:39.680: DbConnection/Regular - Class_Terminate
2021-11-27 03:19:39.680: DbStatement /Regular - Class_Terminate
```

indicating that all objects are properly terminated, including the ObjectStore. Termination of the ObjectStore variable destroys all parent references, and destruction starts from the top. Specifically, once ObectStore and DbManager instances are destructed, DbConnection, the top affected class, can be destructed, as it no longer has a parent holding its reference. Destruction of the DbConnection instance clears its children collection freeing its children.

The ObjectStore class has four methods. AllocateHandle() reserves an object handle. If an optional string argument is provided, AllocateHandle() checks whether the matching dictionary key already exists. If no such key exists, AllocateHandle() adds the new key to the dictionary with an empty value and returns -1 (currency); otherwise, it returns 0. Alternatively, a new timestamp-based key is returned, using the Unix Epoch time with milliseconds retrieved from the Timer() function. FreeHandle() removes a saved object reference. RefSet()/RefGet() saves/retrieves object references. RefSet() returns the object's handle on success and a blank string on failure. RefSet() requires a preallocated key with an empty value, preserving existing object reference.
