---
layout: default
title: ObjectStore
nav_order: 4
permalink: /objectstore
---

The other alternative prevents the formation of circular references by introducing an additional ObjectStore class (located in the *Object Store* folder). This class wraps a dictionary and acts as a global object store, with object addresses being the dictionary keys. The amended class diagram is shown in [Fig. 1](#CircularReferenceResolved).

<a name="CircularReferenceResolved"></a>  
<div align="center"><img src="https://github.com/pchemguy/ObjectStore/raw/develop/Assets/Diagrams/CircularReferenceResolved.svg" alt="Circular References Resolved" width="75%" /></div>
<p align="center"><b>Fig. 1. Simplified database library class diagram with ObjectStore</b></p>  

The important part is how ObjectStore is accessed. Accessing an ObjectStore instance via a regular reference would result in a three-node loop. Instead, a public ObjectStore variable, such as the predeclared instance, should be used. At the same time, the ObjectStore collection should be destroyed during termination to free the stored objects. The simplest way to achieve this goal is to destroy the ObjectStore variable itself. Unfortunately, setting the predeclared instance variable to Nothing is not supported and would crash the application. On the other hand, the similarly behaving auto-assigned variables can be destroyed by setting them to Nothing. Thus, a public auto-assigned variable ObjectStore named after the class (mimicking predeclared instances) is declared in the *ObjectStoreGlobals* regular module located in the same folder.

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

The ObjectStore class has three methods. SetRef() saves an object reference and returns its address, GetRef() retrieves a saved object reference, and DelRef() removes a saved object reference. When ObjectStore is used, the Child's factory signature and the Parent calling code remain the same. The Child code needs **four** changes:  

1) The private parent attribute type changes from parent class to pointer:

```vb
Private Type TDbStatement
    #If VBA7 Then
        DbConn As LongPtr
    #Else
        DbConn As Long
    #End If
    DbStmtID As String
End Type
Private this As TDbStatement
```

2) The constructor should submit the parent object reference to ObjectStore and save the handle in its private field:

```vb
Friend Sub Init(ByVal DbConn As DbConnection, ByVal DbStmtID As String)
    this.DbStmtID = DbStmtID
    this.DbConn = ObjectStore.SetRef(DbConn)
End Sub
```

3) The parent getter needs to be defined or amended (any visibility will work for local use):

```vb
Public Property Get DbConn() As DbConnection
    Set DbConn = ObjectStore.GetRef(this.DbConn)
End Property
```

4) Finally, any references to the `this.DbConn` parent class attribute must be replaced with the getter `DbConn`, possibly caching its value within the local procedural scope as necessary.

Apart from the changes to child classes, the ObjectStore variable itself needs to be destroyed during the termination stage. DbManager class is naturally positioned to perform this task. However, because the ObjectStore variable is global and more than one DbManager instance may exist, ObjectStore should not be destroyed from DbManager.Class_Terminate(). Instead, we add code to make the default DbManager instance count the regular instances. Whenever count goes to zero, the **default** instance of DbManager sets ObjectStore to Nothing.

---

Running Main.Main() with `ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE` setting should produce output similar to this (events due to the dafault instance may or may not appear depending on current execution context):

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
