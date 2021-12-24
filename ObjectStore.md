---
layout: default
title: ObjectStore
nav_order: 5
permalink: /objectstore
---

The ObjectStore class simulates weak references via a Dictionary-based registry ([Patterns of Enterprise Application Architecture][]) and scalar keys as object handles.

### Dictionary map and public auto-assigned store instance

ObjectStore essentially maintains an **ObjPtr&nbsp;&rarr;&nbsp;Obj** map in a wrapped dictionary. With ObjectStore employed, the child class (DbStatement) constructor saves the parent object (DbConnection) address instead of its reference while saving the latter to ObjectStore (see amended class diagram in [Fig. 1](#CircularReferenceResolved)).

<a name="CircularReferenceResolved"></a>  
<div align="center"><img src="https://github.com/pchemguy/ObjectStore/raw/develop/Assets/Diagrams/CircularReferenceResolved.svg" alt="Circular References Resolved" width="75%" /></div>
<p align="center"><b>Fig. 1. Simplified database library class diagram with ObjectStore</b></p>  

The important part is how ObjectStore is accessed. Accessing an ObjectStore instance via a regular reference would result in a three-node loop. Instead, a public ObjectStore variable, such as the predeclared instance, should be used. At the same time, the ObjectStore collection should be destroyed during termination to free the stored objects. The simplest way to achieve this goal is to destroy the ObjectStore variable itself. Unfortunately, setting the predeclared instance variable to Nothing is not supported and would crash the application. On the other hand, the similarly behaving auto-assigned variables can be reset by setting them to Nothing. Thus, the *ObjectStoreGlobals* module (in the same folder) contains a declaration of a public auto-assigned variable ObjectStore named after the class, mimicking predeclared instances.

### Child class adjustments

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

The ObjectStore class has three methods. SetRef() saves an object reference and returns its address, GetRef() retrieves a saved object reference, and DelRef() removes a saved object reference. With the ObjectStore class employed, the Child's factory signature and the Parent calling code remain unchanged. The Child code needs **four** changes:  

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

### Instance counter and ObjectStore destruction

Apart from the changes to child classes, the ObjectStore variable itself needs to be destroyed during the termination stage. DbManager class is naturally positioned to perform this task. However, because the ObjectStore variable is global and more than one DbManager instance may exist, ObjectStore should not be destroyed from DbManager.Class_Terminate(). Instead, the default DbManager instance counts existing regular instances. Whenever this count goes to zero, the **default** instance of DbManager sets ObjectStore to Nothing, destroying the wrapped dictionary object alone with all saved object references and freeing the stored objects.
 
---

Running Main.Main() with `ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE` setting should produce output similar to this (events due to the default instance may or may not appear depending on current execution context):

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

indicating that all objects are properly terminated, including the ObjectStore. Termination of the ObjectStore variable destroys all parent references. Then object destruction proceeds starting from the top. Specifically, once ObectStore and DbManager instances are destroyed, DbConnection, the top affected class, can be destroyed, as it no longer has a parent holding its reference. Destruction of the DbConnection instance clears its children collection freeing its children.

### Alternative implementation

While resetting the ObjectStore object encapsulates its clean-up code, it is sufficient to reset the wrapped dictionary directly. In this case, the predeclared instance can also do the job. Further, the managing class, DbManager in this case, can absorb ObjectStore functionality.


<!-- References -->

[Patterns of Enterprise Application Architecture]: https://books.google.com/books?id=FyWZt5DdvFkC
