---
layout: default
title: Managing circular references
nav_order: 3
permalink: /managed-refs
---

### CleanUp cascade

Because DbManager is not involved in circular references, its destructor can initiate traversal of the object hierarchy unraveling reference loops. For this purpose, the DbManager class and its descendents contain CleanUp routines. Setting *ReferenceLoopManagementMode = REF_LOOP_CLEANUP_CASCADE* and running the Main.Main() again should produce an immediate pane output similar to this:

    2021-11-27 02:51:53.172: DbManager   /Regular - Class_Initialize
    2021-11-27 02:51:53.173: DbConnection/Regular - Class_Initialize
    2021-11-27 02:51:53.174: DbStatement /Regular - Class_Initialize
    2021-11-27 02:51:53.170: DbManager   /Regular - Class_Terminate
    2021-11-27 02:51:53.170: DbStatement /Regular - Class_Terminate
    2021-11-27 02:51:53.170: DbConnection/Regular - Class_Terminate

The message from the initiating DbManager class comes first, followed by the affected classes in the reverse order from bottom to top as expected. Each CleanUp routine should go through all affected children and call their CleanUp routines first, and then complete the clean-up process once the control is returned from the subtree underneath. In particular, parent object references should be set to Nothing, as well as the children collections. The clean-up process essentially severs the links to related objects making subsequent traversing of the hierarchy impossible. However, because the CleanUp routines are called from their parents, the control traverses the call stack back up, cleaning up all the visited objects.

### ObjectStore

The other alternative prevents the formation of circular references by introducing an additional ObjectStore class (located in the *Object Store* folder). This class wraps a dictionary and acts as a global object store, with object handles being the dictionary keys. The amended class diagram is shown in [Fig. 1](#CircularReferenceResolved).

<a name="CircularReferenceResolved"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/ObjectStore/develop/Assets/Diagrams/CircularReferenceResolved.svg" alt="Circular References Resolved" width="75%" /></div>
<p align="center"><b>Fig. 1. Simplified database library class diagram with ObjectStore</b></p>  

The important part is how ObjectStore is accessed. Accessing an ObjectStore instance via a regular reference would result in a three-node loop. Instead, a public ObjectStore variable, such as the predeclared instance, should be used. At the same time, the ObjectStore collection should be destroyed during termination to free the stored objects. The simplest way to achieve this goal is to destroy the ObjectStore variable itself. Unfortunately, an attempt to set a predeclared instance variable to Nothing crashes the application. On the other hand, the similarly behaving auto-assigned variables can be destroyed by setting them to Nothing. Thus, a public auto-assigned variable ObjectStore named after the class is declared in the *ObjectStoreGlobals* regular module located in the same folder, mimicking predeclared instances. Running Main.Main() with `ReferenceLoopManagementMode = REF_LOOP_OBJECT_STORE` setting should produce output similar to this:

    2021-11-27 03:19:39.682: DbManager   /Regular - Class_Initialize
    2021-11-27 03:19:39.683: DbConnection/Regular - Class_Initialize
    2021-11-27 03:19:39.683: DbStatement /Regular - Class_Initialize
    2021-11-27 03:19:39.680: ObjectStore - Class_Initialize
    2021-11-27 03:19:39.680: ObjectStore - Class_Terminate
    2021-11-27 03:19:39.680: DbManager   /Regular - Class_Terminate
    2021-11-27 03:19:39.680: DbConnection/Regular - Class_Terminate
    2021-11-27 03:19:39.680: DbStatement /Regular - Class_Terminate

indicating that all objects are properly terminated, including the ObjectStore. Termination of the ObjectStore variable destroys all parent references, and destruction starts from the top. Specifically, once ObectStore and DbManager instances are destructed, DbConnection, the top affected class, can be destructed, as it no longer has a parent holding its reference. Destruction of the DbConnection instance clears its children collection freeing its children.

The ObjectStore class has four methods. AllocateHandle reserves an object handle. If an optional string argument is provided, AllocateHandle checks whether the matching dictionary key already exists. If no such key exists, AllocateHandle adds the new key to the dictionary with an empty value and returns -1 (currency); otherwise, it returns 0. Alternatively, a new timestamp-based key is generated and returned, using the Unix Epoch time with milliseconds retrieved from the Timer function. FreeHandle removes a saved object reference. RefSet/RefGet saves/retrieves object references. RefSet requires that the key is preallocated and the corresponding value must be empty, preserving existing object reference, if set. If RefSet returns object handle on success and blank string on failure.
