---
layout: default
title: Introduction
nav_order: 1
permalink: /
---

Consider [Fig. 1](#CircularReferences) showing a simplified diagram of database library classes located in the virtual Demo folder in [RDVBA][] CodeExplorer. Here, DbManager is the top API class instantiated by the application. DbManager creates DbConnection class instances handling opening and closing database connections. Thus, DbManager acts as an [abstract factory][OOP in VBA] for the DbConnection class. The DbConnection class is, in turn, an abstract factory for the DbStatement class, handling the preparation and execution of database queries. Both DbManager and DbConnection classes keep object references for all generated children in dictionary-based collections, making it possible to traverse the object hierarchy during the termination clean-up process.  At the same time, DbConnection objects provide database connection handles to their DbStatement children. To access this handle, DbStatement holds a reference to its parent DbConnection, forming a reference loop.

<a name="CircularReferences"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/ObjectStore/develop/Assets/Diagrams/CircularReference.svg" alt="Circular References" width="50%" /></div>
<p align="center"><b>Fig. 1. Simplified class diagram of a database library</b></p>  

VBA cannot handle the disposal of objects with circular references automatically, causing memory leaks. One approach to resolve this issue is via weak references. However, VBA does not support this feature natively, so the code must rely on lower-level APIs to simulate it. The [Lazy Objects post][Weak Reference] presents one implementation of non-native object references invisible to VBA, and the [VBA-WeakReference][Weak Reference CB] project develops a similar but more efficient solution. This demo project illustrates two alternative approaches to managing circular references.


### Demo project overview

The VBA project is a part of the host Excel workbook *ObjectStore.xls* located in the repository root. The *Project* folder contains individual code modules, and the Assets folder contains figures used by the documentation. The virtual Demo folder in the host workbook (as seen from RDVBA project explorer) contains three class modules discussed above. The design of these classes mostly follows the patterns described in this RDVBA [post][Factories]. The fourth regular module runs the demo via the Main.Main() sub. Modules from the *Object Store* folder will be discussed later, and the *Common* folder contains several dependencies. This project also activates five library references, including Rubberduck AddIn referenced by tests modules and four other references needed by the dependencies (Windows Scripting Host Object Model, Microsoft Scripting Runtime, Microsoft Visual Basic for Applications Extensibility 5.3, and Microsoft ActiveX Data Objects 6.0 Library).

The three demo classes contain only the functionality necessary for illustrating the discussed topic, including the debugging code added for the same purpose. The demo includes three alternative regimes: one, showing manifestation of the circular reference issue, and two possible solutions relying either on the CleanUp cascade or the extra ObjectStore class. The first assignment in the Main.Main() runner selects one of these three regimes.


<!-- References -->

[RDVBA]: https://rubberduckvba.com/
[OOP in VBA]: https://rubberduckvba.wordpress.com/2016/01/11/oop-in-vba-immutability-the-factory-pattern/
[Weak Reference]: https://rubberduckvba.wordpress.com/2018/09/11/lazy-object-weak-reference/
[Weak Reference CB]: https://github.com/cristianbuse/VBA-WeakReference
[Factories]: https://rubberduckvba.wordpress.com/2018/04/24/factories-parameterized-object-initialization/
