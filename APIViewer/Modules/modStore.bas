Attribute VB_Name = "modStore"
Option Explicit
'******************************************************************************
'This module contain shared storage information and declarations
'******************************************************************************
Public colDepnd As Collection   'list of unresoved dependency Type entries used by current selections
Public DeclChange As String     'contains added declaratrion if one has been added
Public DeclName As String       'new name for list

Public Fso As FileSystemObject  'file I/O interface
Public colAdded As Collection   'collection to store list of items added to select list
Public colAddFL As Collection   'same collection, but with full definitions
Public colNew As Collection     'Collection of User-Added entries
Public colDelete As Collection  'collection of API file entries to be deleted

''''-------------------
'''Public colConst As Collection
'''Public colDecl As Collection
'''Public colType As Collection
'''Public ColEnum As Collection
''''-------------------

Public Enum DeclareTypes
  Constants = 0
  Declares = 1
  Types = 2
  Enums = 3
End Enum

#If False Then
  Public Constants, Declares, Types, Enums
#End If

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

