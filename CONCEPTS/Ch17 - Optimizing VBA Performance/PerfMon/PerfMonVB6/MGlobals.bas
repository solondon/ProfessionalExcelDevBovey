Attribute VB_Name = "MGlobals"
'
'   Module to store Enums, UDTs and global variables used in the project
'
'   Version Date        Author          Comment
'   0.0.1   09-02-2004  Stephen Bullen  Initial Version
'
Public Enum pmAddRemove
    pmAddRemoveAdd
    pmAddRemoveRemove
End Enum

Public Enum pmScope
    pmScopeAllProjects
    pmScopeSelProject
    pmScopeSelModule
    pmScopeSelProc
End Enum

