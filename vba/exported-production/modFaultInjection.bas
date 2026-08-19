Attribute VB_Name = "modFaultInjection"
Option Explicit

' ============================================================================
' modFaultInjection - deterministic one-shot fault injection at mutation
' boundaries (test-only). NOT part of the operator workflow.
'
' Boundaries (see evidence/vba/atomicity-fault-injection.md):
'   1 = before transaction append
'   2 = immediately after transaction row allocation/append
'   3 = before container mutation
'   4 = during/after partial container mutation
'   5 = after container mutation but before successful workflow completion
'
' ArmFault(point, fireOnNthCall) arms a fault that raises error 2999 on the
' Nth matching-boundary call (default 1). The error propagates through the
' production call chain so each operation's rollback handler runs; FaultAt
' returns False when no fault is armed (zero production overhead).
' ============================================================================

Public Const FAULT_NONE As Long = 0
Public Const FAULT_BEFORE_TXN_APPEND As Long = 1
Public Const FAULT_AFTER_TXN_APPEND As Long = 2
Public Const FAULT_BEFORE_CONTAINER_MUTATION As Long = 3
Public Const FAULT_DURING_CONTAINER_MUTATION As Long = 4
Public Const FAULT_AFTER_CONTAINER_MUTATION_BEFORE_COMPLETE As Long = 5

Private Const FAULT_ERR_BASE As Long = 2999

Private m_armed As Boolean
Private m_point As Long
Private m_consumed As Boolean
Private m_callCount As Long
Private m_fireOnNth As Long

Public Sub ArmFault(ByVal point As Long, Optional ByVal fireOnNthCall As Long = 1)
    m_armed = True
    m_point = point
    m_consumed = False
    m_callCount = 0
    If fireOnNthCall < 1 Then fireOnNthCall = 1
    m_fireOnNth = fireOnNthCall
End Sub

Public Sub DisarmFault()
    m_armed = False
    m_point = FAULT_NONE
    m_consumed = False
    m_callCount = 0
    m_fireOnNth = 1
End Sub

Public Function IsArmed() As Boolean
    IsArmed = m_armed
End Function

Public Function FaultAt(ByVal point As Long) As Boolean
    ' Raises error 2999 on the Nth matching call while armed; otherwise False.
    If Not m_armed Then Exit Function
    If m_consumed Then Exit Function
    If m_point <> point Then Exit Function
    m_callCount = m_callCount + 1
    If m_callCount < m_fireOnNth Then Exit Function
    m_consumed = True
    m_armed = False
    Err.Raise FAULT_ERR_BASE, "modFaultInjection", _
              "INJECTED FAULT at boundary " & point & " (call " & m_callCount & ")"
End Function

Public Function LastFaultPoint() As Long
    LastFaultPoint = m_point
End Function
