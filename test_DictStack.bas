Option Explicit

Public Function TestDictStack() As Boolean
    Debug.Assert pop_stack = True
End Function

Private Function pop_stack() As Boolean
    Dim popped As Variant
    
    Dim stack As DictStack
    Set stack = New DictStack
    
    stack.push "one", 1
    popped = stack.pop
    If popped(0) = "one" Then pop_stack = True
End Function
