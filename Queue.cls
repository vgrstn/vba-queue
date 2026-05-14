'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder("Class")
'@ModuleDescription("Generic Queue Class.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Generic Queue Class
' Enumeration: For Each item in object
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.09.09
'
' Methods
' Clear             Removes all items from the queue
' Enqueue item      Adds an item to the rear of the queue
' Dequeue           Removes and returns the item in front of the queue
' Items([base])     Exports all queue items to an array
'
' Properties (Get)
' Count             Returns the number of items contained in the queue
' IsEmpty           Returns True if the queue is empty or False if not
' Peek              Returns the item in front of the queue (default)
'
' This generic Queue Class is implemented using an encapsulated VB Collection.
'
' Timings (ms) for one enqueue + one dequeue.
'     count      Class       System
' 1      10    0.00046      0.00518
' 2     100    0.00046      0.00516
' 3    1000    0.00046      0.00523
' 4   10000    0.00046      0.00340
' 5  100000    0.00046      0.00338
'
' The results for the Queue Class show a consistent performance independent of
' its actual size. The System.Collections.Queue uses late binding which is most
' likely the explanation for its relative poor performance.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum

' The position of the front of the queue.
Private Const Front As Long = 1

' Wrapper for private data.
Private Type TPRIVATE
    Queue As Collection
End Type
Private this As TPRIVATE


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    Set this.Queue = New Collection

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Enumerator
'@Description "Enumerates queue from front to rear."
Public Property Get Enumerate() As IEnumVARIANT
' Called by For Each by setting Attribute Enumerate.VB_UserMemId = -4.
' [_NewEnum] is the name of the hidden method to enumerate a VB Collection.
'
' Don't mutate the Queue object during the For Each loop. This will result in
' unpredictable behavior which reflects VB's native semantics.

    Set Enumerate = this.Queue.[_NewEnum]

End Property


'@Description "Removes all items from the Queue."
Public Sub Clear()

    Set this.Queue = New Collection

End Sub


'@Description "Returns the number of items contained in the Queue."
Public Property Get Count() As Long

    Count = this.Queue.Count

End Property


'@Description "Returns True if the queue is empty or False if not."
Public Property Get IsEmpty() As Boolean

    IsEmpty = (this.Queue.Count = 0)

End Property


'@DefaultMember
'@Description "Returns the item at the front of the Queue."
Public Property Get Peek() As Variant

    If this.Queue.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Peek", "Queue is empty"
    AssignVariant Peek, this.Queue.Item(Front)

End Property


'@Description "Removes and returns the item at the front of the Queue."
Public Function Dequeue() As Variant

    If this.Queue.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, VBA.TypeName(Me) & ".Dequeue", "Queue is empty"
    AssignVariant Dequeue, this.Queue.Item(Front)
    this.Queue.Remove Front

End Function


'@Description "Adds an item to the rear of the Queue."
Public Sub Enqueue(ByVal Item As Variant)

    this.Queue.Add Item

End Sub


'@Description "Exports all queue items to an array."
Public Function Items(Optional ByVal base As Long) As Variant()

    If this.Queue.Count = 0 Then
        ' Returns initialized empty array allowing For Each on object.Items.
        Items = Array()
        Exit Function
    End If

    Dim arr() As Variant: ReDim arr(base To base + this.Queue.Count - 1)
    Dim i As Long: i = base
    Dim var As Variant
    For Each var In this.Queue
        If VBA.IsObject(var) Then
            Set arr(i) = var
        Else
            arr(i) = var
        End If
        i = i + 1
    Next

    Items = arr

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Assigns a value or an object reference to a variable."
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub AssignVariant(ByRef var As Variant, ByVal value As Variant)
' Prevents getting a Collection item twice at the expense of a stack frame.

    If VBA.IsObject(value) Then
        Set var = value
    Else
        var = value
    End If

End Sub
