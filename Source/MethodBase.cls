VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MethodBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    CopyRight (c) 2004 Kelly Ethridge. All Rights Reserved.
'
'    This file is part of VBMock.
'
'    VBMock is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBMock is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'
'   MethodBase
'
Option Explicit

Private Const INFINITE As Long = 0
Private Const DEF_CAPACITY As Long = 10


Private mSignature As CallSignature
Private mCalls() As IMockCall
Private mCallCount As Long
Private mName As String
Private mCallIndex As Long



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Invoke(ByRef Args() As Variant, ByRef retval As Variant)
    Do While mCallIndex < mCallCount
        If InvokeCall(Args, retval) Then Exit Sub
        mCallIndex = mCallIndex + 1
    Loop
    Err.Raise 5, mName, GetResourceString(RES_UnexpectedCall, mName)
End Sub



Friend Sub Verify()
    Dim i As Long
    Dim Min As Long
    Dim Count As Long
    
    For i = 0 To mCallCount - 1
        Min = Min + mCalls(i).Min
        Count = Count + mCalls(i).Count
    Next i
    If Count < Min Then
        Err.Raise 5, mName, GetResourceString(RES_MissingCall, mName, Min, Count)
    End If
End Sub



Friend Sub Setsignature(ByRef Params() As Variant)
    Set mSignature = New CallSignature
    mSignature.Init Params, mName
End Sub



Friend Sub Init(ByVal Name As String)
    mName = Name
End Sub



Friend Sub AddNewCall(ByRef retval As Variant, ByVal Min As Long, ByVal Max As Long)
    Dim NewCall As MockCall
    Set NewCall = New MockCall
    NewCall.Init Min, Max, retval, mSignature
    AppendCall NewCall
End Sub



Friend Sub AddNewError(ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal Min As Long, ByVal Max As Long)
    Dim NewCall As MockError
    Set NewCall = New MockError
    NewCall.Init Min, Max, Number, Source, Description
    AppendCall NewCall
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AppendCall(ByVal oCall As IMockCall)
    EnsureCapacity mCallCount + 1
    Set mCalls(mCallCount) = oCall
    mCallCount = mCallCount + 1
End Sub



Private Sub EnsureCapacity(ByVal RequiredCapacity As Long)
    Dim NewCapacity As Long
    NewCapacity = UBound(mCalls) + 1
    If RequiredCapacity > NewCapacity Then
        NewCapacity = NewCapacity * 2
        If RequiredCapacity > NewCapacity Then NewCapacity = RequiredCapacity
        ReDim Preserve mCalls(0 To NewCapacity - 1)
    End If
End Sub



Private Function InvokeCall(ByRef Args() As Variant, ByRef retval As Variant) As Boolean
    With mCalls(mCallIndex)
        If .Invoke(Args) Then
            GetReturnValue retval
            InvokeCall = True
        End If
    End With
End Function



Private Sub GetReturnValue(ByRef retval As Variant)
    If TypeOf mCalls(mCallIndex) Is MockCall Then
        Dim oCall As MockCall
        Set oCall = mCalls(mCallIndex)
        oCall.GetRetVal retval
    End If
End Sub


Private Sub Class_Initialize()
    ReDim mCalls(DEF_CAPACITY - 1)
End Sub

