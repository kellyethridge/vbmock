VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsEqual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'   IsEqual
'
Option Explicit
Implements IConstraint

Private mExpectedValue As Variant
Private mMessage As String
Private mDelta As Double



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef ExpectedValue As Variant, Optional ByVal Delta As Double)
    If IsArray(ExpectedValue) Then Err.Raise 13, "IsEqual", "Cannot pass in array."
    
    SetExpectedValue ExpectedValue
    SetDelta Delta
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Eval(ByRef Value As Variant) As Boolean
    If IsArray(Value) Then Exit Function
    Select Case VarType(mExpectedValue)
        Case vbString:                      Eval = TestForString(Value)
        Case vbDouble, vbSingle:            Eval = TestForFloat(Value)
        Case vbDate:                        Eval = TestForDate(Value)
        Case vbObject, vbDataObject:        Eval = TestForObject(Value)
        Case vbNull:                        Eval = TestForNull(Value)
        Case Else:                          Eval = EvalEverythingElse(Value)
    End Select
End Function



Private Property Get Message() As String
    Message = mMessage
End Property



Private Sub SetExpectedValue(ByRef ExpectedValue As Variant)
    If IsObject(ExpectedValue) Then
        Set mExpectedValue = ExpectedValue
    Else
        mExpectedValue = ExpectedValue
    End If
End Sub



Private Sub SetDelta(ByVal Delta As Double)
    If Delta < 0# Then _
        Err.Raise 5, "IsEqual", "Delta must be non negative."
    
    mDelta = Delta
End Sub



Private Function TestForNull(ByRef Value As Variant) As Boolean
    If IsNull(Value) Then
        TestForNull = SetToTrue
    Else
        TestForNull = SetToFalse("Expected Null.")
    End If
End Function



Private Function TestForString(ByRef Value As Variant) As Boolean
    If VarType(Value) = vbString Then
        TestForString = EvalEverythingElse(Value)
    Else
        TestForString = SetToTypeMismatch(Value)
    End If
End Function



Private Function TestForFloat(ByRef Value As Variant) As Boolean
    Select Case VarType(Value)
        Case vbDouble, vbSingle
            TestForFloat = EvalFloat(Value)
        Case Else
            TestForFloat = SetToTypeMismatch(Value)
    End Select
End Function



Private Function TestForDate(ByRef Value As Variant) As Boolean
    If VarType(Value) = vbDate Then
        TestForDate = EvalDate(Value)
    Else
        TestForDate = SetToTypeMismatch(Value)
    End If
End Function



Private Function TestForObject(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        TestForObject = EvalObject(Value)
    Else
        TestForObject = SetToTypeMismatch(Value)
    End If
End Function



Private Function EvalObject(ByRef Value As Variant) As Boolean
    If Value Is mExpectedValue Then
        EvalObject = SetToTrue
    Else
        EvalObject = SetToFalse(GetResourceString(RES_NotSameObjectReference))
    End If
End Function



Private Function EvalDate(ByRef Value As Variant) As Boolean
    EvalDate = SetValueCondition(DateDiff("s", Value, mExpectedValue) = 0, Value)
End Function



Private Function EvalFloat(ByRef Value As Variant) As Boolean
    EvalFloat = SetValueCondition(Abs(mExpectedValue - Value) <= mDelta, Value)
End Function



Private Function EvalEverythingElse(ByRef Value As Variant) As Boolean
    EvalEverythingElse = SetValueCondition(Value = mExpectedValue, Value)
End Function



Private Function SetValueCondition(ByVal IsCorrectValue As Boolean, ByRef Value As Variant) As Boolean
    If IsCorrectValue Then
        SetValueCondition = SetToTrue
    Else
        SetValueCondition = SetToWrongValue(Value)
    End If
End Function



Private Function SetToWrongValue(ByRef Value As Variant) As Boolean
    SetToWrongValue = SetToFalse(GetResourceString(RES_NoValueReceived, mExpectedValue, IIf(IsEmpty(Value), "no value", "'" & Value & "'")))
End Function



Private Function SetToTypeMismatch(ByRef Value As Variant) As Boolean
    SetToTypeMismatch = SetToFalse(GetResourceString(RES_TypeMismatch, TypeName(mExpectedValue), TypeName(Value)))
End Function



Private Function SetToTrue() As Boolean
    SetToTrue = Not SetToFalse("")
End Function



Private Function SetToFalse(ByRef Message As String) As Boolean
    mMessage = Message
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    IConstraint_Eval = Eval(Value)
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = Message
End Property
