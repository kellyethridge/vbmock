VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsEqualIgnoreCase"
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
'   IsEqualIgnoreCase
'
Option Explicit
Implements IConstraint

Private mConstraint As IConstraint
Private mStringCompare As Boolean
Private mOriginalValue As Variant
Private mCompareString As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Value As Variant)
    If VarType(Value) = vbString Then
        mOriginalValue = Value
    End If
    Set mConstraint = Test.IsEqual(Value)
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    If (VarType(Value) = vbString) And (VarType(mOriginalValue) = vbString) Then
        mStringCompare = True
        mCompareString = Value
        IConstraint_Eval = (LCase$(mOriginalValue) = LCase$(Value))
    Else
        mStringCompare = False
        IConstraint_Eval = mConstraint.Eval(Value)
    End If
End Function

Private Property Get IConstraint_Message() As String
    If mStringCompare Then
        IConstraint_Message = GetResourceString(RES_ExpectedValue, mOriginalValue, mCompareString)
    Else
        IConstraint_Message = mConstraint.Message
    End If
End Property
