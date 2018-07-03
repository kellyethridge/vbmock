VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bAnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'   bAnd
'
Option Explicit
Implements IConstraint

Private mMessage As String
Private mA As IConstraint
Private mB As IConstraint



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef a As Variant, ByRef b As Variant)
    SetConstraint a, mA
    SetConstraint b, mB
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConstraint(ByRef Value As Variant, ByRef Constraint As IConstraint)
    If TypeOf Value Is IConstraint Then
        Set Constraint = Value
    Else
        Set Constraint = Test.IsEqual(Value)
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    If Not mA.Eval(Value) Then
        mMessage = "Constraint 1 failed. " & mA.Message
    ElseIf Not mB.Eval(Value) Then
        mMessage = "Constraint 2 failed. " & mB.Message
    Else
        mMessage = ""
        IConstraint_Eval = True
    End If
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = mMessage
End Property
