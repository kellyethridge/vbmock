VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallSignature"
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
'   CallSignature
'
Option Explicit

Private mConstraintcount As Long
Private mConstraints() As IConstraint
Private mName As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Invoke(ByRef Args() As Variant)
    VerifySignatureLength Args
    EvaluateArguments Args
End Sub



Friend Sub Init(ByRef Params() As Variant, ByVal Name As String)
    mConstraintcount = GetLength(Params)
    SetConstraints Params
    mName = Name
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VerifySignatureLength(ByRef Args() As Variant)
    Dim length As Long
    length = GetLength(Args)
    If length <> mConstraintcount Then
        Err.Raise 5, mName, GetResourceString(RES_WrongArgCount, mName, mConstraintcount, length)
    End If
End Sub



Private Sub EvaluateArguments(ByRef Args() As Variant)
    Dim i As Long
    For i = 0 To mConstraintcount - 1
        EvaluateOneArgument i, Args(i)
    Next i
End Sub



Private Sub EvaluateOneArgument(ByVal Index As Long, ByRef Value As Variant)
    With mConstraints(Index)
        If .Eval(Value) = False Then
            Err.Raise 5, mName, GetResourceString(RES_IncorrectParameter, Index + 1, mName, .Message)
        End If
    End With
End Sub



Private Sub SetConstraints(ByRef Params() As Variant)
    Dim i As Long
    
    If mConstraintcount = 0 Then Exit Sub
    ReDim mConstraints(0 To mConstraintcount - 1)
    
    For i = 0 To mConstraintcount - 1
        If TypeOf Params(i) Is IConstraint Then
            Set mConstraints(i) = Params(i)
        Else
            Set mConstraints(i) = CreateDefaultConstraint(Params(i))
        End If
    Next i
End Sub



Private Function CreateDefaultConstraint(ByRef Value As Variant) As IConstraint
    Set CreateDefaultConstraint = Test.IsEqual(Value)
End Function






