VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallBase"
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
'   CallBase
'
Option Explicit

Private mMin As Long
Private mMax As Long
Private mCallCount As Long



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function Invoke(ByRef Args() As Variant) As Boolean
    If HasCall Then
        mCallCount = mCallCount + 1
        Invoke = True
    End If
End Function



Friend Property Get Min() As Long
    Min = mMin
End Property



Friend Property Get Count() As Long
    Count = mCallCount
End Property



Friend Sub Init(ByVal Min As Long, ByVal Max As Long)
    If Min < 0 Then Err.Raise 5, , "Min cannot be negative."
    If Min > Max Then Err.Raise 5, , "Min cannot be less than Max."
    
    mMin = Min
    mMax = Max
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get HasCall() As Boolean
    HasCall = (mMax = 0) Or (mCallCount < mMax)
End Property


