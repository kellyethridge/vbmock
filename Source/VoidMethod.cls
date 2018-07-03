VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoidMethod"
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
'   VoidMethod
'
Option Explicit

Private Const INFINITE As Long = 0

Private mBase As MethodBase



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddCall(Optional ByVal Count As Long = 1)
    mBase.AddNewCall Empty, Count, Count
End Sub

Public Sub AddCallRange(ByVal Min As Long, ByVal Max As Long)
    mBase.AddNewCall Empty, Min, Max
End Sub

Public Sub AddCallMin(ByVal Min As Long)
    mBase.AddNewCall Empty, Min, INFINITE
End Sub

Public Sub AddCallMax(ByVal Max As Long)
    mBase.AddNewCall Empty, 0, Max
End Sub

Public Sub AddCallAny()
    mBase.AddNewCall Empty, 0, INFINITE
End Sub

Public Sub AddError(ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String, Optional ByVal Count As Long = 1)
    mBase.AddNewError Number, Source, Description, Count, Count
End Sub

Public Sub AddErrorRange(ByVal Min As Long, ByVal Max As Long, ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String)
    mBase.AddNewError Number, Source, Description, Min, Max
End Sub

Public Sub AddErrorMin(ByVal Min As Long, ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String)
    mBase.AddNewError Number, Source, Description, Min, INFINITE
End Sub

Public Sub AddErrorMax(ByVal Max As Long, ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String)
    mBase.AddNewError Number, Source, Description, 0, Max
End Sub

Public Sub AddErrorAny(ByVal Number As Long, Optional ByVal Source As String, Optional ByVal Description As String)
    mBase.AddNewError Number, Source, Description, 0, INFINITE
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal base As MethodBase)
    Set mBase = base
End Sub



Friend Sub Invoke(ByRef Args() As Variant)
    mBase.Invoke Args, Empty
End Sub



Friend Sub Newsignature(ByRef Params() As Variant)
    mBase.Setsignature Params
End Sub



