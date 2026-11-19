VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ProviderName() As String
End Property

' contextState: Array(period, ifAdjusted, anchorPublDate, versionKey, asOfDate)
Public Function TryResolve(ByVal symbol As String, _
                           ByVal evalDate As String, _
                           ByVal contextState As Variant, _
                           ByRef outValue As Variant, _
                           ByRef outStatus As String, _
                           ByRef outSourceKey As String) As Boolean
End Function
