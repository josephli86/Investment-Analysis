VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemoryDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDataSetSource
Private m_Data As clsDataSet

Public Sub Initialize(ByVal dataset As clsDataSet)
    If Not m_Data Is Nothing Then Err.Raise 5, , "Source already initialized."
    If dataset Is Nothing Then Err.Raise 5, , "Dataset required."
    Set m_Data = dataset
End Sub
Private Function IDataSetSource_LoadData() As clsDataSet
    If m_Data Is Nothing Then Err.Raise 5, , "Source not initialized."
    Set IDataSetSource_LoadData = m_Data
End Function
