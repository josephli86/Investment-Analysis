Attribute VB_Name = "Mod_SecurityMaster"
Option Explicit

' 日常增量更新（首次运行若表为空会自动全量建表并初始化）
Public Sub UpdateSecurityMaster()
    RunSecurityMasterSync False
End Sub

' 强制全量重建（当聚源字典规则调整或需要重刷全市场时运行）
Public Sub RebuildSecurityMasterFull()
    If MsgBox("确认清空并全量重建本地证券主数据表 (DIM_SECURITY_MASTER)？", vbYesNo + vbQuestion, "证券主数据同步") <> vbYes Then Exit Sub
    RunSecurityMasterSync True
End Sub

Private Sub RunSecurityMasterSync(ByVal forceFull As Boolean)
    Dim service As clsSecurityMasterService, result As clsAnalysisResult
    Dim summary As clsDataTable, t0 As Double
    On Error GoTo Failed
    t0 = Timer
    Set service = New clsSecurityMasterService
    Set result = service.SyncFromOracle(forceFull)
    Set summary = result.table("SYNC_SUMMARY")

    MsgBox "证券主数据同步完成！" & vbCrLf & _
           "目标库: " & CStr(summary.ValueAt(1, "TARGET_SOURCE")) & vbCrLf & _
           "同步模式: " & CStr(summary.ValueAt(1, "SYNC_MODE")) & vbCrLf & _
           "同步起始时间: " & CStr(summary.ValueAt(1, "CUTOFF_TIME")) & vbCrLf & _
           "更新记录数: " & CStr(summary.ValueAt(1, "SYNCED_ROWS")) & " 行" & vbCrLf & _
           "耗时: " & Format$(Timer - t0, "0.00") & " 秒 (RunId=" & result.RunId & ")", _
           vbInformation, "证券主数据同步"
    Exit Sub
Failed:
    MsgBox "证券主数据同步失败: " & Err.description, vbExclamation, "证券主数据同步"
End Sub
