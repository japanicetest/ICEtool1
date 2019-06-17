VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_f_temp_計器リスト出力_差分表記"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public output_rev As String
Public comp_rev As String

Private Sub close_計器リスト出力フォーム_Click()
    Me.Undo
    DoCmd.Close
End Sub

Private Sub cmd_report_preview_Click()

Dim strSQL1 As String
Dim strSQL2 As String
Dim strSQL3 As String
Dim db As Database
Set db = CurrentDb
Dim myQuery1 As QueryDef
Dim myQuery2 As QueryDef
Dim myQuery3 As QueryDef
Dim i As Long
'Dim fieldscount As Long
'Dim fieldsname(999) As String
Dim strSQL As String

DoCmd.Close acQuery, "計器リストデータ"
DoCmd.Close acReport, "r_計器リスト_temp"

If IsNull(Me.combo_output_rev.Value) Then
GoTo nooutput

ElseIf IsNull(Me.combo_comp_rev.Value) Then
GoTo nocomp

Else


'本プロシージャ内で使っているquery名を削除。前回実行時のエラー対策。del_queryはfunction。
del_query ("計器リストデータ出力")
del_query ("計器リストデータ比較")
del_query ("計器リストデータ")

'出力したいrevisionのクエリを作成。revisionはフォームで入力した値
Set myQuery1 = db.CreateQueryDef("計器リストデータ出力", "Select * from t_INST_TAG_archive")

strSQL1 = "Select * FROM t_INST_TAG_archive where revision = " & "'" & Me.combo_output_rev.Value & "';"
db.QueryDefs![計器リストデータ出力].SQL = strSQL1

'比較対象のrevisionのクエリを作成。revisionはフォームで入力した値
Set myQuery2 = db.CreateQueryDef("計器リストデータ比較", "Select * from t_INST_TAG_archive")

strSQL2 = "Select * FROM t_INST_TAG_archive where revision = " & "'" & Me.combo_comp_rev.Value & "';"
db.QueryDefs![計器リストデータ比較].SQL = strSQL2
  
  
'fieldsname_instreport

  
'差分有無をチェックするための準備
'fieldscount_instreport = myQuery1.Fields.Count
'For i = 0 To fieldscount_instreport - 1
'    fieldsname_instreport(i) = myQuery1.Fields(i).Name
'
'  Debug.Print fieldsname_instreport(i)
'
'Next
    
 
'出力Rev／比較Revを統合したクエリを作成
Set myQuery3 = db.CreateQueryDef("計器リストデータ")
strSQL3 = "Select 計器リストデータ出力.*, 計器リストデータ比較.* from 計器リストデータ出力 left join 計器リストデータ比較 on 計器リストデータ出力.ID= 計器リストデータ比較.ID"
db.QueryDefs![計器リストデータ].SQL = strSQL3
  
  
'DoCmd.OpenQuery "計器リストデータ", , acReadOnly

DoCmd.OpenReport "r_計器リスト_temp", acViewPreview, , , , Me.combo_output_rev.Value & "," & Me.combo_comp_rev.Value
'DoCmd.OpenReport "r_計器リスト_temp", acViewPreview, , , , output_rev & "," & comp_rev



'Application.CurrentDb.QueryDefs.Delete "計器リストデータ出力"
'Application.CurrentDb.QueryDefs.Delete "計器リストデータ比較"
'Application.CurrentDb.QueryDefs.Delete "計器リストデータ"
  
'  db.Close
'  Set db = Nothing

End If

nooutput:
    MsgBox "出力Revを選択してください", vbOKOnly + vbInformation, "確認"
    Me.Undo
    Exit Sub
    
nocomp:
    MsgBox "比較Revを選択してください", vbOKOnly + vbInformation, "確認"
    Me.Undo
    Exit Sub
    
End Sub



Private Sub combo_output_rev_AfterUpdate()

output_rev = Me.combo_output_rev.Text

Debug.Print output_rev

End Sub


Private Sub combo_comp_rev_AfterUpdate()

comp_rev = Me.combo_comp_rev.Text

Debug.Print comp_rev

End Sub

