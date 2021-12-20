Excel Alt+F11 매크로 메소드 생성 후 실행

================================================ 전체 시트명 가져오기 ===================================================================

Sub Test()
Dim sht As Worksheet

Dim i As Integer
Dim data As String

'전체 시트를 하나를 읽는다.
For Each sht In Worksheets

'선택한 셀 기준으로 행을 증가하면서
data = sht.Cells(1, "J")

ActiveCell.Offset(i, 0) = data

'다른방법 : 행/열 설정시 반드시 1이상의 값을 지정해야 한다.
'ActiveSheet.Cells(i, 1) = sht.Name

i = i + 1
Next sht
End Sub

=======================================================================================================================================


Sub Test()
Dim sht As Worksheet

Dim i As Integer
Dim data As String

'전체 시트를 하나를 읽는다.
For Each sht In Worksheets

'선택한 셀 기준으로 행을 증가하면서
data = sht.Cells(1, "J")

ActiveCell.Offset(i, 0) = data

'다른방법 : 행/열 설정시 반드시 1이상의 값을 지정해야 한다.
'ActiveSheet.Cells(i, 1) = sht.Name

i = i + 1
Next sht
End Sub

=======================================================================================================================================

Sub Test()
Dim sht As Worksheet

Dim i As Integer
Dim data As String

'전체 시트를 하나를 읽는다.
For Each sht In Worksheets

'선택한 셀 기준으로 행을 증가하면서
data = ActiveCell.Offset(i, 0)

sht.Cells(1, "J") = data

'다른방법 : 행/열 설정시 반드시 1이상의 값을 지정해야 한다.
'ActiveSheet.Cells(i, 1) = sht.Name

i = i + 1
Next sht
End Sub

=======================================================================================================================================


Sub Test()
Dim sht As Worksheet

Dim i As Integer
Dim data As String

'전체 시트를 하나를 읽는다.
For Each sht In Worksheets

data = sht.Cells(5, "C")
'선택한 셀 기준으로 행을 증가하면서
ActiveCell.Offset(i, 0) = data

'다른방법 : 행/열 설정시 반드시 1이상의 값을 지정해야 한다.
'ActiveSheet.Cells(i, 1) = sht.Name

i = i + 1
Next sht
End Sub

=======================================================================================================================================
