Attribute VB_Name = "리스트뷰_세팅"

Sub L4_Set()
    ' ListView4 초기화 및 설정
    With UserForm1.ListView4
        ' 보기 형태를 Report로 설정 (컬럼 헤더 표시)
        .View = lvwReport
        
        ' 전체 행 선택 가능하도록 설정
        .FullRowSelect = True
        
        ' 그리드라인 표시
        .Gridlines = True
        
        ' 기존 컬럼 제거 (재설정 위해)
        .ColumnHeaders.Clear
        
        ' 첫번째 컬럼: 순서
        .ColumnHeaders.Add , , "구분", 40
        .ColumnHeaders.Add , , "의뢰일자", 65
        .ColumnHeaders.Add , , "시료이름", 180
        .ColumnHeaders.Add , , "법적기준", 70
    End With
End Sub
