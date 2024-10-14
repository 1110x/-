Attribute VB_Name = "UPDATE"
Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    
    ' 업데이트할 파일 목록
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "견적서입력하기.bas", "견적종류.bas", "공정시험법및분장.bas", _
        "기본세팅.bas", "단축어골라내기.bas", "디버그세팅.bas", "리스트뷰_세팅.bas", "방류기준_찾아가.bas", _
        "방류기준거시기.bas", "방류기준만들기임시.bas", "법정양식거시기.bas", "분석결과입력_세팅.bas", _
        "분석결과조회.bas", "성적서견적서출력범위설정.bas", "소수점.bas", "수질측정기록부.bas", _
        "의뢰리스트만들기.bas", "출력하기.bas", "트리뷰_세팅.bas", "특정수질견적.bas", "합계.bas")
    
    ' 모든 파일을 다운로드하고 업데이트
    For i = LBound(fileList) To UBound(fileList)
        Call DownloadAndUpdateComponent(fileList(i), Left(fileList(i), InStrRev(fileList(i), ".") - 1), Right(fileList(i), 3))
    Next i
End Sub

