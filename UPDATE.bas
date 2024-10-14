Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    '??
    ' 업데이트할 파일 목록
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "Module3.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "견적서입력하기.bas", "견적종류.bas", "공정시험법및분장.bas", _
        "기본세팅.bas", "단축어골라내기.bas", "디버그세팅.bas", "리스트뷰_세팅.bas", "방류기준_찾아가.bas", _
        "방류기준거시기.bas", "방류기준만들기임시.bas", "법정양식거시기.bas", "분석결과입력_세팅.bas", _
        "분석결과조회.bas", "성적서견적서출력범위설정.bas", "소수점.bas", "수질측정기록부.bas", _
        "의뢰리스트만들기.bas", "출력하기.bas", "트리뷰_세팅.bas", "특정수질견적.bas", "합계.bas")
    
    ' 모든 파일을 다운로드하고 업데이트
    For i = LBound(fileList) To UBound(fileList)
        DownloadAndUpdateComponentFromGitHub fileList(i)
    Next i
End Sub

Sub DownloadAndUpdateComponentFromGitHub(ByVal fileName As String)
    Dim http As Object
    Dim url As String
    Dim fileData As String
    Dim filePath As String
    Dim componentType As String
    Dim stream As Object
    
    ' 파일 확장자 추출
    componentType = Right(fileName, 3)
    
    ' 깃허브에서 파일 가져오기 (Raw URL 사용)
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://raw.githubusercontent.com/1110x/Center/main/" & fileName
    http.Open "GET", url, False
    http.Send
    
    ' HTTP 요청 성공 시 처리
    If http.Status = 200 Then
        fileData = http.responseText
        filePath = "C:\Users\ironu\OneDrive\바탕 화면\" & fileName ' 경로를 명시적으로 설정
        
        ' ADODB.Stream 객체 사용하여 UTF-8로 파일 저장
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' 텍스트
        stream.Charset = "utf-8" ' UTF-8로 설정
        stream.Open
        stream.WriteText fileData
        stream.SaveToFile filePath, 2 ' 2는 덮어쓰기
        stream.Close
        
        ' 기존 VBA 구성요소 제거 후 새로 가져오기
        On Error Resume Next
        If componentType = "bas" Or componentType = "frm" Then
            ' VBProject가 존재하는지 확인
            If ThisWorkbook.VBProject Is Nothing Then
                MsgBox "VBProject를 찾을 수 없습니다."
                Exit Sub
            End If
            
            ' 구성요소가 존재하는지 확인 후 제거
            Dim comp As Object
            Set comp = ThisWorkbook.VBProject.VBComponents(Left(fileName, Len(fileName) - 4))
            If Not comp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove comp
            End If
        End If
        On Error GoTo 0
        
        ' 새 구성요소 추가
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Import filePath
        If Err.Number <> 0 Then
            MsgBox fileName & " 가져오기 실패: " & Err.Description
        Else
            MsgBox fileName & "가 성공적으로 업데이트되었습니다!"
        End If
        On Error GoTo 0
    Else
        MsgBox fileName & " 업데이트 실패: " & http.Status
    End If
End Sub
