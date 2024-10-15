Sub DownloadFromGitHub()
    Dim http As Object
    Dim LocalFilePath As String
    Dim GitRepoUrl As String
    GitRepoUrl = "https://github.com/1110x/Center/archive/refs/heads/main.zip" ' 메인 브랜치의 파일 다운로드 URL
    
    ' 임시로 다운로드할 위치 설정
    LocalFilePath = "C:\Center\update.zip"
    
    ' WinHTTP 객체 생성
    Set http = CreateObject("WinHTTP.WinHTTPRequest.5.1")
    http.Open "GET", GitRepoUrl, False
    http.Send

    ' HTTP 상태 확인
    If http.Status = 200 Then
        ' 파일 다운로드 후 저장
        Set Stream = CreateObject("ADODB.Stream")
        Stream.Open
        Stream.Type = 1
        Stream.Write http.responseBody
        Stream.SaveToFile LocalFilePath, 2
        Stream.Close
        MsgBox "업데이트 파일 다운로드 완료!"
    Else
        MsgBox "업데이트 파일을 다운로드하지 못했습니다. 상태 코드: " & http.Status
    End If
End Sub
Sub DeleteAllModulesExceptCurrent()
    Dim VBComp As Object
    Dim VBProj As Object
    Dim CurrentModule As String
    
    ' 현재 실행 중인 모듈 이름 가져오기
    CurrentModule = "UP" ' 이 부분에 현재 모듈 이름을 적습니다.
    
    Set VBProj = ThisWorkbook.VBProject
    
    ' 현재 모듈을 제외하고 모든 모듈과 유저폼 삭제
    For Each VBComp In VBProj.VBComponents
        If (VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_ClassModule Or VBComp.Type = vbext_ct_MSForm) _
            And VBComp.Name <> CurrentModule Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
    
    ' 나중에 현재 모듈을 삭제할 수 있도록 메시지 출력
    MsgBox "모든 모듈과 유저폼이 삭제되었습니다. 이제 업데이트를 실행합니다."
    
    ' 이후 업데이트 실행 (새 모듈 추가 등)
    ImportModules "C:\Center\Extracted\Center-main\Modules\"
    
    ' 마지막에 현재 모듈 삭제
    VBProj.VBComponents.Remove VBProj.VBComponents(CurrentModule)
    
    MsgBox "현재 모듈도 삭제되었습니다. 업데이트 완료!"
End Sub
Sub ImportModules(FolderPath As String)
    Dim VBProj As Object
    Dim FileName As String
    Dim FilePath As String
    Dim FileExtension As String

    Set VBProj = ThisWorkbook.VBProject
    
    ' 폴더 내 파일 확인
    FileName = Dir(FolderPath & "*.*") ' 모든 파일 목록 불러오기
    
    ' 폴더에 있는 파일을 하나씩 가져오기
    Do While FileName <> ""
        FilePath = FolderPath & FileName
        FileExtension = LCase(Right(FileName, 4)) ' 파일 확장자 확인

        ' .bas 모듈 또는 .frm 유저폼일 경우에만 추가
        If FileExtension = ".bas" Or FileExtension = ".frm" Then
            VBProj.VBComponents.Import FilePath
        End If

        ' 다음 파일로 이동
        FileName = Dir
    Loop
    
    MsgBox "새로운 모듈과 유저폼이 추가되었습니다."
End Sub

