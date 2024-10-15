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
        Dim Stream As Object
        Set Stream = CreateObject("ADODB.Stream")
        Stream.Open
        Stream.Type = 1
        Stream.Write http.responseBody
        Stream.SaveToFile LocalFilePath, 2
        Stream.Close
        Debug.Print "업데이트 파일 다운로드 완료!"
        
        ' ZIP 파일 해제
        Call UnzipFile(LocalFilePath, "C:\Center\Extracted\")
        
    Else
        Debug.Print "업데이트 파일을 다운로드하지 못했습니다. 상태 코드: " & http.Status
    End If
End Sub

Sub UnzipFile(ByVal ZipFilePath As String, ByVal DestinationFolder As String)
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    
    ' 대상 폴더가 존재하는지 확인하고 없으면 생성
    If Dir(DestinationFolder, vbDirectory) = "" Then
        MkDir DestinationFolder
    End If
    
    ' ZIP 파일 해제
    shellApp.Namespace(DestinationFolder).CopyHere shellApp.Namespace(ZipFilePath).Items
    Debug.Print "ZIP 파일이 성공적으로 해제되었습니다."
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
    
    Debug.Print "모든 모듈과 유저폼이 삭제되었습니다. 이제 업데이트를 실행합니다."
    
    ' 새로운 모듈 및 유저폼 추가
    ImportModules "C:\Center\Extracted\Center-main\Modules\"

    ' 마지막에 현재 모듈 삭제
    VBProj.VBComponents.Remove VBProj.VBComponents(CurrentModule)
    
    D
