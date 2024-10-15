Sub UpdateModulesFromGitHub()

    Dim FileUrl As String
    Dim LocalFilePath As String
    Dim HttpReq As Object
    Dim GitRepoUrl As String
    GitRepoUrl = "https://github.com/1110x/Center/archive/refs/heads/main.zip" ' 메인 브랜치의 파일 다운로드 URL
    
    ' 임시로 다운로드할 위치 설정
    LocalFilePath = "C:\Center\update.zip"
    
    ' HTTP 요청으로 파일 다운로드
    Set HttpReq = CreateObject("MSXML2.XMLHTTP")
    HttpReq.Open "GET", GitRepoUrl, False
    HttpReq.Send
    
    If HttpReq.Status = 200 Then
        Dim Stream As Object
        Set Stream = CreateObject("ADODB.Stream")
        Stream.Open
        Stream.Type = 1
        Stream.Write HttpReq.responseBody
        Stream.SaveToFile LocalFilePath, 2
        Stream.Close
        MsgBox "업데이트 파일 다운로드 완료!"
    Else
        MsgBox "업데이트 파일을 다운로드하지 못했습니다."
        Exit Sub
    End If

    ' 다운로드 후 압축 해제 및 파일 적용하는 부분은 추가로 작성 필요

End Sub


Sub DeleteAllModules()

    Dim VBComp As Object
    Dim VBProj As Object
    Set VBProj = ThisWorkbook.VBProject
    
    ' 모든 모듈 삭제
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_ClassModule Or VBComp.Type = vbext_ct_MSForm Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
    
    MsgBox "모든 모듈과 유저폼이 삭제되었습니다."

End Sub

Sub ImportModules()

    Dim VBProj As Object
    Dim FolderPath As String
    FolderPath = "C:\Center\UpdatedModules\" ' 압축 풀린 모듈의 경로
    
    Set VBProj = ThisWorkbook.VBProject
    
    ' 모듈 추가
    VBProj.VBComponents.Import FolderPath & "Module1.bas"
    VBProj.VBComponents.Import FolderPath & "UserForm1.frm"

    MsgBox "새로운 모듈과 유저폼이 추가되었습니다."

End Sub

