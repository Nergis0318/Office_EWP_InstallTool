# Office_EWP_InstallTool

Excel / Word / PowerPoint 설치 도구

Windows에서 Microsoft Office Deployment Tool을 사용해 Excel, Word, PowerPoint를 설치하는 포터블 UI 도구입니다.

## 기능

- 기존 Office 제거를 위한 Windows 앱 및 기능 화면 열기
- OfficeScrubber 다운로드 및 실행
- `C:\Office` 폴더 생성
- Office Deployment Tool `setup.exe` 다운로드
- 내장된 `Configuration.xml` 생성
- 관리자 권한 `cmd.exe`에서 Office 설치 명령 실행
- Office 설치 완료 후 `C:\Office` 삭제

## 사용 방법

1. [Office_EWP_InstallTool_Portable.zip](https://github.com/Nergis0318/Office_EWP_InstallTool/releases/latest/download/Office_EWP_InstallTool_Portable.zip)을 다운로드 합니다.
2. `Office_EWP_InstallTool_Portable.zip`을 압축 해제합니다.
3. `Office_EWP_InstallTool.exe`를 관리자 권한으로 실행합니다.
4. `[1. 앱 및 기능 열기]`를 눌러 기존 Office가 있으면 제거합니다.
5. `[2. OfficeScrubber 실행]`을 누릅니다.
6. 열린 명령창에서 `[R] Remove all Licenses` 옵션을 선택합니다.
7. `[3. 설치 파일 준비]`를 눌러 `C:\Office`에 설치 파일을 준비합니다.
8. `[4. Office 설치 및 정리]`를 눌러 Office 설치를 시작합니다.
9. 설치가 정상 종료되면 `C:\Office` 폴더가 자동으로 삭제됩니다.

## 설치 대상

다음 앱만 설치됩니다.

- Excel
- Word
- PowerPoint

다음 앱은 제외됩니다.

- Access
- Groove
- Lync
- OneDrive
- OneNote
- Outlook
- Outlook for Windows
- Publisher

## 빌드 방법

이 프로젝트는 Windows 기본 .NET Framework 컴파일러로 빌드할 수 있습니다. 별도 Visual Studio나 .NET SDK가 필요하지 않습니다.

```powershell
& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" `
  /nologo `
  /target:winexe `
  /platform:x64 `
  /codepage:65001 `
  /win32manifest:app.manifest `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.IO.Compression.FileSystem.dll `
  /out:Office_EWP_InstallTool.exe `
  Office_EWP_InstallTool.cs
```

배포용 zip을 만들려면 다음 명령을 실행합니다.

```powershell
Compress-Archive `
  -LiteralPath "Office_EWP_InstallTool.exe","manual.txt","README.md","LICENSE","NOTICE" `
  -DestinationPath "Office_EWP_InstallTool_Portable.zip" `
  -Force
```

## 생성 파일

- `Office_EWP_InstallTool.exe`: 설치 없이 실행 가능한 포터블 UI 실행 파일
- `Office_EWP_InstallTool_Portable.zip`: 배포용 압축 파일
- `C:\Office\setup.exe`: Office Deployment Tool, 설치 완료 후 삭제됨
- `C:\Office\Configuration.xml`: Office 설치 구성 파일, 설치 완료 후 삭제됨

## 외부 다운로드

실행 중 다음 파일을 다운로드합니다.

- Office Deployment Tool: `https://officecdn.microsoft.com/pr/wsus/setup.exe`
- OfficeScrubber: `https://gitlab.com/-/project/11037551/uploads/f49f0d69e0aaf92e740a1f694d0438b9/OfficeScrubber_14.zip`

## 주의사항

- Office 설치와 정리 작업에는 관리자 권한이 필요할 수 있습니다.
- OfficeScrubber는 외부 도구이며, 실행 후 옵션 선택은 사용자가 직접 해야 합니다.
- 기존 Office 제거는 자동으로 수행하지 않습니다. Windows 설정에서 사용자가 직접 제거해야 합니다.
- Office 설치에는 인터넷 연결이 필요합니다.
- Office 설치 명령이 실패하거나 중간에 취소되면 `C:\Office`가 자동 삭제되지 않을 수 있습니다.

## 라이선스

Copyright 2026 Nergis

이 프로젝트는 Apache License 2.0에 따라 배포됩니다. 전체 라이선스 전문은 `LICENSE` 파일을 확인하세요.

이 프로젝트에는 Microsoft Office, Office Deployment Tool, OfficeScrubber가 포함되어 있지 않습니다. 해당 외부 도구와 Microsoft 제품은 각 제공자의 라이선스와 약관을 따릅니다.
