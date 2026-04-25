# Excel / Word / PowerPoint 설치 도구

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

1. `OfficeInstallTool.exe`를 실행합니다.
2. `[1. 앱 및 기능 열기]`를 눌러 기존 Office가 있으면 제거합니다.
3. `[2. OfficeScrubber 실행]`을 누릅니다.
4. 열린 명령창에서 `[R] Remove all Licenses` 옵션을 선택합니다.
5. `[3. 설치 파일 준비]`를 눌러 `C:\Office`에 설치 파일을 준비합니다.
6. `[4. Office 설치 및 정리]`를 눌러 Office 설치를 시작합니다.
7. 설치가 정상 종료되면 `C:\Office` 폴더가 자동으로 삭제됩니다.

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
cd "C:\Users\Nergis\Documents\Codex\2026-04-26\files-mentioned-by-the-user-configuration"

& "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" `
  /nologo `
  /target:winexe `
  /platform:x64 `
  /codepage:65001 `
  /win32manifest:app.manifest `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.IO.Compression.FileSystem.dll `
  /out:OfficeInstallTool.exe `
  OfficeInstallTool.cs
```

배포용 zip을 만들려면 다음 명령을 실행합니다.

```powershell
Compress-Archive `
  -LiteralPath "OfficeInstallTool.exe","README.txt" `
  -DestinationPath "OfficeInstallTool_Portable.zip" `
  -Force
```

## 생성 파일

- `OfficeInstallTool.exe`: 설치 없이 실행 가능한 포터블 UI 실행 파일
- `OfficeInstallTool_Portable.zip`: 배포용 압축 파일
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
