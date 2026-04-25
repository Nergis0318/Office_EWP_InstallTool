Office_EWP_InstallTool

사용 순서:
1. Office_EWP_InstallTool.exe를 관리자 권한으로 실행합니다.
2. [1. 앱 및 기능 열기]에서 기존 Office가 있으면 제거합니다.
3. [2. OfficeScrubber 실행]을 누르고 명령창에서 [R] Remove all Licenses 옵션을 선택합니다.
4. [3. 설치 파일 준비]를 눌러 C:\Office, setup.exe, Configuration.xml을 준비합니다.
5. [4. Office 설치 및 정리]를 눌러 setup.exe /configure Configuration.xml을 실행합니다.
6. 설치가 정상 종료되면 C:\Office 폴더가 자동으로 삭제됩니다.

참고:
- setup.exe와 OfficeScrubber는 실행 시 Microsoft/GitLab 에서 다운로드됩니다.

라이선스:
- Copyright 2026 Nergis
- Apache License 2.0
- Microsoft Office, Office Deployment Tool, OfficeScrubber는 이 프로젝트에 포함되지 않으며 각 제공자의 라이선스와 약관을 따릅니다.
