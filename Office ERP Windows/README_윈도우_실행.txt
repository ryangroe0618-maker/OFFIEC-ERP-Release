Office ERP Windows 실행 안내
============================

가장 쉬운 실행 방법
------------------

1. Windows PC에 Git을 설치합니다.
   - https://git-scm.com/download/win
2. 원하는 위치에서 터미널을 열고 아래 명령으로 처음 한 번만 받습니다.
   git clone https://github.com/ryangroe0618-maker/office-erp.git
3. clone 받은 office-erp 폴더 안의 Office_ERP_시작.vbs 를 더블클릭합니다.
   - 검은 CMD 창 없이 조용히 최신 버전을 확인하고 앱을 실행합니다.
   - git clone으로 받은 폴더면 실행 전에 최신본을 자동 확인합니다.
   - Office_ERP.exe 가 없으면 GitHub 최신 EXE 릴리즈를 자동으로 받아 실행합니다.
   - 이후 실행할 때도 OFFIEC-ERP-Release 최신 버전이 있으면 자동으로 교체 후 실행합니다.

실행 파일 설명
--------------

- START_WINDOWS.cmd
  문제가 생겼을 때 확인용으로 눌러볼 실행 파일입니다.
  GitHub 최신본과 최신 EXE 릴리즈를 확인한 뒤 실행합니다.
  Office_ERP.exe 가 있으면 exe를 실행하고, 없으면 최신 EXE 릴리즈를 받은 뒤 실행합니다.

- 업데이트_확인.ps1
  OFFIEC-ERP-Release 저장소의 최신 Release 버전을 확인합니다.
  PC에 저장된 .office_erp_version 값과 Release tag가 다르면 새 ZIP을 받아 적용합니다.

- Office_ERP_시작.vbs
  직원용 기본 실행 파일입니다.
  검은 CMD 창을 띄우지 않고 START_WINDOWS.cmd 를 뒤에서 실행합니다.

- Office_ERP_실행.bat
  Python 설치형 실행 파일입니다.
  PC에 Python 3.11 이상이 있으면 필요한 패키지를 자동 설치하고 ERP를 실행합니다.

- EXE_만들기.bat
  Python 없이 실행 가능한 Office_ERP.exe 를 만드는 파일입니다.
  exe를 만드는 PC에는 Python 3.11 이상이 한 번 필요합니다.
  완료되면 ZIP 파일 2개가 생성됩니다.
  - Office_ERP_Windows.zip: 작은 소스/Python 실행용, Office_ERP.exe 제외
  - Office_ERP_Windows_EXE.zip: 큰 단독 실행 배포용, Office_ERP.exe 포함

기존 안내
---------

1. 이 폴더 전체를 Windows PC에 다운로드합니다.
   - 폴더명이나 파일명이 한글이어도 실행되도록 UTF-8 실행 배치를 포함했습니다.

2. Python 설치형 실행은 Office_ERP_실행.bat 를 더블클릭합니다.
   - Python 가상환경(.venv)이 없으면 자동으로 생성합니다.
   - 필요한 패키지는 requirements.txt 기준으로 자동 설치합니다.

3. Python 없이 실행하려면 Windows PC에서 EXE_만들기.bat 를 한 번 실행합니다.
   - exe를 만드는 PC에는 Python 3.11 이상이 한 번 필요합니다.
   - 생성 완료 후 Office_ERP.exe 가 생깁니다.
   - 생성 완료 후 작은 Office_ERP_Windows.zip 과 큰 Office_ERP_Windows_EXE.zip 이 생깁니다.
   - 이후 다른 Windows PC에서는 Office_ERP_Windows_EXE.zip 압축을 풀고 Office_ERP.exe 로 실행할 수 있습니다.

4. Python이 없다는 메시지가 나오면 Python 3.11 이상을 설치해 주세요.
   - 설치 시 Add python.exe to PATH 체크 권장
   - 다운로드: https://www.python.org/downloads/windows/

5. 구글 API 인증이 필요한 기능은 인증 JSON 파일이 같은 폴더에 있어야 합니다.
   - 현재 포함 파일: spry-smithy-498901-a0-6fee30e17b3e.json
   - 다른 PC에서 실행할 때도 이 파일을 같이 둬야 합니다.

6. 추후 GitHub 자동 업데이트를 쓰려면 GitHub_업데이트.bat 를 사용합니다.
   - Git 설치 필요: https://git-scm.com/download/win
   - ZIP으로 받은 폴더는 자동 업데이트가 안 됩니다.
   - 자동 업데이트를 쓰려면 처음 설치를 아래 명령으로 받아야 합니다.
     git clone https://github.com/ryangroe0618-maker/office-erp.git
   - git clone으로 받은 폴더에서는 START_WINDOWS.cmd 실행 시 자동으로 최신본을 확인합니다.

주의
----
- Office_ERP_실행.bat 는 이 폴더를 기준으로 실행하므로 위치를 옮겨도 괜찮습니다.
- 결과 엑셀/CSV/PDF는 Windows 사용자 바탕화면 아래에 저장됩니다.
  - 엑셀/CSV: 바탕화면\LIST
  - PDF: 바탕화면\PDF
  - KASHION 분류: 바탕화면\KASHION
  - 구분 대기: 바탕화면\구분 대기
- 회사 PC 보안 정책에 따라 PowerShell 실행이 막히면 .bat 파일을 사용하세요.
