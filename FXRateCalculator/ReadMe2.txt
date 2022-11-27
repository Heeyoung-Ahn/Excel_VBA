1) 데이터베이스 설계(ERD)
프로그램에 필요한 데이터베이스를 설계하고 HeidiSQL 프로그램을 이용하여 데이터베이스 구성해보세요.
common
- 구성 테이블
   1) users: 프로그램 사용자 정보, 비밀번호 초기화여부, 작업 DB 접속 정보, 최근 접속일시
   2) logs: 프로그램 사용 로그 기록
fx_calculator
- 구성 테이블
   1) currencies: 화폐 목록
   2) currency_cal: 프로그램 사용자별 환율 조회 목록

2) 데이터베이스에 자료 업로드
fx_calculator 스키마의 currencies 테이블에 엑셀 자료를 업로드 하세요.
- 테이블 설명
   1) currencies: 하나은행 사이트에서 환율조회가 가능한 화폐의 목록161개
   2) currency_cal: 161개의 화폐 목록 중에서 사용자 마다 별도의 목록을 구성할 수 있도록 해주는 테이블
       currency_id와 user_id를 기본키(Primary Key)로 설정해도 무방

엑셀 파일 업로드 방법
- ExcelToDB 프로시저 사용
- HeidiSQL의 Tools > Import CSV file 이용
- 업로드할 엑셀 파일: currencies.xlsx

3) 데이터베이스 사용자 추가 및 권한 부여
데이터베이스 사용자를 추가하고 권한을 부여하세요.

참조사항
- 프로그램 사용자와 데이터베이스 사용자는 다릅니다.
- 프로그램 사용자마다 데이터베이스 사용자를 만들어 주는 것은 아닙니다.
- 프로그램 사용자는 실제 프로그램을 사용하는 모든 사용자에 대해 ID가 발급되지만,
  데이터베이스 사용자는 접근가능한 데이터베이스 스키마와 권한에 따라 공통 계정으로 발급합니다.
  (예: common, task)
- 데이터베이스 'common' 사용자는 사용자를 확인할 수 있는 데이터베이스 접근권한만을 부여해준 사용자로
  1) 엑셀에서 입력한 사용자 정보와 데이터베이스에 입력된 사용자 정보를 비교하기 위해 필요한 select권한과
      비밀번호 변경, 접속 일시 기록 등에 필요한 update 권한을  common 데이터베이스에 한하여 부여합니다.
  2) 'common' 사용자의 DB 접속에 필요한 정보는 애플리케이션(엑셀 VBA)에 기록되어 있어야 하기 때문에
      필요 최소한의 DB권한만 부여합니다.
- 데이터베이스 'task' 사용자는 작업용 데이터베이스에 접근하여 실제 업무 수행(조회, 추가, 삭제, 업데이트)을
   할 수 있도록 작업용 데이터베이스에 해당 권한을 부여합니다.
  1) 프로그램 사용자가 'common' 사용자로 데이터베이스테 접근하여 사용자정보 확인이 완료되면 프로그램 사용자에게
      'task' 사용자로 데이터베이스에 접근할 수 있는 정보를 Public Constant로 넘겨 줍니다.
       # 'common.users' 테이블에  argIP, argDB, argUN, argPW로 'task' 사용자의 접속 정보를 기록해 두고 로그인 확인 시
          엑셀 vba의 Public Constant로 넘겨 줌
      # 엑셀 vba에서 엑셀이 종료될 때까지 데이터베이스 'task' 사용자 접속 정보를 가지고 있으면서 DB접속이 필요할 때 사용
  2) 애플리케이션을 사용하는 동안 항상 DB 에 접속해 있는 것이 아니라 DB에서 작업이 필요한 순간에만 잠시 접속하고
      작업이 마치면 바로 접속을 종료합니다.

common  사용자 만들기
- HeidiSQL에서 사용자 관리 메뉴를 통해 common 사용자를 추가하고
- common 데이터베이스의 users 테이블에 대해 select, update 권한을 부여합니다.
- common 데이터베이스의 logs 테이블에 대해 insert 권한을 부여합니다.
  # common 사용자로 데이터베이스에 접속하여 사용자 정보를 업데이트 하는 내용도 로그에 남겨야 하기 때문

task 사용자 만들기
- HeidiSQL에서 사용자 관리 메뉴를 통해 task 사용자를 추가하고
- common 데이터베이스의 logs 테이블에 대해 insert 권한을 부여합니다.
- fx_calculator 데이터베이스에 대해 Execute, Select, Insert, Update, Delete, Drop, Lock Tables 권한을 부여합니다.

4) 엑셀 파일 만들고 공통 모듈 부착
엑셀 기반 DB를 활용하는 프로그램을 만들때 항상 사용되는 공통 VBA 모듈이 있습니다.
이를 표준화해두고 재활용할 수 있도록 준비해 두면 프로그램 제작 시간이 단축되고 유지보수가 간결해 집니다.

a_common.bas
- 데이터베이스 연결, 조작에 필요한 코드 포함
- common 계정 로그인 정보 포함
- Public Constant, Public Variable 포함
- 기타 공통 코드

a_WriteLog.bas
- 데이터베이스에 로그를 기록하기 위한 프로시저
- 에러 로그: a_ErrHandler.bas에서 로그 기록에 필요한 정보를 담아서 writeLog Sub 프로시저 실행
- 액션 로그: 애플리케이션에서 데이터베이스에 영향을 미치는 실행이 있을 경우 각 프로시저에서 로그 기록 정보를 담아서 writeLog 프로시저 실행
- 로그 기록 정보: 프로시저명, 테이블명, SQL문, 에러코드(0: 정상, 1: 에러), 유저폼이름, 실행내용, 영향받은 레코드수

a_ErrHandler.bas
- 애플리케이션 또는 DB에서 에러 발생 시 처리하는 프로시저
- 사용자에게는 에러 발생내용을 디버깅하여 메시지박스로 보여주고
- 데이터베이스 에러 로그 기록은 CallDBtoRS 프로시저 또는 executeSQL 프로시저에서 로그길에 필요한 정보를 담아서 실행
  (CallDBtoRS, executeSQL 프로시저는 a_common.bas 모듈에 있음)

a_Ribbon.bas
- 엑셀 리본메뉴에 부착되는 추가기능으로 프로그램을 개발할 경우 사용되는 프로시저
- 앞의 3개의 프로시저는 수정사항이 많지 않은 공통 모듈이지만, a_Ribbon.bas 모듈은 그 틀과 형식은 공통이지만 내용은 프로그램에 맞춰서 대부분 수정이 되어야 함
- 구성요소
   1) 리본메뉴 클릭 시 프로시저 실행을 위한 프로시저: run_RibbonControl(Button AS Office.IRibbonControl)
   2) 준비중인 리본메뉴 클릭 시 처리를 위한 프로시저: RibbonButton_Error(sbID As String)
   3) 리본메뉴에 해당하는 실행 처리를 위한 프로시저

기타 공통 모듈
- 로컬PC의 IP주소를 반환하는 fn_GetLocalIPaddress.bas
- VBA InputBox 입력 내용을 '*'로 만들어 주는 fn_InputboxPW.bas
- 데이터베이스에 단방향 암호로 입력된 내용과 VBA에서 입려한 내용을 비교할 수 있도록 VBA입력 내용을 sha 512 암호화 해주는 fn_sha512.bas
- 'common' 사용자로 DB에 연결하여 애플리케이션 사용자 정보 확인 완료 시 애플리케인션에 task 사용자로 DB에 접속할 수 있도록 DB연결정보와 사용자 정보를 Public Constant로 반환해주는 sb_setGlobalVariant.bas

5) 추가기능 만들기
엑셀 VBA로 개발된 프로그램의 접근성을 높이기 위해 엑셀 리본메뉴 형태의 추가기능을 만들 수 있습니다.
엑셀 리본메뉴 추가기능 관련 자세한 내용은 네이버 검색을 통해 별도의 학습을 진행해 주시기 바랍니다.
(오빠두 추가기능, AddIn 리본메뉴의 다양한 방식)
# 추가기능은 'xlam' 확장자로 저장됨

추가기능 기초
- 추가기능으로 파일을 저장 시에도 엑셀 시트를 사용할 수 있지만 권장하지 않습니다.
- 만약 추기가능 파일에 시트를 사용하였고 추후 이 시트에 작성된 내용을 수정하려면 추가기능이 실행된 상태에서 VBA 편집기 메뉴에 진입 후
  프로젝트 탐색기에서 현재통합문서 클릭 후 속성창에서 IsAddin을 False로 변경하여 편집 후 완료시 다시 True로 변경합니다.
- 따라서 추가기능으로 저장할 엑셀 파일에 작성되는 VBA 코드는 엑셀 워크시트를 참조 및 활용하지 않도록 합니다.
   단, DB의 자료를 엑셀로 내보내는 경우 워크북을 새로 만들어서 자료를 내보내고 원하는 위치에 저장하는 방법으로 운영
- 추가기능의 기본 저장위치는 'C:\Users\Administrator\AppData\Roming\Microsoft\AddIns' 이지만 다른 위치에 저장되어도 작동에는 문제가 안됨
- 추가기능으로 저장 및 실행되면 기본적으로 엑셀이 실행될 때 자동으로 추가기능도 함께 시작되지만 만약 자동으로 추가기능이 실행되지 않으면
  '파일 > 옵션 > 추가기능 > 이동 > 찾아보기'  후 목록에서 체크하여 엑셀이 실행될 때마다 자동으로 추가기능이 실행되게 할 수 있습니다.
- 추가기능을 엑셀 리본메뉴 형태로 제작하려면 
  1) VBA 모듈에 a_Ribbon.bas 모듈을 부착하고: 리본메뉴를 컨트롤 하는 모듈
  2) 리본메뉴 편집 프로그램을 설치한 후 추가기능 파일에 리본메뉴를 구성할 코드를 XML 형태로 작성해줘야 합니다(매뉴얼 바로가기).
      # 리본메뉴 편집 프로그램에서 xlam 파일을 리본메뉴로 구성하기 위해서는 xlam 파일을 복호화해야 합니다.
      # 리본메뉴 구성을 위한 XML 코드는 매뉴얼의 샘플코드를 참조하여 작업합니다(Custom UI Editor 사용 방법 - 참조)

리본메뉴 추가기능 제작
- Office 2007
  # Custom UI 추가: Ribbon X12
  # Custom UI Code: <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
- Office 2010 이상
  # Custom UI 추가: Ribbon X14
  # Custom UI Code: <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  
  6) 사용자 정의 폼 만들기
프로그램의 최종적 목표이며 프로그램 사용자들에게 필요한 건

1) 업무의 각 절차에서 발생되는 데이터의 효율적인 수집 및 업무 지원
2) 수집된 데이터로 부터 의미있는 통계 구성입니다.

이를 위해 데이터베이스를 설계하고 데이터의 유통 흐름에 따른 업무 절차를 기획하였습니다.
이 단계는 사용자들이 데이터베이스에 접근하여 자료를 입력 및 조회하기 위한 사용자 인터페이스를 엑셀 VBA의 사용자정의폼으로 구성해 보는 단계입니다.

일반적인 사용자 정의 폼의 작업 순서와 구성은 11강 내용을 참조하여 주세요.
여기서는 로그인 폼과 환율 조회 폼에 대해 설명을 드립니다.
다른형태의 프로그램 제작 시에도 아래 사용된 내용은 기본으로 사용되는 코드들이니 디버깅 하면서 잘 살피시길 바랍니다.

로그인 폼
- VBA 모듈명: f_login.frm(유저 폼 파일), f_login.frx(유저폼 디자인 파일)
- 사용하는 DB 스키마: common.users
- 기능
   # 등록된 사용자 체크: Function checkUserNm(ByVal argUserNM As String) As Boolean
      - 엑셀 사용자(Application.UserName)가 common.users 테이블에 등록된 사용자인지 체크
   # 비빌번호 설정 체크: Sub checkInitialPW()
      - 등록된 사용자의 경우 DB에 비밀번호가 설정되어 있었는지 체크하고 설정하도록 진행
   # 비밀번호 설정: Sub registerNewPW()
      - 사용자는 등록된 사용자이지만 비밀번호 설정이 안되어 있는 경우(pw_initialize = 1) 신규비밀번호 등록
   # 사용자 확인: Sub cmd_query_Click()
      - common.users 테이블에 등록된 정보와 login 폼에 입력되거나 사용자 PC에서 수집한 정보 비교하여 아래의 4가지 사항 체크
      - 사용자 이름 등록 여부 체크
      - 프로그램 최신 버전 체크
      - IP 체크: 최초 접속 시에는 DB에 사용자 PC의 IP 기록
      - 비밀번호 체크(Function checkPW(ByVal argPW As String) As Boolean)
      - 환영인사
        # 모든게 다 맞으면 Public Constant인 'checkLogin'값을 1로 설정
        # 작업용 데이터베이스 접속 시 사용할 task 사용자의 접속 정보를 VBA의 Public Constant로 넘겨 줌
          (SUB setGlobalVariant(Optional ProcedureNM As String = "NULL"))
        # 최근 접속시간 DB에 기록
        # 환영인사
    -  비밀번호 변경: Sub cmd_chgPW_Click()

환율조회
- VBA 모듈명: f_currency_cal.frm(유저 폼 파일), f_currency_cal.frx(유저폼 디자인 파일)
- 사용하는 DB 스키마: fx_calculator.currencies(화폐 목록), fx_calculator.currency_cal(사용자별 환율 조회 화폐 목록)
- 기능
   # cbo_FX: fx_calculator.currencies 에서 화폐 목록 조회 및 선택
      - 화폐추가 명령단추 클릭 시 fx_calculator.currency_cal 테이블에 currency_id, user_id, 환율 등의 정보를 Insert
   # lst1: fx_calculator.currency_cal 테이블에 등록된 화폐 및 환율 정보 조회 및 관리(삭제, 업데이트)
   # 환율 업데이트: Sub cmd_update_Click()
      - fx_calculator.currency_cal 테이블에 등록되어 있는 사용자의 화폐에 대해 조회일에 해당하는 환율 업데이트
   # 환율 조회: Sub cmd_refer_Click()
      - 조회일, From 화폐, To 화폐에 따른 환율 조회
   # 유저 폼 기초 코드
      - Sub UserForm_Initialize(): 폼 열때 실행되는 코드
        # 전역변수 설정, 로그인 체크
        # 개체 선언, 컨트롤 설정, 입력항목 초기화
   # Sub control_initialize1(): 입력항목 초기화
   # Sub lst1_Click(): lst1 클릭 시 클릭한 항목의 정보를 각 컨트롤에 넘겨 줌
   # 날짜관련 처리: 입력 시 날짜 입력 유효성 체크, 라벨 클릭 시 오늘날짜 채워넣기
   # 금액관련 처리: 입력 시 금액 입력 유효성 체크
   # Sub loadDataToList(argListBox As MSForms.ListBox, Optional ByVal queryKey As String)
      - fx_calculator.currency_cal 테이블에서 lst1으로 자료 반환하는 프로시저
   # 화폐 추가, 삭제 관련: 유효성 검사, 중복 체크, 데이터 기록, 로그 기록
   
   7) 사용자 정의 폼 공통 모듈 부착
사용자 정의 폼을 통해 데이터베이스를 조작하기 위해 사용되는 공통 모듈에 대한 설명입니다.
내용을 숙지하고 이를 추기가능에 부착하여 주세요.

a_Type
- 구조체를 모아두는 모듈
- DB의 필드(컬럼)을 효과적으로 사용하기 위해 구조체로 정의하여 사용

fn_checkDoubleInput
- 유저 폼에서 입력한 내용이 데이터베이스 테이블에 기록 시 중복되는지 체크하기 위한 모듈
- 데이터베이스의 중복을 체크하는 것은 다양한 기준과 방법이 있음
  # 특정 필드의 중복 체크: Function checkDoubleInput(fieldNM As String, Data As Variant, tableNM As String, formNM As String, Optional ByVal beforeData As Variant = Empty) As Boolean
    - 입력하려는 값이 특정 컬럼에 있는 지 체크하는 방식으로 진행
  # 관계 테이블 중복 체크: Function checkDoubleInput2(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant, tableNM As String, formNM As String) As Boolean
    - 두 개의 컬럼을 조합하여 중복을 체크하는 경우
  # 기간 관계 테이블 중복 체크: Function checkDoubleInput3(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant,  start_dt As Date, end_dt As Date, tableNM As String, formNM As String) As Boolean
    - 특정 기간 내 두 개의 컬럼을 조합하여 중복을 체크하는 경우
  # 기간 데이터 중복 체크: Function checkDoubleInput4(dataType As Integer, fieldNM As String, Data As Variant, start_dt As Date, end_dt As Date, tableNM As String, formNM As String) As Boolean
    - 특정 기간 내 한 개의 컬럼에 중복을 체크하는 경우

fn_checkTextBox
- 유저 폼의 TextBox에 입력된 값의 유효성 검사 모듈
- 체크 사항
  # 입력 여부
  # 숫자인 경우 숫자 데이터 검증
  # 날짜인 경우 날짜 데이터 검증
  # 입력 길이 검증

sb_loadDataToCBox
- 데이터베이스의 데이터를 ComboBox에 반환하는 모듈
- 콤보박스 컨트롤의 기본 설정은 sb_setCBox 모듈을 통해 진행
  # 하나의 프로젝트(프로그램)에서 특정 콤보박스는 여러 사용자 정의 폼에서 공통으로 사용되는 경우가 많기 때문에 이를 공통 프로시저로 만들어서 사용

sb_returnListPosition
- ListBox에 자료 추가 또는 수정 후에 추가 또는 수정된 리스트 항목에 커서를 두기 위한 모듈