FX Rate Calculator with Log In

1. 프로그램의 내용
	- 리본메뉴에 추가기능으로 메뉴 부착
	- 엑셀 유저폼을 이용하여 DB 로그인
	- 사용자별 환율 조회 환경 구성
	- 하나은행 사이트 환율 조회: 실시간, 특정조건
	- 로그인 시 체크항목
		1) 사용자 이름: Excel Application User Name으로 체크
		2) 비밀번호: 최초 접속 시 비밀번호 설정, 이후 접속 시 비밀번호 체크
		3) 프로그램 버전: DB에 저장되어 있는 프로그램 버전과 VBA Public Constant로 저정된 버전 체크
		4) 사용자 IP 체크: 최초 접속 시 접속 PC의 IP주소를 DB에 저장하고 이후 접속 시 비교
2. Create common DB
	- 설치 DB의 내용
		1) common.users
		2) common.logs
		# common 사용자로 DB에 접속하면 작업용 접속 정보를 VBA로 전달
	- 설치 방법: common.sql 실행 / argIP, argDB, argUN, argPW Default값 설정
	- common.users 테이블에 사용자 추가
3. Create fx_calculator DB
	- 설치 DB의 내용
		1) fx_calculator.currencies
		2) fx_calculator.currency_cal
	- 설치 방법: currency_cal.sql 실행
4. common 사용자 DB에 추가
	- 사용자명: common
	- 권한: logs - insert / users - select, update
	- 추가방법
		1) 코드 수정: grant_privileges_to_common_user.sql에서 비빌번호 수정
		2) 코드 실행
5. task 사용자 DB에 추가
	- 사용자명: task
	- 권한: task DB에 Execute, Select, Insert, Update, Delete, Drop, Lock Tables
	- 추가방법: HeidiSQL에서 GUI 메뉴로 추가
6. 사용하는 Sub, Function code
	- a_Common: Public Constant, Public Variable, DB연결, 기타 공통 코드
		1) Public Constant 값 설정: banner, programv, IPAddress, commonPW
	- a_ErrHandler: 에러 발생 시 보고(MsgBox)
	- a_Ribbon: 추가기능 리본메뉴 관리 및 실행 코드
	- a_Type: VBA에서 사용할 구조체 정의 모듈
	- a.WrigeLog: 로그 기록
	- fn_checkDoubleInput: 데이터 DB 입력 시 다양한 형태의 중복 입력 검토
	- fn_checkTextBox: TextBox 입력 값 데이터형에 따른 검토
	- fn_FXRateC: 환율계산(하나은행)
	- fn_GetLocalIPaddress: 사용자 PC IP주소 조회
	- fn_InputboxPW: InputBox 입력 시 보안처리
	- fn_sha512: DB에서 SHA-512방식으로 단방향 암호화 된 내용과 VBA에서 입력된 내용을 비교하기 위한 VBA 코드
	- s_loadDataToCBox: 콤보박스 리스팅
	- s_returnListPosition: 리스트박스의 원래 위치로 커서 반환
	- s_setCBox: 콤보박스 설정
	- s_setGlobalVariant.bas: 로그인 시 전역변수 설정	
7. Form code
	- f_login.frm: UserForm 파일
	- f_login.frx: UserForm의 디자인 파일
	- f_currency_cal.frm
	- f_currency_cal.frx
8. 완성본(명령단추 추가기능 형태)
	- FX_Calculator V1.0.xlam
9. VBA Library 추가
	- Microsoft ActiveX Data Objects 6.1 Library
	- Microsoft WinHTTP Services, version 5.1
	
