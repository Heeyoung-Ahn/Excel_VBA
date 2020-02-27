-- common 사용자의 비밀번호 변경하여 실행
-- common 사용자에게 부여하는 권한
-- 1) common.users: SELECT, UPDATE
-- 2) common.logs: INSERT
-- DROP USER 'common'@'%'; -- common 유저가 이미 있을 경우
CREATE USER 'common'@'%' IDENTIFIED BY 'password';
GRANT USAGE ON *.* TO 'common'@'%';
GRANT INSERT  ON TABLE common.logs TO 'common'@'%';
GRANT SELECT, UPDATE  ON TABLE common.users TO 'common'@'%';
FLUSH PRIVILEGES;

