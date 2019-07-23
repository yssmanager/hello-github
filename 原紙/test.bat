@echo off

REM �����ɃR�s�[���̃t�@�C���������
set file_name=���㕪��(����).xlsx

REM �p�X���[�h�̐ݒ�
set PASSWORD=sp

REM �p�X���[�h�F��
set /p IN_PASS="�p�X���[�h����͂��Ă�������: "
if %PASSWORD%==%IN_PASS% (
    echo �F��OK
    call :main
    echo;
    echo ���s�͐���I�����܂����B
) else (
    echo �p�X���[�h���Ⴂ�܂��B���s���I�����܂��B
)
goto :end
exit


REM ---�T�u���[�`��---

:main
REM �R�s�[��̃t�H���_�[�ݒ�(��O�̃f�B���N�g���[)
pushd %~dp0..
set put_folder=%CD%
echo %put_folder%
popd

REM �����̒l����
call :input_month
call :input_date_start
call :input_date_end

echo %month%��%date_start%������%month%��%date_end%���܂ł̃t�@�C�����쐬���܂��B

for /l %%i in (%date_start%,1,%date_end%) do (
    call :makefiles %%i
)
exit /b


:makefiles
echo ��%month%��%1���̃t�@�C���쐬�����݂܂��B

REM �����ɍ쐬����t�@�C���������(\�ȍ~)�@
echo f| xcopy /d /y /q %file_name% %put_folder%\���㕪��(%month%��%1��).xlsx
exit /b


:input_month
set /p month="�����w�肵�Ă�������(1~12): "
call :num_check %month% 1 12
exit /b


:input_date_start
set /p date_start="�J�n�̓��t����͂��Ă��������B(1~31): "
call :num_check %date_start% 1 31
exit /b


:input_date_end
set /p date_end="�I���̓��t����͂��Ă��������B(%date_start%~31): "
call :num_check %date_end% %date_start% 31
exit /b


:num_check
REM ��2���� <= ��1���� <= ��3�����@���`�F�b�N
if %1 GEQ %2 (
    if %1 LEQ %3 (
        echo ����OK
    )else (
        goto :message_input_error
    )
)else goto :message_input_error
exit /b


:message_input_error
echo ���͂��ꂽ�l���s�K�؂ł��B���s�������I�����܂��B
goto :end


:end
echo;
echo Enter�L�[�������ƏI�����܂��B
set /p ending=
exit
