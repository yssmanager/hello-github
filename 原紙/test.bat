@echo off

REM ここにコピー元のファイル名を入力
set file_name=売上分析(月日).xlsx

REM パスワードの設定
set PASSWORD=sp

REM パスワード認証
set /p IN_PASS="パスワードを入力してください: "
if %PASSWORD%==%IN_PASS% (
    echo 認証OK
    call :main
    echo;
    echo 実行は正常終了しました。
) else (
    echo パスワードが違います。実行を終了します。
)
goto :end
exit


REM ---サブルーチン---

:main
REM コピー先のフォルダー設定(一つ前のディレクトリー)
pushd %~dp0..
set put_folder=%CD%
echo %put_folder%
popd

REM 月日の値入力
call :input_month
call :input_date_start
call :input_date_end

echo %month%月%date_start%日から%month%月%date_end%日までのファイルを作成します。

for /l %%i in (%date_start%,1,%date_end%) do (
    call :makefiles %%i
)
exit /b


:makefiles
echo ＞%month%月%1日のファイル作成を試みます。

REM ここに作成するファイル名を入力(\以降)　
echo f| xcopy /d /y /q %file_name% %put_folder%\売上分析(%month%月%1日).xlsx
exit /b


:input_month
set /p month="月を指定してください(1~12): "
call :num_check %month% 1 12
exit /b


:input_date_start
set /p date_start="開始の日付を入力してください。(1~31): "
call :num_check %date_start% 1 31
exit /b


:input_date_end
set /p date_end="終了の日付を入力してください。(%date_start%~31): "
call :num_check %date_end% %date_start% 31
exit /b


:num_check
REM 第2引数 <= 第1引数 <= 第3引数　かチェック
if %1 GEQ %2 (
    if %1 LEQ %3 (
        echo 入力OK
    )else (
        goto :message_input_error
    )
)else goto :message_input_error
exit /b


:message_input_error
echo 入力された値が不適切です。実行を強制終了します。
goto :end


:end
echo;
echo Enterキーを押すと終了します。
set /p ending=
exit
