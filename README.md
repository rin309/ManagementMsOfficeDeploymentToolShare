# ManagementMsOfficeDeploymentToolShare
このツールは、ローカルサーバー上への Office 展開ツール (ODT) を簡単に配置できます。
通常 ODT 実行後に古いバージョンが削除されずに残りますが、このツールでは古いバージョンの削除ができます。


# Office 展開ツールとは
従来の Windows インストーラーをベースとしたインストーラーではなく、クイック実行 (Click-to-Run) でインストールされる Office のインストールに必要なツールです。
https://docs.microsoft.com/ja-jp/deployoffice/overview-office-deployment-tool

ライセンス製品以外の家庭向け Office 製品のインストーラーをお探しの場合は、下記サイトから入手してください。
https://setup.office.com/


# 事前に必要な作業
Office 展開ツールにてサポートされる構成ファイルは、事前に下記サイトで作成してください。
https://config.office.com/deploymentsettings


# 使い方
1. https://config.office.com/deploymentsettings で構成ファイルを作成します
2. https://go.microsoft.com/fwlink/p/?LinkID=626065 から ODT のインストーラーをダウンロードします
3. サーバーに C:\Shares フォルダーを作成します
4. サーバーに ManagementMsOfficeDeploymentToolShare を AllUser にインストールします *1
5. PowerShell を管理者として実行し、下記コマンドを実行します
New-MsOfficeDeploymentToolShare -ConfigPath $env:UserProfile\Downloads\Configuration.xml -LocalOfficeDeploymentToolPath $env:UserProfile\Downloads\officedeploymenttool_15128-20224.exe

*1: 後述のタスクを実行する際に SYSTEM ユーザーで実行していることから AllUser と指定していますが、NoRegisterTask スイッチを使用してタスクを登録していない場合は CurrentUser でもかまいません。


# Office\Data フォルダーを見てもダウンロードされている形跡がない・タスクが終了しない
エラーメッセージが GUI で表示されますが、タスクの実行には SYSTEM ユーザーを指定しているため表示されません。
SYSTEM ユーザーがログインせずに実行させるため指定していますが、下記の理由で変更されたい場合もあると思われます。

- 動作しないため、エラー画面を表示させたい
- プロキシなどのネットワーク環境要因によって、ユーザーを指定したい
- 動作する権限を最小限にするため、ユーザーを指定したい


