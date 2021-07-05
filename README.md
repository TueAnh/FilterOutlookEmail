# FilterOutlookEmail
Outlookメール集計ツール
# Feature(機能)
1. 日付を範囲でメールの選択
2. キーワードの定義
3. 指定されたフォルダにメールの移動
# System Requirement(システム要件)
1. Windows PowerShell
2. Microsoft Outlook 
# User Guide (操作方法)
まず、アウトルックが起動していることを確認してください。

0. スクリプト実行方  
  .ps1ファイルには右クリックし、メニューで「PowerShell で実行」クリックします。
1. キーワード定義ファイルを選択します  
  ➡形式 ： CSVファイル(.csv)  
  ➡必要なコラム : 「メール件名」と「差出人」  
「keyword_sample.csv」参考してください。  
![image](https://user-images.githubusercontent.com/32601267/124418093-de4de300-dd84-11eb-9089-57cc847b0608.png)  
  ➡➡空のセルというのは「すべて」選択されます
2. Outlookソースフォルダのパスを入力します(メール選択・フィルターするフォルダ)  
  例： \\\xxx@yyy\Inbox, \\\xxx@yyy\Inbox\Sample  
  ➡Outlookフォルダパスの決定方  
  ➡➡Outlookアプリでフォルダに右クリック----->プロパティ---->パス＝＜場所＞￥＜フォルダ名＞  
  例： 「Inbox」フォルダのパス：＜\\\xxx@yyy＞￥＜Inbox＞＝\\\xxx@yyy\Inbox  
3. Outlookの保存先フォルダのパスを入力します (移動されたメール保存先フォルダ)  
  ➡注意点：移動したくない場合、同じソースパス入力してください  
  例：  
     ソースパス：\\\xxx@yyy\Inbox  
     保存先パス：\\\xxx@yyy\Inbox  
  ➡移動しません。  
4. 日付を範囲で選択します  
  ➡形式： mm/dd/yyyy  
  例： 07/04/2021　(7月4日2021年)  
5. 日付の範囲によって自動的にメール選択した後で、データ抽出ファイルの保存先を選択します  
  ➡注意点： ファイル保存ダイアログが表示されます。  
  ➡形式： CSVファイル(.csv)  
6. キーワードに基づいてデータ抽出するのは時間がかかるのでお待ちください。データ抽出ながらメール移動します。  
7. 最後、集計結果の保存先を選択します。結果一覧表もPowerShellで表示されます。  
  ➡注意点： ファイル保存ダイアログが表示されます。  
  ➡形式： CSVファイル(.csv)  
  ![image](https://user-images.githubusercontent.com/32601267/124418330-6c29ce00-dd85-11eb-920b-8d291948ed68.png)


