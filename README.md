# HSBG-Capture
Hearthstone Battleground Rank/Rating Capture

## Feature
- バトルグラウンドの各試合が終了する度にあなたの最終順位をキャプチャ
- バトルグラウンドのレーティングをキャプチャ
- キャプチャされた順位及びレーティングを自動的にファイルに保存

## Setup
1. 最終順位及びレーティングを格納するディレクトリを作成し'hbc.exe'を保存
2. 'hbc.exe'ファイルを実行
3. OBSの設定 (if needed)  
(a) ソースに'テキスト (GDI+)'を追加  
  
　　<img width="600" alt="OBS画像1" src="https://github.com/takao3org/HSBG-Capture/blob/master/img/obs1.jpg">

　　(b) プロパティの'ファイルから読み取り'にチェック

　　<img width="600" alt="OBS画像2" src="https://github.com/takao3org/HSBG-Capture/blob/master/img/obs2.jpg">

　　(c) テキストファイルに1.のディレクトリの'rank.txt'か'rate.txt'を指定
  
- 'rank.txt'には最終順位の履歴が、'rate.txt'はレーティングが格納されます
- 最終順位の履歴とレーティングの両方を表示したい場合は'テキスト (GDI+)'を  
追加して下さい

## How to Use
1. 'Rank'横のボックスに最終順位の履歴が表示されます(画像の1部分)。  
　\- 'クリア'ボタンを押すと履歴がクリアされます  
　\- 値を編集して'更新'ボタンを押すと'rank.txt'がその値に上書きされます  
  
2. 'Rate'横にあるボックスにレーティングが表示されます(画像の2部分)  
　- 値を編集して'更新'ボタンを押すと'rate.txt'ファイルがその値に上書きされます  

　　<img alt="HBC画像" src="https://github.com/takao3org/HSBG-Capture/blob/master/img/window.jpg">

## License
パブリックドメイン
