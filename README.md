ご覧いただきありがとうございます。<br>
このプログラムはExcelマクロ入門の授業でラクしようと思いよく使いそうな関数（？）を自動入力しようと思い開発しました。<br>
一応備忘録的なものを残そうと思います。<br>

<h1>共通事項</h1>
    [pyautogui]()が必要です。
    各ボタンを押すとディスプレイx=48 y=0　にクリック信号を入れVBAのウィンドウをアクティブにします。<br>
    その位置に他のウィンドウが存在すると予期せぬトラブルになります。<br>
    *TabキーやReturnキーは使用禁止です。*<br>
    関数間のインデントは各項目入力後のReturnキー信号で自動的に入力されます。<br>
    テキストボックスが空白だとエラーになります。（一部例外除く）<br>
    VBAが受け付ける範囲で日本語の利用が可能です。<br>

## Rangeボタン
    このボタンを押すとRangeから始まる関数を自動で入力します。<br>
    実行結果は　Range("xxx")yyy　の形になります。<br>
        xxxはメインのテキストボックスに入力された文字列（主にセル）が入ります。<br>
            A1のような単独セルや「A1,A2」「A1:D4」のように複数選択も可能です。<br>
            メインのテキストボックスが空白だとエラーになります。<br>
        yyyはオプションのテキストボックス内の文字列をコピペします。<br>
            「=100」や「.font~」といったxxxに続くものを入力してください。<br>
            オプションが空白でもエラーにはなりません。<br>
### AutoFillボタン
        このボタンはオートフィルを実行する関数を入力します。<br>
        実行結果は　Range("xxx").Autofill Destination:=Range("yyy")　となります。<br>
        NANDの書き方が理解できていないためオプションが空白でもエラーになりませんが、どのような挙動になるかは未検証です。<br>
#### Subボタン
    VBAで最初に入力するSub～を入力します。<br>
        実行結果はSub xxx ()となります。<br>
        xxxはメインのテキストボックスに入力された値となります。<br>
    オプションのテキストボックスが空白でもエラーとはなりません。<br>
    最後のEnd Subも自動で入力されます。<br>
##### チェックボックス
    SubボタンでSubとEnd Subの間に「Cells.delete」と入力するか否かです。<br>
    チェックがついていると入力されます。（初期値）<br>
###### 削除ボタン
    対応しているテキストボックスの内容を削除します。<br>
    削除後、対応しているテキストボックスが入力待機状態となります。


　
