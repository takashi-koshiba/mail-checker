# mail-checker

Outlookメールの件数を表示します。

<h1>設定方法</h1>

<ol>
  <li>
    アップロードされているファイルをダウンロードします。
  <li>
    Outlookを起動してALT+F11キーを押下します。
  </li>
<li>  マクロのエディタが開かれるので、<br>
  ファイルー＞ファイルのインポートを選択してダウンロードしたファイルをすべてインポートします。 <br>
  <img src="https://github-production-user-asset-6210df.s3.amazonaws.com/173731813/352814371-fd255887-387a-4b98-ab09-12c0f927777e.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIAVCODYLSA53PQK4ZA%2F20240728%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20240728T131504Z&X-Amz-Expires=300&X-Amz-Signature=c9ca8b7360de5a6a462fdbc99bcd2fe02411f6812763ee1e3714decfdcf71b31&X-Amz-SignedHeaders=host&actor_id=173731813&key_id=0&repo_id=822262477">
</li>
  <li>
    ThisOutlookSession.clsはすでに存在しているためリネームされます。<br>
    クラスモジュール内にある「ThisOutlookSession1.cls」のコードををMicrosoft Outlook Objects内にある「ThisOutlookSession.cls」にコピーします。
  </li>
  <li>
    コードをコピーしたらクラスモジュール内にある「ThisOutlookSession1.cls」は不要のため削除します。<br>
    下記のようになればOKです。<br>
    <img src="https://github-production-user-asset-6210df.s3.amazonaws.com/173731813/352815318-1db5dc71-d914-4ccc-9d98-2354e8b72e6f.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIAVCODYLSA53PQK4ZA%2F20240728%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20240728T132549Z&X-Amz-Expires=300&X-Amz-Signature=683bc8d6a37850dd32bfa5afc4f915541c7efc2a753864fa02d983ddc4b1519a&X-Amz-SignedHeaders=host&actor_id=173731813&key_id=0&repo_id=822262477">
    

  </li>

  <li>
    ツール->参照設定を選択して以下にチェックを入れます。<br>
    ・Microsoft XML v6.0<br>
    ・Microsoft VBScript Regular Expression 5.5
  </li>
  <li>
    ツール->その他のコントロールを押下するとコントロールの追加のウィンドウが表示されます。<br>
    以下の項目にチェックを入れます。<br>
    ・Microsoft ListView Control,version 6.0<br>
    ・Microsoft Slider Control, version 6.0<br>
  </li>
  <li>
    表示->ツールボックスを選択してツールボックスウィンドウを表示します。<br>
    ListViewの参照ができないため、<br>
    ツールボックスからListViewを選択してUserForm1にドラッグ&ドロップします。<br>
    ドラッグ&ドロップしたListView は不要のためUserForm1から削除します。
    
  </li>
</ol>

