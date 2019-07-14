# README #
bitbucketに公開していたものをこっちでも登録：https://bitbucket.org/dishida0120/outlook/downloads/
Outlookで指定された@ドメイン以外のアドレスが宛先に含まれていると、**大丈夫？** メッセージをPromptするマクロ

変な動きしても責任とれません。気休め程度にお使いください。

使い方：

0.     スクリプト中15行名。 "@ドメインネームをここに指定"を指定のドメインに置き換える(例：@hogehoge.com)

1.     Outlook > File > options > セキュリティセンター > macro > 全てのマクロに対して警告を表示する

2.     Outlook > File > options > ユーザのリボン設定 > 開発タブを表示

3.     開発タブ > Visual Basic > this outlook sessionを開く > スクリプトをコピー＆ペースト > outlook再起動

outlook2010以上なら動くはずです
