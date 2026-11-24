# Excel｜CSVから2点間の方向別距離を計算する（簡易版）

Ver.2.2 — セル直接参照方式。LET、XMATCH、ヘッダー加工は不要。一般的なExcelで使用可能。

## 1. 準備

1. 既存の集計Excelを別名保存する。
2. 各CSVを同じブックにシートとして追加する（例：CaseA_CSV、CaseB_CSV）。
3. CSVの元のヘッダーは変更しない。
4. 評価する点のID・変位方向・単位を確認し、対応する列と時刻の行を特定する。
5. 集計表の初期距離は既知値を維持する。

## 2. 集計表の例

| 列 | 内容 |
|---|---|
| B | 点AのID |
| C | 点BのID |
| D | 初期時刻 |
| E | 評価時刻1 |
| F | 評価時刻2 |
| G | 初期距離 |
| H | 評価時刻1の距離 |
| I | 評価時刻2の距離 |
| J・K | 旧記入値1・2 |
| L・M | 差分1・2 |

旧H・I列は上書き前にJ・Kへ値貼り付けする。

## 3. 基本の計算式

点Bが初期状態で点Aより正方向にある場合：

初期距離 ＋（点Bの評価時変位 − 点Bの初期変位）−（点Aの評価時変位 − 点Aの初期変位）

以下は架空のセル配置例。実際のCSVの行・列を確認して置き換える。

| 参照先 | 例 |
|---|---|
| 集計表の初期距離 | G4 |
| 点Aの変位列 | B列 |
| 点Bの変位列 | C列 |
| 初期時刻 | 4行目 |
| 評価時刻1 | 7行目 |
| 評価時刻2（早い場合） | 11行目 |
| 評価時刻2（遅い場合） | 12行目 |

時刻列Aを必ず確認する。3行のヘッダー、初期時刻0、一定間隔5の架空例では、4行目=0、7行目=15、8行目=20、11行目=35、12行目=40。評価時刻1が15なら8行目ではなく7行目を参照する。実際の時間単位・出力間隔は個別に確認する。

### ケースA：評価時刻1（H4）

```excel
=$G4+(CaseA_CSV!C7-CaseA_CSV!C4)-(CaseA_CSV!B7-CaseA_CSV!B4)
```

### ケースA：評価時刻2（I4）

```excel
=$G4+(CaseA_CSV!C12-CaseA_CSV!C4)-(CaseA_CSV!B12-CaseA_CSV!B4)
```

### ケースB：評価時刻1（別ブロックの例 H13）

```excel
=$G13+(CaseB_CSV!C7-CaseB_CSV!C4)-(CaseB_CSV!B7-CaseB_CSV!B4)
```

### ケースB：評価時刻2（I13）

```excel
=$G13+(CaseB_CSV!C12-CaseB_CSV!C4)-(CaseB_CSV!B12-CaseB_CSV!B4)
```

評価時刻2が早い場合は、上記の評価時刻2の式の C12 と B12 をそれぞれ C11 と B11 に変更する（初期時刻の4行目は変更しない）。

すべて1行の数式。コピーするのは数式だけ。シート名とセル番地は実際のデータに合わせる。右・下への単純コピーでは対応する点や時刻が変わるため、各セルの参照先を確認する。

### 点Bが初期状態で負方向にある場合

変位差の符号を反転する。

```excel
=$G4-(CaseA_CSV!C7-CaseA_CSV!C4)+(CaseA_CSV!B7-CaseA_CSV!B4)
```

### 初期変位が両点ともゼロの場合のみ使える短縮式

```excel
=$G4+CaseA_CSV!C7-CaseA_CSV!B7
```

原則は初期変位を補正する基本式を使用する。

## 4. Excel上での入力方法

H4を選択し、数式バーに基本式を貼り付ける。あるいは =G4+( と入力し、CSVシートに切り替えて点Bの評価時セルをクリックし、- を入力して点Bの初期セルをクリックする。続けて )-( を入力し、点Aの評価時セル、-、点Aの初期セル、) の順に入力してEnter。

シート名に空白がある場合は 'Case A CSV'!C7 のようにシングルクォートで囲む。

貼り付けで計算されない場合はセルの表示形式を「標準」にし、F2→Enter。先頭が半角の = であること、シート名が一致すること、参照セルが数値であることを確認する。

## 5. 旧値との比較

L4：
```excel
=ROUND(H4,1)-ROUND(J4,1)
```

M4：
```excel
=ROUND(I4,1)-ROUND(K4,1)
```

N4：
```excel
=IF(ROUND(H4,1)=ROUND(J4,1),"OK","NG")
```

O4：
```excel
=IF(ROUND(I4,1)=ROUND(K4,1),"OK","NG")
```

旧値・再計算値がどちらも数値である行に適用する。計算途中では丸めず、表示形式を小数1桁にする。

## 6. 検証チェック

- 初期距離は既知値を使用。
- 各評価時刻の実際のCSV時刻を確認。微小な時刻差は許容範囲を確認し、必要なら補間する。
- 点IDと変位方向が一致する列を参照。重複IDの場合は対象列を確認。
- 時刻と変位の単位を解析設定に基づいて統一。
- 初期座標の大小関係から符号を確認。
- 初期時刻を評価時刻に指定すると初期距離に一致することを確認。
- 架空例：初期距離100、点Aの変位-10、点Bの変位-30、両点の初期変位0、点Bが正方向の場合、結果は80。

この式は指定方向の符号付き間隔であり、3次元距離や接触判定ではない。負値を自動で絶対値にしない。

## 7. 一括貼り付け用（各6行×2列）

シート名は CaseA_CSV と CaseB_CSV に統一する。以下は指定のセル配置に対応する数式例。初期距離はG列、初期データは4行目、評価時刻1は7行目、評価時刻2は先頭の点ペアのみ11行目、残りは12行目。点Bが初期状態で正方向にあること、単位が一致することを確認する。負方向の場合は変位差の符号を反転する。

### Case A：H4を選択して下記6行を一括貼り付け

```text
=$G4+(CaseA_CSV!X7-CaseA_CSV!X4)-(CaseA_CSV!L7-CaseA_CSV!L4)	=$G4+(CaseA_CSV!X11-CaseA_CSV!X4)-(CaseA_CSV!L11-CaseA_CSV!L4)
=$G5+(CaseA_CSV!Q7-CaseA_CSV!Q4)-(CaseA_CSV!J7-CaseA_CSV!J4)	=$G5+(CaseA_CSV!Q12-CaseA_CSV!Q4)-(CaseA_CSV!J12-CaseA_CSV!J4)
=$G6+(CaseA_CSV!S7-CaseA_CSV!S4)-(CaseA_CSV!H7-CaseA_CSV!H4)	=$G6+(CaseA_CSV!S12-CaseA_CSV!S4)-(CaseA_CSV!H12-CaseA_CSV!H4)
=$G7+(CaseA_CSV!T7-CaseA_CSV!T4)-(CaseA_CSV!F7-CaseA_CSV!F4)	=$G7+(CaseA_CSV!T12-CaseA_CSV!T4)-(CaseA_CSV!F12-CaseA_CSV!F4)
=$G8+(CaseA_CSV!U7-CaseA_CSV!U4)-(CaseA_CSV!D7-CaseA_CSV!D4)	=$G8+(CaseA_CSV!U12-CaseA_CSV!U4)-(CaseA_CSV!D12-CaseA_CSV!D4)
=$G9+(CaseA_CSV!V7-CaseA_CSV!V4)-(CaseA_CSV!B7-CaseA_CSV!B4)	=$G9+(CaseA_CSV!V12-CaseA_CSV!V4)-(CaseA_CSV!B12-CaseA_CSV!B4)
```

### Case B：H13を選択して下記6行を一括貼り付け

```text
=$G13+(CaseB_CSV!N7-CaseB_CSV!N4)-(CaseB_CSV!L7-CaseB_CSV!L4)	=$G13+(CaseB_CSV!N11-CaseB_CSV!N4)-(CaseB_CSV!L11-CaseB_CSV!L4)
=$G14+(CaseB_CSV!P7-CaseB_CSV!P4)-(CaseB_CSV!J7-CaseB_CSV!J4)	=$G14+(CaseB_CSV!P12-CaseB_CSV!P4)-(CaseB_CSV!J12-CaseB_CSV!J4)
=$G15+(CaseB_CSV!O7-CaseB_CSV!O4)-(CaseB_CSV!H7-CaseB_CSV!H4)	=$G15+(CaseB_CSV!O12-CaseB_CSV!O4)-(CaseB_CSV!H12-CaseB_CSV!H4)
=$G16+(CaseB_CSV!Y7-CaseB_CSV!Y4)-(CaseB_CSV!F7-CaseB_CSV!F4)	=$G16+(CaseB_CSV!Y12-CaseB_CSV!Y4)-(CaseB_CSV!F12-CaseB_CSV!F4)
=$G17+(CaseB_CSV!R7-CaseB_CSV!R4)-(CaseB_CSV!D7-CaseB_CSV!D4)	=$G17+(CaseB_CSV!R12-CaseB_CSV!R4)-(CaseB_CSV!D12-CaseB_CSV!D4)
=$G18+(CaseB_CSV!W7-CaseB_CSV!W4)-(CaseB_CSV!B7-CaseB_CSV!B4)	=$G18+(CaseB_CSV!W12-CaseB_CSV!W4)-(CaseB_CSV!B12-CaseB_CSV!B4)
```

各ブロックはタブ区切り。Excelの数式バーではなくセルを選択して貼り付ける。H・I列の旧値を保存してから実行する。貼り付け後に各列の点IDと評価時刻を原データで確認する。実際のシート配置が異なる場合は参照先を変更する。

公開リポジトリには実データ、固有ID、画像、社内ファイルを掲載しない。社内規則に従って使用する。
