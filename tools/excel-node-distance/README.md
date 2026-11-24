# Excel｜2点間の方向別距離をCSVから自動計算する

汎用手順 Ver.1.0。Microsoft 365 / Excel 2021以降（LET・XMATCH対応）が対象。実データ・固有識別子は掲載しない。

## 1. 準備

- 既存の集計ブックを別名保存する。
- ケースごとのCSVを同じブックの CaseA_CSV、CaseB_CSV シートに取り込む。
- CSVを上書きせず、取り込み時に1行目を検索用の一意な点ID（架空例 P001、P002）に整える。元のヘッダーは別途保持する。
- A列は時刻、B列以降は各点の同一方向の変位。1行目はID、2行目は項目名、3行目は単位、4行目以降は数値。
- 重複IDは同一データか確認してから別名を付ける。欠損・空白・文字列数値は修正し、時刻は昇順で重複なしとする。
- 時刻・変位の単位は実際の解析設定で確認し、集計表と統一する。表示ラベルや数値の大きさだけで判断しない。

## 2. 集計表

| 列 | 内容 |
|---|---|
| A | 評価部位・点ペア名 |
| B | 点AのID |
| C | 点BのID |
| D | 初期時刻 |
| E | 評価時刻1 |
| F | 評価時刻2 |
| G | 初期距離（既知値） |
| H | 評価時刻1の距離 |
| I | 評価時刻2の距離 |
| J・K | 旧記入値1・2 |
| L・M | 差分1・2 |
| N・O | 一致判定1・2 |
| P | 初期座標の方向符号（点Bが正方向なら1、負方向なら-1） |

旧H・I列は関数入力前にJ・Kへ値貼り付けする。G列は既知の方向別初期距離（非負）を維持する。P列は初期座標で確認する。符号を推測しない。

## 3. 距離の定義

方向別符号付き間隔 = 初期距離 + 方向符号 × [(点Bの変位 - 点Bの初期変位) - (点Aの変位 - 点Aの初期変位)]。

負値は初期の並び順が逆転したことを示す。自動で絶対値にしない。3次元距離や接触の有無を直接表すものではない。

## 4. ケースA：H4に貼り付け

```excel
=LET(
times,CaseA_CSV!$A$4:$A$35,
vals,CaseA_CSV!$B$4:$Y$35,
nodes,CaseA_CSV!$B$1:$Y$1&"",
target,E4,
initial,$D4,
idx,XMATCH(MIN(ABS(times-target)),ABS(times-target),0),
baseIdx,XMATCH(MIN(ABS(times-initial)),ABS(times-initial),0),
nodeA,$B4&"",
nodeB,$C4&"",
colA,XMATCH(nodeA,nodes,0),
colB,XMATCH(nodeB,nodes,0),
IF(OR(ABS(INDEX(times,idx)-target)>0.01,ABS(INDEX(times,baseIdx)-initial)>0.01,SUM(--(nodes=nodeA))<>1,SUM(--(nodes=nodeB))<>1),NA(),
$G4+$P4*((INDEX(vals,idx,colB)-INDEX(vals,baseIdx,colB))-(INDEX(vals,idx,colA)-INDEX(vals,baseIdx,colA))))
)
```

I4へ右コピーし、必要な行まで下コピーする。E4だけが右コピーでF4へ変化する。

## 5. ケースB：別ブロックの例 H13

```excel
=LET(
times,CaseB_CSV!$A$4:$A$35,
vals,CaseB_CSV!$B$4:$Y$35,
nodes,CaseB_CSV!$B$1:$Y$1&"",
target,E13,
initial,$D13,
idx,XMATCH(MIN(ABS(times-target)),ABS(times-target),0),
baseIdx,XMATCH(MIN(ABS(times-initial)),ABS(times-initial),0),
nodeA,$B13&"",
nodeB,$C13&"",
colA,XMATCH(nodeA,nodes,0),
colB,XMATCH(nodeB,nodes,0),
IF(OR(ABS(INDEX(times,idx)-target)>0.01,ABS(INDEX(times,baseIdx)-initial)>0.01,SUM(--(nodes=nodeA))<>1,SUM(--(nodes=nodeB))<>1),NA(),
$G13+$P13*((INDEX(vals,idx,colB)-INDEX(vals,baseIdx,colB))-(INDEX(vals,idx,colA)-INDEX(vals,baseIdx,colA))))
)
```

I13へ右コピーし、必要な行まで下コピーする。

### 参照範囲と時刻の注意

A4:A35、B4:Y35、B1:Y1は例示範囲。実際の最終行・列に合わせ、空白行や列を含めない。近傍時刻との許容差0.01は時刻単位に応じて変更する。許容差外・ID重複・ID未検出は #N/A として扱う。等距離なら先頭の行を選ぶため、必要に応じて補間または別の評価規則を採用する。P列は必ず1または-1を入力する。

## 6. 旧値との比較（H4行の例）

```excel
L4: =ROUND(H4,1)-ROUND(J4,1)
M4: =ROUND(I4,1)-ROUND(K4,1)
N4: =IFERROR(IF(OR(NOT(ISNUMBER(J4)),NOT(ISNUMBER(H4))),"未確認",IF(ROUND(H4,1)=ROUND(J4,1),"OK","NG")),"要確認")
O4: =IFERROR(IF(OR(NOT(ISNUMBER(K4)),NOT(ISNUMBER(I4))),"未確認",IF(ROUND(I4,1)=ROUND(K4,1),"OK","NG")),"要確認")
```

各式のセル名とコロンを除いて入力し、下コピーする。距離は計算途中で丸めず、表示形式を小数1桁にする。

## 7. 動作確認（架空データ）

初期距離100、方向符号1、両点の初期変位0、評価時刻の点A変位-10、点B変位-30なら、期待値は80。評価時刻を初期時刻と同じにすると100。存在しないIDや許容差外の時刻なら #N/A。実データでは数点を手計算してから全行を確定する。

公開リポジトリには実データ、業務上の固有ID、画像、社内ファイルを追加しない。社内規則に従って使用する。
