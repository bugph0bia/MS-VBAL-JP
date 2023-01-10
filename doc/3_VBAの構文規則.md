# 3 VBA プログラ厶の構文規則

VBA プログラムは、モジュール（セクション 4.2）と呼ばれるテキストファイル（またはそれに相当するテキスト単位）を用いて定義する。VBA プログラムの定義におけるモジュールの役割についてはセクション 4 で規定する。ここでは、モジュールのテキストを解釈するために使用される構文規則を説明する。

VBA モジュールの構造は、相互関連する文法の集合によって定義される。それぞれの文法は、個別に VBA モジュールのはっきりとした外観を定義する。このセットの文法は以下の通りです。

- 物理行文法
- 論理行文法
- 字句トークン文法
- 条件付きコンパイル文法
- シンタックス文法

これらの文法のうち、最初の 4 つはこのセクションで定義される。シンタックス文法はセクション 5 で定義される。

文法は ABNF [(RFC4234)](https://go.microsoft.com/fwlink/?LinkId=90462) を使って表現される。これらの文法では、数字コードは [Unicode](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/213ca0c8-6b82-4899-80a3-3c76eb534829#gt_c305d0ab-8b94-461a-bd76-13b40cb8c4d8) コードポイントとして解釈される。

## 3.1 文字エンコーディング

VBA モジュールのテキストを外部エンコーディングするために使用される実際の文字セット規格（セクション 4.2）は実装定義である。この仕様では、VBA モジュールが Unicode を使用してエンコードされているものとして、VBA モジュールの字句構造を記述している。特定の文字は、Unicode コードポイントと文字クラスの観点から識別される。Unicode と実装の特定の文字エンコーディングの間の等価マッピングは実装定義である。Unicode 以外のエンコーディングを使用する実装は、少なくとも Unicode コードポイント U+0009, U+000A, U+000D, U+0020 ～ U+007E と同等なものをサポート<ins>しなければならない</ins>。さらに，固定長の文字列は初期化時に U+0000 で埋められるので、`String` データ値の中ではこれに相当するものをサポート<ins>しなければならない</ins>。

# 3.2 モジュールの行構成

VBA モジュール（セクション 4.2）の本体は、物理行文法で記述された物理行の集合で構成される。この文法の終端記号は、Unicode コードポイントである。

### 3.2.1 物理行文法

```
module-body-physical-structure = *source-line [non-terminated-line]
source-line = *non-line-termination-character line-terminator
non-terminated-line = *non-line-termination-character
line-terminator = (%x000D  %x000A) / %x000D / %x000A / %x2028 / %x2029
non-line-termination-character = <any character other than %x000D / %x000A / %x2028 / %x2029> 
```

実装は、物理行で許可される文字数を制限<ins>してもよい</ins>。実装の制限を超える物理行を含むモジュールの意義はこの仕様では定義されていない。`<module-body-physical-structure>` が `<non-terminated-line>` で終わる場合、実装はそのモジュールを `<non-terminated-line>` の直後に `<line-terminator>` が続くかのように<ins>扱ってもよい</ins>。
