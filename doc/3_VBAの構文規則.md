# 3 VBA プログラ厶の構文規則

VBA プログラムは、モジュール（セクション 4.2）と呼ばれるテキストファイル（またはそれに相当するテキスト単位）を用いて定義する。VBA プログラムの定義におけるモジュールの役割についてはセクション 4 で規定する。ここでは、モジュールのテキストを解釈するために使用される構文規則を説明する。

VBA モジュールの構造は、相互関連する文法の集合によって定義される。それぞれの文法は、個別に VBA モジュールのはっきりとした外観を定義する。このセットの文法は以下の通りです。

- 物理行文法
- 論理行文法
- 字句トークン文法
- 条件付きコンパイル文法
- シンタックス文法

これらの文法のうち、最初の 4 つはこのセクションで定義される。シンタックス文法はセクション 5 で定義される。

文法は ABNF [(RFC4234)](https://go.microsoft.com/fwlink/?LinkId=90462) を使って表現される。これらの文法では、数字コードは [Unicode](./1_はじめに.md) コードポイントとして解釈される。

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

VBA のプログラムテキストとして解釈するために、モジュール（セクション 4.2）は、それぞれが複数の物理行に対応することができる論理行の集合とみなす。この構造は論理行文法で記述される。この文法の終端記号は [Unicode](./1_はじめに.md) 文字コードポイントである。

### 3.2.2 論理行文法

```
module-body-logical-structure = *extended-line
extended-line = *(line-continuation / non-line-termination-character)  line-terminator
line-continuation = *WSC underscore *WSC line-terminator
WSC = (tab-character / eom-character /space-character / DBCS-whitespace / most-Unicode-class-Zs)
tab-character = %x0009
eom-character = %x0019
space-character = %x0020
underscore = %x005F
DBCS-whitespace = %x3000
most-Unicode-class-Zs = <all members of Unicode class Zs which are not CP2-characters>
```

実装は `<extended-line>` の文字数を制限<ins>してもよい</ins>。

仕様を簡単にするために、論理行の開始直前の位置と論理行の最後の `<line-terminator>` 直前の位置を明示的に参照できるようにすると便利である。これは、VBA の文法の終端記号である `<LINE-START>` と `<LINE-END>` を使うことで実現されている。`<LINE-START>` は各論理行の直前、`<LINE-END>` は各論理行の最後の `<line-terminator>` を置き換えるものとして定義される。

```
module-body-lines = *logical-line
logical-line = LINE-START *extended-line LINE-END
```

ABNF ルール定義で使用する場合、`<LINE-START>` と `<LINE-END>` は `<logical-line>` の開始または終了を示すために使用される。

### 3.3 字句トークン

VBA プログラムの構文は、個々の [Unicode](./1_はじめに.md) 文字ではなく、字句トークンの観点から最も簡単に記述することができる。特にほとんどの構文要素間の空白や行の連続は、通常、構文文法とは無関係である。このような空白の出現の可能性を記述する必要がなければ、構文文法は著しく単純化される。これは、空白を抽象化した字句トークン（単にトークンともいう）を構文文法の終端記号として用いることで達成される。

字句解析文法は `<module-body-lines>` をこのような字句トークンの集合として解釈することを定義している。

字句解析文法の終端要素は、Unicode 文字と `<LINE-START>` 要素および `<LINE-END>` 要素である。一般に、すべて大文字で書かれた字句解析文法の規則名は、VBA 構文文法の字句解析トークンおよび終端要素でもある。ABNF 引用リテラルテキスト規則も構文文法の字句トークンであるとみなされる。字句解析トークンには、その直前にある空白文字が含まれる。字句解析文法内で使用される場合、引用されたリテラルテキスト規則はトークンとして扱われないので、先行する空白文字は重要であることに注意すること。

### 3.3.1 セパレータと特殊トークン

```
WS = 1*(WSC / line-continuation) 
special-token = "," / "." / "!" /  "#" / "&" / "(" / ")" / "*" / "+" / "-" / "/" / ":" / ";" / "<" / "=" / ">" / "?" / "\" / "^"
NO-WS = <no whitespace characters allowed here>
NO-LINE-CONTINUATION = <a line-continuation is not allowed here>
EOL = [WS] LINE-END / single-quote comment-body
EOS = *(EOL  /  ":")  ;End Of Statement
single-quote = %x0027  ; '
comment-body = *(line-continuation / non-line-termination-character) LINE-END
```

`<special-token>` は、VBA プログラムの構文において特別な意味を持つ単一文字を識別するために使用される。これらは字句トークン（セクション 3.3）であるため、この文字の前に空白文字を置くことができるが無視される。クォーテーション文字で囲まれた `<special-token>` 要素のいずれかが構文文法の要素として出現する場合、対応するトークン（セクション 3.3）への参照である。

`<NO-WS>` は構文文法の終端要素として用いられ、その直後のトークンの前に空白文字が<ins>あってはならない</ins>ことを示す。`<NO-LINE-CONTINUATION>` は構文文法の終端要素として用いられ、直後のトークンの前に `<linecontinuation>` シーケンスを含む空白文字が<ins>あってはならない</ins>ことを示す。

`<WS>` は構文文法の終端要素として用いられ、その直後のトークンの前に 1 つ以上の空白文字が<ins>なければならない</ins>ことを示す。

`<EOL>` は構文文法の要素として用いられ、論理行の唯一または最後でなければならないステートメントの中の「ステートメントの終端」マーカーとして機能するトークンに指定するために使用される。

`<EOS>` は構文文法の終端要素として用いられ、「ステートメントの終端」マーカーとして機能するトークンを名付けたものである。一般に、ステートメントの終端は `<LINE-END>` かコロン文字でマークされる。`<single-quote>` と `<LINE-END>` の間の文字はコメントテキストとして無視される。

### 3.3.2 数値トークン

```
INTEGER = integer-literal ["%" / "&" / "^"]
integer-literal = decimal-literal / octal-literal / hex-literal
decimal-literal = 1*decimal-digit
octal-literal = "&" [%x004F / %x006F] 1*octal-digit    ; & or &o or &O
hex-literal = "&" (%x0048 / %x0068) 1*hex-digit   ; &h or &H
octal-digit = "0" / "1" / "2" / "3" / "4" / "5" / "6" / "7"
decimal-digit = octal-digit / "8" / "9"
hex-digit = decimal-digit / %x0041-0046 / %x0061-0066 ;A-F / a-f
```

静的セマンティクス

- `<decimal-digit>`, `<octal-digit>`, `<hex-digit>` シーケンスは、それぞれ 10 進数、8 進数、16 進数で表される符号なし整数値として解釈される。
- 各 `<INTEGER>` には定数データ値（セクション 2.1）が関連している。定数のデータ値、データ型（セクション 2.1）、宣言型（セクション 2.2）は次の表で定義される（「有効性」欄が No の場合、 `<INTEGER>` は無効）。

| 基数 | 範囲内の正の `<INTEGER>` | 型サフィックス | `<INTEGER>` の有効性 | 宣言型 | データ型 | 符号付きデータ値 |
| ---- | ---- | ---- | ---- | ---- | ---- | ---- |
| 10進数 | 0 ≤ n ≤ 32767 | なし | Yes | `Integer` | `Integer` | n |
| 10進数 | 0 ≤ n ≤ 32767 | "%" | Yes | `Integer` | `Integer` | n |
| 10進数 | 0 ≤ n ≤ 32767 | "&" | Yes | `Long` | `Integer` | n |
| 10進数 | 0 ≤ n ≤ 32767 | "^" | Yes | `LongLong` | `Integer` | n |
| 8進数 | 0 ≤ n ≤ &o77777 | なし | Yes | `Integer` | `Integer` | n |
| 8進数 | 0 ≤ n ≤ &o77777 | "%" | Yes | `Integer` | `Integer` | n |
| 8進数 | 0 ≤ n ≤ &o77777 | "&" | Yes | `Long` | `Integer` | n |
| 8進数 | 0 ≤ n ≤ &o77777 | "^" | Yes | `LongLong` | `Integer` | n |
| 8進数 | &o100000 ≤ n ≤ &o177777 | なし | Yes | `Integer` | `Integer` | n – 65,536 |
| 8進数 | &o100000 ≤ n ≤ &o177777 | "%" | Yes | `Integer` | `Integer` | n – 65,536 |
| 8進数 | &o100000 ≤ n ≤ &o177777 | "&" | Yes | `Long` | `Integer` | n |
| 8進数 | &o100000 ≤ n ≤ &o177777 | "^" | Yes | `LongLong` | `Integer` | n |
| 16進数 | 0 ≤ n ≤ &H7FFF | なし | Yes | `Integer` | `Integer` | n |
| 16進数 | 0 ≤ n ≤ &H7FFF | "%" | Yes | `Integer` | `Integer` | n |
| 16進数 | 0 ≤ n ≤ &H7FFF | "&" | Yes | `Long` | `Integer` | n |
| 16進数 | 0 ≤ n ≤ &H7FFF | "^" | Yes | `LongLong` | `Integer` | n |
| 16進数 | &H8000 ≤ n ≤ &HFFFF | なし | Yes | `Integer` | `Integer` | n – 65,536 |
| 16進数 | &H8000 ≤ n ≤ &HFFFF | "%" | Yes | `Integer` | `Integer` | n – 65,536 |
| 16進数 | &H8000 ≤ n ≤ &HFFFF | "&" | Yes | `Long` | `Integer` | n |
| 16進数 | &H8000 ≤ n ≤ &HFFFF | "^" | Yes | `LongLong` | `Integer` | n |
| 10進数 | 32768 ≤ n ≤ 2147483647 | なし | Yes | `Long` | `Long` | n |
| 10進数 | n ≥ 32768 | "%" | No |  |  |  |
| 10進数 | 32768 ≤ n ≤ 2147483647 | "&" | Yes | `Long` | `Long` | n |
| 10進数 | 32768 ≤ n ≤ 2147483647 | "^" | Yes | `LongLong` | `Long` | n |
| 10進数 | n ≥ 2147483647 | なし | (注1 参照) | `Double` | `Double` | n# (注1 参照) |
| 10進数 | n ≥ 2147483647 | "&" | No |  |  |  |
| 8進数 | &o200000 ≤ n ≤ &o17777777777 | なし | Yes | `Long` | `Long` | n |
| 8進数 | &o200000 ≤ n ≤ &o17777777777 | "%" | No |  |  |  |
| 8進数 | &o200000 ≤ n ≤ &o17777777777 | "&" | Yes | `Long` | `Long` | n |
| 8進数 | &o200000 ≤ n ≤ &o17777777777 | "^" | Yes | `LongLong` | `Long` | n |
| 8進数 | &o20000000000 ≤ n ≤ &o37777777777 | なし | Yes | `Long` | `Long` | n – 4,294,967,296 |
| 8進数 | &o20000000000 ≤ n ≤ &o37777777777 | "%" | No |  |  |  |
| 8進数 | &o20000000000 ≤ n ≤ &o37777777777 | "&" | Yes | `Long` | `Long` | n – 4,294,967,296 |
| 8進数 | &o20000000000 ≤ n ≤ &o37777777777 | "^" | Yes | `LongLong` | `Long` | n |
| 8進数 | n ≥ &o40000000000 | なし | No |  |  |  |
| 8進数 | n ≥ &o40000000000 | "%" | No |  |  |  |
| 8進数 | n ≥ &o40000000000 | "&" | No |  |  |  |
| 16進数 | &H8000 ≤ n ≤ &H7FFFFFFF | なし | Yes | `Long` | `Long` | n |
| 16進数 | &H8000 ≤ n ≤ &H7FFFFFFF | "%" | No |  |  |  |
| 16進数 | &H8000 ≤ n ≤ &H7FFFFFFF | "&" | Yes | `Long` | `Long` | n |
| 16進数 | &H8000 ≤ n ≤ &H7FFFFFFF | "^" | Yes | `LongLong` | `Long` | n |
| 16進数 | &H80000000 ≤ n ≤ &H7FFFFFFFF | なし | Yes | `Long` | `Long` | n – 4,294,967,296 |
| 16進数 | &H80000000 ≤ n ≤ &H7FFFFFFFF | "%" | No |  |  |  |
| 16進数 | &H80000000 ≤ n ≤ &H7FFFFFFFF | "&" | Yes | `Long` | `Long` | n – 4,294,967,296 |
| 16進数 | &H80000000 ≤ n ≤ &H7FFFFFFFF | "^" | Yes | `LongLong` | `Long` | n |
| 16進数 | n ≥ &H100000000 | なし | No |  |  |  |
| 16進数 | n ≥ &H100000000 | "%" | No |  |  |  |
| 16進数 | n ≥ &H100000000 | "&" | No |  |  |  |
| 10進数 | 2147483648 ≤ n ≤ 9223372036854775807 | "^" | Yes | `LongLong` | `LongLong` | n |
| 10進数 | n ≥ 9223372036854775808 | "^" |  |  |  |  |
| 8進数 | &o40000000000 ≤ n ≤ &o1777777777777777777777 | "^" | Yes | `LongLong` | `LongLong` | n - 232 |
| 8進数 | n ≥ &o2000000000000000000000 | 任意 | No |  |  |  |
| 16進数 | &H100000000 ≤ n ≤ &HFFFFFFFFFFFFFFFF | "^" | Yes | `LongLong` | `LongLong` | n - 232 |
| 16進数 | n ≥ &H10000000000000000 | 任意 | No |  |  |  |

（訳注："注1" とあるが参照先が見つかっていない）

- 64 ビット演算をサポートしない実装において `LongLong` 型に宣言されたリテラルは、静的に無効である。

```
FLOAT = (floating-point-literal [floating-point-type-suffix] ) / (decimal-literal floating-point-type-suffix)
floating-point-literal = (integer-digits exponent) / (integer-digits "." [fractional-digits] [exponent]) / ( "." fractional-digits [exponent])

integer-digits = decimal-literal
fractional-digits = decimal-literal
exponent = exponent-letter  [sign] decimal-literal
exponent-letter = %x0044 / %x0045 / %x0064 / %x0065   ; D / E / d / e
sign = "+" / "-"
floating-point-type-suffix = "!" / "#" / "@"
```

静的セマンティクス

- `<FLOAT>` トークンは、バイナリ浮動小数点または通貨データ値を表す。 `<floatingpoint-type-suffix>` は、下の表に従ってトークンに関連付けられたデータ値の宣言型とデータ型を指定する。
    - iを `<integer-digits>` の整数値、f を `<fractional-digits>` の整数値、d を `<fractional-digits>` の桁数、x を `<exponent>` の符号付き整数値とする。そして、`<floating-point-literal>` は次の式従って数学的実数である r を表す。
        - $r = (i + f 10^-d) 10^x$
    - `<floating-point-literal>` は、その数学的な値が宣言型を使って表現できる最大値より大きい場合は無効となる。

| `<floating-point-type-suffix>` | 宣言型とデータ型 |
| ---- | ---- |
| なし | `Double` |
| ! | `Single` |
| # | `Double` |
| @ | `Currency` |

（訳注：Markdown都合で表の位置を移動した）

- `<floating-point-literal>` の宣言型が `Currency` の場合、r の小数部は偶数丸め（セクション 5.5.1.2.1.1）により有効数字 4 桁で丸められる。

### 3.3.3 日付トークン

```
date-or-time = (date-value 1*WSC time-value) / date-value / time-value

date-value = left-date-value date-separator  middle-date-value [date-separator right-date-value]
left-date-value = decimal-literal / month-name
middle-date-value = decimal-literal / month-name
right-date-value = decimal-literal / month-name
date-separator = 1*WSC / (*WSC ("/" / "-" / ",") *WSC)

month-name = English-month-name / English-month-abbreviation
English-month-name = "january" / "february" / "march" / "april" / "may" / "june" / "august" / "september" / "october" / "november" / "december"
English-month-abbreviation = "jan" / "feb" / "mar" / "apr" / "jun" / "jul" / "aug" / "sep" /  "oct" / "nov" / "dec"

time-value = (hour-value ampm) / (hour-value time-separator minute-value [time-separator second-value] [ampm])
hour-value = decimal-literal
minute-value = decimal-literal
second-value = decimal-literal
time-separator = *WSC (":" / ".") *WSC
ampm = *WSC ("am" / "pm" / "a" / "p")
```

静的セマンティクス

- `<DATE>` トークン（セクション 3.3）は、データ型（セクション 2.1）及び宣言型（セクション 2.2）`Date` のデータ値（セクション 2.1）を持つ。
- `<DATE>` トークンのデータ値の数値は、指定された日付と指定された時刻の和となる。
- `<date-or-time>` が `<time-value>` を含まない場合、その指定時刻は "00:00:00" からなる `<time-value>` が存在するものとして決定される。
- `<date-or-time>` に `<date-value>` が含まれない場合、"1899/12/30"という文字からなる `<date-value>` が存在するものとして日付が決定される。
- `<left-date-value>`, `<middle-date-value>`,  `<right-date-value>` のうち 1 つは `<month-name>` となり得る。
- 「3.3.3.1 日付トークンの解釈方法」に示す内容で値を決定する。（訳注：数式等が混在する関係上、記載内容を別の章に移した）

### 3.3.3.1 日付トークンの解釈方法
$L$ が `<left-date-value>`、 $M$ が `<middle-date-value>`、 $R$ が `<right-date-value>` のデータ値として与えられているとすると、 $L, M, R$ は次のようにカレンダーの日付として解釈される。

まず、下記の通り式と定数を定義する。

$$
LegalMonth(x) = \begin{cases}
  true & 0 \le x \le 12 \\
  false & \text{上記以外}
\end{cases}
$$

$$
LegalDay(month, day, year) = \begin{cases}
  false & \begin{cases}
    \textrm{year < 0 または year > 32767 または} \\
    \textrm{LegalMonth(month) が false または} \\
    \textrm{day が指定された年月において有効ではない}
  \end{cases} \\
  true & \text{上記以外}
\end{cases}
$$

- $CY$ を実装定義のデフォルトの年とする。

$$
Year(x) = \begin{cases}
  x + 2000 & 0 \le x \le 29 \\
  x + 1900 & 30 \le x \le 99 \\
  x & \text{上記以外}
\end{cases}
$$

次に、以下の通り解釈する。

- $L$ と $M$ が数値で $R$ が存在しない場合、
    - もし $LegalMonth(L)$ および $LegalDay(L, M, CY)$ の場合、月は $L$ 、日は $M$ 、年は $CY$ である。
    - それ以外で、もし $LegalMonth(M)$ および $LegalDay(M, L, CY)$ の場合、月は $M$ 、日は $L$ 、年は $CY$ である。
    - それ以外で、もし $LegalMonth(L)$ の場合、月は $L$ 、日は $1$ 、年は $M$ である。
    - それ以外で、もし $LegalMonth(M)$ の場合、月は $M$ 、日は $1$ 、年は $L$ である。
    - それ以外の場合、`<date-value>` は有効ではない。
- $L, M, R$ が数値の場合、
    - $LegalMonth(L)$ および $LegalDay(L, M, Year(R))$ の場合、月は $L$ 、日は $M$ 、年は $Year(R)$ である。
    - それ以外で、もし $LegalMonth(M)$ および $LegalDay(M, R, Year(L))$ の場合、月は $M$ 、日は $R$ 、年は $Year(L)$ である。
    - それ以外で、もし $LegalMonth(M)$ および $LegalDay(M, L, Year(R))$ の場合、月は $M$ 、日は $L$ 、年は $Year(R)$ である。
    - それ以外の場合、`<date-value>` は有効ではない。
- $L, M$ のいずれかが数値ではなく、かつ $R$ が存在しない場合、
    - 次の通りとする。
        - $N$ を $L$ と $M$ いずれかの数値の方とする。
        - $L$ と $M$ のうち数値ではない値の月名または略号に対応する 1～12 の範囲の値を $M$ とする。
    - $LegalDay(M, N, CY)$ ならば、月は $M$ 、日は $N$ 、年は $CY$ である。
    - それ以外の場合、月は $M$ 、日は $1$ 、年は $Year(N)$ である。
- それ以外（ $R$ が存在し、 $L, M, R$ のいずれかが数値ではない）の場合、
    - 次の通りとする。
        - $L, M, R$ のうち数値でない値の月名または略号に対応する 1～12 の範囲の値を $M$ とする。
        - $L, M, R$ のうち数値である値を $N1, N2$ とする。（訳注：原文では両方 $N1$ とされており誤記と思われるので修正した）
    - もし $LegalDay(M, N1, Year(N2))$ の場合、月は $M$ 、日は $N1$ 、年は $Year(N2)$ である。
    - それ以外で、もし $LegalDay(M, N2, Year(N1))$ の場合、月は $M$ 、日は $N2$ 、年は $Year(N1)$ である。
    - それ以外の場合、`<date-value>` は有効ではない。

- `<hour-value>` の要素である `<decimal-literal>` は 0 から 23 の範囲の整数で<ins>なければならない</ins>。
- `<minute-value>` の要素である `<decimal-literal>` は  0 から 59 の範囲の整数で<ins>なければならない</ins>。
- `<second-value>` の要素である `<decimal-literal>` は 0 から 59 の範囲の整数で<ins>なければならない</ins>。

- `<time-value>` が "pm" または "p" からなる `<ampm>` 要素を含み、`<hour-value>` が 0 から 11 の範囲の整数値を持つ場合、`<hour-value>` は実際の値より 12 大きい整数値として使用される。
- `<hour-value>` が 12 より大きい場合、`<ampm>` 要素は無視される。
- `<time-value>` が "am" または "a" からなる `<ampm>` 要素を含み、`<hour-value>` が整数値 12 の場合、`<hour-value>` は 0 として使用される。
- `<time-value>` に `<minute-value>` が含まれていない場合、整数値 0 の `<minute-value>` が存在するものとして扱われる。
- `<time-value>` に `<second-value>` が含まれていない場合、整数値 0 の `<second-value>` が存在するものとして扱われる。
- `<time-value>` の `<hour-value> `要素の整数値を $h$ 、<time-value> の `<minute-value>` 要素の整数値を $m$ 、`<time-value>` の `<second-value>` の整数値を $s$ とすると、`<time-value>` の指定時刻は、 $(3600h+60m+s)/86400$ の式で定義される。

### 3.3.4 文字列トークン

```
STRING = double-quote *string-character (double-quote /  line-continuation / LINE-END)
double-quote = %x0022  ; "
string-character = NO-LINE-CONTINUATION ((double-quote double-quote)  /  non-line-termination-character)
```

静的セマンティクス

- `<STRING>` トークン（セクション 3.3）は、データ型（セクション 2.1）と宣言型（セクション 2.2）が `String` のデータ値（セクション 2.1）を関連付ける。
- 関連付けられた文字列データ値の長さは、`<STRING>` トークンを構成する `<string-character>` 要素の数であり `<STRING>` トークンの長さではない。
- データ値は、`<string-character>` 要素に対応する実装定義の符号化文字列からなり、左から右の順に、最も左の `<string-character>` 要素がその先頭要素を、最も右の `<string-character>` 要素がその最終文字を規定する。
- `<string-character>` 要素のいずれかが実装定義の文字セットでエンコードされていない場合、`<STRING>` トークンは無効である。
- 2 つの `<double-quote>` 文字のシーケンスはデータ値内で文字 U+0022 が 1 つだけ出現することを表す。
- `<string-character>` 要素が存在しない場合、データ値は長さ 0 の空文字列となる。
- `<STRING>` が `<line-continuation>` 要素で終わっている場合、データ値の最終文字は `<line-continuation>` の前の `<WSC>` でない右端の文字となる。
- `<STRING>` が `<LINE-END>` 要素で終わる場合、関連するデータ値の最終文字は `<LINE-END>` の前の `<line-terminator>` でない右端の文字となる。

### 3.3.5 識別子トークン

```
lex-identifier = Latin-identifier / codepage-identifier / Japanese-identifier / Korean-identifier / simplified-Chinese-identifier / traditional-Chinese-identifier

Latin-identifier = first-Latin-identifier-character *subsequent-Latin-identifier-character
first-Latin-identifier-character = (%x0041-005A / %x0061-007A) ; A-Z / a-z
subsequent-Latin-identifier-character = first-Latin-identifier-character / decimal-digit / %x5F    ; underscore
```

静的セマンティクス

- アルファベットの大文字と小文字は、VBA の識別子では同等とみなされる。対応する `<first-Latin-identifier-character>` 文字の大文字/小文字のみが異なる二つの識別子は、同一であるとみなされる。
- 実装は、`<Latin-identifier>` をサポート<ins>しなければならない</ins>。実装は他の識別子形式を 1 つ以上サポートする場合があり、その場合は識別子形式の併用を<ins>制限してもよい</ins>。

#### 3.3.5.1 非アルファベット識別子

```
Japanese-identifier = first-Japanese-identifier-character *subsequent-Japanese-identifier-character
first-Japanese-identifier-character = (first-Latin-identifier-character / CP932-initial-character)
subsequent-Japanese-identifier-character = (subsequent-Latin-identifier-character / CP932-subsequent-character)
CP932-initial-character = < character ranges specified in section 3.3.5.1.1>
CP932-subsequent-character = < character ranges specified in section 3.3.5.1.1>

Korean-identifier = first-Korean-identifier-character *subsequent Korean-identifier-character
first-Korean-identifier-character = (first-Latin-identifier-character / CP949-initial-character)
subsequent-Korean-identifier-character = (subsequent-Latin-identifier-character / CP949-subsequent-character)
CP949-initial-character = < character ranges specified in section 3.3.5.1.2>
CP949-subsequent-character = < character ranges specified in section 3.3.5.1.2>

simplified-Chinese-identifier = first-sChinese-identifier-character *subsequent-sChinese-identifier-character 
first-sChinese-identifier-character = (first-Latin-identifier-character / CP936-initial-character)
subsequent-sChinese-identifier-character = (subsequent-Latin-identifier-character / CP936-subsequent-character)
CP936-initial-character = < character ranges specified in section 3.3.5.1.3>
CP936-subsequent-character = < character ranges specified in section 3.3.5.1.3>

traditional-Chinese-identifier = first-tChinese-identifier-character *subsequent-tChinese-identifier-character
first-tChinese-identifier-character = (first-Latin-identifier-character / CP950-initial-character)
subsequent-tChinese-identifier-character = (subsequent-Latin-identifier-character / CP950-subsequent-character)
CP950-initial-character = < character ranges specified in section 3.3.5.1.4>
CP950-subsequent-character = < character ranges specified in section 3.3.5.1.4>

codepage-identifier = (first-Latin-identifier-character / CP2-character) *(subsequent-Latin-identifier-character / CP2-character)

CP2-character = <any Unicode character that has a mapping to the character range %x80-FF in a Microsoft Windows supported code page>
```

アルファベット以外の表意文字を含む識別子に対する VBA のサポートは [Unicode](./1_はじめに.md) が作成されるより前の文字コード標準に基づいて設計されたため、非アルファベット識別子は類似の Unicode 文字クラスを直接使用するのではなく、これらのレガシー標準のコードポイントに対応する Unicode 文字から指定されています。

Microsoft Windows コードページ内の文字に対応する Unicode 文字で、1 バイトのコードポイントが %x80-FF のものはすべて有効な `<CP2-characters>` となる。このような文字を定義しているコードページは、Windows コードページ 874, 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258 である。これらのコードページの定義と、個々のコードページ固有のコードポイントの Unicode コードポイントへのマッピングは、[[UNICODE-BESTFIT]](https://go.microsoft.com/fwlink/?LinkId=95708) でホストされているファイルによって指定され、[[UNICODE-README]](https://go.microsoft.com/fwlink/?LinkId=95709) によって説明されている。[[CODEPG]](https://go.microsoft.com/fwlink/?LinkId=89840) は、コードページのコードポイントとその対応する Unicode 文字へのマッピング情報なを提供する。
##### 3.3.5.1.1 日本語識別子

日本語を含む識別子に対する VBA サポートは Windows コードページ 932 [[UNICODE-BESTFIT[](https://go.microsoft.com/fwlink/?LinkId=95708) に基づいている。日本語文字は、%x80 から始まるコードポイントを持つ 8 ビットのシングルバイト文字と 16 ビットのダブルバイト文字としてエンコードされている。Windows コードページ 932 のコードポイントに相当する [Unicode](./1_はじめに.md) は [UNICODE-BESTFIT] で提供されているファイル bestfit932.txt で指定されている。%x80-FF の範囲の文字の多くは、コードポイントの 16 ビットエンコーディングの先頭バイトとして機能する。ただし、この範囲内にも有効な文字は存在する。

`<CP932-initial-character>` は、定義済みの[コードページ](./1_はじめに.md) 932 に対応する任意の Unicode 文字にすることができる。この文字の Windows コードページ 932 のコード ポイントは %x7F よりも大きくなる。ただし、先頭バイトが %x80-FF の範囲のコードポイントと、明示的に除外されているコードポイント %x8140, %x8143-8151, %x815E-8197, %x824f-8258 は除く。

`<CP932-subsequent-character>` は コードポイント %x824f-8258 を除いて `<CP932-initial-character>` と同様に定義される。

##### 3.3.5.1.2 韓国語識別子

韓国語を含む識別子に対する VBA のサポートは、Windows コードページ 949 [[UNICODE-BESTFIT]](https://go.microsoft.com/fwlink/?LinkId=95708) に基づいている。韓国語文字は、%x8141 で始まるコードポイントを持つ 16 ビットのダブルバイト文字としてエンコードされている。Windows コードページ 949 のコードポイントに相当する [Unicode](./1_はじめに.md) は [UNICODE-BESTFIT] で提供されているファイル bestfit949.txt で指定されている。%x81-FE の範囲のコードポイントはすべて、コードポイントの 16 ビットエンコーディングの先頭バイトとして機能する。

`<CP949-initial-character>` は、Windows コードページ 949 文字コードポイントのうち、先頭バイトが %xA1 以上 %xAF 未満の定義済み 16 ビットコードポイント、先頭バイトの値に関わらずセカンドバイトが %xA1 以上 %xFE 未満の定義済みコードポイント、%xA3C1-A3DA の範囲のコードポイント、%xA3E1-A3FA の範囲のコードポイント、%xA4A1-A4FE に対応する任意の Unicode 文字で<ins>あってもよい</ins>。

`<CP949-subsequent-character>` は、コードポイント %xA3DF と %xA3B0-A3B9 を加えて `<CP949-initial-character>` と同様に定義される。

##### 3.3.5.1.3 簡体字中国語識別子

簡体字中国語を含む識別子に対する VBA のサポートは、Windows コードページ 936 [[UNICODE-BESTFIT]](https://go.microsoft.com/fwlink/?LinkId=95708) に基づいている。簡体字中国語文字は、%x8140 で始まるコードポイントを持つ 16 ビットのダブルバイト文字としてエンコードされている。Windows コードページ 936 のコードポイントに相当する [Unicode](./1_はじめに.md) は、[UNICODE-BESTFIT] で提供されているファイル bestfit936.txt で指定されている。

`<CP936-initial-character>` は、Windows コードページ 936 のコードポイントのうち、 %xA3C1-A3DA, %xA3E1-A3FA, %xA1A2A1AA, %xA1AC-A1AD, %xA1B2-A1E6, %xA1E8-A1EF, %xA2B1-A2FC, %xA4A1-FE4F に対応する任意の Unicode 文字で<ins>あってもよい</ins>。

`<CP936-subsequent-character>` は、コードポイント %xA3DF と %xA3B0-A3B9 を加えて `<CP949-initial-character>` と同様に定義される。

##### 3.3.5.1.4 繁体字中国語識別子

繁体字中国語を含む識別子に対する VBA のサポートは、Windows コードページ 950 [[UNICODE-BESTFIT]](https://go.microsoft.com/fwlink/?LinkId=95708) に基づいている。繁体字中国語文字は、%xA140 で始まるコードポイントを持つ 16 ビットダブルバイト文字としてエンコードされている。Windows コードページ 950 のコードポイントに相当する [Unicode](./1_はじめに.md) は、[UNICODE-BESTFIT] で提供されているファイル bestfit950.txt で指定されている。

`<CP950-initial-character>` は、Windows コードページ 950 のコードポイントのうち、%xA2CF-A2FE, %xA340-F9DD に対応する任意の Unicode 文字で<ins>あってもよい</ins>。

`<CP950-subsequent-character>` は、コードポイント %xA1C5 と %xA2AF-A2B8 を加えて `<CP950-initial-character>` と同様に定義される。

#### 3.3.5.2 予約済み識別子とそれ以外の識別子

`<reserved-identifier>` は `<Latin-identifier>` に準拠し、VBA 言語内で特別な用途のために予約されているすべての文字列を指定する。キーワードは `<reserved-identifier>` を意味する代替用語である。この仕様の散文セクションで特定のキーワードを指定する必要がある場合、キーワードは太字で強調されて書かれている（訳注：太字ではなくコードブロックで表記する）。すべての VBA 識別子と同様に `<reserved-identifier>` は大文字と小文字を区別しない。`<reserved-identifier>` はトークン（セクション 3.３）である。構文文法内の要素として `<reserved-identifier>` 要素の 1 つが引用されて出現すると、対応するトークンへの参照となる。トークン要素 `<IDENTIFIER>` は `<reserved-identifier>` でない識別子を指定するために構文文法内で使用される。

静的セマンティクス

- `<IDENTIFIER>` の名称値は `<lex-identifier>` のテキストである。
- `<reserved-identifier>` トークンの名称値はその `<Latin-identifier>` のテキストとなる。
- 大文字小文字を区別しないテキスト比較で同じになのであれば、2 つの名称値は同じものとなる。
    - `<reserved-identifier>` は、以下のルールで用途別に分類される。中には複数の用途を持ち複数の規則に登場するものもある。

```
statement-keyword = "Call" / "Case" /"Close" / "Const"/ "Declare" / "DefBool" / "DefByte" / "DefCur" / "DefDate" / "DefDbl" / "DefInt" / "DefLng" / "DefLngLng" / "DefLngPtr" / "DefObj" / "DefSng" / "DefStr" / "DefVar" / "Dim" / "Do" / "Else" / "ElseIf" / "End" / "EndIf" /  "Enum" / "Erase" / "Event" / "Exit" / "For" / "Friend" / "Function" / "Get" / "Global" / "GoSub" / "GoTo" / "If" / "Implements"/ "Input" / "Let" / "Lock" / "Loop" / "LSet" / "Next" / "On" / "Open" / "Option" / "Print" / "Private" / "Public" / "Put" / "RaiseEvent" / "ReDim" / "Resume" / "Return" / "RSet" / "Seek" / "Select" / "Set" / "Static" / "Stop" / "Sub" / "Type" / "Unlock" / "Wend" / "While" / "With" / "Write"

rem-keyword = "Rem"
marker-keyword = "Any" / "As" / "ByRef" / "ByVal" / "Case" / "Each" / "Else" / "In"/ "New" / "Shared" / "Until" / "WithEvents" / "Write" / "Optional" / "ParamArray" / "Preserve" / "Spc" / "Tab" / "Then" / "To"

operator-identifier = "AddressOf" / "And" / "Eqv" / "Imp" / "Is" / "Like" / "New" / "Mod" / "Not" / "Or" / "TypeOf" / "Xor"
```

`<statement-keyword>` は、ステートメントや宣言の最初の構文要素である `<reserved-identifier>` である。`<marker-keyword>` は、文の内部の構文構造の一部として使われる `<reserved-identifier> `である。`<operator-identifier>` は `<reserved-identifier> `で式の中で演算子として使われる。

```
reserved-name = "Abs" / "CBool" / "CByte" / "CCur" / "CDate" / "CDbl" / "CDec" / "CInt" / "CLng" / "CLngLng" / "CLngPtr" / "CSng" / "CStr" / "CVar" / "CVErr" / "Date" / "Debug" / "DoEvents" / "Fix" / "Int" / "Len" / "LenB" / "Me" / "PSet" / "Scale" / "Sgn" / "String"

special-form = "Array" / "Circle" / "Input" / "InputB"  / "LBound" / "Scale" / "UBound"
```

```
literal-identifier = boolean-literal-identifier / object-literal-identifier / variant-literal-identifier
boolean-literal-identifier = "true" / "false"
object-literal-identifier = "nothing"
variant-literal-identifier = "empty" / "null"
```

`<reserved-name>` は `<reserved-identifier>` で、通常のプログラム定義のエンティティ（セクション 2.2）のように式中で使用される。`<special-form>` は `<reserved-identifier> `で、プログラム定義のプロシージャ名のように式中で使われるが、その引数には特別な構文規則がある。`<reserved-type-identifier>` は、宣言内であるエンティティの宣言型（セクション 2.2）を特定するために使用される。

`<literal-identifier>`は、特定の識別されたデータ値（セクション 2.1）を表す `<reserved-identifier>` である。`True` または `False` を指定する `<boolean-literal-identifier>` は `Boolean` の宣言型を持ち、それぞれ `True` または `False` のデータ値を持つ。`<object-literal-identifier>` は、`Object` の宣言型を持ち、`Nothing` のデータ値を持つ。"empty" または "null" を指定する `<variant-literal-identifier>` は `Variant` の宣言型を持ち、それぞれ `Empty` と `Null` のデータ値を持つ。

```
reserved-for-implementation-use = "Attribute" / "LINEINPUT" / "VB_Base" / "VB_Control" / "VB_Creatable" /  "VB_Customizable" / "VB_Description" / "VB_Exposed" / "VB_Ext_KEY " / "VB_GlobalNameSpace" / "VB_HelpID" / "VB_Invoke_Func" / "VB_Invoke_Property " / "VB_Invoke_PropertyPut" / "VB_Invoke_PropertyPutRefVB_MemberFlags" / "VB_Name" / "VB_PredeclaredId" / "VB_ProcData" / "VB_TemplateDerived" / "VB_UserMemId" / "VB_VarDescription" / "VB_VarHelpID" / "VB_VarMemberFlags" / "VB_VarProcData " / "VB_VarUserMemId"

future-reserved = "CDecl" / "Decimal" / "DefDec"
```

`<reserved-for-implementation-use>` は `<reserved-identifier>` であり、現在 VBA 言語には定義されていないが、言語実装者が使用するために予約されている。`<future-reserved>` は `<reserved-identifier>` であり、現在は VBA 言語に対して意味を持たず、将来起こりうる言語の拡張のために予約されたものである。

#### 3.3.5.3 特殊な識別子構文

```
FOREIGN-NAME = "[" foreign-identifier "]"
foreign-identifier = 1*non-line-termination-character
```

`<FOREIGN-NAME>` は識別子のように使用されるが、識別子を形成するための VBA 規則に適合しないテキストシーケンスを表すトークン（セクション 3.3）である。通常、`<FOREIGN-NAME>` は VBA 以外のプログラミング言語を使って作成されたエンティティ（セクション 2.2）を参照するために使用される。

静的セマンティクス

- `<FOREIGN-NAME>` の名称値（セクション 3.3.5.1）は、その `<foreign-identifier>` のテキストとなる。

```
BUILTIN-TYPE = reserved-type-identifier /  ("[" reserved-type-identifier "]") / "object" / "[object]"
```

- VBA のコンテキストによっては、`<reserved-type-identifier>` と同じ名前の `<FOREIGN-NAME>` はその `<reserved-type-identifier>` と同等に使用することができる。名称値が "object" の識別子は`<reserved-identifier>` ではないが、一般に `<reserved-type-identifier>` であるかのように使われる。

静的セマンティクス

- `<BUILTIN-TYPE>` の名称値は、その `<reserved-type-identifier>` 要素があればそのテキストとなる。そうでなければ "object" である。
- `<BUILTIN-TYPE>` 要素の宣言型（セクション 2.2）は、`<BUILTIN-TYPE>` の名称値と同じ名前を持つ宣言型とする。

```
TYPED-NAME = IDENTIFIER  type-suffix

type-suffix = "%" / "&" / "^" / "!" / "#" / "@" / "$"
```

`<TYPED-NAME>` は、`<IDENTIFIER>` の直後に `<type-suffix>` が空白を挟むことなく続くものである。

静的セマンティクス

- `<TYPED-NAME>` の名称値は、その `<IDENTIFIER>` 要素の名称値である。
- `<TYPED-NAME>` の宣言型は，次の表で定義される。

| `<type-suffix>` | 宣言型 |
| ---- | ---- |
| % | `Integer`|
| & | `Long` |
| ^ | `LongLong` |
| ! | `Single` |
| # | `Double` |
| @ | `Currency` |
| $ | `String` |

## 3.4 条件付きコンパイル

モジュール本体には、モジュール（セクション 4.2）で定義された VBA プログラムコードの一部として、条件付きで解釈から除外できる論理行（セクション 3.2 ）を含めることができる。このような行を論理的に除去したモジュール本体（セクション 4.2）を前処理済モジュール本体と呼ぶ。前処理されたモジュール本体は、以下の文法に従ってトークン化された `<module-body-lines>` 内の条件付きコンパイル指示を解釈することによって決定される。

```
conditional-module-body = cc-block
cc-block = *(cc-const / cc-if-block / logical-line)
```

静的セマンティクス

- この文法の規則に従わない `<module-body-logical-structure>` は、有効な VBA モジュールではない。
- `<conditional-module-body>` を直接構成する `<cc-block>` は包含ブロックである。
- 包含ブロックの直接の要素である `<logical-line>` 行は、すべて前処理されたモジュール本体に含まれる。
- 除外ブロック (セクション 3.4.2) の直接の要素である `<logical-line>` 行はすべて前処理されたモジュール本体に含まれない。
- 前処理されたモジュール本体内の `<logical-line>` 行の相対的な順序は，元のモジュール本体内の相対的な順序と同じである。

### 3.4.1 条件付きコンパイル Const ディレクティブ

```
cc-const = LINE-START  "#"  "const" cc-var-lhs "=" cc-expression cc-eol
cc-var-lhs = name
cc-eol = [single-quote *non-line-termination-character] LINE-END
```

静的セマンティクス

- すべての `<cc-const>` 行は、前処理されたモジュール本体（セクション 3.4）から除外される。
- すべての `<cc-const>` 命令は、除外ブロック（セクション 3.4.2）に含まれるものを含めて処理される。
- `<cc-var-lhs>` が `<type-suffix>` を持つ `<TYPED-NAME>` である場合、`<type-suffix>` は無視される。
- `<cc-var-lhs>` の `<name>` の名称値（セクション 3.3.5.1）は、`<conditionalmodule-body>` 内のすべての `<cc-var-lhs>` （`<cc-block>` が除外ブロックのものを含む）について異なって<ins>いなければならない</ins>。
- `<cc-expression>` のデータ値（セクション 2.1）は、`<cc-expression>` の定数値（セクション 5.6.16.2）となる。
- `<cc-expression>` の定数評価で評価エラーが発生した場合、前処理されたモジュール本体の内容は不定となる。
- `<cc-const>` は、包含モジュールの `<cc-expression>` 要素にアクセス可能な定数バインディングを定義する。バインドされた名前は `<cc-var-lhs>` の `<name>` の名称値であり、定数バインディングの宣言型は `Variant` であり、定数バインディングのデータ値は `<cc-expression>` のデータ値である。
- `<cc-var-lhs>` の `<name>` の値は、プロジェクトレベルの条件付きコンパイル定数の束縛名と同じにすることができる。その場合、`<cc-const>` 要素で定義される定数バインディングは、プロジェクトレベルのバインディングを陰で支える。

### 3.4.2 条件付きコンパイル If ディレクティブ

```
cc-if-block = cc-if
    cc-block
    *cc-elseif-block
    [cc-else-block]
    cc-endif

cc-if = LINE-START  "#" "if" cc-expression "then" cc-eol

cc-elseif-block = cc-elseif cc-block
cc-elseif = LINE-START "#" "elseif" cc-expression "then" cc-eol

cc-else-block = cc-else cc-block
cc-else = LINE-START "#" "else" cc-eol

cc-endif = LINE-START "#" ("endif" / ("end" "if")) cc-eol
```
静的セマンティクス

- `<cc-if-block>` の構成要素である `<cc-expression>` は、`<cc-if-block>` が包含ブロック（セクション 3.4）に含まれていない場合でも、すべて次の規則に<ins>従わなければならない</ins>。
- `<cc-if>` 内の `<cc-expression>` と各 `<cc-elseif>` 内のものはそれぞれ評価される。
- 構成要素である `<cc-expression>` のデータ値（セクション 2.1）は、すべて `Boolean` データ型（セクション 2.1）に `Let` 強制可能で<ins>なければならない</ins>。
- 構成要素である `<cc-expression>` のいずれかが評価エラーとなった場合、前処理されたモジュール本体（セクション 3.4）の内容は未定義となる。
- `<cc-if-block>` が包含ブロックに含まれる場合、次の規則の順次適用に従い `<cc-block>` が最大 1 つ包含ブロックとして選択される。
    1. `<cc-if>` 内の `<cc-expression> `の評価値が真の場合、`<cc-if>` の直後に続く `<cc-block>` が包含ブロックとなる。
    2. `<cc-elseif>` 内にある `<cc-expression>` 要素の 1 つ以上の評価値が真である場合、その最初の`<cc-elseif>` の直後にある `<cc-block>` が包含ブロックである。
    3. 評価された` <cc-expression>` 要素のいずれも真の値を持たず、`<cc-else-block>` が存在する場合、`<cc-else-block>` の要素である `<cc-block>` は包含ブロックである。
    4. 評価された` <cc-expression>` のいずれも真の値を持たず、`<cc-else-block>` が存在しない場合、包含ブロックは存在しない。
- `<cc-if-block>`, `<cc-elseif-block>`, `<cc-else-block>` の直接の要素であり、包含ブロックでない `<cc-block>` は除外ブロック（セクション 3.4）となる。
- すべての `<cc-if>`, `<cc-elseif>`, `<cc-else>`, `<cc-endif>` 行は、前処理されたモジュール本体から除外される。
