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

VBA のプログラムテキストとして解釈するために、モジュール（セクション 4.2）は、それぞれが複数の物理行に対応することができる論理行の集合とみなす。この構造は論理行文法で記述される。この文法の終端記号は [Unicode](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/213ca0c8-6b82-4899-80a3-3c76eb534829#gt_c305d0ab-8b94-461a-bd76-13b40cb8c4d8) 文字コードポイントである。

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

VBA プログラムの構文は、個々の [Unicode](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/213ca0c8-6b82-4899-80a3-3c76eb534829#gt_c305d0ab-8b94-461a-bd76-13b40cb8c4d8) 文字ではなく、字句トークンの観点から最も簡単に記述することができる。特にほとんどの構文要素間の空白や行の連続は、通常、構文文法とは無関係である。このような空白の出現の可能性を記述する必要がなければ、構文文法は著しく単純化される。これは、空白を抽象化した字句トークン（単にトークンともいう）を構文文法の終端記号として用いることで達成される。

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
  false & otherwise
\end{cases}
$$

$$
LegalDay(month, day, year) = \begin{cases}
  false & \begin{cases}
    \textrm{year < 0 or year > 32767, or} \\
    \textrm{LegalMonth(month) is false, or} \\
    \textrm{day is not a valid day for the specified month and year}
  \end{cases} \\
  true & otherwise
\end{cases}
$$

- $CY$ を実装定義のデフォルトの年とする。

$$
Year(x) = \begin{cases}
  x + 2000 & 0 \le x \le 29 \\ x + 1900 & 30 \le x \le 99 \\
  x & otherwise
\end{cases}
$$

次に、以下の通り解釈する。

- $L$ と $M$ が数値で $R$ が存在しない場合、
    - もし $LegalMonth(L)$ および $LegalDay(L, M, CY)$ の場合、月は $L$ 、日は $M$ 、年は $CY$ である。
    - それ以外で、もし $LegalMonth(M)$ および $LegalDay(M, L, CY)$ の場合、月は$M$ 、日は $L$ 、年は $CY$ である。
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
- それ以外の（$R$ が存在し、$L, M, R$ のいずれかが数値ではない）場合、
    - 次の通りとする。
        - $L, M, R$ のうち数値でない値の月名または略号に対応する 1～12 の範囲の値を$M$ とする。
        - $L, M, R$ のうち数値である値を $N1, N2$ とする。（訳注：原文では両方 N1 とされており誤記と思われるので修正した）
    - もし $LegalDay(M, N1, Year(N2))$ の場合、月は $M$ 、日は $N1$ 、年は $Year(N2)$ である。
    - それ以外で、もし $LegalDay(M, N2, Year(N1))$ の場合、月は $M$ 、日は $N2$ 、年は $Year(N1)$ である。
    - それ以外の場合、`<date-value>` は有効ではない。
