# 4 VBA のプログラム構成

VBA 環境は、多数のユーザー定義プロジェクトとホストアプリケーション定義プロジェクト（セクション 4.1）に整理することができる。各プロジェクトは 1 つまたは複数のモジュール（セクション 4.2）で構成される。

## 4.1 プロジェクト

プロジェクトは、VBA プログラムコードを定義し VBA 環境に組み込むための単位である。プロジェクトは論理的に、プロジェクト名、モジュール名、プロジェクト参照の順序付きリストから構成される。このリストで先に出現したプロジェクト参照は、リストで後に出現した参照より高い参照優先度を持つと言われる。プロジェクトの物理的な表現と、プロジェクトの命名、保存、アクセスに使用されるメカニズムは実装定義される。

プロジェクト参照は、あるプロジェクトが他のプロジェクトで定義されている公開エンティティ（セクション 2.2）にアクセスすることを指定する。プロジェクトの参照元を特定する仕組みは実装定義される。

VBA プロジェクトには、ソースプロジェクト、ホストプロジェクト、ライブラリプロジェクトの 3 種類がある。ソースプロジェクトは、ソースコードの形で存在する VBA プログラムコードで構成される。ライブラリプロジェクトは、VBA 言語のソースコード形式では存在せず、VBA 言語を使用して実装されていない可能性があることを除けばソースプロジェクトが定義する可能性があるすべての種類のエンティティを定義可能な、実装定義されたプロジェクトである。

ホストプロジェクトは、ホストアプリケーションによって VBA 環境に導入されるライブラリプロジェクトである。導入手段は実装依存である。ホストプロジェクトで定義された公開変数（セクション 5.2.3.1）、定数、プロシージャ、クラス（セクション 2.5）、UDT は、ホストプロジェクトがソースプロジェクトであるかのように、同じ VBA 環境内の VBA ソースプロジェクトからアクセス可能である。オープンホストプロジェクトとは、ホストアプリケーション以外のエージェントがモジュールを追加できるプロジェクトのことである。オープンホストプロジェクトを指定する方法とモジュールを追加する方法は、実装定義されている。

静的セマンティクス

- プロジェクト名は \<IDENTIFIER\> として有効で<ins>なければならない</ins>。
- プロジェクト名は "VBA" に<ins>しないほうがよい</ins>。この名前は VBA 標準ライブラリ（セクション 2.7.1）にアクセスするために予約されている。
- プロジェクト名は \<reserved-identifier\> に<ins>しないほうがよい</ins>。
- 特定のプロジェクトのプロジェクトリ参照は、個別の名称を持つプロジェクトを<ins>識別しなければならない</ins>。
- ソースプロジェクトが、参照するプロジェクトと同じ名称を持つ別のプロジェクトを参照するかどうかは、実装依存である。

## 4.2 モジュール

モジュールは VBA ソースコードの基本的な構文単位である。モジュールの物理的な表現は実装依存だが、論理的には VBA 言語文法に準拠した Unicode 文字列である。

モジュールは、モジュールヘッダとモジュール本体の 2 つの部分から構成される。

モジュールヘッダは名前と値のペアで構成される属性のセットで、モジュールの特定の言語特性を指定する。モジュールヘッダは人間が直接書くこともできるが、より一般的には、プログラマーが実装固有のツールを使用することに基づいて VBA 実装により機械的に生成される。

モジュール本体は実際の VBA 言語ソースコードからなり、最も一般的には人間のプログラマによって直接記述される。

VBA では、プロシージャモジュールとクラスモジュールの 2 種類のモジュールをサポートしており、その内容はそれぞれ \<procedural-module\> と \<class-module\> というそれぞれの構文に<ins>従わなければならない</ins>。

```
procedural-module = LINE-START procedural-module-header EOS
                    LINE-START procedural-module-body
class-module = LINE-START class-module-header 
               LINE-START class-module-body

procedural-module-header = attribute "VB_Name" attr-eq quoted-identifier attr-end

class-module-header = 1*class-attr

class-attr = attribute "VB_Name" attr-eq quoted-identifier attr-end
     /  attribute "VB_GlobalNameSpace" attr-eq "False" attr-end
     /  attribute "VB_Creatable" attr-eq "False" attr-end
     /  attribute "VB_PredeclaredId" attr-eq boolean-literal-identifier attr-end
     /  attribute "VB_Exposed" attr-eq boolean-literal-identifier attr-end
     /  attribute "VB_Customizable" attr-eq boolean-literal-identifier attr-end
attribute = LINE-START "Attribute"
attr-eq = "="
attr-end = LINE-END

quoted-identifier = double-quote NO-WS IDENTIFIER NO-WS double-quote
```

静的セマンティクス

- \<attribute\> 要素に続く \<IDENTIFIER\> の名称値（セクション 3.3.5.1）は属性名となる。
- \<attr-eq\> 要素に続く要素は、同じ \<attr-eq\> に先行する属性名に対する属性値を定義する。
- \<quoted-identifier\> で定義される属性値は、それに含まれる識別子の名称値である。
- 与えられた \<class-module-header\> 内の特定の属性名に対する最後の \<class-attr\> は、その属性名に対する属性値を提供する。
- \<class-module-header\> 内に特定の属性名のための \<class-attr\> が存在しない場合、次の表に従ってデフォルトの属性値が属性名と関連付けられたものと仮定される。

| 属性名 | デフォルト属性値 |
| ---- | ---- |
| VB_Creatable | `False` |
| VB_Customizable | `False` |
| VB_Exposed | `False` |
| VB_GlobalNameSpace | `False` |
| VB_PredeclaredId | `False` |

- モジュール名は、モジュールの VB_NAME 属性値である。
- モジュール名の最大長は 31 文字である。
- モジュール名は \<reserved-identifier\> に<ins>しないほうがよい</ins>。
- モジュール名は、そのモジュールを含むプロジェクトのプロジェクト名（セクション 4.1）や、それを含むプロジェクト（セクション 4.1）から参照されるプロジェクト名と同じでない場合がある。
- プロジェクトに含まれるすべてのモジュールは、個別のモジュール名を<ins>持たなければならない</ins>。
- ソースプロジェクト（セクション 4.1）に含まれるモジュールでは、VB_GlobalNamespace 属性と VB_Creatable 属性の両方が "False" で<ins>なければならない</ins>。ただし、ライブラリプロジェクト（セクション 4.1）は、これらの属性が "True" であるモジュールを含むことができる。
- その他にも、クラスモジュールの定義で使用される特定の属性や属性の組み合わせの意味をセクション 5.2.4.1 に定義している。それ以外の属性の使い方や意味は実装依存とする。

### 4.2.1 モジュール拡張

オープンなホストプロジェクト（セクション 4.1）には、拡張可能モジュールを含めることができる。拡張可能モジュールとは、ホストプロジェクトに追加された、同じ名前の外部提供の拡張モジュールによって拡張することができるモジュール（セクション 4.2）である。拡張モジュールとは、変数（セクション 2.3）、定数、プロシージャ、UDT エンティティ（セクション 2.2）を追加定義したモジュールのことである。追加された拡張モジュールのエンティティは、対応する拡張可能モジュール内で直接定義されているかのように動作する。このとき、拡張可能モジュールがその中のイベントハンドラプロシージャのターゲットとなる `WithEvents` 変数を定義できることに注意すること。

ホストプロジェクト（セクション 4.1）に拡張モジュールを追加できる仕組みは実装定義である。

静的セマンティクス

- 拡張モジュールのモジュール名（セクション 4.2）は、それが拡張している拡張可能モジュールと<ins>同一でなければならない</ins>。
- 拡張モジュールは、対応する拡張可能モジュールで既に定義されている変数、定数、プロシージャ、列挙型、UDT を定義したり再定義したりすることはできない。拡張モジュールの要素が対応する拡張モジュールのモジュール本体（セクション 4.2）に物理的に含まれている場合と同じ名前衝突ルールが適用される。
- 拡張モジュールに含まれるオプション指示は、その拡張モジュールにのみ適用され、対応する拡張可能モジュールには適用されない。
- 特定の拡張モジュールに対して一つの拡張可能プロジェクト内に複数の拡張可能モジュールが存在する可能性があるかどうかは実装定義である。
