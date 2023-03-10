# MS-VBAL-JP
VBA Language Specification 日本語訳

## このドキュメントについて
VBA の言語仕様ドキュメントを探していたところ Microsoft 公式のドキュメントを見つけましたので、少しずつ日本語訳して読み進めていきます。 
私の英語能力は拙いので、基本的には各種翻訳サービスを利用させてもらうつもりですが、機械翻訳による一貫性のない表現や回りくどい言い回しについては極力修正できればと思っています。  

## 原文
- Web ページ : [[MS-VBAL]: VBA Language Specification](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74?redirectedfrom=MSDN)
- PDF : [[MS-VBAL].pdf](./org/[MS-VBAL].pdf)
- DOCX : [[MS-VBAL]-210216.docx](./org/[MS-VBAL]-210216.docx)

上記リンク先で公開されている原文のバージョン 1.7 （2021年2月16日公開）を対象にしています。  
PDF版およびDOCX版はWebページからダウンロードできるものと同じファイルです（リンク切れになったときのためにコピーを置いてあります）。  

## 翻訳済みセクション
- [1 はじめに](./doc/1_はじめに.md)
    - 1.1 用語の解説
    - 1.2 参考文献
        - 1.2.1 規範となる参考文献
        - 1.2.2 情報に関する参考文献
    - 1.3 VBA 言語仕様の概要
    - 1.4 仕様の規定
- [2 VBA の実行環境](./doc/2_VBAの実行環境.md)
    - 2.1 データ値とデータ型
    - 2.2 エンティティと宣言型
    - 2.3 変数
        - 2.3.1 集合変数
    - 2.4 プロシージャ
    - 2.5 オブジェクト
        - 2.5.1 オブジェクトの自動インスタンス化
    - 2.6 プロジェクト
    - 2.7 拡張環境
        - 2.7.1 VBA 標準ライブラリ
        - 2.7.2 外部の変数、プロシージャ、オブジェクト
        - 2.7.3 ホスト 環境
- [3 VAB の構文規則](./doc/3_VBAの構文規則.md)
    - 3.1 文字エンコーディング
    - 3.2 モジュールの行構成
        - 3.2.1 物理行文法
        - 3.2.2 論理行文法
    - 3.3 字句トークン
        - 3.3.1 セパレータと特殊トークン
        - 3.3.2 数値トークン
        - 3.3.3 日付トークン
            - 3.3.3.1 日付トークンの解釈方法
        - 3.3.4 文字列トークン
        - 3.3.5 識別子トークン
            - 3.3.5.1 非アルファベット識別子
                - 3.3.5.1.1 日本語識別子
                - 3.3.5.1.2 韓国語識別子
                - 3.3.5.1.3 簡体字中国語識別子
                - 3.3.5.1.4 繁体字中国語識別子
            - 3.3.5.2 予約済み識別子とそれ以外の識別子
            - 3.3.5.3 特殊な識別子構文
    - 3.4 条件付きコンパイル
        - 3.4.1 条件付きコンパイル Const ディレクティブ
        - 3.4.2 条件付きコンパイル If ディレクティブ
- [4 VBA のプログラム構成](./doc/4_VBAのプログラム構成.md)
    - 4.1 プロジェクト
    - 4.2 モジュール
        - 4.2.1 モジュール拡張
- [5 モジュール本体](./doc/5_モジュール本体.md)
    - 5.1 モジュール本体の構造
    - 5.2 モジュール宣言部の構造
        - 5.2.1 Option ディレクティブ
            - 5.2.1.1 Option Compare ディレクティブ
            - 5.2.1.2 Option Base ディレクティブ
            - 5.2.1.3 Option Explicit ディレクティブ
            - 5.2.1.4 Option Private ディレクティブ
        - 5.2.2 暗黙定義ディレクティブ
        - 5.2.3 モジュール宣言部
            - 5.2.3.1 モジュール宣言の変数リスト
                - 5.2.3.1.1 変数宣言
                - 5.2.3.1.2 WithEvents 変数宣言
                - 5.2.3.1.3 配列の次元と境界
                - 5.2.3.1.4 変数の型宣言
                - 5.2.3.1.5 暗黙の型判定
            - 5.2.3.2 Const 宣言
            - 5.2.3.3 ユーザ定義型（UDT）宣言
            - 5.2.3.4 列挙型宣言
            - 5.2.3.5 外部プロシージャ宣言
            - 5.2.3.6 循環モジュールの依存関係
        - 5.2.4 クラスモジュール宣言
            - 5.2.4.1 構文外で定義されるクラス特性
                - 5.2.4.1.1 クラスのアクセシビリティとインスタンス化
                - 5.2.4.1.2 デフォルトインスタンス変数の静的セマンティクス
            - 5.2.4.2 実装ディレクティブ
            - 5.2.4.3 イベント宣言
    - 5.3 モジュールのコード部の構造
        - 5.3.1 プロシージャ宣言
            - 5.3.1.1 プロシージャスコープ
            - 5.3.1.2 静的プロシージャ
            - 5.3.1.3 プロシージャ名
            - 5.3.1.4 関数の型宣言
            - 5.3.1.5 仮引数リスト
            - 5.3.1.6 サブルーチンと関数の宣言
            - 5.3.1.7 プロパティ宣言
            - 5.3.1.8 イベントハンドラ宣言
            - 5.3.1.9 実装名宣言
            - 5.3.1.10 ライフサイクルハンドラ宣言
            - 5.3.1.11 プロシージャの実引数処理
