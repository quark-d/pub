## Chapter 2 OpenAI APIの概要

* **OpenAI APIとは**
  OpenAI社が提供するAPIで、ChatGPTと同様にLLM（GPT-3.5やGPT-4）へプロンプトを渡して回答を得られる。開発者は自分のシステムに生成AI機能を組み込める。

* **APIの基本**

  * API = Application Programming Interface
  * コンピュータ同士のやり取りのルールや方法を定義。
  * リクエスト（要求）を送り、処理結果がレスポンス（応答）として返ってくる。
  * ATMや銀行システムの連携など、世の中の多くのサービスがAPIを利用している。

![APIとは](6_2_aboutApi.png)

* **OpenAI APIのモデル**

  * GPT-3 → GPT-3.5 → GPT-4 の順に進化。
  * 呼び出しは具体的なモデル名（例: `gpt-3.5-turbo`, `gpt-4o`, `gpt-4o-mini`）を指定。
  * 最新は GPT-4 系（`gpt-4o`, `gpt-4o-mini`）。
  * `gpt-4o-mini` は安価かつ高性能で推奨。
  * モデルには日付付きバージョンがあり、指定しない場合は最新が呼び出される。
  * OpenAIの[公式サイト](https://openai.com/ja-JP/api/pricing/)を見ると、「gpt-4o-2024-08-06」や「gpt-4o-2024-05-13」など日付の付いたモデルがあり、新しい日付のものがより新しいモデル。

* **利用料金**

  * モデルごとに異なり、トークン単位で計算。
  * 入力（プロンプト）と出力（回答）で料金が分かれる。
  * 例: `gpt-4o-mini` → 入力\$0.15 / 出力\$0.60（100万トークン単位）。
  * 1トークンは単語や記号の最小単位（例:「こんにちは」=1トークン）。
  * 最大トークン数にも制限があり、出力が少ないほど入力に使えるトークンが増える。

* **補足**

  * GPT-3 はすでにレガシー扱い。
  * 当コースでは主に `gpt-4o-mini` を使用し、トークン節約方法も解説予定。


### OpenAI API 概要まとめ表

| 項目                 | 内容                                                                                                                                                       |
| -------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **OpenAI APIとは**   | OpenAI社が提供するAPI。ChatGPT同様にLLMへプロンプトを渡して回答を生成。開発者は自分のシステムに組み込める。                                                |
| **APIの基本**        | API = Application Programming Interface。コンピュータ間のやり取りのルール。リクエスト送信 → 相手で処理 → レスポンス返却。ATMや金融システムなど幅広く利用。 |
| **利用できるモデル** | GPT-3（レガシー）、GPT-3.5（`gpt-3.5-turbo`）、GPT-4（`gpt-4o`, `gpt-4o-mini`）。最新はGPT-4系。日付付きバージョンあり。                                   |
| **推奨モデル**       | `gpt-4o-mini`（安価かつ高性能）。用途により `gpt-4o` を選択。                                                                                              |
| **利用料金**         | トークン単位で課金。入力と出力で料金が異なる。例: `gpt-4o-mini` 入力\$0.15／出力\$0.60（100万トークン単位）。                                              |
| **トークンとは**     | 文章を分割した最小単位。例：「こんにちは」= 1トークン。                                                                                                    |
| **最大トークン数**   | モデルごとに制限あり。例: `gpt-4o-mini` → 入出力合計128,000、出力最大16,384。出力が少ないほど入力に使えるトークンが増加。                                  |
| **コースでの利用**   | 基本的に `gpt-4o-mini` を使用。消費トークン数を節約する実装方法も解説予定。                                                                                |

### モデル別 料金比較表（100万トークンあたり）

| モデル      | モデル名      | 入力料金（USD） | 出力料金（USD） | 入力料金（JPY, \$1=150円） | 出力料金（JPY, \$1=150円） | 最大トークン数（入出力合計 / 出力） |
| ----------- | ------------- | --------------- | --------------- | -------------------------- | -------------------------- | ----------------------------------- |
| **GPT-4**   | gpt-4o        | \$2.50          | \$10.00         | 約¥375                     | 約¥1,500                   | 128,000 / 16,384                    |
| **GPT-4**   | gpt-4o-mini   | \$0.15          | \$0.60          | 約¥22.5                    | 約¥90                      | 128,000 / 16,384                    |
| **GPT-3.5** | gpt-3.5-turbo | \$3.00          | \$6.00          | 約¥450                     | 約¥900                     | 16,385 / 4,096                      |
| **GPT-3**   | davinci-002   | \$12.00         | \$12.00         | 約¥1,800                   | 約¥1,800                   | 16,384 / 16,384                     |
| **GPT-3**   | babbage-002   | \$1.60          | \$1.60          | 約¥240                     | 約¥240                     | 16,384 / 16,384                     |


参照1: https://openai.com/ja-JP/api/pricing/
参照2: https://platform.openai.com/docs/models/

👉 この比較を見ると、**性能とコストのバランスが良いのは「gpt-4o-mini」** だと分かります。
一方で、高精度な応答が必要な場面では「gpt-4o」を選択するケースもあります。

## Chapter 3 APIキーの取得方法

* **APIキーとは**

  * OpenAIなど外部サービスにアクセスするための「鍵」。
  * 認証や利用量管理に使われ、図書カードやログインパスワードに例えられる。
  * 利用時は必ずAPIキーを一緒に送信する。
  * 絶対に他者に共有してはいけない。
  
* **利用準備**

  * API利用にはクレジットカード登録が必須（従量課金制）。
  * 以前は無料枠（\$5分）があったが現在は廃止。
  
    OpenAIのGPTモデル（LLM）にプロンプトを渡す際は、このAPIキーも**毎回**一緒に渡す。
    APIキーを使ったLLM活用の流れは、基本的に以下の通り。
    1. APIキーを渡して認証を受ける
    2. プロンプトを渡して回答を生成してもらう
    3. プログラム上で回答を受け取り、画面上に表示するなどの処理を行う

* **アカウント作成・ログイン手順**

  1. OpenAIアカウントを[作成](https://platform.openai.com/signup)（メールアドレス・パスワード入力、本人確認）。
  2. 既にChatGPTを利用している場合はアカウント済み。
  3. [ログイン](https://platform.openai.com/login)はメール/パスワード、またはGoogleなど外部連携で可能。

* **クレジットカード登録方法**

  * 設定 → Billing → Add payment details。
  * 個人利用か会社利用を選び、カード情報を入力。
  * <del>登録時に一時的に\$5の請求がかかるが後で取り消される。</del>
  * 2024年のはじめまではアカウント作成時に5ドル分の無料枠が付与されてたが、
    無料枠の付与がなくなり、OpenAI APIの利用にあたりクレジットカードの登録が必須となった。

* **APIキー取得手順**

  1. 設定 → API Keys → 「Create new secret key」。
  2. 任意で名前を付け、キーを作成。
  3. 表示されたキーをコピーして保管（再表示不可）。
     - 画面を閉じた後、再度APIキーをコピーしようとしてもできないため、メモを取り忘れた場合は作り直す必要がある

> APIキーは絶対に他者に共有してはいけません。
> APIキーが外部に漏れてしまうと、悪用されて利用制限の満額まで使われる可能性があります。
> APIキーは絶対に外部に漏れないよう厳重に保管しましょう。
> 友人や同僚から頼まれて一時的に貸すのも避けるべきです。

* **トークンと料金**

  * 料金はトークン単位で計算。
  * 例: `gpt-4o-mini` → 入力 \$0.15 / 出力 \$0.60（100万トークンあたり）。
  * 日本語は英語よりトークン消費が多い。
  * 節約のためプロンプトを英訳 → 回答を和訳する方法もある。
  * 入力と出力のトークン数がどれくらいだったかはLLMからの回答結果と一緒に返ってくる
  * 

* **トークン数確認方法**

  * Webアプリ「Tokenizer」（簡易確認用）。
  * Pythonライブラリ「tiktoken」（高精度）。


* **トークン数確認方法例**
- Tokenizer（ https://platform.openai.com/tokenizer ）のサイトにアクセス
- 「現在はOpenAI APIについての勉強中です。」という文字列のトークン数は「13」
- 入力を削除し、続けて英語訳した「I'm currently learning about OpenAI API.」のトークン数は「8」

👉 まとめると、**OpenAI APIを使うにはアカウント作成 → クレジットカード登録 → APIキー取得 → トークン単位で従量課金** という流れになります。


了解しました！Chapter 3 の要約表に **注意点** を追加しました。

---

### Chapter 3 APIキーの取得方法（まとめ表）

| 項目                     | 内容                                                                                                                                                                                              |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **APIキーとは**          | OpenAIなど外部サービスにアクセスするための「鍵」。認証や利用量管理に利用。図書カードやログインパスワードに例えられる。必ずプロンプトと一緒に送信。外部共有は禁止。                                |
| **利用準備**             | API利用には **クレジットカード登録が必須**（従量課金制）。以前の無料枠（\$5）は廃止済み。                                                                                                         |
| **アカウント作成手順**   | 1. [signupページ](https://platform.openai.com/signup) で登録<br>2. メール・パスワード入力<br>3. 確認メールのリンクをクリック<br>4. 氏名・生年月日入力で完了<br>※ChatGPT利用者は既にアカウント有り |
| **ログイン方法**         | [loginページ](https://platform.openai.com/login) でメール＋パスワード、またはGoogleログイン可                                                                                                     |
| **クレジットカード登録** | 1. 設定（右上の歯車）<br>2. 「Billing」をクリック<br>3. 「Add payment details」<br>4. 個人 or 会社利用を選択<br>5. カード情報入力で完了<br>※一時的に\$5請求（7日以内に取消）                      |
| **APIキー取得手順**      | 1. 設定 → 「API Keys」<br>2. 「Create new secret key」クリック<br>3. 任意の名前を入力し作成<br>4. 表示されたキーをコピーし保管（再表示不可、紛失時は再発行）                                      |
| **料金体系**             | トークン単位で従量課金。例: `gpt-4o-mini` 入力 \$0.15／出力 \$0.60（100万トークン単位）。                                                                                                         |
| **トークンの特徴**       | - 文字列を分割した最小単位（例：「こんにちは」= 1トークン）<br>- 日本語は英語よりトークン消費が多い<br>- 節約のため英訳プロンプト＋和訳回答の手法あり                                             |
| **トークン数確認方法**   | - Webアプリ「Tokenizer」（簡易確認）<br>- Pythonライブラリ「tiktoken」（高精度）                                                                                                                  |
| **注意点**               | - APIキーは厳重に管理し、外部共有禁止<br>- 漏洩すると第三者に不正利用され高額請求のリスクあり<br>- 紛失時は再発行が必要（再表示不可）<br>- 利用量は常に確認し、上限を超えないよう注意             |



## Chapter 4 Chat Completions APIを使う準備

* **教材ファイル**

  * 記述用ファイル：「【記述用】Lesson6: OpenAI API」
  * ソースコード確認用ファイル：「【教材ソースコード】Lesson6: OpenAI API」

* **Chat Completions APIとは**

  * OpenAI APIには2種類：

    * **Completions API**（旧：GPT-3用、単一回答取得、Legacy扱い）
    * **Chat Completions API**（新：GPT-3.5 / GPT-4用、会話形式対応）
    * 今後はChat Completions APIを使用。推奨モデルは **gpt-4o-mini**（高性能・低コスト）。

| 項目               | Completions API                                                       | Chat Completions API                                                                                                                                                      |
| :----------------- | :-------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **使用可能モデル** | GPT-3 の古いモデル                                                    | GPT-3.5、GPT-4（例: gpt-3.5-turbo, gpt-4o-mini）                                                                                                                          |
| **目的**           | 単一の回答取得                                                        | 会話形式でのやり取り、履歴を保持した応答                                                                                                                                  |
| **特徴**           | - 単発の入力に対して単一回答を返す<br>- すでに **Legacy（旧式）扱い** | - 過去の会話を保持し、文脈に沿った応答が可能<br>- 過去の会話をLLMが理解しやすい形式で保持し、プロンプトに含めて渡すことができる<br>- 新しいモデルに対応、料金・性能が優秀 |
| **今後の利用方針** | 使う機会は少ない                                                      | コースの開発手法でメイン活用                                                                                                                                              |
| **推奨モデル**     | ―                                                                     | 高性能かつ安価な **gpt-4o-mini**（推奨）                                                                                                                                  |


* **ライブラリのインストール**

  * `openai` ライブラリ（教材推奨: `openai==1.47.0`）
  * `httpx` ライブラリ（推奨: `httpx==0.27.2`）
  * Google Colabでは先頭に「!」を付けて `!pip install`。
  * 依存関係エラーが出ても無視して問題なし。

```py
!pip install openai==1.47.0
```

```py
!pip install httpx==0.27.2
```

```py
# Google Colaboratory上での生成AI開発においてライブラリのインストール時に発生することが多い、パッケージの依存関係に関するエラー。
# 無視しても後続のコード実行に影響はないため、解決を試みず無視して進めてよい
ERROR: pip's dependency resolver does not currently take into account all the packages that are installed. <省略>
```

* **APIキーの設定方法**

  * APIキーは外部に漏らさないよう **シークレット機能** を利用する。
  * 手順：

    1. Colabのサイドバーで「鍵」アイコン → 「新しいシークレットを追加」
    2. 名前を `OPENAI_API_KEY`、値に取得したAPIキーを入力
    3. ノートブックからのアクセスを許可
  * 環境変数設定コード例：

    ```python
    import os
    from google.colab import userdata
    os.environ["OPENAI_API_KEY"] = userdata.get("OPENAI_API_KEY")
    ```
  * これによりAPI呼び出し時にキーが自動で参照される。


👉 まとめると、**Chat Completions APIを使う準備 = 教材ファイル利用 → ライブラリ導入 → シークレット機能でAPIキーを環境変数に設定** です。


### まとめ表

| 項目                       | 内容                                                                                                                                    |
| -------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| **教材ファイル**           | - 記述用：「【記述用】Lesson6: OpenAI API」<br>- ソースコード確認用：「【教材ソースコード】Lesson6: OpenAI API」                        |
| **APIの種類**              | - **Completions API**：GPT-3用、単一回答取得、Legacy（旧式）<br>- **Chat Completions API**：GPT-3.5 / GPT-4用、会話形式対応（推奨）     |
| **使用モデル**             | - `gpt-4o-mini`（高性能・低コスト、推奨）<br>- `gpt-3.5-turbo` も利用可                                                                 |
| **ライブラリインストール** | - `!pip install openai==1.47.0`<br>- `!pip install httpx==0.27.2`<br>※Colabではコマンド先頭に「!」必須<br>※依存関係エラーは無視してOK   |
| **APIキー管理方法**        | - 外部漏洩を防ぐため **シークレット機能** を利用<br>- 名前: `OPENAI_API_KEY`、値: 取得済みAPIキー<br>- ノートブックからのアクセスを許可 |
| **環境変数設定コード例**   | `python<br>import os<br>from google.colab import userdata<br>os.environ["OPENAI_API_KEY"] = userdata.get("OPENAI_API_KEY")<br>`         |
| **ポイント**               | - APIキーは直接コードに書かずシークレット機能で管理<br>- 環境変数に設定するとAPI呼び出し時に自動で参照される                            |

### Chat Completion というニュアンス

by Gemini
```md
確かに、
「Chat」（おしゃべり、会話）
と
「Completion」（完成、完了）
を組み合わせると、少し不思議な感じがしますよね。

この言葉を理解する鍵は、**「文章の完成」**という視点にあります。

OpenAIの「Chat Completions」は、一連の会話（Chat）を入力として受け取り、
その会話の続きとなる「応答文」（Completion）を生成し、**「会話を完成させる」**という仕組みです。

つまり、以下のように解釈すると意味合いがしっくりくるでしょう。

**「会話の続きを完成させる」**

たとえば、あなたが「昨日は何を食べた？」と尋ねたとします。
AIは「昨日はカレーを食べました」と答えます。

この時、AIはあなたの質問に対して、意味のある答え（文章）を「完成」させています。
この一連のプロセス全体を指して「Chat Completions」と呼んでいるのです。

単なる「おしゃべり」というよりも、
**「対話というタスクを、言葉を生成することで完成させる」**というニュアンスで捉えると、
より正確に理解できるかと思います。

```

### hello world

```py
!pip install openai==1.47.0 httpx==0.27.2
```
```py
Collecting openai==1.47.0
  Downloading openai-1.47.0-py3-none-any.whl.metadata (24 kB)
Collecting httpx==0.27.2
  Downloading httpx-0.27.2-py3-none-any.whl.metadata (7.1 kB)
Requirement already satisfied: anyio<5,>=3.5.0 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (4.10.0)
Requirement already satisfied: distro<2,>=1.7.0 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (1.9.0)
Requirement already satisfied: jiter<1,>=0.4.0 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (0.10.0)
Requirement already satisfied: pydantic<3,>=1.9.0 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (2.11.7)
Requirement already satisfied: sniffio in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (1.3.1)
Requirement already satisfied: tqdm>4 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (4.67.1)
Requirement already satisfied: typing-extensions<5,>=4.11 in /usr/local/lib/python3.12/dist-packages (from openai==1.47.0) (4.15.0)
Requirement already satisfied: certifi in /usr/local/lib/python3.12/dist-packages (from httpx==0.27.2) (2025.8.3)
Requirement already satisfied: httpcore==1.* in /usr/local/lib/python3.12/dist-packages (from httpx==0.27.2) (1.0.9)
Requirement already satisfied: idna in /usr/local/lib/python3.12/dist-packages (from httpx==0.27.2) (3.10)
Requirement already satisfied: h11>=0.16 in /usr/local/lib/python3.12/dist-packages (from httpcore==1.*->httpx==0.27.2) (0.16.0)
Requirement already satisfied: annotated-types>=0.6.0 in /usr/local/lib/python3.12/dist-packages (from pydantic<3,>=1.9.0->openai==1.47.0) (0.7.0)
Requirement already satisfied: pydantic-core==2.33.2 in /usr/local/lib/python3.12/dist-packages (from pydantic<3,>=1.9.0->openai==1.47.0) (2.33.2)
Requirement already satisfied: typing-inspection>=0.4.0 in /usr/local/lib/python3.12/dist-packages (from pydantic<3,>=1.9.0->openai==1.47.0) (0.4.1)
Downloading openai-1.47.0-py3-none-any.whl (375 kB)
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 375.6/375.6 kB 4.7 MB/s eta 0:00:00
Downloading httpx-0.27.2-py3-none-any.whl (76 kB)
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 76.4/76.4 kB 5.4 MB/s eta 0:00:00
Installing collected packages: httpx, openai
  Attempting uninstall: httpx
    Found existing installation: httpx 0.28.1
    Uninstalling httpx-0.28.1:
      Successfully uninstalled httpx-0.28.1
  Attempting uninstall: openai
    Found existing installation: openai 1.106.1
    Uninstalling openai-1.106.1:
      Successfully uninstalled openai-1.106.1
ERROR: pip's dependency resolver does not currently take into account all the packages that are installed. This behaviour is the source of the following dependency conflicts.
firebase-admin 6.9.0 requires httpx[http2]==0.28.1, but you have httpx 0.27.2 which is incompatible.
google-genai 1.33.0 requires httpx<1.0.0,>=0.28.1, but you have httpx 0.27.2 which is incompatible.
Successfully installed httpx-0.27.2 openai-1.47.0
```
```py
import os
from google.colab import userdata

os.environ["OPENAI_API_KEY"] = userdata.get("OPENAI_API_KEY")

print(os.environ["OPENAI_API_KEY"])

```

## Chapter 5：Chat Completions APIの基本

### できること（代表タスク）

* **質問応答**／**編集（誤字修正）**／**要約**／**翻訳**
* **文章の続きを生成**／**文章の“間”をつなぐ**（ブリッジ）
* **コード生成・リファクタ・エラーフィックス**

### 基本フロー

1. `OpenAI(api_key=...)` でクライアント生成
2. `client.chat.completions.create(model, messages, temperature, ...)` 実行
3. 応答本文は `completion.choices[0].message.content` から取得

#### 基本的な実装

```py
from openai import OpenAI

# APIキーは環境変数から安全に取得する（.envやシークレット管理を推奨）
client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

# Chat Completions APIで会話形式の応答を生成
completion = client.chat.completions.create(
  model="gpt-4o-mini",  # 高性能・低コストな推奨モデル
  messages=[
    {"role": "system", "content": "あなたは健康に関するアドバイザーです。安全なアドバイスを提供してください。"},  # AIの役割を指定
    {"role": "user", "content": "最近眠れないのですが、どうしたらいいですか？"},  # ユーザーからの質問
    {"role": "assistant", "content": "寝る前にリラックスできる音楽を聴く、またはスクリーンタイムを減らすとよいでしょう。"},  # AIの前回回答
    {"role": "user", "content": "カフェインも避けたほうがいいですか？"}  # ユーザーの追加質問
  ],
  temperature=0.5  # 応答の多様性（0に近いほど決定的）
)

# messagesに渡す各要素（辞書）のキーの仕様
# 主なキーは以下の3つです：
# - role（必須）："system"、"user"、"assistant" のいずれか
# - content（必須）：メッセージ本文（文字列）
# - name（任意）：話者の名前（必要な場合のみ）
# これ以外のキーは基本的に無視されるか、エラー。  
# roleとcontentは必須で、正しい値を指定する必要がある。

# 生成された応答を出力
print(completion.choices[0].message.content)
```

### messages設計（会話文脈は自前で渡す）


* `role` は **system / user / assistant** の3種類のみ
* 会話履歴を保持したい場合は、**過去分も含めて `messages` に積む**（ChatGPTのように自動保持されない）
* **system** で振る舞いを定義（編集者／要約家／翻訳者／作家／プログラマ など）

#### 【参考】messages要素のPlantUMLクラス図
```plantuml
@startuml
enum Role {
    SYSTEM = "system"
    USER = "user"
    ASSISTANT = "assistant"

}
class Message {
    -- requarement --
  - role: string 
  - content: string <<requarement>>
  -- optional --
  - name: string  <<optional>>// 話者名
}

Message [role] .> Role

class Messages {
  - messages: List<Message>
}

Messages "1" *-- "1..*" Message
@enduml
```

### client.chat.completions.create 主要パラメータ

* `model`：例 **gpt-4o-mini**（推奨：高性能・低コスト）
* `messages`：会話履歴（必須）
* `temperature`：出力のランダムさ。defaul=1, 0=固定, 2=完全にランダム
* **追加**：

  * `max_tokens`：最大出力長（超過時 `finish_reason="length"`）
  * `n`：候補数（`choices` に複数格納、比較に便利）
  * `stream`：逐次出力（本章では概念のみ）

パラメーターは膨大な数がある。
[公式](https://platform.openai.com/docs/api-reference/chat/create)


#### max_tokens

| 項目         | 内容                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| :----------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| パラメータ名 | `max_tokens`                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| 役割         | 出力されるテキストの最大トークン数を指定する                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| 効果         | 回答の長さに制限を設け、トークン消費を節約できる                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| デフォルト値 | モデルの最大出力トークン数（例：`gpt-4o-mini` = 16,384）                                                                                                                                                                                                                                                                                                                                                                                                                               |
| 使用例コード | `python<br>from openai import OpenAI<br><br>client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])<br><br>completion = client.chat.completions.create(<br>    model="gpt-4o-mini",<br>    messages=[<br>        {"role": "system", "content": "あなたは料理のアドバイザーです。"},<br>        {"role": "user", "content": "ヘルシーで簡単なレシピを教えてください。"}<br>    ],<br>    temperature=0.5,<br>    max_tokens=100<br>)<br><br>print(completion.choices[0].message.content)` |
| 実行結果例   | 「野菜たっぷりスープ」のレシピが出力されたが、トークン数制限により途中で打ち切り表示された                                                                                                                                                                                                                                                                                                                                                                                             |
| メリット     | 無駄に長い回答を避けられる・コスト削減につながる                                                                                                                                                                                                                                                                                                                                                                                                                                       |

#### temperature
| 項目         | 内容                                                                                                                                    |
| :----------- | :-------------------------------------------------------------------------------------------------------------------------------------- |
| パラメータ名 | `temperature`                                                                                                                           |
| 役割         | 出力のランダムさを制御する                                                                                                              |
| 効果         | 値が低いほど安定した回答、値が高いほど多様性のある回答を生成する                                                                        |
| デフォルト値 | `1`                                                                                                                                     |
| 使用例コード | `python<br>completion = client.chat.completions.create(<br>    model="gpt-4o-mini",<br>    messages=[...],<br>    temperature=0.2<br>)` |
| 実行結果例   | 値が 0.2 → 似たような回答を繰り返す／値が 1.5 → バリエーションのある回答が返る                                                          |
| メリット     | 目的に応じて「正確さ重視」「創造性重視」を切り替えられる                                                                                |

#### n
| 項目         | 内容                                                                                                                                                                                                                              |
| :----------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| パラメータ名 | `n`                                                                                                                                                                                                                               |
| 役割         | 一度に生成する回答数を指定する                                                                                                                                                                                                    |
| 効果         | 複数の候補を比較し、最適な回答を選べる                                                                                                                                                                                            |
| デフォルト値 | `1`                                                                                                                                                                                                                               |
| 使用例コード | `python<br>completion = client.chat.completions.create(<br>    model="gpt-4o-mini",<br>    messages=[...],<br>    temperature=0.5,<br>    n=3<br>)<br>for choice in completion.choices:<br>    print(choice.message.content)<br>` |
| 実行結果例   | 「日本旅行プラン」を3種類同時に生成（東京＋京都＋奈良、大阪＋沖縄、広島＋沖縄など）                                                                                                                                               |
| メリット     | 複数のアイデアを効率よく取得できる                                                                                                                                                                                                |



### レスポンスの読み方（抜粋）

* 本文：`choices[0].message.content`
* 終了理由：`choices[0].finish_reason`（`stop` / `length` など）
* トークン利用：`usage.prompt_tokens` / `usage.completion_tokens` / `usage.total_tokens`

### 料金とトークン最適化の意識

* 例：**gpt-4o-mini**（100万トークンあたり）

  * 入力 **\$0.15**、出力 **\$0.60**
* 短文なら微少だが、長文・大量処理で増加
* 節約例：**会話履歴の圧縮**、`max_tokens` の活用、`temperature` 調整、プロンプトの簡潔化

### コードパターン例（要旨）

* **質問応答**：役割＝健康アドバイザー等の system 指定 → user 質問 → assistant 回答を取得
* **編集／要約／翻訳**：system をタスク特化にし、user に対象文を渡す
* **続き生成／間生成**：冒頭や前後文を user で提示し、system で作家指示
* **コード生成・改善**：要件やエラー内容を user に、system で「**コードのみ出力**」等の制約を付与。`n` で複数案比較も可

---

**ポイント**：

* 望む振る舞いは **system メッセージで明確化**。
* **必要なコンテキストは必ず `messages` に同梱**。
* レスポンスの `usage` / `finish_reason` を確認し、**コストと品質を継続的に改善**する。

### 編集
```py
    messages=[
        {"role": "system", "content": "あなたは編集者です。与えられたテキストの誤字脱字を直してください。"},
        {"role": "user", "content": "きょうははれで、きぶんもすっきりんした。"}
    ],
```
### 要約
```py
    messages=[
        {"role": "system", "content": "あなたは文章要約の専門家です。元の意味を保ちながら、文章の要点を100文字程度で簡潔にまとめてください。"},
        {"role": "user", "content": input_text}
    ],
```

### 翻訳

```py
    messages=[
        {"role": "system", "content": "あなたは翻訳の専門家です。与えられたテキストを日本語から英語に正確に翻訳してください。"},
        {"role": "user", "content": "新しい技術を活用して、効率的に業務を改善する方法"}
    ],
```

### 文章の続きを生成

```py
    messages=[
        {"role": "system", "content": "あなたは物語の続きを描くプロの作家です。与えられた冒頭に合わせて、一貫した語り口で続きを創作してください。"},
        {"role": "user", "content": "彼は静かな湖のほとりに立ち、遠くの山々を見つめていた。水面は穏やかで、彼の心もそれに合わせるように次第に落ち着いていったが、ふと何かに気づいたように視線を移し、"}
    ],
```

### 文章の“間”を生成（ブリッジ）

```py
    messages=[
        {"role": "system", "content": "あなたは物語をつなぐ文章を創作するプロの作家です。以下の文章Aと文章Bの間に自然に繋がる中間の文章を生成してください。"},
        {"role": "user", "content": f"文章A: {before_text}　文章B: {after_text}"}
    ],
```

### コード生成（要件から新規作成）

```py
    messages=[
        {"role": "system", "content": "あなたは熟練したPythonプログラマーです。ユーザーからの要件に基づいて、効率的で正確なPythonコードを作成してください。ただし、コード部分のみを出力するものとします。"},
        {"role": "user", "content": input_text}
    ],
```

### リファクタリング（動作同じで可読性・効率改善）

```py
    messages=[
        {"role": "system", "content": "あなたは熟練したPythonプログラマーです。コードの動作を変えずに、可読性と実行効率を向上させるリファクタリングを行ってください。ただし、コード部分のみを出力するものとします。"},
        {"role": "user", "content": input_text}
    ],
```

### エラーハンドリング修正（IndexError回避）

```py
    messages=[
        {"role": "system", "content": "あなたは熟練したPythonプログラマーです。与えられたコードでは、リストのインデックスが範囲外になると「IndexError」が発生します。リストの長さを考慮してエラーが発生しないように修正してください。ただし、コード部分のみを出力するものとします。"},
        {"role": "user", "content": input_text}
    ],
```

## Chapter 6 Chat Completions APIの応用

### 概要

* **目的**：Chat Completions APIの“応用”——会話履歴を踏まえた応答、トークン節約、ストリーミング表示。
* **到達点**：ChatGPTライクな多ターン対話を実装しつつ、コスト（トークン）を管理し、UXを高める。

### 会話履歴を踏まえて回答させる

* **APIは履歴を自動保持しない**：毎回のリクエスト`messages`に、過去の`system / user / assistant`を自前で積む必要。
* **会話ループの基本形**

  1. `conversation_history`を用意（`system`は通常1つ）
  2. `while`でユーザー入力→履歴に`append`→API呼び出し→返答を履歴に`append`
  3. `"quit"`で終了

```py
from openai import OpenAI

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

# 最初にシステムメッセージ（AIの役割や回答条件）を設定
conversation_history = [
    {"role": "system", "content": "あなたは掃除に関するアドバイザーです。質問に対して200文字以内で回答してください。"}
]

while True:
    input_data = input("掃除に関する質問（quitで終了）: ")

    if input_data == "quit":
        break

    # 入力を会話履歴に追加（ユーザーからの発言として記録）
    conversation_history.append({"role": "user", "content": input_data})

    # Chat Completions APIにリクエストを送信
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=conversation_history, # これまでの会話履歴をすべて渡す
        temperature=0.5
    )

    # APIレスポンスからアシスタントの回答テキストを取り出す
    result = completion.choices[0].message.content

    # 生成された回答も会話履歴に追加（次回以降の文脈として利用）
    conversation_history.append({"role": "assistant", "content": result})

    # 回答を出力
    print(result)
    print("=======================")  # 区切り線を表示
```

### トークン数を節約する方法

* **トークン数の事前計測**：`tiktoken`でテキストをトークン化し、**履歴の合計トークン**を管理可能。
* **トークン節約の定石**

  * 古い履歴を削除（`system`は残し、古い`user/assistant`から間引く）
  * 古い履歴を**要約**して短縮
  * 関連の高い履歴のみ抽出して同梱（検索型の圧縮）
  * `max_tokens`は**出力**制御であり、**入力**は`tiktoken`などで管理する。

```py
from openai import OpenAI
import os
import tiktoken  # トークン数を数えるためのライブラリを利用

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

# tiktokenでエンコーディングを取得（cl100k_baseはGPT-3.5/4系列で使用するトークナイザー）
enc = tiktoken.get_encoding("cl100k_base")

# 会話履歴に含める最大トークン数（入力＋出力の合計）
MAX_ALLOWED_TOKENS = 300

# 会話全体で使用しているトークン数を管理する変数
total_tokens = 0

conversation_history = [
    {"role": "system", "content": "あなたは掃除に関するアドバイザーです。質問に対して200文字以内で回答してください。"}
]

while True:
    input_data = input("掃除に関する質問（quitで終了）: ")

    if input_data == "quit":
        break

    # 入力テキストのトークン数をカウント
    input_tokens = len(enc.encode(input_data))

    # 入力だけで上限を超える場合は再入力を促す
    if input_tokens > MAX_ALLOWED_TOKENS:
        print(f"入力する文字列は{MAX_ALLOWED_TOKENS}以下にしてください。")
        continue

    # 入力トークン数を合計に加算
    total_tokens += input_tokens

    conversation_history.append({"role": "user", "content": input_data})

    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=conversation_history,
        temperature=0.5
    )

    # AIの回答を取得
    result = completion.choices[0].message.content
    conversation_history.append({"role": "assistant", "content": result})

    # 回答のトークン数をカウントし合計に加算
    response_tokens = len(enc.encode(result))
    total_tokens += response_tokens

    # 現在のトークン数を表示
    print(f"現在の会話履歴のトークン数: {total_tokens}")

    # フラグ（削除処理を行ったかどうかの記録）
    remove_flg = False

    # トークン数が上限を超えている間は古い履歴を削除
    # pop(1)とすることで「最も古いユーザー/アシスタントの発言」を削除
    # （index=0はsystemメッセージなので残す）
    while total_tokens > MAX_ALLOWED_TOKENS:
        removed_message = conversation_history.pop(1)  # 履歴の2番目を削除
        removed_tokens = len(enc.encode(removed_message['content']))  # 削除分のトークン数を取得
        total_tokens -= removed_tokens  # 合計から差し引き
        remove_flg = True

    # 削除処理が行われた場合、削除後のトークン数を表示
    if remove_flg:
        print(f"削除処理後の会話履歴のトークン数: {total_tokens}")

    # AIの回答を表示
    print(result)
    print("=======================")  # 区切り線
```

### トークン数を節約する方法 要約バージョン

上限超過時に古い数ターン分をひとかたまりで要約→要約1メッセージに置換→トークン再計算

使い方とポイント
- 超過時の挙動：古い2往復（最大4メッセージ）をひとつの要約メッセージに置換。必要なら複数回繰り返して上限内に収めます。
- 文脈の温存：systemは必ず保持、直近の発言は極力そのまま残す設計にしています。
- コスト最適化：単純削除よりも文脈を保ちやすく、回答品質を下げにくいのが利点です。

```py
from openai import OpenAI
import os
import tiktoken  # トークン数カウント用

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

# GPT-3.5/4系で使用するトークナイザー
enc = tiktoken.get_encoding("cl100k_base")

# 会話履歴に含める最大トークン数（入力＋出力の合計）
MAX_ALLOWED_TOKENS = 300

# 会話全体のトークン数を管理
total_tokens = 0

conversation_history = [
    {"role": "system", "content": "あなたは掃除に関するアドバイザーです。質問に対して200文字以内で回答してください。"}
]

def count_history_tokens(messages) -> int:
    """会話全体のトークン数を概算（各contentだけを合算）。
       厳密にはroleや枠組み分もありますが、簡易な上限管理として十分です。"""
    return sum(len(enc.encode(m["content"])) for m in messages)

def format_chunk_for_summary(messages_chunk) -> str:
    """ユーザー/アシスタントの順で、要約しやすいテキストに整形"""
    lines = []
    for m in messages_chunk:
        if m["role"] == "user":
            lines.append(f"User: {m['content']}")
        elif m["role"] == "assistant":
            lines.append(f"Assistant: {m['content']}")
        else:
            # systemは通常要約対象に含めない
            continue
    return "\n".join(lines)

def summarize_messages(client: OpenAI, messages_chunk) -> str:
    """古い会話の塊を短く要約する。
       - 重要な指示・固有名詞・数値は保持
       - 200字目安で簡潔に
    """
    prompt_user_text = format_chunk_for_summary(messages_chunk)
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "system",
                "content": (
                    "あなたは優秀な会話要約者です。重要な事実・指示・数値・固有名詞を残しつつ、"
                    "後続の回答生成に必要な文脈が損なわれないように、日本語で200字程度に要約してください。"
                ),
            },
            {
                "role": "user",
                "content": (
                    "以下の会話ログを要約してください。形式はプレーンテキストで構いません。\n\n"
                    + prompt_user_text
                ),
            },
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()

while True:
    input_data = input("掃除に関する質問（quitで終了）: ")

    if input_data == "quit":
        break

    # 入力テキストのトークン数をカウント
    input_tokens = len(enc.encode(input_data))

    # 入力だけで上限を超える場合は再入力を促す
    if input_tokens > MAX_ALLOWED_TOKENS:
        print(f"入力する文字列は{MAX_ALLOWED_TOKENS}以下にしてください。")
        continue

    # 入力トークン数を合計に加算
    total_tokens += input_tokens

    # ユーザー発言を履歴に追加
    conversation_history.append({"role": "user", "content": input_data})

    # LLM呼び出し
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=conversation_history,
        temperature=0.5
    )

    # 回答を履歴に追加
    result = completion.choices[0].message.content
    conversation_history.append({"role": "assistant", "content": result})

    # 入出力を加味してトークン再計算（簡易）
    total_tokens = count_history_tokens(conversation_history)

    print(f"現在の会話履歴のトークン数: {total_tokens}")

    # ======== ここから「要約による圧縮」処理 ========
    # 上限を超えたら、古いターンの塊を要約→置換→再計算、を繰り返す
    summarized_any = False
    # system(0) は残し、直近の文脈は残すため最低でも末尾2メッセージ(user/assistant)は保持目安
    while total_tokens > MAX_ALLOWED_TOKENS and len(conversation_history) > 3:
        # 古い方から数ターン分を塊として抽出（例：最大4メッセージ＝2往復分）
        # 対象は index=1 以降（systemを避ける）
        # 末尾は直近の文脈保護のため取り除く（-2 くらいまで）
        start_idx = 1
        end_idx = min(len(conversation_history) - 2, start_idx + 4)  # 上限: 1..4件だけを要約
        if end_idx <= start_idx:
            # 要約できる十分な塊がなければ中断（安全弁）
            break

        chunk = conversation_history[start_idx:end_idx]

        # 要約を生成
        summary_text = summarize_messages(client, chunk)

        # 要約を会話に差し戻し（assistantロールの「要約メッセージ」に置換）
        # 置換＝古い塊を削除し、その位置に1件の要約メッセージを挿入
        del conversation_history[start_idx:end_idx]
        conversation_history.insert(
            start_idx,
            {"role": "assistant", "content": f"【これまでの要約】\n{summary_text}"}
        )

        # トークンを再計算
        total_tokens = count_history_tokens(conversation_history)
        summarized_any = True

    if summarized_any:
        print(f"要約圧縮後の会話履歴のトークン数: {total_tokens}")

    # 応答表示
    print(result)
    print("=======================")

```

### ストリーミングで回答を得る方法

* **ユーザー体験を優先する対話型アプリ**（チャット、UI付きサービス、音声読み上げ） → **ストリーミングが有利**
* **完全な構造化出力が必要なバッチ処理やAPI連携**（JSON出力、整形レポート生成） → **通常の同期出力が安定**

👉 実務的には、**チャットUIではストリーミング、裏側の処理では通常出力**を組み合わせるのが最適解になることが多い。

#### ✅ メリット

1. **ユーザー体験の向上**

   * 回答が一気に出るのを待たず、ChatGPTのように「少しずつ文字が出てくる」ため、待ち時間のストレスを減らせる。
   * 長文の回答でも「すぐに出始める」ので、ユーザーは安心感を得られる。

2. **応答の即時利用**

   * 途中までの回答でもユーザー側で処理可能（例：部分的にUIに表示、音声読み上げを開始、チャット画面にタイプ中のように表示）。

3. **リソース効率（UX面）**

   * 「大きな回答が完成するまで何も表示されない」よりも、リアルタイムで情報が得られるのでユーザーが離脱しにくい。

#### ❌ デメリット

1. **実装の複雑化**

   * ストリームから逐次チャンクを受け取り、整形して画面に出す処理が必要。
   * 途中でエラーや接続切断があった場合、半端な出力になる可能性がある。

2. **部分的にしか利用できない**

   * 「最終的な完成した回答テキスト」が揃う前に処理を走らせると、後で整合性を取る必要が出る（例：JSONや表形式の出力を期待する場合、ストリーミングだとパースできないことがある）。

3. **制御が難しい**

   * 完全なレスポンスを受け取るまで「トークン数の合計」「finish\_reason」などのメタ情報が確定しない。
   * そのため、料金計算やログ記録はストリーム終了後にまとめて行う必要がある。


### ストリーミングで回答を得る方法 - 使い分けのガイド

- チャットUI/タイプ表示・音声読み上げ：ストリーミング
- 厳密なJSON/CSV/Markdown表の生成・自動処理：非ストリーミング
- ハイブリッド：ストリーミングで表示しつつバッファを集めて最後に検証（実務でよくやる方式）


| 観点                 | ストリーミング（stream=True）      | 非ストリーミング（通常レスポンス） |
| :------------------- | :--------------------------------- | :--------------------------------- |
| 体感速度/UX          | ◎：すぐ出始める／離脱防止          | △：出るまで待つ                    |
| 実装の楽さ           | △：逐次受信・描画の処理が必要      | ◎：`content`を一発取得             |
| エラー時の扱い       | △：途中で途切れる可能性            | ○：原子的に結果が来る              |
| 構造化出力（JSON等） | △：まず全文をバッファ収集が必要    | ◎：そのままパース可能              |
| ログ/請求算定        | △：最後に集計                      | ◎：レスポンスにまとまる            |
| 向く用途             | チャットUI/音声読み上げ/ライブ出力 | レポート生成/バッチ/厳密パース     |


#### 1) ストリーミング版 超シンプル版（まずは挙動確認）
```py
from openai import OpenAI
import os

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

stream = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "あなたは気の利くアシスタントです。"},
        {"role": "user", "content": "ヘルシーで簡単な朝食を3つ、理由付きで教えて。"}
    ],
    temperature=0.5,
    stream=True
)

for chunk in stream:
    choice = chunk.choices[0]
    # 最終チャンクはfinish_reason=stop 等になるので非表示
    if choice.finish_reason is None and choice.delta and choice.delta.content:
        print(choice.delta.content, end="", flush=True)
```

#### 1) ストリーミング版 実務向け（全文バッファ→必要ならパース）

> ポイント：ライブ表示と同時にバッファし、最後に厳密処理（JSONパース等）を行うのが定石。

```py
from openai import OpenAI
import os, json

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

def stream_completion(messages):
    buf = []
    stream = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=messages,
        temperature=0.2,
        stream=True
    )
    for chunk in stream:
        ch = chunk.choices[0]
        if ch.finish_reason is None:
            part = (ch.delta.content or "")
            buf.append(part)
            print(part, end="", flush=True)  # 画面にも出す
    print()  # 改行
    return "".join(buf)

messages = [
    {"role": "system", "content": "あなたはJSONで厳密に出力するアシスタントです。"},
    {"role": "user", "content": "朝食アイデアをJSON配列([{name,why}...])で出力して"}
]

text = stream_completion(messages)

# 逐次表示しつつ、最後に厳密処理（構造化が必要ならここでパース）
try:
    data = json.loads(text)
    print(f"\nJSON items: {len(data)}")
except json.JSONDecodeError:
    print("\n⚠ JSONとして不完全。プロンプトを厳格化/再試行してください。")
```

#### 2) 非ストリーミング版 超シンプル版（最短）
```py
from openai import OpenAI
import os

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

resp = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "あなたは気の利くアシスタントです。"},
        {"role": "user", "content": "ヘルシーで簡単な朝食を3つ、理由付きで教えて。"}
    ],
    temperature=0.5
)

print(resp.choices[0].message.content)
```

#### 2) 非ストリーミング版 実務向け（構造化出力＋検証)
```py
from openai import OpenAI
import os, json

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

resp = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "あなたはJSONで厳密に出力するアシスタントです。"},
        {"role": "user", "content": "朝食アイデアをJSON配列([{name,why}...])だけで出力して"}
    ],
    temperature=0.2
)

text = resp.choices[0].message.content
print(text)

# 構造化チェック
try:
    data = json.loads(text)
    assert isinstance(data, list), "JSONは配列であるべきです"
    for item in data:
        assert "name" in item and "why" in item
    print(f"\nOK: {len(data)}件のアイテムを検証しました。")
except Exception as e:
    print(f"\n⚠ 構造化エラー: {e}")
```


#### 「stream=True を集約して非ストリーミングと同等に扱える」ラッパ関数

ポイント
- stream_then_aggregate が 逐次受信 → 文字列結合 → ChatCompletion 風のオブジェクトに整形します。
- usage はストリームでは取れないことが多いので None。コスト見積りが必要なら tiktoken で概算してください。
- 必要に応じて raw_chunks にチャンクを保持しておくと、デバッグや再構成に便利です。

```py
from openai import OpenAI
import os
from dataclasses import dataclass
from typing import List, Optional, Any, Dict

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

# --- 便宜的な「完成レスポンス風」データ構造 -----------------------------

@dataclass
class Message:
    role: str
    content: str

@dataclass
class Choice:
    index: int
    message: Message
    finish_reason: Optional[str] = None

@dataclass
class ChatCompletionLike:
    id: Optional[str]
    model: str
    created: Optional[int]
    choices: List[Choice]
    usage: Optional[Dict[str, int]] = None
    raw_chunks: Optional[List[Any]] = None  # 必要ならチャンク原本も保持

# --- ストリーミング → 結合 → ChatCompletionLike に整形 --------------------

def stream_then_aggregate(
    client: OpenAI,
    *,
    model: str,
    messages: List[Dict[str, str]],
    temperature: float = 0.2,
    **kwargs
) -> ChatCompletionLike:
    """
    stream=True で受け取ったチャンクを全部つなげ、
    非ストリーミングの ChatCompletion に近い形 (ChatCompletionLike) で返す。
    """
    text_fragments: List[str] = []
    collected_chunks = []
    last_chunk_meta = {"id": None, "created": None, "model": model, "finish_reason": None}

    stream = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=temperature,
        stream=True,
        **kwargs
    )

    for chunk in stream:
        collected_chunks.append(chunk)
        ch = chunk.choices[0]

        # メタ情報を覚えておく（最後のチャンクの値でよい）
        last_chunk_meta["id"] = getattr(chunk, "id", last_chunk_meta["id"])
        last_chunk_meta["created"] = getattr(chunk, "created", last_chunk_meta["created"])
        last_chunk_meta["model"] = getattr(chunk, "model", last_chunk_meta["model"])

        if ch.finish_reason is None:
            part = (ch.delta.content or "")
            text_fragments.append(part)
        else:
            last_chunk_meta["finish_reason"] = ch.finish_reason

    full_text = "".join(text_fragments)

    # usage はストリーミング中は逐次は来ないことが多いので None のままにする
    # 必要なら後段で tiktoken などで概算可
    choice = Choice(
        index=0,
        message=Message(role="assistant", content=full_text),
        finish_reason=last_chunk_meta["finish_reason"]
    )

    return ChatCompletionLike(
        id=last_chunk_meta["id"],
        model=last_chunk_meta["model"],
        created=last_chunk_meta["created"],
        choices=[choice],
        usage=None,
        raw_chunks=collected_chunks
    )

# --- 使い方（例） ----------------------------------------------------------

messages = [
    {"role": "system", "content": "あなたは気の利くアシスタントです。"},
    {"role": "user", "content": "ヘルシーで簡単な朝食を3つ、理由付きで教えて。"}
]

# 1) 非ストリーミング
completion = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=messages,
    temperature=0.5
)
print("非ストリーミング:", completion.choices[0].message.content)

# 2) ストリーミング → 結合（非ストリーミング風に扱える）
agg = stream_then_aggregate(client, model="gpt-4o-mini", messages=messages, temperature=0.5)
print("ストリーミング結合:", agg.choices[0].message.content)

# 以降は completion と agg をほぼ同様に扱える（usage などは差異あり）

```

#### Class Diagram

- 非ストリーミングは 完成形 ChatCompletion を一発で取得。
- ストリーミングは ChatCompletionStreamIterator が ChatCompletionChunk を逐次返す構造。
- チャンクは choices[].delta.content に 差分テキストを持ち、finish_reason により終了がわかります。

```plantuml
@startuml

title ChatCompletion (非ストリーミング)

class ChatCompletion {
  +id: String
  +object: String
  +created: Number
  +model: String
  +choices: List<Choice>
  +usage: Usage
}

class Choice {
  +index: Number
  +message: Message
  +finish_reason: String
}

class Message {
  +role: String
  +content: String
}

class Usage {
  +prompt_tokens: Number
  +completion_tokens: Number
  +total_tokens: Number
}

ChatCompletion "1" *-- "many" Choice
Choice "1" *-- "1" Message
ChatCompletion "1" *-- "1" Usage
@enduml
```

title ChatCompletion (非ストリーミング)
```plantuml
@startuml
title ChatCompletion Streaming (ストリーミング: イテレーター + ChatCompletionChunk)

class ChatCompletionStreamIterator {
  +__iter__(): ChatCompletionStreamIterator
  +__next__(): ChatCompletionChunk
}

class ChatCompletionChunk {
  +id: String
  +object: String           << "chat.completion.chunk" >>
  +created: Number
  +model: String
  +choices: List<ChunkChoice>
}

class ChunkChoice {
  +index: Number
  +delta: Delta
  +finish_reason: String?   // None か "stop" 等
}

class Delta {
  +role: String?            // 先頭チャンク等で付く場合あり
  +content: String?         // 差分テキスト（部分）
}

ChatCompletionStreamIterator --> ChatCompletionChunk : yields
ChatCompletionChunk "1" *-- "many" ChunkChoice
ChunkChoice "1" *-- "1" Delta
@enduml

```

### 実装の型（要点だけ）

* **履歴管理**

  * `conversation_history = [{"role":"system","content":"…"}]`
  * ループで`input()`→`messages=conversation_history`でAPI→返答を履歴に追加。
* **事前トークン計測と上限**

  * `enc = tiktoken.get_encoding("cl100k_base")`
  * 入出力を都度`len(enc.encode(text))`で加算し、上限超過なら**古いメッセージを削除**（`pop(1)`でsystem以外から）。
* **ストリーミング描画**

  * `stream=True`で呼び、`for chunk in completion:`内で`delta.content`を随時`print(end="")`。

### 実務Tips

* **冗長な二重実装を避ける**：履歴保持と入出力、削除/要約を**関数化**して再利用。
* **コスト見積り**：履歴・入出力のトークンを合算して**料金試算**→上限や圧縮方針を決める。
* **削除より要約優先**：文脈維持が必要な場面では古い履歴を短く**要約**して残す。
* **Memory/RAGの活用**（後続章）：履歴の保存・検索を仕組み化すると、精度とコストのバランスが取りやすい。

### 学習のゴール

* 自前履歴で**多ターン対話**を実装できる。
* `tiktoken`で**入力側のトークン管理**と**コスト最適化**ができる。
* `stream=True`で**タイピング風の逐次表示**を提供できる。


## Chapter 7 テキスト以外のコンテンツ生成（学習時間目安：0.5時間）

OpenAI APIを使った **画像生成** と **音声生成** の基本を解説。

### 画像生成（DALL-E）

* **テキストから画像生成**

  * `client.images.generate()` を利用
  * パラメータ：`model`（dall-e-2/3）、`prompt`、`size`、`quality`、`style`
  * DALL-E 3 → 高品質、1枚のみ
  * DALL-E 2 → 安価・最大10枚生成可（ただし quality/style 指定不可）

```py
from openai import OpenAI

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

response = client.images.generate(
    model="dall-e-3",
    prompt="恐竜と人間が住む世界",
    size="1024x1024",
    quality="hd",
    style="vivid"
)

# 画像URLを取得
image_url = response.data[0].url

print(image_url)
```

> https://oaidalleapiprodscus.blob.core.windows.net/private/org-UW1pAwKfpOoxwIoPDFpHoRBH/user-BHljUAN1FM5GDt5aHD89rmz3/img-PjXUfuwZtTIKNFgziUP3ufm9.png?st=2024-11-24T19%3A44%3A41Z&se=2024-11-24T21%3A44%3A41Z&sp=r&sv=2024-08-04&sr=b&rscd=inline&rsct=image/png&skoid=d505667d-d6c1-4a0a-bac7-5c84a87759f8&sktid=a48cca56-e6da-484e-a814-9c849652bcb3&skt=2024-11-24T00%3A12%3A41Z&ske=2024-11-25T00%3A12%3A41Z&sks=b&skv=2024-08-04&sig=2joKheLaPYdefnLhtVpBU8IEyS1RyJxIT6y5We3o9ls%3D


* **画像から似た画像生成**

  * `client.images.create_variation()` を利用
  * `image=open("original.png","rb")` で入力画像を渡す
  * モデルは DALL-E 2 のみ対応

### OpenAI API テキスト以外のコンテンツ生成 比較表

| 項目           | 画像生成（テキスト入力）                                                                                                                       | 画像生成（画像入力：バリエーション）            | 音声生成（Text-to-Speech）                           |
| :------------- | :--------------------------------------------------------------------------------------------------------------------------------------------- | :---------------------------------------------- | ---------------------------------------------------- |
| 使用メソッド   | `client.images.generate()`                                                                                                                     | `client.images.create_variation()`              | `client.audio.speech.create()`                       |
| 主な入力       | `prompt`（テキスト）                                                                                                                           | `image=open("xxx.png","rb")`                    | `input`（テキスト）                                  |
| 利用モデル     | DALL-E 3（高品質・1枚のみ）<br>DALL-E 2（安価・最大10枚）                                                                                      | DALL-E 2のみ                                    | tts-1（標準）<br>tts-1-hd（高品質）                  |
| 出力内容       | 新規生成画像                                                                                                                                   | 入力画像に似た画像                              | 音声ファイル（例：mp3）                              |
| 主なパラメータ | - `size`（1024x1024 / 1024x1792 / 1792x1024）<br>- `quality`（standard(標準) / hd(詳細を強調)）<br>- `style`（natural(自然) / vivid(鮮やか)）<br>- `n`（DALL-E 2のみ複数生成） | - `size`（1024x1024 など）<br>- `n`（生成枚数） | - `voice`（alloy, echo, fable, onyx, nova, shimmer） |
| 特徴           | 高品質 or 安価を選択可能                                                                                                                       | 元画像に基づく類似画像を生成                    | テキストを自然な音声に変換                           |
| 出力保存       | `response.data[0].url` から画像URL取得                                                                                                         | `response.data[0].url` から画像URL取得          | `response.stream_to_file("voice.mp3")`               |


### DALL-E 3 と DALL-E 2 の比較表

| 項目               | **DALL-E 3**                                                    | **DALL-E 2**                                                               |
| :----------------- | :-------------------------------------------------------------- | :------------------------------------------------------------------------- |
| メソッド           | `client.images.generate()`                                      | `client.images.generate()` / `client.images.create_variation()`            |
| 生成枚数 (`n`)     | **不可**（1枚のみ生成）                                         | **可能**（最大10枚まで生成可）                                             |
| サイズ (`size`)    | 1024×1024<br>1024×1792<br>1792×1024                             | 1024×1024 など（固定的）                                                   |
| 品質 (`quality`)   | 指定可能：`standard` または `hd`                                | **不可**                                                                   |
| スタイル (`style`) | 指定可能：`natural` または `vivid`                              | **不可**                                                                   |
| バリエーション生成 | **不可**                                                        | `create_variation()` で可能                                                |
| 価格感             | 高め                                                            | 安価                                                                       |
| 特徴               | - 高品質な画像生成<br>- 枚数は制限（1枚）<br>- 表現が自然で詳細 | - 安価で枚数を多く生成可能<br>- バリエーション作成対応<br>- 品質はやや劣る |

✅ まとめると：

* **高品質が欲しい → DALL-E 3**
* **コストを抑えたい / 複数枚 / バリエーションが欲しい → DALL-E 2**

### 音声生成（Text-to-Speech）

```py
from openai import OpenAI

client = OpenAI(api_key=os.environ["OPENAI_API_KEY"])

prompt = "光の速さは約30万キロメートル毎秒です。この速度で地球を一周すると、わずか0.13秒しかかかりません。"

response = client.audio.speech.create(
    model="tts-1",
    voice="shimmer",
    input=prompt
)

response.stream_to_file("voice.mp3")
```

* `client.audio.speech.create()` を利用
* パラメータ：`model`（tts-1 / tts-1-hd）、`voice`（6種類）、`input`（テキスト）
* 出力は mp3 などの音声ファイル
* `tts-1` は安価、`tts-1-hd` は高品質だが料金高め
* `response.stream_to_file("voice.mp3")` で保存可能

「voice」は声の種類
「alloy」「echo」「fable」「onyx」「nova」「shimmer」のうちいずれかを選ぶ。
各種類の声質の違いについては、[OpenAI公式サイト](https://platform.openai.com/docs/guides/text-to-speech)から確認。


### Text-to-Speech モデル比較表（tts-1 / tts-1-hd）

| 項目               | **tts-1**                                                  | **tts-1-hd**                   |
| :----------------- | :--------------------------------------------------------- | :----------------------------- |
| メソッド           | `client.audio.speech.create()`                             | `client.audio.speech.create()` |
| 入力 (`input`)     | テキスト（任意の文章）                                     | テキスト（任意の文章）         |
| 音声保存           | `response.stream_to_file("xxx.mp3")` などで保存            | 同左                           |
| 出力形式           | mp3 など（指定可能）                                       | mp3 など（指定可能）           |
| 声の種類 (`voice`) | 6種類：`alloy`, `echo`, `fable`, `onyx`, `nova`, `shimmer` | 同左                           |
| 品質               | 標準的                                                     | 高品質（より自然・滑らか）     |
| 料金（目安）       | **約 \$0.015 / 1,000 文字**                                | **約 \$0.030 / 1,000 文字**    |
| 利用用途           | コストを抑えたい場合、テスト用                             | 高品質な音声が必要な本番環境   |


✅ ポイントまとめ

* **コスト重視** → `tts-1`
* **品質重視** → `tts-1-hd`
* どちらも **声の種類は同じ6種類** を選択可能



### ポイント

* API呼び出しメソッドはタスクごとに異なる

  * Chat → `chat.completions.create()`
  * 画像生成 → `images.generate()`
  * 画像バリエーション → `images.create_variation()`
  * 音声生成 → `audio.speech.create()`
* モデルやパラメータを調整することで、品質・料金・用途を使い分け可能。

