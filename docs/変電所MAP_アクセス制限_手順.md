# 変電所MAP アクセス制限（許可した人だけが見られるようにする）

最終更新：2026-11-20

この文書は、いま誰でも見られる状態の変電所MAP（および DDチェック・条例リスクDB）を、
**余野さんが許可したメールアドレスの人だけが見られる状態**に変えるための手順書です。
プログラミングの知識は要りません。ブラウザの操作だけで終わります。

---

## 1. いまの状態と、何が問題か

| | いまの状態 |
|---|---|
| 公開先 | `https://yonomotozumi-spec.github.io/battery-storage-dd/`（GitHub Pages） |
| 見られる人 | **URLを知っている人は誰でも。ログインなし** |
| リポジトリ | `public`（GitHubの画面で誰でもファイルを読める） |

2026-11-20 に外部から確認したところ、公開中の地図データに**接続検討回答の実績**
（案件の所在地・申込者の社名・負担金・工期）がそのまま含まれていました。
回答書には目的外使用・第三者開示の禁止が明記されているため、同日、公開版のデータ・画面・
Excel・CSVから実績を外しました（実績は社内版のみ）。

ただし、これで直ったのは「中身」だけです。次の2つはまだ残っています。

1. **地図そのものは誰でも見られる**（変電所の空容量・判定・スコアが全部見える）
2. **GitHubの履歴には過去の実績入りファイルが残っている**（公開リポジトリなので誰でも遡れる）

この手順書は、この2つを閉じるためのものです。

---

## 2. 完成後の姿

```
  外部事業者・PowerX社内の人
        │
        │ ① 地図のURLを開く
        ▼
  ┌─────────────────────────┐
  │ ログイン画面（Cloudflare）  │  メールアドレスを入力 → 6桁のコードがメールで届く
  │                         │  → コードを入れると地図が開く
  └─────────────────────────┘
        │
        │ ② 余野さんが許可したメールアドレスだけ通る
        ▼
  変電所MAP（公表データのみ。実績は入っていない）
```

- **許可制**：余野さんが管理画面にメールアドレスを1行ずつ登録します。登録していない人は、
  コードが届かないので入れません（画面の見た目は同じで、不正な人に「登録が無い」と教えません）。
- **外部事業者はアカウント作成不要**：メールが受け取れればそれだけで入れます。
- **誰がいつ見たかのログが残ります**（アプリ名・メールアドレス・IP・国・許可/拒否・時刻）。
- 公開は GitHub Pages から **Cloudflare Pages** に引っ越します。GitHub Pages は止めます。

### 実績（接続検討回答）の扱い

外部事業者にも地図を見せる以上、**実績は地図に戻しません**。回答書の第三者開示にあたるためです。
実績を見たいときは、リポジトリの外に置いた実績JSONから社内版を作って使います（README「接続検討回答の実績は社内版だけに入れる」）。

```bash
python scripts/build_internal_map.py --results <外>/実績DB.json --ledger <外>/回答書取込台帳.csv --out <外>/変電所マップ_社内版.html
```

将来、社内版も Web で見たくなった場合は、同じ Cloudflare に「社内の人だけ」の別アプリとして
追加できます（本書 7章）。

---

## 3. 用意するもの・費用

| | |
|---|---|
| 必要なもの | Cloudflare の無料アカウント（メールアドレスだけで作れます）、GitHubのログイン |
| 独自ドメイン | **不要**。`変電所MAP.pages.dev` のような Cloudflare のURLが無料でもらえます |
| 費用 | Cloudflare Zero Trust には Free プランがあります。**ただし「無料で何人まで使えるか」「ログが何日残るか」は公式ドキュメントに数字の記載が見つかりませんでした。** 契約前に Cloudflare の管理画面の表示で確認してください（本書 8章） |
| 所要時間 | 30〜60分程度 |

---

## 4. 手順A：Cloudflare に地図を置く

1. <https://dash.cloudflare.com/sign-up> でアカウントを作る（メールアドレスとパスワード）。
2. 左メニューの **Workers & Pages** → **Create** → **Pages** タブ →
   **Connect to Git**。
3. GitHub との接続を許可し、リポジトリ `yonomotozumi-spec/battery-storage-dd` を選ぶ。
   - Cloudflare Pages は**プライベートリポジトリにも対応**しています（公式ドキュメントに明記）。
     後でリポジトリを非公開にしても、ここは動き続けます。
4. ビルドの設定は**何も入れない**でそのまま進める（このリポジトリはHTMLを置くだけで動くため）。
   - Production branch: `main`
   - Build command: 空欄
   - Build output directory: 空欄（または `/`）
5. **Save and Deploy**。数分で `https://<プロジェクト名>.pages.dev` が出来ます。
6. そのURLを開いて、地図がいつもどおり表示されることを確認する。

> この時点ではまだ**誰でも見られます**。次の手順Bで閉じます。

---

## 5. 手順B：メール確認コードのログインを付ける

### B-1. ログイン方法に「One-time PIN」を足す

1. Cloudflare の左メニューから **Zero Trust** に入る（初回はチーム名の設定と、プランの選択を求められます。Free を選ぶ）。
2. **Integrations → Identity providers** を開く。
3. **Add new identity provider** → **One-time PIN** を選んで保存。

これで「メールアドレスを入れると6桁のコードが届く」ログインが使えるようになります。

### B-2. 地図にログインを義務づける（本番URLまでかける）

Cloudflare Pages の「Enable access policy」ボタンは、**そのままだとプレビュー用URLしか守りません**
（公式ドキュメントに明記）。本番の `<プロジェクト名>.pages.dev` まで守るには、次の順で操作します。

1. **Workers & Pages** → 対象プロジェクト → **Settings** → **General** →
   **Enable access policy** を選ぶ。
2. 作られたポリシーの **Manage** を押す。
3. **Access → Applications** で、このプロジェクトを選び **Configure**。
4. **Public hostname** の **Subdomain** 欄にある `*`（ワイルドカード）を**消して保存**する。
   （アプリ名の変更を求められたら、分かる名前を付ける。例：`変電所MAP`）
5. もう一度、プロジェクトの **Settings → General → Enable access policy** を選ぶ。
6. ポリシーが**2つ**（`*.pages.dev` 用 と プレビュー用）できていることを確認する。

### B-3. 誰が入れるかを決める（＝許可制の本体）

1. **Zero Trust → Access → Applications** で、作ったアプリ（例：`変電所MAP`）を開く。
2. **Policies** → ポリシーを編集（または **Add a policy**）。
3. Action は **Allow**。
4. ルールを次のどちらかで作る。
   - **メールアドレスで1人ずつ**：Selector `Emails` → 許可する人のメールを1行ずつ入れる（**推奨**。
     外部事業者はこれ）
   - **会社のドメインごと**：Selector `Emails ending in` → `@powerx.jp` のように入れる
     （PowerX社内を一括で入れたい場合。ドメインは実際のものに合わせる）
5. 保存。

> **注意**：ポリシーを作らないと、ログイン画面は出ても誰も入れません（公式ドキュメントの注意書き）。
> 必ず Allow のポリシーを1つ以上作ってください。

---

## 6. 手順C：動作確認と、GitHub側の後始末

### C-1. 動作確認（この3つを必ずやる）

1. **自分のメールで入れるか**：シークレットウィンドウで `https://<プロジェクト名>.pages.dev` を開く
   → メールアドレスを入れる → 届いた6桁コードを入れる → 地図が開く。
2. **許可していないメールで入れないか**：許可リストに入れていない別のメールアドレスで試す
   → コードが届かない（＝入れない）ことを確認する。
3. **ログが残るか**：**Zero Trust → Insights → Logs → Access authentication logs** に、
   いま入ったログイン記録（メール・時刻・Connection=`onetimepin`・Allow）が出ることを確認する。

> コードが届かないときは、メールの迷惑メール設定を確認してください。送信元は
> `noreply@notify.cloudflare.com`（ドメイン `notify.cloudflare.com`）です。会社のメールで
> ブロックされる場合は、情報システム担当にこの送信元の許可を依頼します。
> コードの有効期限は10分、1回使うと無効になります。

### C-2. GitHub Pages を止める

Cloudflare 側がちゃんと動くことを確認**してから**にしてください。順番を逆にすると、しばらく地図が見られなくなります。

- GitHub のリポジトリ → **Settings → Pages** → Source を **None** にする。

### C-3. リポジトリを非公開（private）にする

履歴に残っている過去の実績入りファイルを外から読めなくするために行います。

- GitHub のリポジトリ → **Settings** → 一番下の **Danger Zone** →
  **Change repository visibility** → **Change visibility** → **Make this repository private**。

**先に知っておくこと（GitHub公式ドキュメントに明記）**

- private にすると、**GitHub Pages のサイトは自動的に公開停止**になります（C-2をやっていなくても止まります）。
- **スターとウォッチャーは消えます**。このリポジトリは現在スター0・フォーク0なので実害はありません。
- 既存のフォークがあると、そのフォークは public のまま残ります（＝履歴も残る）。
  **2026-11-20時点でこのリポジトリのフォークは0件**なので、この問題は起きません。
- GitHub Actions（毎月の座標更新・毎月のスコア元データ更新・毎週の条例監視）は private でも動きます。
  ただし private リポジトリの Actions は無料利用枠が有限（public は無制限）なので、
  使用量が増えたら GitHub の Settings → Billing で確認してください。

---

## 7. 日々の運用（権限者＝余野さんの操作）

### 人を増やす

**Zero Trust → Access → Applications → `変電所MAP` → Policies → 編集 → Emails にメールを追加 → 保存。**
相手に `https://<プロジェクト名>.pages.dev` を伝えるだけで、相手はメールのコードで入れます。

### 人を外す（契約終了・退職時）

同じ場所で、そのメールアドレスを**1行消して保存**。即座に入れなくなります。
パスワード共有と違い、他の人への影響はありません。

### 誰が見たかを確認する

**Zero Trust → Insights → Logs → Access authentication logs**。
アプリ名・メールアドレス・IP・国・許可/拒否・時刻が出ます。
（メールアドレスを入れただけでは記録されず、コードを入力した時点で記録されます）

### 地図を更新したとき

これまでどおり `main` ブランチに変更が入れば、Cloudflare Pages が自動で作り直します。
操作は不要です。

### 社内版（実績入り）も Web で見たくなったら

**Zero Trust → Access → Applications → Add an application → Self-hosted** で別アプリを作り、
社内版を置いたパス（例：`https://<プロジェクト名>.pages.dev/internal/`）を Public hostname にして、
**PowerX社内のメールだけ**を許可するポリシーを付けます。
外向けのアプリとはポリシーが別々になるので、外部事業者には実績が見えません。
（この場合、社内版のHTMLをリポジトリに置くことになるため、実績をリポジトリ外に保つ原則との
兼ね合いを余野さんと相談してから実施してください）

---

## 8. 公式ドキュメントで確認できなかったこと

次の3点は、Cloudflareの公開ドキュメントに数字の記載を見つけられませんでした。
**管理画面の表示を見て確認してください**。想定と違っていたら教えてください。

1. Zero Trust Free プランで、Access を使える**利用者数の上限**
   （「50」という数字はドキュメント上ではアプリ数の上限であって、人数ではありませんでした）
2. 認証ログの**保持期間**
3. Free プランで認証ログが見られるか（画面自体の有無）

---

## 9. この手順を選んだ理由（他の方法との比較）

| 方法 | 外部事業者 | 許可制 | 誰が見たかの記録 | 費用 | 判断 |
|---|---|---|---|---|---|
| **Cloudflare Access ＋ メール確認コード** | アカウント作成不要 | ○ メール1行で追加・削除 | ○ 残る | 無料（上限は要確認） | **採用** |
| Google/Microsoft の SSO | 相手の会社アカウントが要る | ○ ドメイン単位も可 | ○ | 無料 | 初期設定にIdP側の作業が要るため見送り。後から足せる |
| 共通パスワード1つ | 簡単 | × 全員同じ | × 残らない | 無料 | 回答書の開示管理には弱く、人が抜けるたび全員に配り直しになるため不採用 |
| GitHub Enterprise Cloud の private Pages | 相手もGitHubアカウントが要る | ○ | ○ | 1人あたり月額 | 社内の契約手続きが要るため不採用 |

---

## 10. 困ったときの連絡先・調べ先

- Cloudflare Pages と Access の公式手順：
  <https://developers.cloudflare.com/pages/configuration/preview-deployments/>（Access の有効化）
  <https://developers.cloudflare.com/pages/platform/known-issues/>（本番 `*.pages.dev` にかける手順）
- One-time PIN の仕様：<https://developers.cloudflare.com/cloudflare-one/identity/one-time-pin/>
- GitHub のリポジトリ公開設定：
  <https://docs.github.com/en/repositories/managing-your-repositorys-settings-and-features/managing-repository-settings/setting-repository-visibility>
