const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, VerticalAlign
} = require('docx');
const fs = require('fs');

const font = "MS Gothic";
const sz = 22;

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after !== undefined ? opts.after : 100 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, size: opts.size || sz, font, bold: opts.bold || false, color: opts.color || undefined })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text, size: sz, font })]
  });
}

function subbullet(text) {
  return new Paragraph({
    numbering: { reference: "subbullets", level: 0 },
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, size: 20, font, color: "444444" })]
  });
}

function h1(text) {
  return new Paragraph({
    spacing: { before: 360, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F5C9E", space: 1 } },
    children: [new TextRun({ text, size: 28, font, bold: true, color: "1F5C9E" })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, size: 24, font, bold: true })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text: "◆ " + text, size: sz, font, bold: true, color: "1F5C9E" })]
  });
}

function note(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 300 },
    children: [new TextRun({ text: "※ " + text, size: 19, font, color: "666666" })]
  });
}

const borderGray = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const bordersGray = { top: borderGray, bottom: borderGray, left: borderGray, right: borderGray };
const borderBlue = { style: BorderStyle.SINGLE, size: 4, color: "1F5C9E" };
const bordersBlue = { top: borderBlue, bottom: borderBlue, left: borderBlue, right: borderBlue };

function cell(text, bg, bold = false, width = 4513, color) {
  const lines = text.split('\n');
  const children = [];
  lines.forEach((line, i) => {
    children.push(new Paragraph({
      spacing: { before: 0, after: i < lines.length - 1 ? 60 : 0 },
      children: [new TextRun({ text: line, size: sz, font, bold, color: color || (bg === "1F5C9E" ? "FFFFFF" : undefined) })]
    }));
  });
  return new TableCell({
    borders: bg === "1F5C9E" ? bordersBlue : bordersGray,
    width: { size: width, type: WidthType.DXA },
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    verticalAlign: VerticalAlign.TOP,
    children
  });
}

function headerRow(cells) {
  return new TableRow({ children: cells });
}

function dataRow(cells) {
  return new TableRow({ children: cells });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "・",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } }
        }]
      },
      {
        reference: "subbullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "－",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 240 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font, size: sz } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1701, right: 1701, bottom: 1701, left: 1701 }
      }
    },
    children: [

      // ══════════════════════════
      // タイトルブロック
      // ══════════════════════════
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "サイト制作に関するご要望書", bold: true, size: 36, font })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」補足資料", size: 20, font, color: "666666" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 500 },
        children: [new TextRun({ text: "2026年4月22日", size: 20, font, color: "666666" })]
      }),

      p("岡崎社長　御中", { after: 60 }),
      p("Desert Bamboo Global 日本事務局　怒田裕也", { align: AlignmentType.RIGHT, after: 500 }),

      // 前文
      p("先日ご提案いただきました「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」（2026年4月22日付）について、コンセプトおよびサイトの基本方針に賛同いたします。", { after: 80 }),
      p("つきましては、Phase 1（方向性確定）の議論に先立ち、本サイトを通じて達成すべきゴール・コンテンツ設計・技術基盤について、日本事務局側の要望および提案を具体的にまとめました。本書を叩き台として、制作・設計の方向性を合意できればと存じます。", { after: 80 }),
      p("なお、本要望は提案書を否定するものではなく、制作クオリティと集客効果を最大化するための補足として位置づけてください。", { after: 60 }),

      // ══════════════════════════
      h1("1．本サイトで達成したいゴール"),
      // ══════════════════════════

      p("サイトの成否を測る指標として、以下の定量・定性ゴールを設定することを提案します。制作会社への発注・内製判断においても、これらのゴールを前提として設計・見積もりを依頼いただけると幸いです。", { after: 100 }),

      h2("■ 定量ゴール（公開後6〜12ヶ月以内）"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 3013, 3013],
        rows: [
          headerRow([
            cell("指標", "1F5C9E", true, 3000),
            cell("6ヶ月目標", "1F5C9E", true, 3013),
            cell("12ヶ月目標", "1F5C9E", true, 3013),
          ]),
          dataRow([
            cell("月間オーガニックPV", "EBF3FB", false, 3000),
            cell("500〜1,000 PV", null, false, 3013),
            cell("3,000 PV以上", null, false, 3013),
          ]),
          dataRow([
            cell("会員申請数（月間）", "EBF3FB", false, 3000),
            cell("3〜5件", null, false, 3013),
            cell("10件以上", null, false, 3013),
          ]),
          dataRow([
            cell("商談化件数（累計）", "EBF3FB", false, 3000),
            cell("1〜2件", null, false, 3013),
            cell("5件以上", null, false, 3013),
          ]),
          dataRow([
            cell("AI検索での引用回数", "EBF3FB", false, 3000),
            cell("主要KWで引用確認", null, false, 3013),
            cell("月10回以上の引用", null, false, 3013),
          ]),
          dataRow([
            cell("Google検索順位", "EBF3FB", false, 3000),
            cell("主要KW 20位以内", null, false, 3013),
            cell("主要KW 10位以内", null, false, 3013),
          ]),
        ]
      }),
      p("", { after: 100 }),

      h2("■ 定性ゴール"),
      bullet("「サウジアラビア 建築・不動産・投資」の日本語情報領域において、Desert Bamboo Globalを唯一無二の一次情報源として確立する"),
      bullet("ChatGPT・Perplexity・Google SGE等のAI検索エンジンが、サウジアラビア関連の質問に対してDesert Bambooのコンテンツを引用・推薦する状態をつくる"),
      bullet("富裕層・機関投資家が「このサイトに書いてある情報は信頼できる」と判断し、会員登録・問い合わせへのハードルを自然に下げる設計にする"),
      bullet("将来的な英語・アラビア語展開の起点となる、拡張可能なコンテンツ資産を蓄積する"),

      // ══════════════════════════
      h1("2．狙うキーワード戦略（SEO / AEO）"),
      // ══════════════════════════

      p("日本語でのサウジアラビア関連キーワードは検索ボリューム自体は多くありませんが、競合コンテンツがほぼ存在しないため、少量の良質コンテンツで上位表示・AI引用を狙えるブルーオーシャン領域です。以下の優先度に従いコンテンツを配置してください。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 3000, 2313, 2313],
        rows: [
          headerRow([
            cell("優先度", "1F5C9E", true, 1400),
            cell("キーワード群", "1F5C9E", true, 3000),
            cell("想定ターゲット", "1F5C9E", true, 2313),
            cell("コンテンツ形式", "1F5C9E", true, 2313),
          ]),
          dataRow([
            cell("★★★\n最優先", "EBF3FB", true, 1400),
            cell("サウジアラビア 不動産 購入\nサウジアラビア 投資 日本人\nサウジアラビア 会社設立\nVision 2030 投資機会 日本", null, false, 3000),
            cell("富裕層個人投資家\n法人投資家\n進出検討企業", null, false, 2313),
            cell("解説記事\nFAQ\n事例紹介", null, false, 2313),
          ]),
          dataRow([
            cell("★★☆\n準優先", "EBF3FB", true, 1400),
            cell("NEOM 最新情報 日本語\nAMALA サウジアラビア\nジェッダ 観光 投資\nサウジアラビア ビザ 投資家", null, false, 3000),
            cell("中東ビジネス関心層\nメディア・研究者\n初期情報収集層", null, false, 2313),
            cell("ニュース記事\n動向レポート\n解説記事", null, false, 2313),
          ]),
          dataRow([
            cell("★☆☆\nブランドKW", "EBF3FB", true, 1400),
            cell("竹建築 中東\nサステナブル建築 サウジアラビア\nBamboo Lab サウジアラビア\nVision 2030 建築デザイン", null, false, 3000),
            cell("建築・デザイン業界\n協業検討先\nDesert Bamboo指名検索", null, false, 2313),
            cell("実績紹介\nプロジェクト解説\nコンセプト記事", null, false, 2313),
          ]),
        ]
      }),
      p("", { after: 100 }),

      h2("■ AEO（AI検索最適化）対応方針"),
      p("Google・ChatGPT・Perplexity等のAI検索では、「日本語で書かれた専門性の高い一次情報」が引用されやすい傾向があります。以下の方針でコンテンツを整備してください。", { after: 80 }),

      h3("FAQコンテンツの大量設置"),
      bullet("公開ページにFAQを最低20問以上設置。各回答は300〜500字の実質的な内容とする"),
      bullet("質問形式の例："),
      subbullet("「サウジアラビアで外国人が不動産を購入できますか？」"),
      subbullet("「Vision 2030とは何ですか？日本企業にどんなメリットがありますか？」"),
      subbullet("「NEOMプロジェクトの現在の進捗状況は？」"),
      subbullet("「サウジアラビアへの投資リスクにはどんなものがありますか？」"),
      subbullet("「竹建築はサウジアラビアの気候環境に対応できますか？」"),
      note("FAQは「会員制ページへの誘導フック」としても機能する。概要をFAQで公開し、詳細は会員限定とする二段構えも有効"),

      h3("構造化データ・E-E-A-T対策"),
      bullet("著者プロフィール（専門家としての岡崎社長・日本事務局）を明示し、AI検索の「信頼性シグナル」を強化する"),
      bullet("各記事に「最終更新日」「情報ソース（MISA・PIF・Vision 2030公式等）」を明記する"),
      bullet("JSON-LDによるFAQSchemaをページに埋め込み、Google検索・AI検索の構造化理解を促進する"),

      h3("更新頻度の担保"),
      bullet("公開コンテンツは月2〜4本ペースで更新。更新が止まるとAI検索での引用優先度が下がるため、運用体制を明確にしてください"),
      bullet("日本事務局側でコラム・FAQ原稿の叩き台を作成し、岡崎社長側で一次情報を加筆する分業体制を提案します"),

      // ══════════════════════════
      h1("3．公開コンテンツ・非公開コンテンツの設計"),
      // ══════════════════════════

      p("提案書では「情報の深さ」で公開/非公開を区分けする設計になっていましたが、「情報の種類」で区分けすることを強く推奨します。理由は、概要だけの公開ページではSEO・AEO効果が薄く、AI検索に引用される一次情報量が確保できないためです。", { after: 100 }),

      h2("■ 区分けの基本方針"),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2000, 3513, 3513],
        rows: [
          headerRow([
            cell("", "1F5C9E", true, 2000),
            cell("一般公開ページ", "1F5C9E", true, 3513),
            cell("会員制ページ（審査招待制）", "1F5C9E", true, 3513),
          ]),
          dataRow([
            cell("目的", "EBF3FB", true, 2000),
            cell("SEO・AEO集客 / ブランド認知 / 会員誘導", null, false, 3513),
            cell("商談化 / 一次情報提供 / 信頼構築 / クロージング", null, false, 3513),
          ]),
          dataRow([
            cell("情報の性質", "EBF3FB", true, 2000),
            cell("市場・プロジェクトに関する「一般論・解説情報」\n（Desert Bamboo独自でなくても書ける内容）", null, false, 3513),
            cell("Desert Bamboo固有の「案件情報・現地一次情報」\n（他では入手できない情報）", null, false, 3513),
          ]),
          dataRow([
            cell("コンテンツ例", "EBF3FB", true, 2000),
            cell("・Vision 2030全体解説（NEOM・AMAALA・Red Sea・Diriyah・Qiddiyah等の各プロジェクト紹介）\n・サウジアラビア不動産市場の基礎知識（外国人購入制度・エスクロー・登記手続き等）\n・サウジ進出の法規制・会社設立ガイド\n・竹建築×中東の可能性コラム（Bamboo Labのブランドコンテンツ）\n・FAQ 20問以上（AI検索対応）\n・Bamboo Lab企業紹介・実績ハイライト\n・会員制ページへの誘導導線", null, false, 3513),
            cell("・現地ネットワーク経由の未公開プロジェクト情報（開発エリア・事業者・スケジュール）\n・投資案件の個別詳細（物件概要・価格・条件・エスクロー先・デューデリジェンス資料）\n・「永遠の砂時計」等の構想プロジェクト全容\n・Desert Bamboo独自の市場レポート（四半期更新）\n・現地パートナー・政府機関とのネットワーク情報\n・入居・購入・投資参加のプロセス資料\n・個別案件のQ&A・商談窓口", null, false, 3513),
          ]),
          dataRow([
            cell("更新頻度", "EBF3FB", true, 2000),
            cell("月2〜4本（コラム・FAQ・ニュース解説）", null, false, 3513),
            cell("月1〜2本（案件情報・現地レポート）", null, false, 3513),
          ]),
          dataRow([
            cell("想定読者", "EBF3FB", true, 2000),
            cell("サウジアラビアに関心を持つ一般ビジネス層・投資検討初期層・メディア・研究者", null, false, 3513),
            cell("富裕層個人投資家・ファミリーオフィス・法人投資家（審査通過者のみ）", null, false, 3513),
          ]),
          dataRow([
            cell("流入経路", "EBF3FB", true, 2000),
            cell("Google検索・AI検索・SNSシェア・PR記事", null, false, 3513),
            cell("公開ページからの会員申請・紹介・直接招待", null, false, 3513),
          ]),
        ]
      }),
      p("", { after: 100 }),

      h2("■ 会員審査フロー（案）"),
      bullet("申請フォームにて氏名・所属・投資関心領域・概算資産規模（任意）を取得"),
      bullet("日本事務局が一次スクリーニング（2営業日以内に返答）"),
      bullet("審査通過後、招待URLをメール送付・パスワード発行"),
      bullet("審査基準（仮）：個人の場合は金融資産1億円以上目安、法人の場合は役員・投資担当者"),
      note("スクリーニング基準はPhase 1にて岡崎社長と合意のうえ確定させてください"),

      // ══════════════════════════
      h1("4．ドメイン・サーバー・インフラ構成"),
      // ══════════════════════════

      h2("■ ドメイン方針"),
      p("本サイトはDesert Bambooのコーポレートサイト配下ではなく、独立した専用ドメインで立ち上げることを強く推奨します。", { after: 80 }),
      bullet("推奨候補ドメイン（優先順）：saudiarabia.jp ／ saudi-vision.jp ／ saudivision.jp ／ saudi-lab.jp"),
      bullet("「saudiarabia.jp」は現時点でアクティブなサイトが存在しない可能性が高く、取得できれば集客・ブランド両面で最大の資産になります。Phase 1確定を待たず、早期に空き確認・確保を並行して進めてください"),
      bullet("独立ドメインを推奨する理由："),
      subbullet("施工会社サイトと混在させると、富裕層・機関投資家向けの「情報メディア」としての信頼性が低下する"),
      subbullet("将来的な英語・アラビア語展開、第三者との提携・売却時に独立ドメインの方が圧倒的に動かしやすい"),
      subbullet("SEO上も、特化ドメインはGoogleから専門性シグナルが高く評価される傾向がある"),
      p("", { after: 60 }),

      h2("■ サーバー・インフラ構成（推奨3層構成）"),
      p("以下の3層構成を推奨します。コストと拡張性のバランスが最も優れており、将来のサウジ現地展開にも対応できます。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [600, 1600, 2400, 2213, 2213],
        rows: [
          headerRow([
            cell("層", "1F5C9E", true, 600),
            cell("レイヤー", "1F5C9E", true, 1600),
            cell("採用サービス", "1F5C9E", true, 2400),
            cell("役割", "1F5C9E", true, 2213),
            cell("費用目安", "1F5C9E", true, 2213),
          ]),
          dataRow([
            cell("①", "EBF3FB", true, 600),
            cell("CDN / WAF\n（フロント）", null, false, 1600),
            cell("Cloudflare\n（Free〜Pro）", null, false, 2400),
            cell("日本・サウジ双方からの表示速度を担保。DDoS対策・WAF・SSL証明書を一元管理。Desert Bamboo LLC（サウジ法人）のITインフラと構成を統一できる", null, false, 2213),
            cell("無料〜月$20", null, false, 2213),
          ]),
          dataRow([
            cell("②", "EBF3FB", true, 600),
            cell("ホスティング\n（バックエンド）", null, false, 1600),
            cell("Vercel（Webflow利用時）\nまたは\nAWS 東京リージョン\n（WordPress利用時）", null, false, 2400),
            cell("ターゲットが日本の富裕層のため国内リージョンを優先。Cloudflare経由で中東からのアクセス速度も補完できる。会員制コンテンツは国内リージョンに置くことでPDPL対応も容易", null, false, 2213),
            cell("月$20〜50\n（Vercel）\n月¥3,000〜\n（AWS EC2）", null, false, 2213),
          ]),
          dataRow([
            cell("③", "EBF3FB", true, 600),
            cell("会員管理\n（オプション）", null, false, 1600),
            cell("Memberstack\nまたは\nMemberful", null, false, 2400),
            cell("既存CMS（WordPress・Webflow）に乗せる形で会員審査・招待制・コンテンツのアクセス制御を実装。申請フォームからの審査フロー・メール通知まで一元管理できる", null, false, 2213),
            cell("月$49〜99\n（Memberstack）", null, false, 2213),
          ]),
        ]
      }),
      p("", { after: 100 }),

      h2("■ プラットフォーム（CMS）の選定"),
      p("Note（note.com）の採用はお控えください。以下の理由から、富裕層・機関投資家向けのサイトには適しません。", { after: 80 }),
      bullet("有料マガジン・情報商材のプラットフォームとして一般に認知されており、「専門性の高い情報メディア」としての信頼性が担保しにくい"),
      bullet("会員審査・招待制の仕組みを実装することが構造上困難"),
      bullet("独自ドメイン・デザインの自由度に制限があり、ブランド統一が難しい"),
      p("", { after: 80 }),
      p("推奨プラットフォームは以下の通りです。", { after: 80 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 2413, 2413, 2400],
        rows: [
          headerRow([
            cell("プラットフォーム", "1F5C9E", true, 1800),
            cell("メリット", "1F5C9E", true, 2413),
            cell("デメリット", "1F5C9E", true, 2413),
            cell("推奨シナリオ", "1F5C9E", true, 2400),
          ]),
          dataRow([
            cell("Webflow", "EBF3FB", true, 1800),
            cell("デザイン自由度が高く高級感のあるUI設計が可能。Memberstack連携で会員制を実装しやすい", null, false, 2413),
            cell("コンテンツ更新の学習コストがやや高い。プラグイン文化がない", null, false, 2413),
            cell("外注制作・デザイン品質を最優先する場合", null, false, 2400),
          ]),
          dataRow([
            cell("WordPress", "EBF3FB", true, 1800),
            cell("プラグインで機能拡張が豊富。SEO・会員制・多言語に強い。運用コストが低い", null, false, 2413),
            cell("デザインの差別化にはテーマ・カスタマイズが必要。セキュリティ管理が必要", null, false, 2413),
            cell("内製・ハイブリッド制作・長期運用コスト重視の場合", null, false, 2400),
          ]),
        ]
      }),
      p("", { after: 60 }),

      // ══════════════════════════
      h1("5．制作体制・費用に関する要望"),
      // ══════════════════════════

      p("提案書記載の費用目安（外注80〜200万円）は妥当な水準と認識しています。以下の点を制作会社への発注条件・確認事項として含めていただけると幸いです。", { after: 100 }),

      h2("■ 発注時の必須要件"),
      bullet("独立ドメインでの構築（コーポレートサイト配下への設置は不可）"),
      bullet("会員制機能（申請フォーム・審査フロー・コンテンツアクセス制御）の実装"),
      bullet("Cloudflare CDNとの連携・設定を含む"),
      bullet("FAQ構造化データ（JSON-LD / FAQSchema）の実装"),
      bullet("日本語・英語の二言語対応（英語は主要ページのみ）"),
      bullet("管理画面での記事・FAQ・案件情報の自社更新が可能なこと"),
      p("", { after: 60 }),

      h2("■ 確認・検討事項"),
      bullet("本郷さま（株式会社SAC Technologies）への相談の可否：同社が既存のシステム構築実績を持つ場合、見積もりのベンチマークとして活用できます。初期費用・月額・保守体制の比較検討をお願いします"),
      bullet("コンテンツ制作の分業：FAQ・コラム記事の原稿は日本事務局側で叩き台を作成することが可能です。岡崎社長側での一次情報・加筆・承認という分業体制を提案します"),
      bullet("ドメイン取得：Phase 1の合意を待たず、候補ドメインの空き確認および仮確保を早期に実施してください（費用：年間数千円程度）"),
      p("", { after: 60 }),

      // ══════════════════════════
      h1("6．フェーズ別アクション提案（補足）"),
      // ══════════════════════════

      p("提案書記載のロードマップに加え、以下のアクションを各フェーズに組み込んでいただくことを提案します。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1600, 2200, 5226],
        rows: [
          headerRow([
            cell("フェーズ", "1F5C9E", true, 1600),
            cell("時期", "1F5C9E", true, 2200),
            cell("追加アクション（本要望書より）", "1F5C9E", true, 5226),
          ]),
          dataRow([
            cell("Phase 1\n方向性確定", "EBF3FB", true, 1600),
            cell("〜2026年5月上旬", null, false, 2200),
            cell("・候補ドメインの空き確認・仮確保\n・CMS（Webflow or WordPress）の方向性決定\n・会員スクリーニング基準の仮合意\n・KW戦略・コンテンツ設計の合意", null, false, 5226),
          ]),
          dataRow([
            cell("Phase 2\n設計", "EBF3FB", true, 1600),
            cell("2026年5〜6月", null, false, 2200),
            cell("・FAQ初稿20問の作成（日本事務局担当）\n・公開/非公開コンテンツの詳細区分確定\n・Cloudflare・会員管理ツール選定\n・制作会社（SAC Technologies含む）への見積もり依頼", null, false, 5226),
          ]),
          dataRow([
            cell("Phase 3\n制作", "EBF3FB", true, 1600),
            cell("2026年6〜7月", null, false, 2200),
            cell("・FAQSchemaの実装確認\n・会員申請フォーム・審査フローのテスト\n・Cloudflare設定・SSL証明書の確認\n・公開前のSEOチェック（メタタグ・構造化データ・サイトマップ）", null, false, 5226),
          ]),
          dataRow([
            cell("Phase 4\n公開・運用", "EBF3FB", true, 1600),
            cell("2026年7月〜", null, false, 2200),
            cell("・月次でPV・会員申請数・AI引用状況をモニタリング\n・コンテンツ更新（月2〜4本）の継続\n・6ヶ月時点でKPI達成状況を評価し、英語展開・広告投資の判断", null, false, 5226),
          ]),
        ]
      }),
      p("", { after: 80 }),

      // 締め
      p("以上が日本事務局からのご要望です。本書に記載した内容はあくまで提案・要望であり、最終的な判断は岡崎社長にお任せします。Phase 1の議論において、本書を参考に方向性を合意できますと幸いです。", { before: 300, after: 80 }),
      p("ご不明な点・追加のご議論が必要な場合は、いつでもご連絡ください。引き続きよろしくお願いいたします。", { after: 400 }),
      p("以上", { align: AlignmentType.RIGHT }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/yobo3.docx', buf);
  console.log('Done');
});
