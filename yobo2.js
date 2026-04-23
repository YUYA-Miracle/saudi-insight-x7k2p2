const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, VerticalAlign
} = require('docx');
const fs = require('fs');

const font = "MS Gothic";
const sz = 22;
const szH = 24;

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after || 100 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, size: opts.size || sz, font, bold: opts.bold || false })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text, size: sz, font })]
  });
}

function h1(text) {
  return new Paragraph({
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 1 } },
    children: [new TextRun({ text, size: 26, font, bold: true, color: "2E75B6" })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, size: szH, font, bold: true })]
  });
}

const border = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, bg, bold = false, width = 4683) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 140, right: 140 },
    children: [new Paragraph({ children: [new TextRun({ text, size: sz, font, bold })] })]
  });
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

      // タイトル
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "サイト制作に関するご要望書", bold: true, size: 34, font })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "（「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」補足）", size: 20, font, color: "666666" })]
      }),

      // 日付・宛先・差出人
      p("2026年4月22日", { align: AlignmentType.RIGHT }),
      p("岡崎社長　御中", { after: 60 }),
      p("Desert Bamboo Global 日本事務局　怒田裕也", { align: AlignmentType.RIGHT, after: 400 }),

      // 前文
      p("先日ご提案いただきましたサイト開設提案について、方向性に賛同いたします。制作・設計に着手いただくにあたり、具体的なイメージと要望を以下にまとめましたので、Phase 1の議論の叩き台としてお使いください。", { after: 80 }),

      // ━━━━━━━━━━
      h1("1．達成したいゴール"),

      h2("■ 集客目標"),
      bullet("オーガニック検索（Google・AI検索）経由で、月間500〜1,000PVを6ヶ月以内に達成"),
      bullet("問い合わせ・会員申請を月3〜5件ペースで獲得（うち1件以上を商談化）"),
      bullet("「Desert Bamboo = サウジアラビア建築・投資の一次情報源」としてAI検索（ChatGPT・Perplexity等）に認知させる"),
      p("", { after: 60 }),

      h2("■ ブランド目標"),
      bullet("「サウジアラビア × 日本語 × 建築・投資情報」の領域で、競合のない独占ポジションを確立"),
      bullet("富裕層・機関投資家が「信頼できる専門家」として認識するサイト品質を担保する"),
      p("", { after: 60 }),

      // ━━━━━━━━━━
      h1("2．狙うキーワード（SEO / AEO）"),

      p("以下のキーワード群を優先的に取りに行く設計をお願いします。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 3626, 3200],
        rows: [
          new TableRow({
            children: [
              cell("優先度", "2E75B6", true, 2200),
              cell("キーワード例", "2E75B6", true, 3626),
              cell("狙い", "2E75B6", true, 3200),
            ]
          }),
          new TableRow({ children: [
            cell("★★★", "EBF3FB", false, 2200),
            cell("サウジアラビア 投資 日本人\nサウジアラビア 不動産 購入\nVision 2030 投資機会", null, false, 3626),
            cell("実需層への直接リーチ。競合ほぼゼロ", null, false, 3200),
          ]}),
          new TableRow({ children: [
            cell("★★☆", "EBF3FB", false, 2200),
            cell("NEOM プロジェクト 最新情報\nAMALA サウジアラビア\nサウジアラビア 建築 設計", null, false, 3626),
            cell("認知獲得・メディア露出。検索数は中程度だが一次情報で差別化可", null, false, 3200),
          ]}),
          new TableRow({ children: [
            cell("★☆☆", "EBF3FB", false, 2200),
            cell("竹建築 中東\nサステナブル建築 サウジアラビア\nVision 2030 日本語", null, false, 3626),
            cell("Bamboo Lab固有の差別化ワード。AEO（AI引用）に有効", null, false, 3200),
          ]}),
        ]
      }),
      p("", { after: 120 }),

      h2("■ AEO（AI検索最適化）対応"),
      bullet("「サウジアラビアの投資環境はどうですか？」「NEOMとは何ですか？」など、AI検索で問われやすい質問形式のFAQコンテンツを公開ページに多数設置する"),
      bullet("FAQは最低20問以上を想定。各回答は300〜500字の一次情報ベースで記述"),
      bullet("AI検索に引用されるには「日本語・専門性・具体的な数値・独自情報」の4要素が重要"),
      p("", { after: 60 }),

      // ━━━━━━━━━━
      h1("3．公開コンテンツ・非公開コンテンツの設計イメージ"),

      p("「概要は公開、詳細は非公開」ではなく、「情報の種類」で区分けすることを推奨します。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 3613, 3613],
        rows: [
          new TableRow({ children: [
            cell("", "2E75B6", true, 1800),
            cell("一般公開ページ", "2E75B6", true, 3613),
            cell("会員制ページ（審査制）", "2E75B6", true, 3613),
          ]}),
          new TableRow({ children: [
            cell("目的", "EBF3FB", true, 1800),
            cell("SEO・AEO対策 / ブランド認知 / 会員誘導", null, false, 3613),
            cell("商談化 / 一次情報の提供 / 信頼構築", null, false, 3613),
          ]}),
          new TableRow({ children: [
            cell("コンテンツ例", "EBF3FB", true, 1800),
            cell("・Vision 2030 解説記事（NEOM・AMAALA・Diriyah等）\n・竹建築×中東の可能性コラム\n・Bamboo Lab 企業紹介・実績\n・FAQ（AI検索対応・20問以上）\n・サウジ不動産市場の基礎知識", null, false, 3613),
            cell("・現地ネットワーク経由の未公開プロジェクト情報\n・投資案件の個別詳細（価格・条件・スケジュール）\n・永遠の砂時計など構想プロジェクトの全容\n・入居・購入のプロセス資料\n・Desert Bamboo独自の市場レポート", null, false, 3613),
          ]}),
          new TableRow({ children: [
            cell("更新頻度", "EBF3FB", true, 1800),
            cell("月2〜4本（コラム・FAQ追記）", null, false, 3613),
            cell("月1〜2本（案件情報・現地レポート）", null, false, 3613),
          ]}),
          new TableRow({ children: [
            cell("想定読者", "EBF3FB", true, 1800),
            cell("サウジアラビアに興味を持つ一般層・投資検討初期層・メディア関係者", null, false, 3613),
            cell("富裕層個人投資家・ファミリーオフィス・法人投資家（審査通過者のみ）", null, false, 3613),
          ]}),
        ]
      }),
      p("", { after: 120 }),

      // ━━━━━━━━━━
      h1("4．ドメイン・サーバー構成について"),

      h2("■ ドメイン"),
      bullet("Desert Bamboo コーポレートサイト配下ではなく、独立ドメインでの立ち上げを希望します"),
      bullet("推奨候補：saudiarabia.jp / saudi-vision.jp / saudivision.jp（早期に空き確認・確保をお願いします）"),
      bullet("理由：富裕層・機関投資家への訴求において、施工会社サイトと混在するとブランドが分散するため"),
      p("", { after: 60 }),

      h2("■ サーバー・インフラ構成（推奨）"),
      p("以下の3層構成を推奨します。", { after: 80 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 3000, 3826],
        rows: [
          new TableRow({ children: [
            cell("レイヤー", "2E75B6", true, 2200),
            cell("採用サービス", "2E75B6", true, 3000),
            cell("役割・理由", "2E75B6", true, 3826),
          ]}),
          new TableRow({ children: [
            cell("① CDN / WAF\n（フロント）", "EBF3FB", false, 2200),
            cell("Cloudflare（無料〜Pro）", null, false, 3000),
            cell("日本・サウジ双方からの表示速度を担保。DDoS対策・WAF機能つき。Desert Bamboo LLC（サウジ法人）のIT構成と統一可能", null, false, 3826),
          ]}),
          new TableRow({ children: [
            cell("② ホスティング\n（バックエンド）", "EBF3FB", false, 2200),
            cell("AWS 東京リージョン\nまたは Vercel", null, false, 3000),
            cell("ターゲットが日本の富裕層のため国内リージョン優先。Webflow利用の場合はVercelが最適", null, false, 3826),
          ]}),
          new TableRow({ children: [
            cell("③ 会員管理\n（オプション）", "EBF3FB", false, 2200),
            cell("Memberstack\nまたは Memberful", null, false, 3000),
            cell("既存CMS（WordPress・Webflow）に乗せる形で会員審査・招待制を実装。月額数千円〜", null, false, 3826),
          ]}),
        ]
      }),
      p("", { after: 120 }),

      h2("■ プラットフォームについて"),
      bullet("Note（note.com）の採用はお控えください。情報商材プラットフォームとしての認知が強く、富裕層・機関投資家向けのブランドイメージと乖離があります。また会員審査・招待制の実装が困難です"),
      bullet("推奨：Webflow（デザイン自由度◎・会員制対応）または WordPress（プラグインで柔軟に拡張可）"),
      p("", { after: 60 }),

      // ━━━━━━━━━━
      h1("5．その他確認事項"),
      bullet("ドメインの空き確認・早期取得：Phase 1確定前に並行して着手いただけると助かります"),
      bullet("コンテンツ制作体制：FAQ・コラム記事の初期制作は日本事務局側で原稿を用意することも可能です。ご相談ください"),
      bullet("会員スクリーニング基準：「資産規模○億円以上」「法人限定」など基準の仮置きをPhase 1で議論したいと思います"),
      p("", { after: 60 }),

      // 締め
      p("以上の点についてご確認・ご検討のほど、よろしくお願いいたします。", { before: 200, after: 400 }),
      p("以上", { align: AlignmentType.RIGHT }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/yobo2.docx', buf);
  console.log('Done');
});
