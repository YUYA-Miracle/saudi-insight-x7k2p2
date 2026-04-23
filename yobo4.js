const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, VerticalAlign, PageBreak
} = require('docx');
const fs = require('fs');

const font = "MS Gothic";
const sz = 22;

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after !== undefined ? opts.after : 100 },
    alignment: opts.align || AlignmentType.LEFT,
    indent: opts.indent ? { left: opts.indent } : undefined,
    children: [new TextRun({ text, size: opts.size || sz, font, bold: opts.bold || false, color: opts.color || undefined })]
  });
}

function bullet(text, indent = 480) {
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
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F5C9E", space: 1 } },
    children: [new TextRun({ text, size: 28, font, bold: true, color: "1F5C9E" })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, size: 24, font, bold: true })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text: "◆ " + text, size: sz, font, bold: true, color: "1F5C9E" })]
  });
}

function note(text) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    indent: { left: 300 },
    children: [new TextRun({ text: "※ " + text, size: 19, font, color: "666666" })]
  });
}

function gap(n = 80) {
  return new Paragraph({ spacing: { before: 0, after: n }, children: [] });
}

function pb() {
  return new Paragraph({ children: [new TextRun({ break: 1 })], pageBreakBefore: true });
}

const bGray = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const bBlue = { style: BorderStyle.SINGLE, size: 4, color: "1F5C9E" };
const bsGray = { top: bGray, bottom: bGray, left: bGray, right: bGray };
const bsBlue = { top: bBlue, bottom: bBlue, left: bBlue, right: bBlue };

function cell(lines, bg, bold = false, width = 4513) {
  const texts = typeof lines === 'string' ? lines.split('\n') : lines;
  const isHeader = bg === "1F5C9E";
  const children = texts.map((line, i) =>
    new Paragraph({
      spacing: { before: 0, after: i < texts.length - 1 ? 60 : 0 },
      children: [new TextRun({ text: line, size: i === 0 && bold ? sz : (isHeader ? sz : 20), font, bold: bold && i === 0 || isHeader, color: isHeader ? "FFFFFF" : undefined })]
    })
  );
  return new TableCell({
    borders: isHeader ? bsBlue : bsGray,
    width: { size: width, type: WidthType.DXA },
    shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 100, bottom: 100, left: 160, right: 160 },
    verticalAlign: VerticalAlign.TOP,
    children
  });
}

function row(cells) { return new TableRow({ children: cells }); }

// ══════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "・", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } } }]
      },
      {
        reference: "subbullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "－", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 240 } } } }]
      }
    ]
  },
  styles: { default: { document: { run: { font, size: sz } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1701, right: 1701, bottom: 1701, left: 1701 }
      }
    },
    children: [

      // ─────────────────────────────────────────
      // タイトル
      // ─────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "サイト制作に関するご要望書", bold: true, size: 38, font })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」補足・要望資料", size: 20, font, color: "666666" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 500 },
        children: [new TextRun({ text: "2026年4月22日　Desert Bamboo Global 日本事務局　怒田裕也", size: 20, font, color: "666666" })]
      }),
      p("岡崎社長　御中", { after: 500 }),

      // ─────────────────────────────────────────
      // 前文
      // ─────────────────────────────────────────
      p("先日ご提案いただきました「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」（2026年4月22日付）について、コンセプトおよびサイトの基本方針に賛同いたします。", { after: 80 }),
      p("本書は、制作・設計に着手いただくにあたり、日本事務局として明確にしておきたいゴール設定・サイト構成・コンテンツ設計・技術基盤・収益モデルについて、具体的な要望と提案をまとめたものです。提案書を否定するものではなく、制作クオリティと集客効果を最大化するための補足として位置づけてください。", { after: 80 }),
      p("Phase 1（方向性確定）の議論において本書を叩き台とし、合意形成を図っていただけますと幸いです。", { after: 60 }),

      // ─────────────────────────────────────────
      h1("1．本サイトで達成したいゴール"),
      // ─────────────────────────────────────────

      p("サイトの成否を測る指標として、以下の定量・定性ゴールを設定することを提案します。これらのゴールを前提として、制作会社への発注仕様・内製判断・運用体制を設計してください。", { after: 100 }),

      h2("■ 定量ゴール"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 3013, 3013],
        rows: [
          row([cell("指標", "1F5C9E", true, 3000), cell("6ヶ月目標", "1F5C9E", true, 3013), cell("12ヶ月目標", "1F5C9E", true, 3013)]),
          row([cell("月間オーガニックPV", "EBF3FB", true, 3000), cell("500〜1,000 PV", null, false, 3013), cell("3,000 PV以上", null, false, 3013)]),
          row([cell("会員申請数（月間）", "EBF3FB", true, 3000), cell("3〜5件", null, false, 3013), cell("10件以上", null, false, 3013)]),
          row([cell("商談化件数（累計）", "EBF3FB", true, 3000), cell("1〜2件", null, false, 3013), cell("5件以上", null, false, 3013)]),
          row([cell("AI検索（GPT・Perplexity等）での引用", "EBF3FB", true, 3000), cell("主要KWで引用確認", null, false, 3013), cell("月10回以上", null, false, 3013)]),
          row([cell("主要KW Google検索順位", "EBF3FB", true, 3000), cell("20位以内", null, false, 3013), cell("10位以内", null, false, 3013)]),
          row([cell("メールマガジン登録数（累計）", "EBF3FB", true, 3000), cell("100名", null, false, 3013), cell("500名", null, false, 3013)]),
        ]
      }),
      gap(100),

      h2("■ 定性ゴール"),
      bullet("「サウジアラビア × 建築・不動産・投資 × 日本語」の領域において、Desert Bamboo Globalを唯一無二の一次情報源として確立する"),
      bullet("ChatGPT・Perplexity・Google AI Overviewが、サウジアラビア関連の質問に対してDesert Bambooのコンテンツを引用・推薦する状態をつくる"),
      bullet("富裕層・機関投資家が「このサイトは信頼できる専門家の情報だ」と感じ、会員登録・問い合わせへのハードルを自然に下げるサイト品質を担保する"),
      bullet("将来的な英語・アラビア語展開の起点となる拡張可能なコンテンツ資産と技術基盤を初期段階から設計しておく"),

      gap(60),

      // ─────────────────────────────────────────
      h1("2．競合分析と差別化ポイント"),
      // ─────────────────────────────────────────

      p("本サイトを設計するにあたり、競合となりうる既存サイト・メディアを整理し、差別化ポイントを明確化します。制作会社へのブリーフィングおよびコンテンツ戦略の基軸として活用してください。", { after: 100 }),

      h2("■ 競合サイト分析"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 2000, 2413, 2413],
        rows: [
          row([cell("カテゴリ", "1F5C9E", true, 2200), cell("代表例", "1F5C9E", true, 2000), cell("強み", "1F5C9E", true, 2413), cell("弱み・Desert Bambooが勝てる点", "1F5C9E", true, 2413)]),
          row([cell("政府系情報サイト\n（英語）", "EBF3FB", true, 2200), cell("MISA公式\nNeom.com\nVision2030.gov.sa", null, false, 2000), cell("情報の一次性・公式性", null, false, 2413), cell("日本語なし。投資家向けの実務情報が薄い。問い合わせ窓口が不明瞭", null, false, 2413)]),
          row([cell("日本の中東情報メディア", "EBF3FB", true, 2200), cell("JETRO中東情報\nアラブニュース日本版", null, false, 2000), cell("ニュース速報性\n日本語対応", null, false, 2413), cell("投資・建築・開発案件の実務情報がない。Desert Bamboo独自の一次情報・案件紹介は皆無", null, false, 2413)]),
          row([cell("不動産・投資コンサル系", "EBF3FB", true, 2200), cell("ドバイ・UAEの\n不動産仲介サイト各社", null, false, 2000), cell("案件数の多さ\n問い合わせ導線", null, false, 2413), cell("サウジアラビア特化なし。建築デザイン×投資という希少な切り口なし。信頼性の根拠が薄い", null, false, 2413)]),
          row([cell("建築・デザイン系メディア", "EBF3FB", true, 2200), cell("Dezeen\nArchDaily\n日経アーキテクチュア", null, false, 2000), cell("デザイン品質\n業界認知", null, false, 2413), cell("投資・商談機能なし。サウジ特化なし。日本語での一次情報発信なし", null, false, 2413)]),
        ]
      }),
      gap(100),

      h2("■ Desert Bamboo Globalの圧倒的差別化ポイント"),
      bullet("「サウジアラビア政府プロジェクト情報 × 建築設計実績 × 投資案件紹介」を日本語で一元提供できるサイトは現時点で世界に存在しない"),
      bullet("竹建築という独創的コンセプト × Vision 2030のサステナブル政策との親和性という、他社が模倣できない独自軸を持つ"),
      bullet("ロンボク島での実績（世界初竹建築スターバックス等）がブランドの信頼性を担保し、単なる情報メディアではなく「実績ある設計・開発者」として訴求できる"),
      bullet("現地ネットワーク経由の一次情報を保有しており、ネット上に流通していない情報を継続的に提供できるポジションにある"),
      note("この差別化優位は「早く動いた者が総取り」の構造。競合が追いつく前にポジションを確立することが最優先"),

      gap(60),

      // ─────────────────────────────────────────
      h1("3．サイト構成・ページ設計・導線設計"),
      // ─────────────────────────────────────────

      p("提案書のサイト構成案をベースに、SEO・AEO・商談化率の最大化を念頭に置いた具体的なページ構成・導線設計を提案します。制作会社への仕様書として本節の内容を活用してください。", { after: 100 }),

      h2("■ 推奨ページ構成（全体マップ）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [800, 2000, 1613, 4613],
        rows: [
          row([cell("No.", "1F5C9E", true, 800), cell("ページ名", "1F5C9E", true, 2000), cell("公開区分", "1F5C9E", true, 1613), cell("主な役割・設置コンテンツ", "1F5C9E", true, 4613)]),
          row([cell("01", "EBF3FB", true, 800), cell("TOP / ホーム", null, false, 2000), cell("一般公開", null, false, 1613), cell("ファーストビューでブランドの世界観を提示。「サウジアラビア × Desert Bamboo」の二軸コンセプトを30秒で理解させる設計。会員登録CTAを最上部と最下部の両方に配置。最新ニュース・注目記事のカードを3〜5件表示", null, false, 4613)]),
          row([cell("02", "EBF3FB", true, 800), cell("サウジアラビア市場解説", null, false, 2000), cell("一般公開", null, false, 1613), cell("Vision 2030の全体像から各ギガプロジェクト（NEOM・AMAALA・Red Sea Global・Diriyah・Qiddiyah）までをカテゴリ別に解説。各プロジェクトに個別詳細ページを設け、SEOの間口を最大化する。更新頻度：月2〜3本", null, false, 4613)]),
          row([cell("03", "EBF3FB", true, 800), cell("投資・不動産ガイド", null, false, 2000), cell("一般公開", null, false, 1613), cell("外国人のサウジ不動産購入制度・エスクロー・登記・RERA規制・会社設立手続きをガイド形式で解説。最も検索意図が高い実務情報を網羅。各記事末尾に会員申請CTAを設置", null, false, 4613)]),
          row([cell("04", "EBF3FB", true, 800), cell("FAQ", null, false, 2000), cell("一般公開", null, false, 1613), cell("AI検索対応の最重要ページ。「サウジアラビアで外国人は不動産を買えますか？」等の質問形式で最低30問を設置。各回答300〜500字・FAQSchemaを実装。月次でFAQを追加し続ける運用を前提とする", null, false, 4613)]),
          row([cell("05", "EBF3FB", true, 800), cell("Desert Bamboo紹介", null, false, 2000), cell("一般公開", null, false, 1613), cell("企業概要・実績（ロンボク島プロジェクト）・竹建築のコンセプト・サウジへの参入意義を掲載。信頼性の根拠として写真・実績データを豊富に使用。ここでの「実力の証明」が会員申請率に直結する", null, false, 4613)]),
          row([cell("06", "EBF3FB", true, 800), cell("ニュース・コラム", null, false, 2000), cell("一般公開", null, false, 1613), cell("サウジ最新動向・Bamboo Lab動向を月2〜4本発信。記事末尾に関連FAQ・会員申請CTAを設置。タグ・カテゴリ設計でSEOの網を広げる", null, false, 4613)]),
          row([cell("07", "EBF3FB", true, 800), cell("会員申請・お問い合わせ", null, false, 2000), cell("一般公開", null, false, 1613), cell("会員申請フォーム（氏名・所属・投資関心・概算資産規模）と一般問い合わせフォームを分離配置。申請完了後の自動返信メール・審査プロセスをページ上で明示しておく", null, false, 4613)]),
          row([cell("08", "EBF3FB", true, 800), cell("非公開案件情報", null, false, 2000), cell("会員限定", null, false, 1613), cell("進行中開発案件・投資機会の詳細（価格・条件・スケジュール・デューデリジェンス資料）。現地一次情報レポート。案件ごとに個別ページを設け、問い合わせCTAを必ず設置", null, false, 4613)]),
          row([cell("09", "EBF3FB", true, 800), cell("構想プロジェクト", null, false, 2000), cell("会員限定", null, false, 1613), cell("永遠の砂時計等のDesert Bamboo独自構想プロジェクトを掲載。投資家・コラボレーター向けに全容を開示。ここへの誘導が会員申請の最大のモチベーションになるよう、公開ページでの「チラ見せ」設計が重要", null, false, 4613)]),
          row([cell("10", "EBF3FB", true, 800), cell("会員限定レポート", null, false, 2000), cell("会員限定", null, false, 1613), cell("Desert Bamboo独自の市場分析レポートを四半期で発行。「このレポートのために会員であり続ける」価値を設計する", null, false, 4613)]),
        ]
      }),
      gap(100),

      h2("■ 導線設計（ユーザージャーニー）"),
      p("訪問者が「会員申請」「問い合わせ」に至るまでの導線を以下の通り設計してください。", { after: 80 }),

      h3("① オーガニック流入層（検索・AI検索経由）"),
      subbullet("「サウジアラビア 不動産 購入」等で検索 → 市場解説・投資ガイド・FAQページへ着地"),
      subbullet("記事を読む → 記事末尾の「続きは会員限定」または「詳しい案件情報はこちら」CTAへ誘導"),
      subbullet("会員申請フォームへ → 自動返信メールで審査フロー説明 → 審査通過後に会員ページへ招待"),

      h3("② TOP直接訪問層（紹介・SNS経由）"),
      subbullet("TOPのファーストビューでコンセプトを理解 → Desert Bamboo紹介ページで信頼性を確認"),
      subbullet("「会員になると何が見えるのか」のティーザー訴求（構想プロジェクトの冒頭だけ公開）"),
      subbullet("会員申請 or お問い合わせへ"),

      h3("③ 会員化後の商談フロー"),
      subbullet("会員ログイン後、非公開案件情報・構想プロジェクトを閲覧"),
      subbullet("各案件ページに「個別相談を申し込む」CTAを設置 → 日本事務局への商談問い合わせへ"),
      subbullet("商談化 → 岡崎社長・各部門へエスカレーション"),

      note("CTAは「会員申請する」と「まず資料を請求する」の2種類を用意し、申請ハードルの高低で受け皿を分散させることを推奨"),

      gap(60),

      h2("■ デザイン方針（制作会社へのブリーフィング用）"),
      bullet("トンマナ：「高級感・信頼性・静謐さ」。情報商材サイト・ニュースサイトとは明確に一線を画す。参考：Four Seasons・東京中目黒系高級ブランドサイトのトーン"),
      bullet("カラーパレット：砂漠のベージュ・金・深緑・白を基調。ネオンや鮮やか系の色は使用しない"),
      bullet("フォント：日本語は明朝系またはゴシック系の上質なもの。英語はSerif系を推奨（信頼感の演出）"),
      bullet("写真・ビジュアル：ロンボク島の竹建築実績写真・サウジの砂漠景観・プロジェクトのCGを高品質で使用。フリー素材の多用は禁止"),
      bullet("モバイル対応：富裕層はスマートフォンでの閲覧も多い。レスポンシブデザインは必須"),

      gap(60),

      // ─────────────────────────────────────────
      h1("4．キーワード戦略（SEO / AEO）"),
      // ─────────────────────────────────────────

      p("日本語でのサウジアラビア関連キーワードは検索ボリューム自体は多くありませんが、競合コンテンツがほぼ存在しないブルーオーシャンです。少量の良質コンテンツで上位表示・AI引用を狙えます。", { after: 100 }),

      h2("■ 優先KW一覧"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1200, 2800, 2013, 3013],
        rows: [
          row([cell("優先度", "1F5C9E", true, 1200), cell("キーワード群", "1F5C9E", true, 2800), cell("想定ターゲット", "1F5C9E", true, 2013), cell("対応コンテンツ", "1F5C9E", true, 3013)]),
          row([cell("★★★", "EBF3FB", true, 1200), cell("サウジアラビア 不動産 購入\nサウジアラビア 投資 日本人\nサウジアラビア 会社設立 外国人\nVision 2030 投資機会 日本", null, false, 2800), cell("富裕層個人\n法人投資家\n進出検討企業", null, false, 2013), cell("投資・不動産ガイド\nFAQ\n市場解説記事", null, false, 3013)]),
          row([cell("★★☆", "EBF3FB", true, 1200), cell("NEOM 最新情報 日本語\nAMALA サウジアラビア\nジェッダ 不動産 投資\nサウジアラビア ビザ 投資家\nRed Sea Global プロジェクト", null, false, 2800), cell("中東ビジネス関心層\nメディア・研究者\n初期情報収集層", null, false, 2013), cell("プロジェクト別解説\nニュース記事\n動向レポート", null, false, 3013)]),
          row([cell("★☆☆", "EBF3FB", true, 1200), cell("竹建築 中東\nサステナブル建築 サウジアラビア\nBamboo Lab サウジアラビア\nVision 2030 建築 デザイン\nDesert Bamboo", null, false, 2800), cell("建築・デザイン業界\n協業検討先\n指名検索層", null, false, 2013), cell("実績紹介\nプロジェクト解説\nブランドコラム", null, false, 3013)]),
        ]
      }),
      gap(100),

      h2("■ AEO（AI検索最適化）実施事項"),
      h3("FAQの大量・高品質設置"),
      bullet("公開FAQページに最低30問を設置。各回答は300〜500字・具体的な数値・出典（MISA・PIF・Vision 2030公式等）を明記"),
      bullet("質問形式の例："),
      subbullet("「サウジアラビアで外国人が不動産を購入できますか？」"),
      subbullet("「Vision 2030とは何ですか？日本企業にどんな投資機会がありますか？」"),
      subbullet("「NEOMの現在の進捗状況と縮小計画の実態は？」"),
      subbullet("「サウジアラビアでの会社設立にかかる費用と期間は？」"),
      subbullet("「竹建築はサウジアラビアの高温・砂漠環境に対応できますか？」"),
      subbullet("「AMALAプロジェクトへの外国人投資は可能ですか？」"),
      note("FAQは「会員制ページへの誘導フック」としても機能させる。概要をFAQで公開し、詳細（具体的案件・価格・条件）は会員限定とする二段構えを徹底する"),

      h3("構造化データ・E-E-A-T対策"),
      bullet("著者プロフィール（岡崎社長の経歴・実績・専門性）を全記事に付与し、AI検索の「信頼性シグナル」を強化"),
      bullet("各記事に「最終更新日」「情報ソース」を必ず明記"),
      bullet("JSON-LDによるFAQSchema・OrganizationSchema・BreadcrumbSchemaをページに実装"),
      bullet("Googleサーチコンソール・Bing Webmaster Toolsへのサイトマップ送信を公開初日に実施"),

      h3("更新頻度と運用体制"),
      bullet("公開コンテンツは月2〜4本を厳守。更新が止まるとAI検索での引用優先度が低下する"),
      bullet("日本事務局がFAQ・コラム原稿の叩き台を作成 → 岡崎社長が一次情報・現地知見を加筆 → 承認・公開という分業体制を提案する"),

      gap(60),

      // ─────────────────────────────────────────
      h1("5．公開・非公開コンテンツの詳細設計"),
      // ─────────────────────────────────────────

      p("提案書の「情報の深さで区分け」する設計を修正し、「情報の種類で区分け」することを強く推奨します。概要だけの公開ページではSEO・AEO効果が出ず、AI検索に引用されるだけの情報量が確保できないためです。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 3613, 3613],
        rows: [
          row([cell("", "1F5C9E", true, 1800), cell("一般公開", "1F5C9E", true, 3613), cell("会員制（審査招待制）", "1F5C9E", true, 3613)]),
          row([cell("目的", "EBF3FB", true, 1800), cell("SEO・AEO集客 / ブランド認知 / 信頼構築 / 会員誘導", null, false, 3613), cell("商談化 / 一次情報提供 / 長期関係構築 / クロージング", null, false, 3613)]),
          row([cell("情報の性質", "EBF3FB", true, 1800), cell("市場・プロジェクトの「一般論・解説情報」\n（Desert Bamboo独自でなくても書ける内容）", null, false, 3613), cell("Desert Bamboo固有の「案件情報・現地一次情報」\n（他では入手できない、流出を避けたい情報）", null, false, 3613)]),
          row([cell("コンテンツ例", "EBF3FB", true, 1800),
            cell("・Vision 2030全体解説\n・各ギガプロジェクト解説（NEOM・AMAALA・Red Sea・Diriyah・Qiddiyah）\n・サウジ不動産市場の基礎知識\n（外国人購入制度・エスクロー・登記・RERA等）\n・サウジ進出の法規制・会社設立ガイド\n・竹建築×中東のコンセプトコラム\n・FAQ 30問以上（AI検索対応）\n・Desert Bamboo企業紹介・実績\n・メールマガジン登録フォーム", null, false, 3613),
            cell("・現地ネットワーク経由の未公開PJ情報\n（開発エリア・事業者・スケジュール等）\n・投資案件の個別詳細\n（物件概要・価格・条件・エスクロー先・DD資料）\n・「永遠の砂時計」等の構想PJ全容\n・Desert Bamboo独自市場レポート（四半期）\n・現地パートナー・政府機関のネットワーク情報\n・購入・投資参加のプロセス資料\n・個別案件のQ&A・商談申込窓口", null, false, 3613)]),
          row([cell("更新頻度", "EBF3FB", true, 1800), cell("月2〜4本", null, false, 3613), cell("月1〜2本", null, false, 3613)]),
          row([cell("流入経路", "EBF3FB", true, 1800), cell("Google・AI検索・SNS・PR記事・紹介", null, false, 3613), cell("公開ページ → 会員申請 → 招待\n直接招待（営業・紹介）", null, false, 3613)]),
          row([cell("想定読者", "EBF3FB", true, 1800), cell("サウジに関心を持つ一般ビジネス層\nメディア・研究者・初期検討層", null, false, 3613), cell("富裕層個人投資家\nファミリーオフィス・法人投資家\n（審査通過者のみ）", null, false, 3613)]),
        ]
      }),
      gap(100),

      h2("■ 公開コンテンツの「チラ見せ」設計"),
      p("会員限定コンテンツへの申請モチベーションを高めるため、公開ページ上で以下の「チラ見せ」を実施してください。", { after: 80 }),
      bullet("構想プロジェクト（永遠の砂時計等）の冒頭500字・コンセプトビジュアルのみ公開し、「全容は会員限定」として誘導"),
      bullet("現地レポートのサマリー（1〜2段落）を公開し、詳細PDFは会員限定ダウンロード"),
      bullet("直近の非公開案件について「現在○件の投資案件を会員向けに掲載中」とカウントのみ表示"),
      note("「見えそうで見えない」設計が会員申請の最大の動機になる。提案書のFAQ「会員になると何が見えますか？」への回答を公開ページ上で明確に示すこと"),

      gap(60),

      // ─────────────────────────────────────────
      h1("6．会員スクリーニング基準・フロー"),
      // ─────────────────────────────────────────

      p("会員制の品質を担保し、商談化率を高めるために、スクリーニング基準とフローを以下の通り設計することを提案します。", { after: 100 }),

      h2("■ スクリーニング基準（案）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1800, 3613, 3613],
        rows: [
          row([cell("属性", "1F5C9E", true, 1800), cell("通過基準（案）", "1F5C9E", true, 3613), cell("確認方法", "1F5C9E", true, 3613)]),
          row([cell("個人投資家", "EBF3FB", true, 1800), cell("金融資産1億円以上（目安）\n不動産・海外投資の経験あり\n投資目的が明確", null, false, 3613), cell("申請フォームの自己申告\n職業・資産規模の記入欄を設ける\n必要に応じてオンライン面談", null, false, 3613)]),
          row([cell("法人・機関投資家", "EBF3FB", true, 1800), cell("役員・投資担当者・ファミリーオフィス\nM&A・海外投資の実績あり\n中東・サウジアラビアへの関心が具体的", null, false, 3613), cell("会社名・役職の記入\n法人の事業概要確認\n必要に応じてLinkedIn等で確認", null, false, 3613)]),
          row([cell("メディア・研究者", "EBF3FB", true, 1800), cell("信頼性の高い媒体・機関に所属\n取材・調査目的が明確", null, false, 3613), cell("所属媒体・機関名の確認\n掲載方針について個別協議", null, false, 3613)]),
          row([cell("不承認ケース", "EBF3FB", true, 1800), cell("目的・属性が不明確な申請\n競合他社と判断される場合\n投資能力が著しく低いと判断される場合", null, false, 3613), cell("丁寧にお断りするメールテンプレートを用意しておく", null, false, 3613)]),
        ]
      }),
      note("スクリーニング基準の最終決定はPhase 1にて岡崎社長と合意してください。初期は基準を緩めに設定し、商談パイプラインの母数を確保することも戦略の一つ"),
      gap(100),

      h2("■ 審査フロー"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [800, 2200, 3013, 3013],
        rows: [
          row([cell("Step", "1F5C9E", true, 800), cell("アクション", "1F5C9E", true, 2200), cell("担当", "1F5C9E", true, 3013), cell("目安時間", "1F5C9E", true, 3013)]),
          row([cell("1", "EBF3FB", true, 800), cell("申請フォーム送信", null, false, 2200), cell("申請者", null, false, 3013), cell("即時", null, false, 3013)]),
          row([cell("2", "EBF3FB", true, 800), cell("自動返信メール送信\n（審査中の旨・所要日数を通知）", null, false, 2200), cell("システム自動", null, false, 3013), cell("即時", null, false, 3013)]),
          row([cell("3", "EBF3FB", true, 800), cell("一次スクリーニング\n（申請内容の確認・判断）", null, false, 2200), cell("日本事務局", null, false, 3013), cell("2営業日以内", null, false, 3013)]),
          row([cell("4a", "EBF3FB", true, 800), cell("承認の場合：招待URL・\nパスワードをメール送付", null, false, 2200), cell("日本事務局", null, false, 3013), cell("審査完了後即日", null, false, 3013)]),
          row([cell("4b", "EBF3FB", true, 800), cell("保留の場合：\nオンライン面談を打診", null, false, 2200), cell("日本事務局", null, false, 3013), cell("3営業日以内に連絡", null, false, 3013)]),
          row([cell("4c", "EBF3FB", true, 800), cell("不承認の場合：\nお断りメール送付", null, false, 2200), cell("日本事務局", null, false, 3013), cell("5営業日以内", null, false, 3013)]),
          row([cell("5", "EBF3FB", true, 800), cell("会員ページ閲覧・案件情報確認", null, false, 2200), cell("会員", null, false, 3013), cell("招待後随時", null, false, 3013)]),
          row([cell("6", "EBF3FB", true, 800), cell("個別案件相談の申し込み\n→ 商談化・岡崎社長へエスカレーション", null, false, 2200), cell("会員 → 日本事務局", null, false, 3013), cell("会員自身のタイミング", null, false, 3013)]),
        ]
      }),
      gap(60),

      // ─────────────────────────────────────────
      h1("7．ドメイン・サーバー・インフラ構成"),
      // ─────────────────────────────────────────

      h2("■ ドメイン方針"),
      p("本サイトはDesert Bambooコーポレートサイト配下ではなく、独立した専用ドメインでの立ち上げを強く要望します。", { after: 80 }),
      bullet("推奨候補（優先順）：saudiarabia.jp → saudi-vision.jp → saudivision.jp → saudi-lab.jp"),
      bullet("「saudiarabia.jp」はアクティブなサイトが現時点で存在せず、取得できれば集客・ブランド両面で最大の資産になります。Phase 1の合意を待たず早期に空き確認・仮確保を並行して進めてください（年間費用：数千円程度）"),
      bullet("独立ドメインを推奨する理由："),
      subbullet("施工会社サイトと混在すると、富裕層向け「情報メディア」としての信頼性が低下する"),
      subbullet("将来の英語・アラビア語展開、第三者提携・売却時に独立ドメインの方が圧倒的に動かしやすい"),
      subbullet("SEO上も特化ドメインはGoogleから専門性シグナルが高く評価される"),
      gap(100),

      h2("■ 推奨3層インフラ構成"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [500, 1600, 2400, 3126, 1400],
        rows: [
          row([cell("層", "1F5C9E", true, 500), cell("レイヤー", "1F5C9E", true, 1600), cell("採用サービス", "1F5C9E", true, 2400), cell("役割・理由", "1F5C9E", true, 3126), cell("費用目安", "1F5C9E", true, 1400)]),
          row([cell("①", "EBF3FB", true, 500), cell("CDN / WAF", null, false, 1600), cell("Cloudflare（Free〜Pro）", null, false, 2400), cell("日本・サウジ双方からの表示速度を担保。DDoS・WAF・SSL証明書を一元管理。Desert Bamboo LLC（サウジ法人）のIT構成と統一可能", null, false, 3126), cell("無料〜月$20", null, false, 1400)]),
          row([cell("②", "EBF3FB", true, 500), cell("ホスティング", null, false, 1600), cell("Vercel（Webflow時）\nAWS東京（WordPress時）", null, false, 2400), cell("ターゲットが日本の富裕層のため国内リージョン優先。Cloudflare経由でサウジからのアクセス速度も補完。会員制コンテンツはPDPL対応の観点からも国内リージョンが望ましい", null, false, 3126), cell("月$20〜50\n（Vercel）\n月¥3,000〜\n（AWS）", null, false, 1400)]),
          row([cell("③", "EBF3FB", true, 500), cell("会員管理", null, false, 1600), cell("Memberstack\nまたはMemberful", null, false, 2400), cell("既存CMSに会員審査・招待制・コンテンツアクセス制御を実装。申請フォームから審査フロー・招待メールまで一元管理", null, false, 3126), cell("月$49〜99", null, false, 1400)]),
        ]
      }),
      gap(100),

      h2("■ CMS（制作プラットフォーム）選定"),
      p("Note（note.com）の採用はお控えください。情報商材プラットフォームとしての認知が強く、富裕層・機関投資家向けのブランドに不適合です。会員審査・招待制の実装も構造上困難です。", { after: 80 }),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1600, 2613, 2413, 2400],
        rows: [
          row([cell("CMS", "1F5C9E", true, 1600), cell("メリット", "1F5C9E", true, 2613), cell("デメリット", "1F5C9E", true, 2413), cell("推奨シナリオ", "1F5C9E", true, 2400)]),
          row([cell("Webflow", "EBF3FB", true, 1600), cell("デザイン自由度が高く高級感のあるUI設計が可能。Memberstack連携で会員制を実装しやすい。ノーコードで運用可能", null, false, 2613), cell("コンテンツ更新の学習コストがやや高い。プラグイン拡張性がない", null, false, 2413), cell("外注制作・デザイン品質を最優先する場合", null, false, 2400)]),
          row([cell("WordPress", "EBF3FB", true, 1600), cell("プラグインで機能拡張が豊富。SEO・会員制・多言語に強い。長期運用コストが低い", null, false, 2613), cell("デザイン差別化にはカスタマイズが必要。セキュリティ管理が必要", null, false, 2413), cell("内製・ハイブリッド制作・長期運用コスト重視の場合", null, false, 2400)]),
        ]
      }),
      gap(60),

      // ─────────────────────────────────────────
      h1("8．収益モデル・ROI試算"),
      // ─────────────────────────────────────────

      p("本サイトへの投資対効果を試算します。制作費用の意思決定および運用継続の判断材料としてご活用ください。", { after: 100 }),

      h2("■ 収益モデル"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [2200, 3413, 3413],
        rows: [
          row([cell("収益タイプ", "1F5C9E", true, 2200), cell("内容", "1F5C9E", true, 3413), cell("想定単価・規模", "1F5C9E", true, 3413)]),
          row([cell("①設計・開発案件の受注", "EBF3FB", true, 2200), cell("会員向けに紹介した案件の設計・施工受注。本サイトの主収益源", null, false, 3413), cell("案件単価：数億〜数十億円\n年1件受注で投資回収", null, false, 3413)]),
          row([cell("②投資仲介・コンサルティング", "EBF3FB", true, 2200), cell("サウジ不動産・プロジェクトへの投資仲介フィー・コンサル契約", null, false, 3413), cell("仲介フィー：投資額の1〜3%\nコンサル：月50〜100万円", null, false, 3413)]),
          row([cell("③会員制レポートの有料化\n（将来オプション）", "EBF3FB", true, 2200), cell("現在は無料会員制だが、将来的に上位会員向け有料プランの設計も可能", null, false, 3413), cell("月額3〜10万円/会員\n×50名＝月150〜500万円", null, false, 3413)]),
          row([cell("④メディア価値の活用", "EBF3FB", true, 2200), cell("サイトの権威性が上がった段階での取材対応・登壇・PR記事等によるブランド価値向上", null, false, 3413), cell("間接効果として定性評価", null, false, 3413)]),
        ]
      }),
      gap(100),

      h2("■ ROI試算（保守的シナリオ）"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3000, 3013, 3013],
        rows: [
          row([cell("項目", "1F5C9E", true, 3000), cell("金額（目安）", "1F5C9E", true, 3013), cell("備考", "1F5C9E", true, 3013)]),
          row([cell("【投資】制作費（外注）", "EBF3FB", true, 3000), cell("100〜150万円", null, false, 3013), cell("Webflow + 会員制機能実装含む", null, false, 3013)]),
          row([cell("【投資】月額維持費", "EBF3FB", true, 3000), cell("5〜10万円/月", null, false, 3013), cell("サーバー・Cloudflare・Memberstack・コンテンツ制作", null, false, 3013)]),
          row([cell("【投資】年間運用コスト合計", "EBF3FB", true, 3000), cell("初年度160〜270万円", null, false, 3013), cell("制作費＋維持費12ヶ月", null, false, 3013)]),
          row([cell("【回収】投資仲介1件（保守的）", "EBF3FB", true, 3000), cell("300〜500万円", null, false, 3013), cell("5億円案件の仲介フィー1%想定", null, false, 3013)]),
          row([cell("【回収】設計・施工受注1件（保守的）", "EBF3FB", true, 3000), cell("数億〜数十億円", null, false, 3013), cell("規模により大幅に上振れ", null, false, 3013)]),
          row([cell("投資回収の目安", "EBF3FB", true, 3000), cell("仲介1件成立で初年度回収", null, false, 3013), cell("設計受注が1件でも入れば数十倍のROI", null, false, 3013)]),
        ]
      }),
      note("上記はあくまで参考試算です。案件規模・成約率により大幅に変動します。重要なのは「1件の大型案件で初期投資が回収できる」構造であり、本サイトへの投資はリスクに対してリターンが明らかに大きい"),
      gap(60),

      // ─────────────────────────────────────────
      h1("9．将来の多言語展開（英語・アラビア語）"),
      // ─────────────────────────────────────────

      p("本サイトは将来的な英語・アラビア語展開を見据えた設計にしておくことを要望します。初期は日本語メインでOKですが、技術基盤と構造だけは多言語対応できるように設計しておいてください。", { after: 100 }),

      h2("■ フェーズ別言語展開計画"),
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1800, 3413, 2413],
        rows: [
          row([cell("フェーズ", "1F5C9E", true, 1400), cell("時期", "1F5C9E", true, 1800), cell("言語展開", "1F5C9E", true, 3413), cell("想定ターゲット", "1F5C9E", true, 2413)]),
          row([cell("Phase 1〜2\n（現在）", "EBF3FB", true, 1400), cell("〜2026年末", null, false, 1800), cell("日本語メイン。主要ページ（TOP・企業紹介・お問い合わせ）のみ英語を併記", null, false, 3413), cell("日本の富裕層・機関投資家", null, false, 2413)]),
          row([cell("Phase 3\n（拡張）", "EBF3FB", true, 1400), cell("2027年〜", null, false, 1800), cell("英語ページを全面展開。サウジ・GCC地域の開発関係者・政府機関・グローバル投資家向けにコンテンツを追加", null, false, 3413), cell("GCC富裕層\nグローバル投資家\n現地開発関係者", null, false, 2413)]),
          row([cell("Phase 4\n（現地化）", "EBF3FB", true, 1400), cell("2028年〜\n（NOD上場前後）", null, false, 1800), cell("アラビア語対応。サウジ政府・地場デベロッパー・現地富裕層向けに現地化。Desert Bamboo LLCのIR資料との連携も視野に", null, false, 3413), cell("サウジアラビア現地法人\nNOMU/MAIN上場に向けた\n投資家向け情報発信", null, false, 2413)]),
        ]
      }),
      gap(80),

      h2("■ 初期設計時の必須要件"),
      bullet("CMSは多言語対応可能なもの（WordPress + WPML、またはWebflow多言語プラン）を選定する"),
      bullet("URLは言語ごとにサブディレクトリ構造（/ja/ /en/ /ar/）で設計し、hreflangタグを実装できるようにしておく"),
      bullet("RTL（右から左）テキスト対応が必要なアラビア語のために、CSSのdirection設定を初期設計から考慮しておく"),
      bullet("コンテンツ管理は言語ごとに独立しているが、案件・会員データは一元管理できる構造にする"),

      gap(60),

      // ─────────────────────────────────────────
      h1("10．フェーズ別アクションプラン（補足）"),
      // ─────────────────────────────────────────

      p("提案書記載のロードマップに本要望書の内容を統合した、より詳細なアクションプランを提示します。", { after: 100 }),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [1400, 1600, 6026],
        rows: [
          row([cell("フェーズ", "1F5C9E", true, 1400), cell("時期", "1F5C9E", true, 1600), cell("具体的アクション", "1F5C9E", true, 6026)]),
          row([cell("Phase 1\n方向性確定", "EBF3FB", true, 1400), cell("〜5月上旬", null, false, 1600),
            cell("・候補ドメインの空き確認・仮確保（並行着手）\n・CMS（Webflow or WordPress）の方向性決定\n・会員スクリーニング基準の仮合意\n・KW戦略・ページ構成の合意（本要望書を叩き台に）\n・制作体制（外注 or 内製 or SAC Technologies）の決定\n・収益モデル・ROI目線の合意", null, false, 6026)]),
          row([cell("Phase 2\n設計", "EBF3FB", true, 1400), cell("5〜6月", null, false, 1600),
            cell("・ページ構成詳細設計（本要望書10ページ構成を仕様書化）\n・FAQ初稿30問作成（日本事務局担当）\n・公開/非公開コンテンツの詳細区分確定\n・会員申請フォーム・審査フロー設計\n・Cloudflare・Memberstack選定・設定方針確定\n・デザインブリーフィング資料作成（トンマナ・参考サイト選定）\n・制作会社（SAC Technologies含む）への見積もり依頼・比較", null, false, 6026)]),
          row([cell("Phase 3\n制作", "EBF3FB", true, 1400), cell("6〜7月", null, false, 1600),
            cell("・ドメイン本取得・DNS設定・Cloudflare接続\n・CMSセットアップ・会員管理ツール連携\n・初期コンテンツ（FAQ30問・解説記事5本・企業紹介）投入\n・FAQSchema・OrganizationSchema実装確認\n・会員申請フォーム・自動返信メール・審査フローのテスト\n・Google Search Console・サイトマップ送信\n・公開前SEOチェック（メタタグ・OGP・サイト速度）", null, false, 6026)]),
          row([cell("Phase 4\n公開・運用", "EBF3FB", true, 1400), cell("7月〜", null, false, 1600),
            cell("・公開と同時にLinkedIn・SNSでのPR発信\n・月次でFAQ追加・コラム2〜4本更新継続\n・月次KPIモニタリング（PV・会員申請数・AI引用状況）\n・6ヶ月時点でKPI達成状況を評価し英語展開・有料会員化の判断\n・商談パイプラインのCRM管理との連携（Notion等）", null, false, 6026)]),
        ]
      }),
      gap(80),

      // ─────────────────────────────────────────
      // 締め
      // ─────────────────────────────────────────
      p("以上が日本事務局からの要望および提案です。本書に記載した内容はあくまで叩き台であり、最終的な判断は岡崎社長にお任みします。", { before: 300, after: 80 }),
      p("Phase 1の議論において本書の各項目について確認・合意できますと幸いです。ご不明な点や追加のご議論が必要な場合は、いつでもご連絡ください。引き続きよろしくお願いいたします。", { after: 400 }),
      p("以上", { align: AlignmentType.RIGHT }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/yobo4.docx', buf);
  console.log('Done');
});
