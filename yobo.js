const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, HeadingLevel
} = require('docx');
const fs = require('fs');

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "・",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 240 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "MS Gothic", size: 22 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "MS Gothic" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 0 }
      }
    ]
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
        spacing: { before: 0, after: 320 },
        children: [new TextRun({ text: "サイト制作に関するご要望", bold: true, size: 32, font: "MS Gothic" })]
      }),

      // 日付・宛先・差出人
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "2026年4月22日", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "岡崎社長　御中", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 400 },
        children: [new TextRun({ text: "Desert Bamboo Global 日本事務局　怒田裕也", size: 22, font: "MS Gothic" })]
      }),

      // 前文
      new Paragraph({
        spacing: { before: 0, after: 200 },
        children: [new TextRun({
          text: "先日ご提案いただきました「サウジアラビア向け情報×営業ハイブリッドサイト開設提案」（2026年4月22日付）について、コンセプトおよびサイトの方向性に賛同いたします。つきましては、制作・設計に着手いただくにあたり、以下の点についてご要望申し上げます。",
          size: 22, font: "MS Gothic"
        })]
      }),

      // 区切り線代わりのスペース
      new Paragraph({ spacing: { before: 0, after: 100 }, children: [] }),

      // 要望1
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "1．独立ドメインでの立ち上げ", bold: true, size: 24, font: "MS Gothic" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({
          text: "本サイトは、Desert Bamboo のコーポレートサイト配下ではなく、独立ドメインにて立ち上げることを希望します。",
          size: 22, font: "MS Gothic"
        })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "推奨ドメイン候補：saudiarabia.jp、saudi-vision.jp、saudivision.jp 等（早期に空き確認・確保をお願いします）", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "理由：ターゲットである富裕層・機関投資家への訴求において、コーポレートサイトとの混在はブランドイメージの分散につながるため", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "将来的な英語・アラビア語展開や第三者との提携においても、独立ドメインの方が柔軟に対応可能", size: 22, font: "MS Gothic" })]
      }),

      // 要望2
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "2．制作プラットフォームについて", bold: true, size: 24, font: "MS Gothic" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({
          text: "制作プラットフォームとして、Note（note.com）の採用はお控えいただくようお願いします。",
          size: 22, font: "MS Gothic"
        })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Note は有料記事・情報商材のプラットフォームとして認知されており、富裕層・機関投資家向けのブランドイメージと乖離があります", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "また、会員審査・招待制の仕組みを実装することが困難です", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "推奨プラットフォーム：Webflow または WordPress（会員制プラグイン対応）", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "会員管理機能が必要な場合は Memberstack / Memberful 等との組み合わせも検討ください", size: 22, font: "MS Gothic" })]
      }),

      // 要望3
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "3．公開コンテンツの充実（SEO / AEO 対策）", bold: true, size: 24, font: "MS Gothic" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({
          text: "一般公開ページには、Google 検索および AI 検索（ChatGPT・Perplexity 等）に対応した、専門性の高いコンテンツを一定量配置することをお願いします。",
          size: 22, font: "MS Gothic"
        })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "Vision 2030 解説記事、NEOM・AMAALA 等の最新動向コラム、竹建築×中東の可能性に関するオリジナル記事など", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "一般公開と会員制の区分は「情報の深さ（概要 vs. 詳細）」ではなく、「情報の種類（一般論 vs. Desert Bamboo 独自の案件情報）」で設計することを推奨します", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "公開コンテンツが薄いと AI 検索に引用されず、「サウジ建築の権威」としての認知形成が遅れる懸念があります", size: 22, font: "MS Gothic" })]
      }),

      // 要望4
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "4．サーバー・インフラ構成について", bold: true, size: 24, font: "MS Gothic" })]
      }),
      new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [new TextRun({
          text: "ホスティングおよびインフラについては、以下の構成を推奨します。",
          size: 22, font: "MS Gothic"
        })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "サーバー：日本リージョン（AWS 東京 または Vercel）", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "CDN / WAF：Cloudflare を前段に設置（サウジアラビアからのアクセス速度担保および将来の現地展開に対応）", size: 22, font: "MS Gothic" })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 0, after: 200 },
        children: [new TextRun({ text: "Desert Bamboo LLC（サウジ法人）の IT インフラで検討済みの Cloudflare 構成をそのまま活用可能です", size: 22, font: "MS Gothic" })]
      }),

      // 締め
      new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [new TextRun({
          text: "以上の点について、Phase 1（方向性確定）の議論の中でご確認いただけますと幸いです。引き続きよろしくお願いいたします。",
          size: 22, font: "MS Gothic"
        })]
      }),

      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 400, after: 0 },
        children: [new TextRun({ text: "以上", size: 22, font: "MS Gothic" })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/yobo.docx', buf);
  console.log('Done');
});
