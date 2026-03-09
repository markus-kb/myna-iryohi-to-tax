import fs from "fs/promises";
import { TextDecoder } from "util";

// --- なぜ `csv` ライブラリを使わず手動でパースするのか ---
// 入力ファイルは単純な2カラムのキー・バリュー形式 CSV のみ。
// フル CSV パーサーは不要なオーバーヘッドであり、npm 依存も増やす。
// `npx tsx index.ts` だけで動かせるよう、ゼロインストールを維持する。

/** RFC 4180 に従い、フィールドを囲む二重引用符を取り除き、連続した引用符をアンエスケープする。 */
function unquote(field: string): string {
  const trimmed = field.trim();
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replace(/""/g, '"');
  }
  return trimmed;
}

/**
 * 2カラム CSV 文字列を [キー, 値] の行配列にパースする。
 * 内側にカンマを含むクォートフィールド（例："78,040"）は処理できるが、
 * 複数行にまたがるクォート値には対応しない（入力 CSV には存在しないため）。
 */
function parseKeyValueCsv(text: string): Array<[string, string]> {
  const rows: Array<[string, string]> = [];
  for (const line of text.split(/\r?\n/)) {
    if (line.trim() === "") continue;

    // クォート内にない最初のカンマで分割する。
    // 戦略：クォート内かどうかを追跡しながら1文字ずつ走査する。
    let inQuotes = false;
    let splitIndex = -1;
    for (let i = 0; i < line.length; i++) {
      if (line[i] === '"') {
        inQuotes = !inQuotes;
      } else if (line[i] === "," && !inQuotes) {
        splitIndex = i;
        break;
      }
    }

    if (splitIndex === -1) continue; // 1カラム行（セクションヘッダーなど）はスキップ

    const key = unquote(line.slice(0, splitIndex));
    const value = unquote(line.slice(splitIndex + 1));
    rows.push([key, value]);
  }
  return rows;
}

/**
 * エンコーディングを自動判別する：まず UTF-8 を試み、失敗したら cp932（Shift-JIS）にフォールバック。
 *
 * この順番にする理由：実際には Shift-JIS なのに UTF-8 として `TextDecoder` に渡すと、
 * strict モードでエラーになる（逆も然り）。`fatal: true` オプションで
 * 不正なバイト列を拒否させることを利用している。
 */
function decodeBuffer(buf: Buffer): string {
  // UTF-8 BOM があれば確実に UTF-8
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return new TextDecoder("utf-8").decode(buf);
  }

  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buf);
  } catch {
    // UTF-8 デコード失敗 → cp932（日本の政府・保険機関の CSV で標準的な
    // Windows Shift-JIS 上位互換エンコーディング）と判断する。
    return new TextDecoder("shift_jis").decode(buf);
  }
}

interface MedicalRecord {
  person: string;
  hospital: string;
  type: string;
  amount: string;
  date: string; // YYYY/MM — 診療年月（例："2025年2月" → "2025/02"）からパース
}

/**
 * "2025年2月" のような日本語の年月表記を "2025/02" に変換する。
 * 形式が想定外の場合は空文字を返し、不正な値を出力しないようにする。
 */
function parseJapaneseYearMonth(value: string): string {
  const match = value.match(/^(\d{4})年(\d{1,2})月$/);
  if (!match) return "";
  return `${match[1]}/${match[2].padStart(2, "0")}`;
}

/**
 * 1つの CSV ファイルからすべての医療費レコードを抽出する。
 *
 * CSV は論理的に2つのセクションに分かれている：
 *   1〜30行目  — ヘッダー・集計部（氏名などを含む）
 *   31行目以降 — 医療費情報明細の繰り返しブロック。各ブロックは以下を含む：
 *                  診療年月、診療区分、日数、医療機関等名称、
 *                  医療費の総額、保険者の負担額、その他の公費の負担額、
 *                  窓口相当負担額
 *
 * ヘッダーから氏名を取得したあと、明細行を順に読み込んでレコードを収集する。
 * 窓口相当負担額（各ブロックの最終フィールド）に達した時点でレコードが「完成」する。
 */
function extractRecords(text: string): MedicalRecord[] {
  const allRows = parseKeyValueCsv(text);

  // --- ヘッダーセクション（最初の30行）から氏名を取得 ---
  let person = "";
  for (const [key, value] of allRows) {
    if (key === "氏名") {
      person = value;
      break;
    }
  }

  // --- 医療費情報明細マーカー以降の明細行を抽出 ---
  let inDetailSection = false;
  const records: MedicalRecord[] = [];
  let type = "";
  let hospital = "";
  let date = "";

  for (const [key, value] of allRows) {
    // 明細セクションは "医療費情報明細" 行（1カラム行のためパーサーにスキップされる）の
    // 直後から始まる。最初の明細キーは必ず "診療年月"。
    // 集計セクションの窓口相当負担額を過ぎてから 診療年月 が現れた時点で
    // 明細セクションへの突入を検知する。
    if (key === "診療年月") {
      inDetailSection = true;
      date = parseJapaneseYearMonth(value);
      continue;
    }
    if (!inDetailSection) continue;

    switch (key) {
      case "診療区分":
        type = value;
        break;
      case "医療機関等名称":
        hospital = value;
        break;
      case "窓口相当負担額（円）": {
        // 桁区切りカンマを除去 — xlsx は純粋な数値を期待する。
        const amount = value.replace(/,/g, "");
        records.push({ person, hospital, type, amount, date });
        break;
      }
      // その他のキー（日数、医療費の総額など）は意図的に無視する —
      // xlsx フォームが必要とするフィールドは上記のみ。
    }
  }

  return records;
}

/**
 * MedicalRecord を xlsx の列レイアウト（B〜J列）にマッピングする。
 *
 * xlsx 列：
 *   A: No               — =ROW()-8 で自動生成されるため出力しない
 *   B: 医療を受けた人       — 氏名
 *   C: 病院・薬局などの名称  — 医療機関等名称
 *   D: 診療・治療           — 医科外来または歯科外来の場合に「該当する」
 *   E: 医薬品購入           — 調剤の場合に「該当する」
 *   F: 介護保険サービス      — （空欄：元データになし）
 *   G: その他の医療費        — （空欄：元データになし）
 *   H: 支払った医療費の金額  — 窓口相当負担額の金額
 *   I: 左のうち、補填される金額 — 0（元データに補填情報なし）
 *   J: 支払年月日           — 診療年月から YYYY/MM 形式（日は元データになし）
 */
function toTsvRow(record: MedicalRecord): string {
  const columns = [
    record.person,
    record.hospital,
    // 医科外来・歯科外来はともに「診療・治療」に該当。調剤は「医薬品購入」に該当。
    record.type === "医科外来" || record.type === "歯科外来"
      ? "該当する"
      : "",
    record.type === "調剤" ? "該当する" : "",
    "", // 介護保険サービス
    "", // その他の医療費
    record.amount,
    "0",
    record.date,
  ];
  return columns.join("\t");
}

async function main() {
  let files = process.argv.slice(2);

  // 引数がない場合、カレントディレクトリの .csv ファイルを自動検出する。
  // ダウンロードした CSV と同じフォルダで `npx tsx index.ts` を実行するのが
  // 主な使い方。
  if (files.length === 0) {
    const entries = await fs.readdir(".");
    files = entries
      .filter((f) => f.toLowerCase().endsWith(".csv"))
      .sort();

    if (files.length === 0) {
      console.error(
        "カレントディレクトリに .csv ファイルが見つかりません。\n" +
          "使い方: npx tsx index.ts [file1.csv file2.csv ...]\n" +
          "  または CSV と同じフォルダにスクリプトを置いて引数なしで実行してください。"
      );
      process.exit(1);
    }

    console.error(`CSV ファイルが ${files.length} 件見つかりました：`);
    for (const f of files) {
      console.error(`  ${f}`);
    }
  }

  const results = await Promise.all(
    files.map(async (file) => {
      const buf = await fs.readFile(file);
      const records = extractRecords(decodeBuffer(buf));
      if (records.length === 0) {
        console.error(`警告: ${file} に医療費レコードが見つかりませんでした`);
      }
      return records;
    })
  );
  const allRecords = results.flat();

  if (allRecords.length === 0) {
    console.error("どのファイルにもレコードが見つかりませんでした。");
    process.exit(1);
  }

  // TSV を標準出力に書き出す — ファイルにリダイレクトするかターミナルからコピーする。
  // ヘッダー行なし：xlsx にはすでにヘッダーがあり、9行目以降に貼り付けるだけでよい。
  const tsv = allRecords.map(toTsvRow).join("\n") + "\n";
  process.stdout.write(tsv);

  // タイムスタンプ付きの .tsv ファイルにも書き出す（後から参照しやすいよう）。
  const now = new Date();
  const stamp = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    "_",
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ].join("");
  const outFile = `iryouhi_${stamp}.tsv`;
  await fs.writeFile(outFile, tsv, "utf-8");
  console.error(`${allRecords.length} 件のレコードを ${outFile} に書き出しました`);
}

main();
