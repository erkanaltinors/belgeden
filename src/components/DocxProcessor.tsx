import { type ChangeEvent, useState } from "react";
import mammoth from "mammoth";
import { Document, Paragraph, Packer, TextRun, AlignmentType, convertMillimetersToTwip } from "docx";
import * as FileSaver from "file-saver";
import JSZip from "jszip";

const headingRegex = /Turu(?:.*\(\s*[İi]le\s*\))?$/;
const bodyTextSize = 18; // 9pt

function isValidOutputName(name: string): boolean {
  return Boolean(name?.trim().length);
}

function getListLevel(item: Element): number {
  let level = 0;
  let parent: HTMLElement | null = item.parentElement;

  while (parent && parent.tagName !== "BODY") {
    if (parent.tagName === "UL" || parent.tagName === "OL") {
      level += 1;
    }
    parent = parent.parentElement;
  }

  return Math.max(0, level - 1);
}

function getTextFromNode(node: Node): string {
  const textParts: string[] = [];

  function collectText(current: Node): void {
    if (current.nodeType === Node.TEXT_NODE) {
      const value = current.textContent?.trim() ?? "";
      if (value) {
        textParts.push(value);
      }
      return;
    }

    if (current.nodeType === Node.ELEMENT_NODE) {
      const element = current as Element;
      if (element.tagName === "SCRIPT" || element.tagName === "STYLE") {
        return;
      }

      current.childNodes.forEach(collectText);
    }
  }

  collectText(node);
  return textParts.join(" ").replace(/\s+/g, " ").trim();
}

function normalizeImportedText(value: string): string {
  const normalized = normalizeHeaderText(value || "").replace(/\s+/g, " ").trim();
  const stripped = normalized
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[Çç]/g, "c")
    .replace(/[Ğğ]/g, "g")
    .replace(/[İı]/g, "i")
    .replace(/[Öö]/g, "o")
    .replace(/[Şş]/g, "s")
    .replace(/[Üü]/g, "u")
    .toLowerCase();

  if (
    stripped.includes("pakettur") &&
    stripped.includes("fiyatina") &&
    stripped.includes("dahil") &&
    (stripped.includes("servis") || stripped.includes("hizmet")) &&
    !stripped.includes("olmayan")
  ) {
    return "Fiyatlara Dahil Olan Hizmetler";
  }

  if (
    stripped.includes("pakettur") &&
    stripped.includes("fiyatina") &&
    stripped.includes("dahil") &&
    stripped.includes("olmayan") &&
    (stripped.includes("servis") || stripped.includes("hizmet"))
  ) {
    return "Fiyatlara Dahil Olmayan Hizmetler";
  }

  const insurancePhrasePattern = /^(S|s)eyahat (V|v)e (S|s)a(Ğ|G|g|ğ)lık (S|s)igortası/i;

  if (insurancePhrasePattern.test(normalized)) {
    return "Seyahat ve Sağlık Sigortası";
  }

  return normalized;
}

function walkNodes(node: Node, callback: (node: Node) => void): void {
  callback(node);
  node.childNodes.forEach((child) => walkNodes(child, callback));
}

function getLocalName(node: Node): string | null {
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return null;
  }

  const element = node as Element;
  return element.localName ?? element.tagName.split(":").pop()?.toLowerCase() ?? null;
}

type HeaderParagraph = {
  text: string;
  bold: boolean;
  size: number;
  words: number;
  digits: number;
  containsPrice: boolean;
};

function normalizeHeaderText(value: string): string {
  return value.replace(/[\u2018\u2019\u201C\u201D]/g, "'").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
}

function cleanHeaderText(value: string): string {
  let text = normalizeHeaderText(value);

  text = text.replace(/(?:^|\s)(?:\d+(?:\s+\d+)*)\s*(?:euro|tl|usd|\$|€)\s*(?:'?(?:\s*(?:dan|den)\s+itibaren))?(?=\s|$)/gi, " ");
  text = text.replace(/\b(?:euro|tl|usd|\$|€|fiyat|tutar)\b/gi, " ");
  text = text.replace(/\b(?:dan|den)\s+itibaren\b/gi, " ");
  text = text.replace(/\b\d+\b/g, " ");
  text = text.replace(/\b([A-Za-zÇçĞğİıÖöŞşÜü]{1,2})\s+(?=[A-Za-zÇçĞğİıÖöŞşÜü]{2,})/g, "$1");
  text = text.replace(/["']/g, "");
  text = text.replace(/\s+/g, " ").trim();

  return text.replace(/^[\s\-–—:;,.]+|[\s\-–—:;,.]+$/g, "").trim();
}

function toTitleCaseWord(value: string): string {
  const normalized = value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[Çç]/g, "c")
    .replace(/[Ğğ]/g, "g")
    .replace(/[İı]/g, "i")
    .replace(/[Öö]/g, "o")
    .replace(/[Şş]/g, "s")
    .replace(/[Üü]/g, "u")
    .replace(/[^a-zA-Z0-9\s]/g, "");

  if (!normalized) {
    return "";
  }

  const lower = normalized.toLowerCase();
  return lower.charAt(0).toUpperCase() + lower.slice(1);
}

function formatMainTitle(value: string): string {
  const normalized = normalizeHeaderText(value || "");
  if (!normalized) {
    return "";
  }

  const beforeTuru = normalized.split(/\bTURU\b/i)[0].trim();
  const base = beforeTuru
    .replace(/\bDOLU\s+DOLU\b/gi, "Premium")
    .replace(/[&/\\-]+/g, " ")
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!base) {
    return "Turu";
  }

  const titleWords = base
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => toTitleCaseWord(word))
    .filter(Boolean);

  return titleWords.length ? `${titleWords.join(" ")} Turu` : "Turu";
}

function formatDayHeading(value: string): string {
  const normalized = normalizeHeaderText(value || "");
  if (!normalized) {
    return "";
  }

  const dayMatch = normalized.match(/^(\d+)\s*\.\s*Gün\s*(.*)$/i) || normalized.match(/^(\d+)\s*Gün\s*(.*)$/i);
  if (!dayMatch) {
    return normalized;
  }

  const [, day, routePart] = dayMatch;
  const cleanedRoute = routePart
    .replace(/[–—-]/g, " - ")
    .replace(/\s+/g, " ")
    .trim()
    .split(" - ")
    .map((segment) => segment.trim())
    .filter(Boolean)
    .map((segment) => {
      const words = segment.split(/\s+/).filter(Boolean);
      return words
        .map((word) => {
          const cleaned = word.replace(/[^\p{L}\p{N}]/gu, "");
          if (!cleaned) {
            return word;
          }
          const lower = cleaned.toLowerCase();
          return lower.charAt(0).toUpperCase() + lower.slice(1);
        })
        .join(" ");
    })
    .join(" – ");

  return `${day}.Gün:${cleanedRoute}`;
}

function formatSecondaryTitle(value: string): string {
  const normalized = normalizeHeaderText(value || "");
  if (!normalized) {
    return "";
  }

  const parts = normalized.split("|").map((part) => part.trim()).filter(Boolean);
  if (parts.length < 2) {
    return "";
  }

  const leftPart = parts[0];
  const rightPart = parts[1];
  const airlineName = leftPart
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\bile\b/gi, "")
    .trim();

  const airlineTitle = airlineName
    ? airlineName
      .split(/\s+/)
      .filter(Boolean)
      .map((word) => toTitleCaseWord(word))
      .join(" ")
    : "";

  if (!airlineTitle) {
    return "";
  }

  if (airlineTitle.toLowerCase().includes("turk")) {
    return "Türk\nHavayolları Tarifeli Seferi İle";
  }

  return `${airlineTitle} Tarifeli Seferi İle`.trim();
}

function parseHeaderParagraph(element: Element): HeaderParagraph | null {
  const textParts: string[] = [];
  let bold = false;
  const sizes: number[] = [];

  function collectRunInfo(node: Node): void {
    if (node.nodeType !== Node.ELEMENT_NODE) {
      return;
    }

    const element = node as Element;
    const localName = getLocalName(element);
    if (localName === "t") {
      const value = element.textContent?.trim();
      if (value) {
        textParts.push(value);
      }
      return;
    }

    if (localName === "b") {
      const val = element.getAttribute("w:val");
      if (val === null || val.toLowerCase() !== "false") {
        bold = true;
      }
    }

    if (localName === "sz" || localName === "szCs") {
      const val = element.getAttribute("w:val");
      const numeric = val ? parseInt(val, 10) : NaN;
      if (!Number.isNaN(numeric)) {
        sizes.push(numeric / 2);
      }
    }

    node.childNodes.forEach(collectRunInfo);
  }

  collectRunInfo(element);
  const text = cleanHeaderText(textParts.join(" "));
  if (!text) {
    return null;
  }

  const words = text.split(/\s+/).length;
  const digits = (text.match(/\d/g) || []).length;
  const containsPrice = /\b(Euro|TL|USD|\$|€|Fiyat|den itibaren|dan itibaren|tutar|baiteribaren)\b/i.test(text);
  const size = sizes.length ? Math.max(...sizes) : 0;

  return {
    text,
    bold,
    size,
    words,
    digits,
    containsPrice,
  };
}

function collectHeaderParagraphs(node: Node, paragraphs: HeaderParagraph[]): void {
  if (node.nodeType !== Node.ELEMENT_NODE) {
    return;
  }

  const element = node as Element;
  const localName = getLocalName(element);
  if (localName === "p") {
    const parsed = parseHeaderParagraph(element);
    if (parsed) {
      paragraphs.push(parsed);
    }
    return;
  }

  node.childNodes.forEach((child) => collectHeaderParagraphs(child, paragraphs));
}

function scoreHeaderParagraph(paragraph: HeaderParagraph): number {
  let score = paragraph.words * 10;
  score += paragraph.bold ? 80 : 0;
  score += paragraph.size * 8;
  score -= paragraph.digits * 6;
  score -= paragraph.containsPrice ? 200 : 0;
  score -= paragraph.text.length > 100 ? 20 : 0;
  score -= paragraph.words < 3 ? 100 : 0;
  return score;
}

function buildMainTitle(paragraphs: HeaderParagraph[]): string {
  const cleanedParagraphs = paragraphs
    .map((paragraph) => ({ ...paragraph, text: cleanHeaderText(paragraph.text) }))
    .filter((paragraph) => Boolean(paragraph.text))
    .filter((paragraph) => !/\d/.test(paragraph.text))
    .filter((paragraph) => !/\b(?:euro|tl|usd|\$|€|fiyat|tutar)\b/i.test(paragraph.text));

  if (!cleanedParagraphs.length) {
    return "";
  }

  const sorted = cleanedParagraphs.slice().sort((a, b) => scoreHeaderParagraph(b) - scoreHeaderParagraph(a));
  const [best, second] = sorted;

  if (!best) {
    return "";
  }

  if (second && second.text.split(/\s+/).length <= 6 && !best.text.includes(second.text) && !second.text.includes(best.text)) {
    return `${best.text} ${second.text}`.trim();
  }

  return best.text;
}

async function extractHeaderTextFromDocx(arrayBuffer: ArrayBuffer): Promise<{ textBlocks: string[]; mainTitle: string; secondaryTitle: string }> {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const headerEntries = Object.keys(zip.files).filter((name) => /^word\/header\d*\.xml$/i.test(name));
  const paragraphs: HeaderParagraph[] = [];

  for (const entryName of headerEntries) {
    const entry = zip.files[entryName];
    if (!entry || entry.dir) continue;

    const xml = await entry.async("string");
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xml, "application/xml");
    collectHeaderParagraphs(xmlDoc.documentElement, paragraphs);
  }

  const cleanedParagraphs = paragraphs
    .map((paragraph) => ({ ...paragraph, text: cleanHeaderText(paragraph.text) }))
    .filter((paragraph) => Boolean(paragraph.text));
  const normalizedBlocks = cleanedParagraphs.map((paragraph) => paragraph.text).filter(Boolean);
  const matchingBlock = normalizedBlocks.find((text) => /.+\s*\|\s*[A-ZÇĞİÖŞÜ]{3,}(-?\s*[A-ZÇĞİÖŞÜ]{3,})?/i.test(text));

  return {
    textBlocks: normalizedBlocks,
    mainTitle: formatMainTitle(buildMainTitle(cleanedParagraphs)),
    secondaryTitle: formatSecondaryTitle(matchingBlock ?? ""),
  };
}

type DocxItem =
  | { type: "heading"; text: string }
  | { type: "subheading"; text: string }
  | { type: "list"; text: string; level: number }
  | { type: "paragraph"; text: string; bold?: boolean; spacingBefore?: boolean };

function DocxProcessor() {
  const [error, setError] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [outputName, setOutputName] = useState("");
  const [nightInput, setNightInput] = useState("");
  const [dayInput, setDayInput] = useState("");
  const [isProcessing, setIsProcessing] = useState(false);

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    setError("");
    const selectedFile = event.target.files?.[0] ?? null;

    if (!selectedFile) {
      setFile(null);
      return;
    }

    if (!selectedFile.name.toLowerCase().endsWith(".docx")) {
      setFile(null);
      setError("Lütfen .docx uzantılı bir dosya seçin.");
      return;
    }

    setFile(selectedFile);
  }

  async function handleProcessClick() {
    if (!file) {
      setError("Lütfen önce bir DOCX dosyası seçin.");
      return;
    }

    if (!isValidOutputName(outputName)) {
      setError("Dosya adı en az bir harf veya rakam içermelidir.");
      return;
    }

    setError("");
    setIsProcessing(true);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { value: html } = await mammoth.convertToHtml({ arrayBuffer });
      const headerResult = await extractHeaderTextFromDocx(arrayBuffer);
      // Print full converted HTML of the uploaded DOCX to the browser console
      console.log("Converted HTML:", html);
      console.log("Header blocks:", headerResult.textBlocks);
      console.log("Main header title:", headerResult.mainTitle);
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");
      const items: DocxItem[] = [];

      const bodyChildren = Array.from(doc.body?.children ?? []);
      let serviceListPending = false;

      for (let index = 0; index < bodyChildren.length; index += 1) {
        const element = bodyChildren[index];
        const tag = element.tagName;
        const rawText = getTextFromNode(element);
        const text = normalizeImportedText(rawText);

        if (text && headingRegex.test(text)) {
          items.push({ type: "heading", text });
          continue;
        }

        const isServiceHeading = text === "Fiyatlara Dahil Olan Hizmetler" || text === "Fiyatlara Dahil Olmayan Hizmetler";
        if (isServiceHeading) {
          if (text === "Fiyatlara Dahil Olmayan Hizmetler") {
            items.push({ type: "paragraph", text: "", spacingBefore: true });
          }

          items.push({ type: "paragraph", text, bold: true });
          serviceListPending = true;
          continue;
        }

        const strongText = element.querySelector("strong")?.textContent?.trim() ?? "";
        const isDayHeading = /^(\d+\s*\.\s*Gün|\d+\s*Gün)/i.test(strongText);
        if (tag === "P" && isDayHeading) {
          const formattedDayTitle = formatDayHeading(strongText);
          items.push({ type: "subheading", text: formattedDayTitle });

          const nextElement = bodyChildren[index + 1];
          if (nextElement && nextElement.tagName === "P") {
            const nextText = normalizeImportedText(getTextFromNode(nextElement));
            if (nextText) {
              items.push({ type: "paragraph", text: nextText });
              items.push({ type: "paragraph", text: "" });
            }
          }

          index += 1;
          continue;
        }

        const className = ((element as HTMLElement).className || "").toString();
        const isHtmlHeading = /^H[1-6]$/.test(tag) || /heading|title|header/i.test(className) || (element.getAttribute && element.getAttribute("role") === "heading");
        if (isHtmlHeading) {
          continue;
        }

        if ((tag === "UL" || tag === "OL") && element.children.length) {
          Array.from(element.children).forEach((child) => {
            if (child.tagName === "LI") {
              const listText = normalizeImportedText(getTextFromNode(child));
              if (listText) {
                if (serviceListPending) {
                  items.push({ type: "paragraph", text: listText });
                } else {
                  items.push({ type: "list", text: listText, level: getListLevel(child as Element) });
                }
              }
            }
          });
          serviceListPending = false;
          continue;
        }

        if (tag === "LI") {
          if (text) {
            items.push({ type: "list", text, level: getListLevel(element) });
          }
          continue;
        }

        if (tag === "P") {
          if (text) {
            items.push({ type: "paragraph", text });
          }
        }

        if (tag === "TD" || tag === "TH") {
          const cellText = normalizeImportedText(getTextFromNode(element));
          if (cellText) {
            items.push({ type: "paragraph", text: cellText });
          }
        }
      }

      if (!items.length) {
        setError('Eşleşen "Turu" başlıkları veya liste maddesi bulunamadı.');
        setIsProcessing(false);
        return;
      }



      const thirdTitle = [nightInput.trim(), dayInput.trim()].filter(Boolean);
      const titleParagraphs = [
        ...(headerResult.mainTitle ? [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: headerResult.mainTitle,
              size: 28,
              bold: true,
              font: "Arial",
            }),
          ],
        })] : []),
        ...(headerResult.secondaryTitle ? [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: headerResult.secondaryTitle,
              size: 28,
              bold: true,
              font: "Arial",
            }),
          ],
        })] : []),
        ...(thirdTitle.length === 2 ? [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `${nightInput.trim()} Gece ${dayInput.trim()} Gün`,
                size: 28,
                bold: true,
                font: "Arial",
              }),
            ],
          }),
          new Paragraph({}),
          new Paragraph({}),
        ] : []),
      ];

      const docxDocument = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: convertMillimetersToTwip(25),
                  right: convertMillimetersToTwip(25),
                  bottom: convertMillimetersToTwip(25),
                  left: convertMillimetersToTwip(25),
                  gutter: 0,
                },
              },
            },
            children: [
              ...titleParagraphs,
              ...items.map((item) => {
                if (item.type === "heading") {
                  return new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [
                      new TextRun({
                        text: item.text,
                        size: 28, // 14pt
                        bold: true,
                        font: "Arial",
                      }),
                    ],
                  });
                }

                if (item.type === "subheading") {
                  return new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    children: [
                      new TextRun({
                        text: item.text,
                        size: bodyTextSize,
                        bold: true,
                        font: "Arial",
                      }),
                    ],
                  });
                }

                if (item.type === "list") {
                  return new Paragraph({
                    bullet: { level: item.level },
                    alignment: AlignmentType.JUSTIFIED,
                    children: [
                      new TextRun({
                        text: item.text,
                        size: bodyTextSize,
                        font: "Arial",
                      }),
                    ],
                  });
                }

                // paragraph
                return new Paragraph({
                  spacing: item.spacingBefore ? { before: 240 } : undefined,
                  alignment: AlignmentType.JUSTIFIED,
                  children: [
                    new TextRun({
                      text: item.text,
                      size: bodyTextSize,
                      font: "Arial",
                      bold: item.bold ?? false,
                    }),
                  ],
                });
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(docxDocument);
      FileSaver.saveAs(blob, `${outputName}.docx`);
      setIsProcessing(false);
    } catch (err) {
      console.error(err);
      setError("Dosya işlenirken bir hata oluştu.");
      setIsProcessing(false);
    }
  }

  const isStartDisabled = !file || !outputName.trim().length || isProcessing;
  const disabledReason = !file
    ? "Lütfen önce bir DOCX dosyası seçin."
    : !outputName.trim().length
      ? "Lütfen yeni dosya adı alanını doldurun."
      : "";

  return (
    <div>
      <div className="max-w-2xl mx-auto rounded-xs border border-slate-200 p-4 shadow-sm bg-sky-50">
        <h2 className="text-xl text-center font-semibold mb-4">
          DOSYALA BEYBİ
        </h2>
        <label className="block mb-4" htmlFor="fileInput">
          Yüklenecek dosya
          <input
            id="fileInput"
            name="fileInput"
            type="file"
            accept=".docx"
            onChange={handleFileChange}
            className="mt-2 block w-full text-sm text-slate-700 file:mr-4 file:py-2 file:px-4 file:rounded file:bg-violet-300 file:text-slate-700 hover:file:bg-violet-500 hover:file:text-white hover:file:cursor-pointer transition-all duration-200"
          />
        </label>
        <label htmlFor="nightInput" className="block mb-2">Gece<input type="text" id="nightInput" value={nightInput} onChange={(event) => setNightInput(event.target.value)} className="bg-white block w-full" /></label>
        <label htmlFor="dayInput" className="block mb-2">Gün<input type="text" id="dayInput" value={dayInput} onChange={(event) => setDayInput(event.target.value)} className="bg-white block w-full" /></label>
        <label className="block mb-4">
          <span className="block text-sm font-medium text-slate-700">
            Yeni Dosya Adı
          </span>
          <input
            type="text"
            value={outputName}
            onChange={(event) => setOutputName(event.target.value)}
            placeholder="örnek-dosya-adi"
            className="mt-2 block w-full rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm text-slate-700 shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200"
          />
        </label>

        <button
          type="button"
          onClick={handleProcessClick}
          disabled={isStartDisabled}
          className="inline-flex items-center justify-center rounded bg-violet-900 px-5 py-3 text-sm font-semibold text-white disabled:bg-neutral-500 disabled:opacity-50 transition-all disabled:cursor-not-allowed"
        >
          {isProcessing ? "İşleniyor..." : "Başlat"}
        </button>
      </div>
      {error && <p className="mt-4 mb-2 text-sm text-red-600">{error}</p>}
      {!error && disabledReason ? (
        <p className="text-center mt-2 text-md text-red-400">{disabledReason}</p>
      ) : null}
    </div>
  );
}

export default DocxProcessor;
