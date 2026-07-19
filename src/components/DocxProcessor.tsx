import { type ChangeEvent, useState } from "react";
import mammoth from "mammoth";
import { Document, Paragraph, Packer, TextRun, AlignmentType, convertMillimetersToTwip } from "docx";
import * as FileSaver from "file-saver";

const dayHeadingRegex = /^(\d+)\s*\.\s*Gün|^(\d+)\s*Gün/i; 
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
  const normalized = value.replace(/[\u2018\u2019\u201C\u201D]/g, "'").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
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
    stripped.includes("fiyatina") &&
    stripped.includes("dahil") &&
    (stripped.includes("servis") || stripped.includes("hizmet")) &&
    !stripped.includes("olmayan")
  ) {
    return "Fiyatlara Dahil Olan Hizmetler";
  }

  if (
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

function formatDayHeading(value: string): string {
  const normalized = value.replace(/[\u2018\u2019\u201C\u201D]/g, "'").replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
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

type DocxItem =
  | { type: "subheading"; text: string }
  | { type: "list"; text: string; level: number }
  | { type: "paragraph"; text: string; bold?: boolean }
  | { type: "empty-line" };

function DocxProcessor() {
  const [error, setError] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [outputName, setOutputName] = useState("");
  const [nightInput, setNightInput] = useState("");
  const [dayInput, setDayInput] = useState("");
  const [mainTitleInput, setMainTitleInput] = useState("");
  const [airline, setAirline] = useState("Türk Hava Yolları");
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
      
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");
      const items: DocxItem[] = [];

      const bodyChildren = Array.from(doc.body?.children ?? []);
      let serviceListPending = false;
      let lastDayHeadingIndex = -1;

      // Orijinal dökümandaki eski üst başlıkları (Tur adı, havayolu vb.) atlamak için kilit mekanizması.
      // İlk resmi gün başlığı görünene kadar dökümanın en tepesindeki düz yazıları listeye eklemiyoruz.
      let standardContentStarted = false;

      for (let index = 0; index < bodyChildren.length; index += 1) {
        const element = bodyChildren[index];
        const tag = element.tagName;
        const rawText = getTextFromNode(element);
        const text = normalizeImportedText(rawText);

        const isDayHeading = dayHeadingRegex.test(rawText);
        if (isDayHeading) {
          standardContentStarted = true; // İlk gün başlığını gördük, içerik alımı başladı.
          const formattedDayTitle = formatDayHeading(rawText);
          items.push({ type: "subheading", text: formattedDayTitle });
          lastDayHeadingIndex = items.length - 1;
          continue;
        }

        // Eğer dökümanın başındaki eski başlık alanındaysak bu elementleri es geç
        if (!standardContentStarted) {
          continue;
        }

        const isServiceHeading = text === "Fiyatlara Dahil Olan Hizmetler" || text === "Fiyatlara Dahil Olmayan Hizmetler";
        if (isServiceHeading) {
          if (text === "Fiyatlara Dahil Olmayan Hizmetler") {
            items.push({ type: "empty-line" });
          }

          items.push({ type: "paragraph", text, bold: true });
          serviceListPending = true;
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
                  
                  if (lastDayHeadingIndex !== -1 && items.length - 2 === lastDayHeadingIndex) {
                    items.push({ type: "empty-line" });
                    lastDayHeadingIndex = -1; 
                  }
                }
              }
            }
          });
          continue;
        }

        if (tag === "LI") {
          if (text) {
            if (serviceListPending) {
              items.push({ type: "paragraph", text });
            } else {
              items.push({ type: "list", text, level: getListLevel(element) });
              if (lastDayHeadingIndex !== -1 && items.length - 2 === lastDayHeadingIndex) {
                items.push({ type: "empty-line" });
                lastDayHeadingIndex = -1;
              }
            }
          }
          continue;
        }

        if (tag === "P") {
          if (text) {
            items.push({ type: "paragraph", text });
            
            if (lastDayHeadingIndex !== -1 && items.length - 2 === lastDayHeadingIndex) {
              items.push({ type: "empty-line" });
              lastDayHeadingIndex = -1; 
            }
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
        setError('Eşleşen gün başlıkları veya liste maddesi bulunamadı.');
        setIsProcessing(false);
        return;
      }

      const titleParagraphs: Paragraph[] = [];

      if (mainTitleInput.trim()) {
        titleParagraphs.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: mainTitleInput.trim(),
                size: 28,
                bold: true,
                font: "Arial",
              }),
            ],
          })
        );
      }

      if (airline.trim()) {
        titleParagraphs.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `${airline.trim()} Tarifeli Seferi İle`,
                size: 28,
                bold: true,
                font: "Arial",
              }),
            ],
          })
        );
      }

      if (nightInput.trim() && dayInput.trim()) {
        titleParagraphs.push(
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
          new Paragraph({})  
        );
      }

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
                if (item.type === "empty-line") {
                  return new Paragraph({}); 
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

                return new Paragraph({
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
        
        <label htmlFor="mainTitle" className="block mb-2">
          Tur Başlığı
          <input 
            className="block w-full rounded bg-white px-3 py-2 border border-slate-300 mt-1" 
            id="mainTitle" 
            value={mainTitleInput} 
            onChange={(event) => setMainTitleInput(event.target.value)} 
          />
        </label>

        <label htmlFor="airlines" className="block mb-2">Havayolu</label>
        <select 
          id="airlines" 
          className="bg-white w-full block rounded border border-slate-300 px-3 py-2 mb-2" 
          value={airline} 
          onChange={(event) => setAirline(event.target.value)}
        >
          <option value="Türk Hava Yolları">THY</option>
          <option value="Pegasus Hava Yolları">Pegasus</option>
        </select>

        <label htmlFor="nightInput" className="block mb-2">
          Gece
          <input 
            type="text" 
            id="nightInput" 
            value={nightInput} 
            onChange={(event) => setNightInput(event.target.value)} 
            className="bg-white block w-full rounded border border-slate-300 px-3 py-2 mt-1" 
          />
        </label>

        <label htmlFor="dayInput" className="block mb-2">
          Gün
          <input 
            type="text" 
            id="dayInput" 
            value={dayInput} 
            onChange={(event) => setDayInput(event.target.value)} 
            className="bg-white block w-full rounded border border-slate-300 px-3 py-2 mt-1" 
          />
        </label>

        <label className="block mb-4">
          <span className="block text-sm font-medium text-slate-700">
            Yeni Dosya Adı
          </span>
          <input
            type="text"
            value={outputName}
            onChange={(event) => setOutputName(event.target.value)}
            placeholder="ornek-dosya-adi"
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