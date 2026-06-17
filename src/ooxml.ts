export type NumberingMode = "plain" | "chapter";

export interface EquationOptions {
  mode: NumberingMode;
  separator: string;
  bookmarkName: string;
}

const WORD_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MATH_NS =
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fieldRun(instruction: string): string {
  return `<w:fldSimple w:instr="${escapeXml(instruction)}"><w:r><w:t>?</w:t></w:r></w:fldSimple>`;
}

function chapterField(separator: string): string {
  const sep = separator === "." ? "." : separator === ":" ? ":" : "-";
  return `${fieldRun('STYLEREF 1 \\s')}<w:r><w:t>${escapeXml(sep)}</w:t></w:r>`;
}

function captionRuns(options: EquationOptions): string {
  const chapter = options.mode === "chapter" ? chapterField(options.separator) : "";
  return [
    '<w:r><w:t>(</w:t></w:r>',
    `<w:bookmarkStart w:id="0" w:name="${escapeXml(options.bookmarkName)}"/>`,
    chapter,
    fieldRun("SEQ Equation \\* ARABIC"),
    '<w:bookmarkEnd w:id="0"/>',
    '<w:r><w:t>)</w:t></w:r>'
  ].join("");
}

function emptyEquation(): string {
  return `
    <m:oMathPara>
      <m:oMath>
        <m:r>
          <w:rPr>
            <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>
          </w:rPr>
          <m:t>□</m:t>
        </m:r>
      </m:oMath>
    </m:oMathPara>`;
}

export function buildEquationParagraphOoxml(options: EquationOptions): string {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${WORD_NS} ${MATH_NS}>
      <w:body>
        <w:p>
          <w:pPr>
            <w:tabs>
              <w:tab w:val="center" w:pos="4680"/>
              <w:tab w:val="right" w:pos="9360"/>
            </w:tabs>
            <w:spacing w:before="80" w:after="80"/>
          </w:pPr>
          <w:r><w:tab/></w:r>
          <w:r>${emptyEquation()}</w:r>
          <w:r><w:tab/></w:r>
          ${captionRuns(options)}
        </w:p>
      </w:body>
    </w:document>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
        <pkg:xmlData>${documentXml}</pkg:xmlData>
      </pkg:part>
    </pkg:package>`;
}

export function buildReferenceText(bookmarkName: string): string {
  return `REF ${bookmarkName} \\h`;
}
