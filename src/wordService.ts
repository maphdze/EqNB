import { buildEquationParagraphOoxml, buildReferenceText, NumberingMode } from "./ooxml";

export interface EquationReference {
  id: string;
  bookmarkName: string;
  label: string;
  text: string;
}

export interface InsertEquationRequest {
  mode: NumberingMode;
  separator: string;
}

const EQUATION_TAG = "word-equation-numbering:equation";

function createBookmarkName(): string {
  return `_Eqn${Date.now().toString(36)}${Math.floor(Math.random() * 1000).toString(36)}`;
}

function assertWordReady(): void {
  if (Office.context.host !== Office.HostType.Word) {
    throw new Error("请在 Word 桌面版中打开插件任务窗格。");
  }
}

export async function insertEquationLine(request: InsertEquationRequest): Promise<string> {
  assertWordReady();
  const bookmarkName = createBookmarkName();

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const inserted = selection.insertOoxml(
      buildEquationParagraphOoxml({ ...request, bookmarkName }),
      Word.InsertLocation.replace
    );
    const contentControl = inserted.insertContentControl(Word.ContentControlType.richText);
    contentControl.tag = EQUATION_TAG;
    contentControl.title = `公式 ${bookmarkName}`;
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    context.document.fields.load("items");
    await context.sync();
    context.document.fields.items.forEach((field) => field.updateResult());
    inserted.select();
    await context.sync();
  });

  return bookmarkName;
}

export async function loadEquationReferences(): Promise<EquationReference[]> {
  assertWordReady();

  return Word.run(async (context) => {
    const controls = context.document.contentControls.getByTag(EQUATION_TAG);
    controls.load("items/id,title,text");
    await context.sync();

    return controls.items.map((control, index) => {
      const bookmarkName = control.title.replace(/^公式\s+/, "");
      const text = control.text.trim().replace(/\s+/g, " ");
      return {
        id: String(control.id),
        bookmarkName,
        label: text || `公式 ${index + 1}`,
        text
      };
    });
  });
}

export async function insertEquationReference(bookmarkName: string): Promise<void> {
  assertWordReady();

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText("公式 ", Word.InsertLocation.replace);
    const afterPrefix = context.document.getSelection();
    afterPrefix.insertField(Word.InsertLocation.end, Word.FieldType.ref, buildReferenceText(bookmarkName), false);
    context.document.fields.load("items");
    await context.sync();
    context.document.fields.items.forEach((field) => field.updateResult());
    await context.sync();
  });
}

export async function refreshAllFields(): Promise<number> {
  assertWordReady();

  return Word.run(async (context) => {
    const fields = context.document.fields;
    fields.load("items");
    await context.sync();
    fields.items.forEach((field) => field.updateResult());
    await context.sync();
    return fields.items.length;
  });
}
