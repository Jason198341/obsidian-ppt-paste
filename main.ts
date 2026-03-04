import { Plugin, MarkdownView, Notice, TFile, normalizePath } from "obsidian";

export default class PPTSlidePaste extends Plugin {
  async onload() {
    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        const items = evt.clipboardData?.items;
        if (!items) return;

        // Collect all image items from clipboard
        const imageItems: DataTransferItem[] = [];
        for (let i = 0; i < items.length; i++) {
          if (items[i].type.startsWith("image/")) {
            imageItems.push(items[i]);
          }
        }

        // Also check for HTML with embedded images (PPT often sends this)
        let htmlContent = "";
        for (let i = 0; i < items.length; i++) {
          if (items[i].type === "text/html") {
            // We'll handle this async
            items[i].getAsString((html) => {
              htmlContent = html;
            });
          }
        }

        // If multiple images or at least one image, intercept
        if (imageItems.length === 0) return;

        evt.preventDefault();

        // Process all clipboard image blobs
        this.handleImagePaste(imageItems, htmlContent, editor);
      })
    );
  }

  private async handleImagePaste(
    imageItems: DataTransferItem[],
    _htmlContent: string,
    editor: any
  ) {
    const activeFile = this.app.workspace.getActiveFile();
    if (!activeFile) {
      new Notice("No active file");
      return;
    }

    // Get attachment folder path
    const attachFolder = await this.getAttachmentFolder(activeFile);

    const blobs: Blob[] = [];
    for (const item of imageItems) {
      const blob = item.getAsFile();
      if (blob) blobs.push(blob);
    }

    if (blobs.length === 0) {
      new Notice("No images found in clipboard");
      return;
    }

    new Notice(`Pasting ${blobs.length} slide image${blobs.length > 1 ? "s" : ""}...`);

    const lines: string[] = [];
    const timestamp = Date.now();

    for (let i = 0; i < blobs.length; i++) {
      const blob = blobs[i];
      const ext = this.getExtension(blob.type);
      const fileName = `slide-${timestamp}-${String(i + 1).padStart(2, "0")}.${ext}`;
      const filePath = normalizePath(`${attachFolder}/${fileName}`);

      // Read blob as ArrayBuffer and save
      const buffer = await blob.arrayBuffer();
      await this.app.vault.createBinary(filePath, buffer);

      // Build markdown image embed
      lines.push(`![[${fileName}]]`);
    }

    // Insert all image embeds at cursor
    const cursor = editor.getCursor();
    const insertText = lines.join("\n\n") + "\n";
    editor.replaceRange(insertText, cursor);

    new Notice(`${blobs.length} slide${blobs.length > 1 ? "s" : ""} pasted!`);
  }

  private getExtension(mimeType: string): string {
    if (mimeType === "image/png") return "png";
    if (mimeType === "image/jpeg") return "jpg";
    if (mimeType === "image/gif") return "gif";
    if (mimeType === "image/webp") return "webp";
    if (mimeType === "image/bmp") return "bmp";
    if (mimeType === "image/svg+xml") return "svg";
    return "png";
  }

  private async getAttachmentFolder(activeFile: TFile): Promise<string> {
    // Respect Obsidian's attachment folder setting
    // @ts-ignore — internal API
    const attachPath = this.app.vault.getConfig("attachmentFolderPath");

    if (!attachPath || attachPath === "/") {
      // Root of vault
      return "";
    }

    if (attachPath === "./") {
      // Same folder as current file
      return activeFile.parent?.path || "";
    }

    if (attachPath.startsWith("./")) {
      // Subfolder relative to current file
      const parentPath = activeFile.parent?.path || "";
      const subFolder = attachPath.slice(2);
      const fullPath = parentPath ? `${parentPath}/${subFolder}` : subFolder;
      await this.ensureFolder(fullPath);
      return fullPath;
    }

    // Absolute path in vault
    await this.ensureFolder(attachPath);
    return attachPath;
  }

  private async ensureFolder(path: string) {
    const normalPath = normalizePath(path);
    if (!this.app.vault.getAbstractFileByPath(normalPath)) {
      await this.app.vault.createFolder(normalPath);
    }
  }
}
