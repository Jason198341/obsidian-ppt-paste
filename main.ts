import { Plugin, Notice, TFile, normalizePath } from "obsidian";

export default class PPTSlidePaste extends Plugin {
  async onload() {
    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        // Strategy 1: Check text/html for multiple embedded base64 images (PPT's format)
        const html = evt.clipboardData.getData("text/html");
        if (html) {
          const base64Images = this.extractBase64Images(html);
          if (base64Images.length > 1) {
            // Multiple images in HTML — this is PPT multi-slide paste
            evt.preventDefault();
            this.saveAndInsertImages(base64Images, editor);
            return;
          }
        }

        // Strategy 2: Check text/rtf — PPT also embits images in RTF
        const rtf = evt.clipboardData.getData("text/rtf");
        if (rtf) {
          const rtfImages = this.extractRtfImages(rtf);
          if (rtfImages.length > 1) {
            evt.preventDefault();
            this.saveAndInsertImages(rtfImages, editor);
            return;
          }
        }

        // Strategy 3: Fallback — if HTML had exactly 1 image and there's also
        // a blob, just let Obsidian handle it normally (single image paste)
      })
    );
  }

  /**
   * Extract base64-encoded images from HTML clipboard content.
   * PPT puts <img src="data:image/png;base64,..."> or
   * <v:imagedata src="data:image/png;base64,..."> for each slide.
   */
  private extractBase64Images(html: string): { data: Uint8Array; ext: string }[] {
    const images: { data: Uint8Array; ext: string }[] = [];

    // Match data URI patterns: data:image/TYPE;base64,DATA
    // These appear in src="..." or src='...'
    const regex = /data:image\/(png|jpeg|jpg|gif|bmp|webp|emf|wmf);base64,([A-Za-z0-9+/=\s]+)/g;
    let match;

    while ((match = regex.exec(html)) !== null) {
      const mimeSubtype = match[1];
      const b64 = match[2].replace(/\s/g, "");

      try {
        const binary = atob(b64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }

        // Skip tiny images (likely icons/bullets, not slides)
        if (bytes.length < 5000) continue;

        const ext = (mimeSubtype === "jpeg" || mimeSubtype === "jpg") ? "jpg" : "png";
        images.push({ data: bytes, ext });
      } catch {
        // Invalid base64, skip
      }
    }

    return images;
  }

  /**
   * Extract images embedded in RTF content.
   * RTF embeds images as hex-encoded data after \pngblip or \jpegblip
   */
  private extractRtfImages(rtf: string): { data: Uint8Array; ext: string }[] {
    const images: { data: Uint8Array; ext: string }[] = [];

    // Match {\pict ... \pngblip HEXDATA} or \jpegblip
    const regex = /\\(pngblip|jpegblip)\s*\r?\n?([0-9a-fA-F\s]+)/g;
    let match;

    while ((match = regex.exec(rtf)) !== null) {
      const type = match[1];
      const hex = match[2].replace(/\s/g, "");

      // Must have substantial data to be a slide
      if (hex.length < 10000) continue;

      try {
        const bytes = new Uint8Array(hex.length / 2);
        for (let i = 0; i < hex.length; i += 2) {
          bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
        }

        const ext = type === "jpegblip" ? "jpg" : "png";
        images.push({ data: bytes, ext });
      } catch {
        // Invalid hex, skip
      }
    }

    return images;
  }

  private async saveAndInsertImages(
    images: { data: Uint8Array; ext: string }[],
    editor: any
  ) {
    const activeFile = this.app.workspace.getActiveFile();
    if (!activeFile) {
      new Notice("No active file");
      return;
    }

    const attachFolder = await this.getAttachmentFolder(activeFile);
    const timestamp = Date.now();
    const lines: string[] = [];

    new Notice(`Pasting ${images.length} slide images...`);

    for (let i = 0; i < images.length; i++) {
      const img = images[i];
      const fileName = `slide-${timestamp}-${String(i + 1).padStart(2, "0")}.${img.ext}`;
      const filePath = normalizePath(
        attachFolder ? `${attachFolder}/${fileName}` : fileName
      );

      await this.app.vault.createBinary(filePath, img.data.buffer);
      lines.push(`![[${fileName}]]`);
    }

    const cursor = editor.getCursor();
    editor.replaceRange(lines.join("\n\n") + "\n", cursor);

    new Notice(`${images.length} slides pasted!`);
  }

  private async getAttachmentFolder(activeFile: TFile): Promise<string> {
    // @ts-ignore — internal API
    const attachPath: string = this.app.vault.getConfig("attachmentFolderPath") || "/";

    if (attachPath === "/") return "";

    if (attachPath === "./") {
      return activeFile.parent?.path || "";
    }

    if (attachPath.startsWith("./")) {
      const parentPath = activeFile.parent?.path || "";
      const sub = attachPath.slice(2);
      const full = parentPath ? `${parentPath}/${sub}` : sub;
      await this.ensureFolder(full);
      return full;
    }

    await this.ensureFolder(attachPath);
    return attachPath;
  }

  private async ensureFolder(path: string) {
    const p = normalizePath(path);
    if (!this.app.vault.getAbstractFileByPath(p)) {
      await this.app.vault.createFolder(p);
    }
  }
}
