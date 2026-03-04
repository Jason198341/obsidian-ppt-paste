import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const html = cd.getData("text/html");
        const rtf = cd.getData("text/rtf");

        // Synchronous detection — must decide before event propagates
        if (!this.hasMultipleImages(cd, html, rtf)) return;

        evt.preventDefault();
        this.extractAndInsert(cd, html, rtf, editor);
      })
    );
  }

  /**
   * Synchronous quick-check: does the clipboard likely contain multiple slide images?
   * Checks all possible sources without heavy parsing.
   */
  private hasMultipleImages(cd: DataTransfer, html: string, rtf: string): boolean {
    // Check 1: Multiple image files in clipboard
    let fileCount = 0;
    for (let i = 0; i < cd.files.length; i++) {
      if (cd.files[i].type.startsWith("image/") && cd.files[i].size >= 3000) fileCount++;
    }
    if (fileCount > 1) return true;

    // Check 2: Multiple <img> tags or data URIs in HTML
    if (html) {
      const imgTags = (html.match(/<img[\s>]/gi) || []).length;
      if (imgTags > 1) return true;

      const dataUris = (html.match(/data:image\/[\w+]+;base64,/gi) || []).length;
      if (dataUris > 1) return true;

      // Check for multiple file:// or blob: image references
      const fileRefs = (html.match(/src=["'](?:file:\/\/\/|blob:)[^"']*\.(?:png|jpg|jpeg|gif|bmp|emf|wmf)/gi) || []).length;
      if (fileRefs > 1) return true;

      // Check for VML imagedata tags (PPT uses these)
      const vmlImages = (html.match(/<v:imagedata[\s>]/gi) || []).length;
      if (vmlImages > 1) return true;
    }

    // Check 3: Multiple image markers in RTF
    if (rtf) {
      const rtfImages = (rtf.match(/\\(pngblip|jpegblip|emfblip)/g) || []).length;
      if (rtfImages > 1) return true;
    }

    return false;
  }

  /**
   * Async extraction — try all strategies, use the one that yields the most images.
   */
  private async extractAndInsert(cd: DataTransfer, html: string, rtf: string, editor: any) {
    const candidates: SlideImage[][] = [];

    // Strategy 1: Clipboard files (most reliable on Windows)
    candidates.push(await this.fromFiles(cd));

    // Strategy 2: HTML data: URIs (base64)
    if (html) candidates.push(this.fromHtmlBase64(html));

    // Strategy 3: HTML file:///blob: URLs (PPT temp files)
    if (html) candidates.push(await this.fromHtmlUrls(html));

    // Strategy 4: RTF embedded images
    if (rtf) candidates.push(this.fromRtf(rtf));

    // Pick the strategy that found the most images
    const images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice("Could not extract slide images from clipboard");
    }
  }

  // ─── Strategy 1: Clipboard Files ───────────────────────────

  private async fromFiles(cd: DataTransfer): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    for (let i = 0; i < cd.files.length; i++) {
      const file = cd.files[i];
      if (!file.type.startsWith("image/")) continue;
      try {
        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);
        if (data.length < 3000) continue;
        images.push({ data, ext: this.extFromMime(file.type) });
      } catch { /* skip */ }
    }
    return images;
  }

  // ─── Strategy 2: HTML Base64 Data URIs ─────────────────────

  private fromHtmlBase64(html: string): SlideImage[] {
    const images: SlideImage[] = [];
    const regex = /data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+)/g;
    let match;

    while ((match = regex.exec(html)) !== null) {
      const mimeType = match[1];
      const b64 = match[2].replace(/\s/g, "");
      try {
        const binary = atob(b64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        if (bytes.length < 3000) continue;
        images.push({ data: bytes, ext: this.extFromMime(`image/${mimeType}`) });
      } catch { /* skip */ }
    }

    return images;
  }

  // ─── Strategy 3: HTML file:///blob: URLs ───────────────────

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    // Match src="file:///..." or src="blob:..." in <img> or <v:imagedata>
    const regex = /src=["']((?:file:\/\/\/|blob:)[^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(html)) !== null) {
      urls.push(match[1]);
    }

    for (const url of urls) {
      try {
        const resp = await fetch(url);
        if (!resp.ok) continue;
        const buf = await resp.arrayBuffer();
        const data = new Uint8Array(buf);
        if (data.length < 3000) continue;
        const ct = resp.headers.get("content-type") || "image/png";
        images.push({ data, ext: this.extFromMime(ct) });
      } catch { /* skip */ }
    }

    return images;
  }

  // ─── Strategy 4: RTF Embedded Images ───────────────────────

  private fromRtf(rtf: string): SlideImage[] {
    const images: SlideImage[] = [];
    const regex = /\\(pngblip|jpegblip)\s*\r?\n?([0-9a-fA-F\s]+)/g;
    let match;

    while ((match = regex.exec(rtf)) !== null) {
      const type = match[1];
      const hex = match[2].replace(/\s/g, "");
      if (hex.length < 6000) continue;
      try {
        const bytes = new Uint8Array(hex.length / 2);
        for (let i = 0; i < hex.length; i += 2) {
          bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
        }
        images.push({ data: bytes, ext: type === "jpegblip" ? "jpg" : "png" });
      } catch { /* skip */ }
    }

    return images;
  }

  // ─── Shared Utilities ──────────────────────────────────────

  private extFromMime(mime: string): string {
    if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg";
    if (mime.includes("gif")) return "gif";
    if (mime.includes("webp")) return "webp";
    return "png";
  }

  private async saveAndInsertImages(images: SlideImage[], editor: any) {
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
    editor.replaceRange(lines.join("\n") + "\n", cursor);

    new Notice(`${images.length} slides pasted!`);
  }

  private async getAttachmentFolder(activeFile: TFile): Promise<string> {
    // @ts-ignore — internal API
    const attachPath: string =
      this.app.vault.getConfig("attachmentFolderPath") || "/";

    if (attachPath === "/") return "";
    if (attachPath === "./") return activeFile.parent?.path || "";

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
