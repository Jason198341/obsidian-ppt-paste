import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    console.log("[PPT Paste] Plugin loaded v1.2.0");

    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const html = cd.getData("text/html");
        const rtf = cd.getData("text/rtf");

        // Diagnostic logging — visible in Obsidian developer console (Ctrl+Shift+I)
        console.log("[PPT Paste] === Paste event ===");
        console.log("[PPT Paste] types:", Array.from(cd.types).join(", "));
        console.log(
          "[PPT Paste] files:",
          cd.files.length,
          Array.from({ length: cd.files.length }, (_, i) =>
            `${cd.files[i].type}(${cd.files[i].size}b)`
          ).join(", ")
        );
        console.log("[PPT Paste] html:", html.length, "chars");
        console.log("[PPT Paste] rtf:", rtf.length, "chars");
        if (html.length > 0) {
          console.log("[PPT Paste] html preview:", html.substring(0, 3000));
        }

        if (!this.hasMultipleImages(cd, html, rtf)) {
          console.log("[PPT Paste] Not multi-image, passing through");
          return;
        }

        console.log("[PPT Paste] Multi-image detected → intercepting");
        evt.preventDefault();
        this.extractAndInsert(cd, html, rtf, editor);
      })
    );
  }

  /**
   * Synchronous quick-check: does the clipboard likely contain multiple slide images?
   */
  private hasMultipleImages(cd: DataTransfer, html: string, rtf: string): boolean {
    // Check 1: Multiple image files
    let fileCount = 0;
    for (let i = 0; i < cd.files.length; i++) {
      if (cd.files[i].type.startsWith("image/") && cd.files[i].size >= 3000) fileCount++;
    }
    if (fileCount > 1) {
      console.log("[PPT Paste] detect: files =", fileCount);
      return true;
    }

    if (html) {
      // Check 2: Multiple <img> tags
      const imgTags = (html.match(/<img[\s>]/gi) || []).length;
      if (imgTags > 1) {
        console.log("[PPT Paste] detect: img tags =", imgTags);
        return true;
      }

      // Check 3: Multiple base64 data URIs
      const dataUris = (html.match(/data:image\/[\w+]+;base64,/gi) || []).length;
      if (dataUris > 1) {
        console.log("[PPT Paste] detect: data URIs =", dataUris);
        return true;
      }

      // Check 4: Multiple src attributes with image extensions (any protocol)
      const srcRefs = (html.match(/src=["'][^"']*\.(?:png|jpg|jpeg|gif|bmp|emf|wmf|tif|tiff)/gi) || []).length;
      if (srcRefs > 1) {
        console.log("[PPT Paste] detect: src image refs =", srcRefs);
        return true;
      }

      // Check 5: VML imagedata tags (PPT uses these)
      const vmlImages = (html.match(/<v:imagedata[\s>]/gi) || []).length;
      if (vmlImages > 1) {
        console.log("[PPT Paste] detect: VML imagedata =", vmlImages);
        return true;
      }

      // Check 6: PPT-specific clip_image pattern (unique filenames)
      const clipImages = new Set(html.match(/clip_image\d+/gi) || []).size;
      if (clipImages > 1) {
        console.log("[PPT Paste] detect: clip_image =", clipImages);
        return true;
      }
    }

    if (rtf) {
      // Check 7: RTF blip markers
      const rtfImages = (rtf.match(/\\(pngblip|jpegblip|emfblip)/g) || []).length;
      if (rtfImages > 1) {
        console.log("[PPT Paste] detect: RTF blips =", rtfImages);
        return true;
      }
    }

    console.log("[PPT Paste] detect: no multi-image signals found");
    return false;
  }

  /**
   * Async extraction — try all strategies, use the one that yields the most images.
   */
  private async extractAndInsert(cd: DataTransfer, html: string, rtf: string, editor: any) {
    const candidates: SlideImage[][] = [];

    // Strategy 1: Clipboard files
    const s1 = await this.fromFiles(cd);
    console.log("[PPT Paste] S1 files:", s1.length);
    candidates.push(s1);

    // Strategy 2: HTML base64 data URIs
    if (html) {
      const s2 = this.fromHtmlBase64(html);
      console.log("[PPT Paste] S2 base64:", s2.length);
      candidates.push(s2);
    }

    // Strategy 3: HTML src URLs → Node.js fs for file:// paths
    if (html) {
      const s3 = await this.fromHtmlUrls(html);
      console.log("[PPT Paste] S3 URLs:", s3.length);
      candidates.push(s3);
    }

    // Strategy 4: RTF embedded images (robust parser)
    if (rtf) {
      const s4 = this.fromRtf(rtf);
      console.log("[PPT Paste] S4 RTF:", s4.length);
      candidates.push(s4);
    }

    const images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    console.log("[PPT Paste] Best result:", images.length, "images");

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice("Could not extract slide images from clipboard.\nOpen console (Ctrl+Shift+I) for diagnostics.");
      console.log("[PPT Paste] FAILED — no images from any strategy");
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
      } catch (e) {
        console.log("[PPT Paste] S1 error:", e);
      }
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
      } catch (e) {
        console.log("[PPT Paste] S2 error:", e);
      }
    }

    return images;
  }

  // ─── Strategy 3: HTML src URLs → Node.js fs ───────────────

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];

    // Match ALL src attributes (not just file:// or blob:)
    const regex = /src=["']([^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(html)) !== null) {
      const url = match[1];
      if (url.startsWith("data:")) continue; // handled by S2
      urls.push(url);
    }

    console.log("[PPT Paste] S3 found URLs:", urls.length,
      urls.map(u => u.substring(0, 120)));

    for (const url of urls) {
      try {
        if (url.startsWith("file:///") || url.startsWith("file://")) {
          const data = this.readLocalFile(url);
          if (data && data.length >= 3000) {
            images.push({ data, ext: this.extFromPath(url) });
          }
        } else if (url.startsWith("blob:")) {
          const resp = await fetch(url);
          if (!resp.ok) continue;
          const buf = await resp.arrayBuffer();
          const data = new Uint8Array(buf);
          if (data.length < 3000) continue;
          const ct = resp.headers.get("content-type") || "image/png";
          images.push({ data, ext: this.extFromMime(ct) });
        }
      } catch (e) {
        console.log("[PPT Paste] S3 error:", url.substring(0, 80), e);
      }
    }

    return images;
  }

  /**
   * Read a file:// URL using Node.js fs (Electron has access).
   * fetch("file:///...") is blocked in many Electron configs,
   * but fs.readFileSync always works.
   */
  private readLocalFile(fileUrl: string): Uint8Array | null {
    try {
      let filePath: string;
      try {
        const urlMod = require("url");
        filePath = urlMod.fileURLToPath(fileUrl);
      } catch {
        // Manual fallback
        filePath = decodeURIComponent(fileUrl.replace(/^file:\/\/\//, ""));
      }

      console.log("[PPT Paste] Reading:", filePath);
      const fs = require("fs");
      const buffer: Buffer = fs.readFileSync(filePath);
      // Safe copy — Node.js Buffer shares an ArrayBuffer pool
      return Uint8Array.from(buffer);
    } catch (e) {
      console.log("[PPT Paste] fs read failed:", e);
      return null;
    }
  }

  // ─── Strategy 4: RTF Embedded Images (robust parser) ──────

  private fromRtf(rtf: string): SlideImage[] {
    const images: SlideImage[] = [];
    const blipTypes = ["pngblip", "jpegblip"];

    for (const blipType of blipTypes) {
      const marker = "\\" + blipType;
      let searchFrom = 0;

      while (true) {
        const pos = rtf.indexOf(marker, searchFrom);
        if (pos === -1) break;
        searchFrom = pos + marker.length;

        // Extract hex block — handles control words between blip marker and hex data
        const hex = this.extractHexBlock(rtf, searchFrom);
        if (!hex || hex.length < 6000) {
          console.log("[PPT Paste] S4 skip: hex too short", hex?.length || 0);
          continue;
        }

        try {
          const bytes = new Uint8Array(hex.length / 2);
          for (let i = 0; i < hex.length; i += 2) {
            bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
          }
          images.push({ data: bytes, ext: blipType === "jpegblip" ? "jpg" : "png" });
        } catch (e) {
          console.log("[PPT Paste] S4 decode error:", e);
        }
      }
    }

    return images;
  }

  /**
   * Extract hex data block from RTF after a blip marker.
   *
   * PowerPoint RTF structure:
   *   \jpegblip\bliptag-123456789\blipupi96
   *   ffd8ffe000104a464946...
   *   }
   *
   * The hex data comes AFTER control words (\keyword, \keyword123, \keyword-123).
   * We skip those and find the continuous hex block.
   */
  private extractHexBlock(rtf: string, startPos: number): string | null {
    let i = startPos;
    const limit = Math.min(startPos + 5000, rtf.length);

    // Phase 1: Skip whitespace and control words to find hex data start
    while (i < limit) {
      const ch = rtf[i];

      // End of pict group
      if (ch === "}") return null;

      // Skip whitespace
      if (ch === " " || ch === "\r" || ch === "\n" || ch === "\t") {
        i++;
        continue;
      }

      // Skip RTF control words: \keyword or \keyword123 or \keyword-123
      if (ch === "\\") {
        i++;
        // Skip letter sequence
        while (i < limit && rtf[i] >= "a" && rtf[i] <= "z") i++;
        // Skip optional numeric parameter (with optional minus sign)
        if (i < limit && (rtf[i] === "-" || (rtf[i] >= "0" && rtf[i] <= "9"))) {
          if (rtf[i] === "-") i++;
          while (i < limit && rtf[i] >= "0" && rtf[i] <= "9") i++;
        }
        // Skip single space delimiter after control word
        if (i < limit && rtf[i] === " ") i++;
        continue;
      }

      // Possible hex data start — verify it's a long hex run
      if (this.isHexChar(ch)) {
        let hexCount = 0;
        let j = i;
        while (j < rtf.length) {
          const c = rtf[j];
          if (this.isHexChar(c)) {
            hexCount++;
            j++;
          } else if (c === " " || c === "\r" || c === "\n" || c === "\t") {
            j++;
          } else {
            break;
          }
        }

        if (hexCount >= 100) {
          // Found real hex data block
          const endPos = Math.min(j, rtf.indexOf("}", i) === -1 ? j : rtf.indexOf("}", i));
          const raw = rtf.substring(i, endPos);
          return raw.replace(/[\s\r\n]/g, "");
        }

        // Short hex — part of something else, skip it
        i = j;
        continue;
      }

      // Skip any other characters (braces in groups, etc.)
      i++;
    }

    return null;
  }

  private isHexChar(ch: string): boolean {
    return (
      (ch >= "0" && ch <= "9") ||
      (ch >= "a" && ch <= "f") ||
      (ch >= "A" && ch <= "F")
    );
  }

  // ─── Shared Utilities ──────────────────────────────────────

  private extFromMime(mime: string): string {
    if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg";
    if (mime.includes("gif")) return "gif";
    if (mime.includes("webp")) return "webp";
    return "png";
  }

  private extFromPath(filePath: string): string {
    const m = filePath.match(/\.(png|jpg|jpeg|gif|webp)$/i);
    if (!m) return "png";
    const ext = m[1].toLowerCase();
    return ext === "jpeg" ? "jpg" : ext;
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

      // Safe buffer — ensure clean ArrayBuffer for Obsidian vault API
      const buf = img.data.buffer.byteLength === img.data.byteLength
        ? img.data.buffer
        : img.data.buffer.slice(img.data.byteOffset, img.data.byteOffset + img.data.byteLength);

      await this.app.vault.createBinary(filePath, buf);
      lines.push(`![[${fileName}]]`);
    }

    const cursor = editor.getCursor();
    editor.replaceRange(lines.join("\n") + "\n", cursor);

    new Notice(`${images.length} slides pasted!`);
    console.log("[PPT Paste] Done:", images.length, "images inserted");
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
