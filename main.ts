import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    console.log("[PPT Paste] Plugin loaded v1.3.0");

    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const types = Array.from(cd.types);
        const html = cd.getData("text/html");
        const rtf = cd.getData("text/rtf");
        const svg = cd.getData("image/svg+xml");

        console.log("[PPT Paste] === Paste event ===");
        console.log("[PPT Paste] types:", types.join(", "));
        console.log(
          "[PPT Paste] files:",
          cd.files.length,
          Array.from({ length: cd.files.length }, (_, i) =>
            `${cd.files[i].type}(${cd.files[i].size}b)`
          ).join(", ")
        );
        console.log("[PPT Paste] html:", html.length, "| rtf:", rtf.length, "| svg:", svg.length);
        if (svg.length > 0) {
          console.log("[PPT Paste] svg preview:", svg.substring(0, 3000));
        }

        const isPpt = types.includes("ppt/slides");
        const hasMulti = this.hasMultipleImages(cd, html, rtf, svg);

        if (!isPpt && !hasMulti) {
          console.log("[PPT Paste] Not PPT / not multi-image → pass through");
          return;
        }

        console.log("[PPT Paste] Intercepting (isPpt=%s, hasMulti=%s)", isPpt, hasMulti);
        evt.preventDefault();
        this.extractAndInsert(cd, html, rtf, svg, isPpt, editor);
      })
    );
  }

  // ─── Detection ─────────────────────────────────────────────

  private hasMultipleImages(
    cd: DataTransfer,
    html: string,
    rtf: string,
    svg: string
  ): boolean {
    // 1. Multiple image files
    let fileCount = 0;
    for (let i = 0; i < cd.files.length; i++) {
      if (cd.files[i].type.startsWith("image/") && cd.files[i].size >= 3000)
        fileCount++;
    }
    if (fileCount > 1) {
      console.log("[PPT Paste] detect: files =", fileCount);
      return true;
    }

    // 2. SVG with multiple embedded images
    if (svg) {
      const svgImageTags = (svg.match(/<image[\s>]/gi) || []).length;
      if (svgImageTags > 1) {
        console.log("[PPT Paste] detect: SVG <image> =", svgImageTags);
        return true;
      }
      const svgDataUris = (svg.match(/data:image\/[\w+]+;base64,/gi) || []).length;
      if (svgDataUris > 1) {
        console.log("[PPT Paste] detect: SVG data URIs =", svgDataUris);
        return true;
      }
    }

    // 3. HTML checks
    if (html) {
      const imgTags = (html.match(/<img[\s>]/gi) || []).length;
      if (imgTags > 1) return true;

      const dataUris = (html.match(/data:image\/[\w+]+;base64,/gi) || []).length;
      if (dataUris > 1) return true;

      const srcRefs = (
        html.match(
          /src=["'][^"']*\.(?:png|jpg|jpeg|gif|bmp|emf|wmf|tif|tiff)/gi
        ) || []
      ).length;
      if (srcRefs > 1) return true;

      const vmlImages = (html.match(/<v:imagedata[\s>]/gi) || []).length;
      if (vmlImages > 1) return true;

      const clipImages = new Set(html.match(/clip_image\d+/gi) || []).size;
      if (clipImages > 1) return true;
    }

    // 4. RTF blip markers
    if (rtf) {
      const rtfImages = (
        rtf.match(/\\(pngblip|jpegblip|emfblip)/g) || []
      ).length;
      if (rtfImages > 1) return true;
    }

    console.log("[PPT Paste] detect: no multi-image signals");
    return false;
  }

  // ─── Extraction orchestrator ───────────────────────────────

  private async extractAndInsert(
    cd: DataTransfer,
    html: string,
    rtf: string,
    svg: string,
    isPpt: boolean,
    editor: any
  ) {
    const candidates: SlideImage[][] = [];

    // Strategy 1: SVG embedded base64 images (PPT primary path)
    if (svg) {
      const s1 = this.fromSvgBase64(svg);
      console.log("[PPT Paste] S1 SVG base64:", s1.length);
      candidates.push(s1);
    }

    // Strategy 2: Clipboard files
    const s2 = await this.fromFiles(cd);
    console.log("[PPT Paste] S2 files:", s2.length);
    candidates.push(s2);

    // Strategy 3: HTML base64 data URIs
    if (html) {
      const s3 = this.fromHtmlBase64(html);
      console.log("[PPT Paste] S3 HTML base64:", s3.length);
      candidates.push(s3);
    }

    // Strategy 4: HTML src URLs → Node.js fs
    if (html) {
      const s4 = await this.fromHtmlUrls(html);
      console.log("[PPT Paste] S4 HTML URLs:", s4.length);
      candidates.push(s4);
    }

    // Strategy 5: RTF embedded images
    if (rtf) {
      const s5 = this.fromRtf(rtf);
      console.log("[PPT Paste] S5 RTF:", s5.length);
      candidates.push(s5);
    }

    // Strategy 6: SVG src/href URLs → Node.js fs
    if (svg) {
      const s6 = await this.fromSvgUrls(svg);
      console.log("[PPT Paste] S6 SVG URLs:", s6.length);
      candidates.push(s6);
    }

    // Pick the strategy with the most images
    let images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    console.log("[PPT Paste] Best:", images.length, "images");

    // Fallback for PPT: if no multi-image found, paste the single file
    if (images.length === 0 && isPpt && cd.files.length > 0) {
      console.log("[PPT Paste] PPT fallback: pasting single file");
      images = await this.fromFiles(cd, /* allowSmall */ true);
    }

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice(
        "Could not extract slide images.\nOpen console (Ctrl+Shift+I) for diagnostics."
      );
      console.log("[PPT Paste] FAILED — no images from any strategy");
    }
  }

  // ─── Strategy 1: SVG Embedded Base64 ──────────────────────

  private fromSvgBase64(svg: string): SlideImage[] {
    const images: SlideImage[] = [];
    // Match href="data:image/..." or xlink:href="data:image/..."
    const regex = /(?:xlink:)?href=["'](data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+))["']/g;
    let match;

    while ((match = regex.exec(svg)) !== null) {
      const mimeType = match[2];
      const b64 = match[3].replace(/\s/g, "");
      try {
        const binary = atob(b64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        if (bytes.length < 3000) continue;
        images.push({ data: bytes, ext: this.extFromMime(`image/${mimeType}`) });
      } catch (e) {
        console.log("[PPT Paste] S1 decode error:", e);
      }
    }

    // Fallback: also try the generic data URI pattern (no href wrapper)
    if (images.length === 0) {
      return this.fromGenericBase64(svg);
    }

    return images;
  }

  // ─── Strategy 6: SVG URL References ───────────────────────

  private async fromSvgUrls(svg: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    // Match href="file:///..." or xlink:href="file:///..."
    const regex = /(?:xlink:)?href=["']((?:file:\/\/\/)[^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(svg)) !== null) {
      urls.push(match[1]);
    }

    if (urls.length > 0) {
      console.log("[PPT Paste] S6 SVG URLs found:", urls.length);
    }

    for (const url of urls) {
      try {
        const data = this.readLocalFile(url);
        if (data && data.length >= 3000) {
          images.push({ data, ext: this.extFromPath(url) });
        }
      } catch (e) {
        console.log("[PPT Paste] S6 error:", e);
      }
    }

    return images;
  }

  // ─── Strategy 2: Clipboard Files ──────────────────────────

  private async fromFiles(
    cd: DataTransfer,
    allowSmall = false
  ): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    for (let i = 0; i < cd.files.length; i++) {
      const file = cd.files[i];
      if (!file.type.startsWith("image/")) continue;
      try {
        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);
        if (!allowSmall && data.length < 3000) continue;
        images.push({ data, ext: this.extFromMime(file.type) });
      } catch (e) {
        console.log("[PPT Paste] S2 error:", e);
      }
    }
    return images;
  }

  // ─── Strategy 3: HTML Base64 Data URIs ────────────────────

  private fromHtmlBase64(html: string): SlideImage[] {
    return this.fromGenericBase64(html);
  }

  private fromGenericBase64(text: string): SlideImage[] {
    const images: SlideImage[] = [];
    const regex = /data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+)/g;
    let match;

    while ((match = regex.exec(text)) !== null) {
      const mimeType = match[1];
      const b64 = match[2].replace(/\s/g, "");
      try {
        const binary = atob(b64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
          bytes[i] = binary.charCodeAt(i);
        }
        if (bytes.length < 3000) continue;
        images.push({
          data: bytes,
          ext: this.extFromMime(`image/${mimeType}`),
        });
      } catch (e) {
        console.log("[PPT Paste] base64 decode error:", e);
      }
    }

    return images;
  }

  // ─── Strategy 4: HTML src URLs → Node.js fs ───────────────

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    const regex = /src=["']([^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(html)) !== null) {
      const url = match[1];
      if (url.startsWith("data:")) continue;
      urls.push(url);
    }

    if (urls.length > 0) {
      console.log(
        "[PPT Paste] S4 URLs:",
        urls.length,
        urls.map((u) => u.substring(0, 120))
      );
    }

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
        console.log("[PPT Paste] S4 error:", url.substring(0, 80), e);
      }
    }

    return images;
  }

  private readLocalFile(fileUrl: string): Uint8Array | null {
    try {
      let filePath: string;
      try {
        const urlMod = require("url");
        filePath = urlMod.fileURLToPath(fileUrl);
      } catch {
        filePath = decodeURIComponent(fileUrl.replace(/^file:\/\/\//, ""));
      }

      console.log("[PPT Paste] Reading:", filePath);
      const fs = require("fs");
      const buffer: Buffer = fs.readFileSync(filePath);
      return Uint8Array.from(buffer);
    } catch (e) {
      console.log("[PPT Paste] fs read failed:", e);
      return null;
    }
  }

  // ─── Strategy 5: RTF Embedded Images ──────────────────────

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

        const hex = this.extractHexBlock(rtf, searchFrom);
        if (!hex || hex.length < 6000) continue;

        try {
          const bytes = new Uint8Array(hex.length / 2);
          for (let i = 0; i < hex.length; i += 2) {
            bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
          }
          images.push({
            data: bytes,
            ext: blipType === "jpegblip" ? "jpg" : "png",
          });
        } catch (e) {
          console.log("[PPT Paste] S5 decode error:", e);
        }
      }
    }

    return images;
  }

  private extractHexBlock(rtf: string, startPos: number): string | null {
    let i = startPos;
    const limit = Math.min(startPos + 5000, rtf.length);

    while (i < limit) {
      const ch = rtf[i];

      if (ch === "}") return null;

      if (ch === " " || ch === "\r" || ch === "\n" || ch === "\t") {
        i++;
        continue;
      }

      if (ch === "\\") {
        i++;
        while (i < limit && rtf[i] >= "a" && rtf[i] <= "z") i++;
        if (
          i < limit &&
          (rtf[i] === "-" || (rtf[i] >= "0" && rtf[i] <= "9"))
        ) {
          if (rtf[i] === "-") i++;
          while (i < limit && rtf[i] >= "0" && rtf[i] <= "9") i++;
        }
        if (i < limit && rtf[i] === " ") i++;
        continue;
      }

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
          const endPos = Math.min(
            j,
            rtf.indexOf("}", i) === -1 ? j : rtf.indexOf("}", i)
          );
          return rtf.substring(i, endPos).replace(/[\s\r\n]/g, "");
        }

        i = j;
        continue;
      }

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

      const buf =
        img.data.buffer.byteLength === img.data.byteLength
          ? img.data.buffer
          : img.data.buffer.slice(
              img.data.byteOffset,
              img.data.byteOffset + img.data.byteLength
            );

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
