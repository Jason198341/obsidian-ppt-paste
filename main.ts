import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    console.log("[PPT Paste] Plugin loaded v1.5.0");

    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const types = Array.from(cd.types);
        const isPpt = types.includes("ppt/slides");

        console.log("[PPT Paste] === Paste ===");
        console.log("[PPT Paste] types:", types.join(", "));
        for (let i = 0; i < cd.items.length; i++) {
          const item = cd.items[i];
          if (item.kind === "file") {
            const f = item.getAsFile();
            console.log(`[PPT Paste] item[${i}] FILE type=${item.type} size=${f?.size}`);
          } else {
            console.log(`[PPT Paste] item[${i}] STRING type=${item.type}`);
          }
        }

        const html = cd.getData("text/html");
        const rtf = cd.getData("text/rtf");
        console.log("[PPT Paste] html:%d rtf:%d files:%d isPpt:%s",
          html.length, rtf.length, cd.files.length, isPpt);

        // Collect File objects synchronously (persist after event)
        const collectedFiles: File[] = [];
        for (let i = 0; i < cd.items.length; i++) {
          if (cd.items[i].kind === "file") {
            const f = cd.items[i].getAsFile();
            if (f) collectedFiles.push(f);
          }
        }

        const hasMulti = this.hasMultipleImages(cd, html, rtf);

        if (!isPpt && !hasMulti) {
          console.log("[PPT Paste] Not PPT / not multi → pass through");
          return;
        }

        console.log("[PPT Paste] Intercepting (isPpt=%s hasMulti=%s)", isPpt, hasMulti);
        evt.preventDefault();
        this.handlePaste(html, rtf, collectedFiles, isPpt, editor);
      })
    );
  }

  // ─── Detection ─────────────────────────────────────────────

  private hasMultipleImages(cd: DataTransfer, html: string, rtf: string): boolean {
    let fileCount = 0;
    for (let i = 0; i < cd.files.length; i++) {
      if (cd.files[i].type.startsWith("image/") && cd.files[i].size >= 3000)
        fileCount++;
    }
    if (fileCount > 1) return true;

    if (html) {
      if ((html.match(/<img[\s>]/gi) || []).length > 1) return true;
      if ((html.match(/data:image\/[\w+]+;base64,/gi) || []).length > 1) return true;
      if ((html.match(/src=["'][^"']*\.(?:png|jpg|jpeg|gif|bmp|emf|wmf)/gi) || []).length > 1) return true;
      if ((html.match(/<v:imagedata[\s>]/gi) || []).length > 1) return true;
      if (new Set(html.match(/clip_image\d+/gi) || []).size > 1) return true;
    }

    if (rtf) {
      if ((rtf.match(/\\(pngblip|jpegblip|emfblip)/g) || []).length > 1) return true;
    }

    return false;
  }

  // ─── Main extraction ──────────────────────────────────────

  private async handlePaste(
    html: string,
    rtf: string,
    collectedFiles: File[],
    isPpt: boolean,
    editor: any
  ) {
    const candidates: SlideImage[][] = [];

    // ── S1: PPT binary data — scan for embedded images ──
    if (isPpt) {
      const s1 = this.fromPptBinary();
      console.log("[PPT Paste] S1 ppt/slides binary:", s1.length);
      candidates.push(s1);
    }

    // ── S2: SVG via Electron (base64 extraction) ──
    if (isPpt) {
      const svg = this.readSvgViaElectron();
      if (svg) {
        const s2 = this.fromGenericBase64(svg);
        console.log("[PPT Paste] S2 SVG base64:", s2.length);
        candidates.push(s2);
      }
    }

    // ── S3: Collected image files ──
    const imageFiles = await this.fromCollectedFiles(collectedFiles);
    console.log("[PPT Paste] S3 files:", imageFiles.length);
    candidates.push(imageFiles);

    // ── S4: HTML base64 ──
    if (html) {
      const s4 = this.fromGenericBase64(html);
      console.log("[PPT Paste] S4 HTML base64:", s4.length);
      candidates.push(s4);
    }

    // ── S5: HTML URLs ──
    if (html) {
      const s5 = await this.fromHtmlUrls(html);
      console.log("[PPT Paste] S5 HTML URLs:", s5.length);
      candidates.push(s5);
    }

    // ── S6: RTF ──
    if (rtf) {
      const s6 = this.fromRtf(rtf);
      console.log("[PPT Paste] S6 RTF:", s6.length);
      candidates.push(s6);
    }

    // Pick best
    let images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    // Fallback: paste single file
    if (images.length === 0 && imageFiles.length > 0) {
      images = imageFiles;
    }

    console.log("[PPT Paste] Final:", images.length, "images");

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice("Could not extract slides.\nCheck console (Ctrl+Shift+I).");
    }
  }

  // ─── S1: PPT Binary — scan ppt/slides for images ─────────

  /**
   * Read the ppt/slides clipboard buffer via Electron.
   * If ZIP (OOXML) → extract media files.
   * Otherwise → scan for PNG/JPEG signatures.
   */
  private fromPptBinary(): SlideImage[] {
    try {
      const electron = require("electron");
      const clipboard = electron.clipboard || (electron.remote && electron.remote.clipboard);
      if (!clipboard) return [];

      const buf: Buffer = clipboard.readBuffer("ppt/slides");
      console.log("[PPT Paste] ppt/slides:", buf.length, "bytes");

      if (buf.length < 100) return [];

      // Log header for format identification
      const headerHex = Array.from(buf.slice(0, 20))
        .map((b: number) => b.toString(16).padStart(2, "0"))
        .join(" ");
      console.log("[PPT Paste] ppt/slides header:", headerHex);

      // Check if ZIP format (OOXML)
      if (buf[0] === 0x50 && buf[1] === 0x4b) {
        console.log("[PPT Paste] ppt/slides is ZIP (OOXML)");
        const zipImages = this.extractImagesFromZip(buf);
        if (zipImages.length > 0) return zipImages;
      }

      // Check if OLE2
      if (buf[0] === 0xd0 && buf[1] === 0xcf && buf[2] === 0x11 && buf[3] === 0xe0) {
        console.log("[PPT Paste] ppt/slides is OLE2");
      }

      // Brute-force: scan for PNG/JPEG image signatures
      return this.scanForImages(buf);
    } catch (e) {
      console.log("[PPT Paste] ppt/slides error:", e);
      return [];
    }
  }

  private extractImagesFromZip(buf: Buffer): SlideImage[] {
    const images: SlideImage[] = [];

    try {
      const zlib = require("zlib");

      // Find End of Central Directory Record (PK\x05\x06)
      let eocdPos = -1;
      for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
        if (buf[i] === 0x50 && buf[i + 1] === 0x4b && buf[i + 2] === 0x05 && buf[i + 3] === 0x06) {
          eocdPos = i;
          break;
        }
      }
      if (eocdPos === -1) {
        console.log("[PPT Paste] ZIP: EOCD not found");
        return images;
      }

      const entryCount = buf.readUInt16LE(eocdPos + 10);
      const cdOffset = buf.readUInt32LE(eocdPos + 16);
      console.log("[PPT Paste] ZIP: entries=%d cdOffset=%d", entryCount, cdOffset);

      // Parse Central Directory
      let pos = cdOffset;
      for (let i = 0; i < entryCount && pos + 46 <= buf.length; i++) {
        // Verify CD signature (PK\x01\x02)
        if (buf[pos] !== 0x50 || buf[pos + 1] !== 0x4b || buf[pos + 2] !== 0x01 || buf[pos + 3] !== 0x02) break;

        const method = buf.readUInt16LE(pos + 10);
        const compSize = buf.readUInt32LE(pos + 20);
        const uncompSize = buf.readUInt32LE(pos + 24);
        const nameLen = buf.readUInt16LE(pos + 28);
        const extraLen = buf.readUInt16LE(pos + 30);
        const commentLen = buf.readUInt16LE(pos + 32);
        const localOffset = buf.readUInt32LE(pos + 42);

        const name = buf.slice(pos + 46, pos + 46 + nameLen).toString("utf-8");

        // Log all entries for diagnosis
        if (i < 30 || /\.(png|jpg|jpeg|gif|emf|wmf)$/i.test(name)) {
          console.log(`[PPT Paste] ZIP[${i}]: ${name} (${uncompSize}b, method=${method})`);
        }

        // Extract image files
        if (/\.(png|jpg|jpeg|gif)$/i.test(name) && uncompSize >= 3000) {
          try {
            // Read local file header to find data
            const localNameLen = buf.readUInt16LE(localOffset + 26);
            const localExtraLen = buf.readUInt16LE(localOffset + 28);
            const dataStart = localOffset + 30 + localNameLen + localExtraLen;
            const compData = buf.slice(dataStart, dataStart + compSize);

            let data: Buffer;
            if (method === 0) {
              data = compData; // Stored
            } else if (method === 8) {
              data = zlib.inflateRawSync(compData); // Deflated
            } else {
              continue;
            }

            const ext = name.match(/\.(png|jpg|jpeg|gif)$/i)?.[1]?.toLowerCase() || "png";
            images.push({
              data: Uint8Array.from(data),
              ext: ext === "jpeg" ? "jpg" : ext,
            });
            console.log("[PPT Paste] ZIP extracted:", name, data.length, "bytes");
          } catch (e) {
            console.log("[PPT Paste] ZIP extract error:", name, e);
          }
        }

        pos += 46 + nameLen + extraLen + commentLen;
      }
    } catch (e) {
      console.log("[PPT Paste] ZIP parse error:", e);
    }

    return images;
  }

  private scanForImages(buf: Buffer): SlideImage[] {
    const images: SlideImage[] = [];

    // PNG: 89 50 4E 47 0D 0A 1A 0A ... 49 45 4E 44 AE 42 60 82
    const pngSig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const pngEnd = Buffer.from([0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82]);

    let pos = 0;
    while (pos < buf.length) {
      const sigPos = buf.indexOf(pngSig, pos);
      if (sigPos === -1) break;

      const endPos = buf.indexOf(pngEnd, sigPos + 8);
      if (endPos === -1) { pos = sigPos + 1; continue; }

      const imgEnd = endPos + pngEnd.length;
      const imgData = buf.slice(sigPos, imgEnd);

      if (imgData.length >= 3000) {
        console.log("[PPT Paste] Scan: PNG at %d (%d bytes)", sigPos, imgData.length);
        images.push({ data: Uint8Array.from(imgData), ext: "png" });
      }
      pos = imgEnd;
    }

    // JPEG: FF D8 FF ... FF D9
    pos = 0;
    while (pos < buf.length - 3) {
      if (buf[pos] === 0xff && buf[pos + 1] === 0xd8 && buf[pos + 2] === 0xff) {
        let endPos = pos + 3;
        let found = false;
        while (endPos < buf.length - 1) {
          if (buf[endPos] === 0xff && buf[endPos + 1] === 0xd9) {
            endPos += 2;
            found = true;
            break;
          }
          endPos++;
        }
        if (found) {
          const imgData = buf.slice(pos, endPos);
          if (imgData.length >= 3000) {
            console.log("[PPT Paste] Scan: JPEG at %d (%d bytes)", pos, imgData.length);
            images.push({ data: Uint8Array.from(imgData), ext: "jpg" });
          }
          pos = endPos;
          continue;
        }
      }
      pos++;
    }

    console.log("[PPT Paste] Binary scan:", images.length, "images found");
    return images;
  }

  // ─── S2: SVG via Electron ─────────────────────────────────

  private readSvgViaElectron(): string {
    try {
      const electron = require("electron");
      const clipboard = electron.clipboard || (electron.remote && electron.remote.clipboard);
      if (!clipboard) return "";

      const formats: string[] = clipboard.availableFormats();
      const svgFmt = formats.find((f: string) => f.includes("svg"));
      if (!svgFmt) return "";

      const buf: Buffer = clipboard.readBuffer(svgFmt);
      return buf && buf.length > 0 ? buf.toString("utf-8") : "";
    } catch {
      return "";
    }
  }

  // ─── S3: Collected Files ──────────────────────────────────

  private async fromCollectedFiles(files: File[]): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    for (const file of files) {
      if (!file.type.startsWith("image/") || file.type === "image/svg+xml") continue;
      try {
        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);
        if (data.length < 3000) continue;
        images.push({ data, ext: this.extFromMime(file.type) });
      } catch {}
    }
    return images;
  }

  // ─── S4/S5: Generic Base64 + HTML URLs ────────────────────

  private fromGenericBase64(text: string): SlideImage[] {
    const images: SlideImage[] = [];
    const regex = /data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+)/g;
    let match;
    while ((match = regex.exec(text)) !== null) {
      const b64 = match[2].replace(/\s/g, "");
      try {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        if (bytes.length < 3000) continue;
        images.push({ data: bytes, ext: this.extFromMime(`image/${match[1]}`) });
      } catch {}
    }
    return images;
  }

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    const regex = /src=["']([^"']+)["']/gi;
    let match;
    while ((match = regex.exec(html)) !== null) {
      if (match[1].startsWith("data:")) continue;
      const url = match[1];
      try {
        if (url.startsWith("file:///")) {
          const data = this.readLocalFile(url);
          if (data && data.length >= 3000)
            images.push({ data, ext: this.extFromPath(url) });
        } else if (url.startsWith("blob:")) {
          const resp = await fetch(url);
          if (!resp.ok) continue;
          const buf = await resp.arrayBuffer();
          const data = new Uint8Array(buf);
          if (data.length < 3000) continue;
          images.push({ data, ext: this.extFromMime(resp.headers.get("content-type") || "image/png") });
        }
      } catch {}
    }
    return images;
  }

  private readLocalFile(fileUrl: string): Uint8Array | null {
    try {
      let fp: string;
      try { fp = require("url").fileURLToPath(fileUrl); }
      catch { fp = decodeURIComponent(fileUrl.replace(/^file:\/\/\//, "")); }
      return Uint8Array.from(require("fs").readFileSync(fp));
    } catch { return null; }
  }

  // ─── S6: RTF ──────────────────────────────────────────────

  private fromRtf(rtf: string): SlideImage[] {
    const images: SlideImage[] = [];
    for (const blipType of ["pngblip", "jpegblip"]) {
      const marker = "\\" + blipType;
      let from = 0;
      while (true) {
        const pos = rtf.indexOf(marker, from);
        if (pos === -1) break;
        from = pos + marker.length;
        const hex = this.extractHexBlock(rtf, from);
        if (!hex || hex.length < 6000) continue;
        try {
          const bytes = new Uint8Array(hex.length / 2);
          for (let i = 0; i < hex.length; i += 2)
            bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
          images.push({ data: bytes, ext: blipType === "jpegblip" ? "jpg" : "png" });
        } catch {}
      }
    }
    return images;
  }

  private extractHexBlock(rtf: string, start: number): string | null {
    let i = start;
    const lim = Math.min(start + 5000, rtf.length);
    while (i < lim) {
      const ch = rtf[i];
      if (ch === "}") return null;
      if (" \r\n\t".includes(ch)) { i++; continue; }
      if (ch === "\\") {
        i++;
        while (i < lim && rtf[i] >= "a" && rtf[i] <= "z") i++;
        if (i < lim && (rtf[i] === "-" || (rtf[i] >= "0" && rtf[i] <= "9"))) {
          if (rtf[i] === "-") i++;
          while (i < lim && rtf[i] >= "0" && rtf[i] <= "9") i++;
        }
        if (i < lim && rtf[i] === " ") i++;
        continue;
      }
      if (this.isHex(ch)) {
        let cnt = 0, j = i;
        while (j < rtf.length) {
          const c = rtf[j];
          if (this.isHex(c)) { cnt++; j++; }
          else if (" \r\n\t".includes(c)) j++;
          else break;
        }
        if (cnt >= 100) {
          const end = Math.min(j, rtf.indexOf("}", i) === -1 ? j : rtf.indexOf("}", i));
          return rtf.substring(i, end).replace(/[\s\r\n]/g, "");
        }
        i = j; continue;
      }
      i++;
    }
    return null;
  }

  private isHex(ch: string): boolean {
    return (ch >= "0" && ch <= "9") || (ch >= "a" && ch <= "f") || (ch >= "A" && ch <= "F");
  }

  // ─── Utilities ────────────────────────────────────────────

  private extFromMime(m: string): string {
    if (m.includes("jpeg") || m.includes("jpg")) return "jpg";
    if (m.includes("gif")) return "gif";
    if (m.includes("webp")) return "webp";
    return "png";
  }

  private extFromPath(p: string): string {
    const m = p.match(/\.(png|jpg|jpeg|gif|webp)$/i);
    if (!m) return "png";
    return m[1].toLowerCase() === "jpeg" ? "jpg" : m[1].toLowerCase();
  }

  private async saveAndInsertImages(images: SlideImage[], editor: any) {
    const activeFile = this.app.workspace.getActiveFile();
    if (!activeFile) { new Notice("No active file"); return; }

    const folder = await this.getAttachmentFolder(activeFile);
    const ts = Date.now();
    const lines: string[] = [];

    new Notice(`Pasting ${images.length} slide images...`);

    for (let i = 0; i < images.length; i++) {
      const img = images[i];
      const name = `slide-${ts}-${String(i + 1).padStart(2, "0")}.${img.ext}`;
      const path = normalizePath(folder ? `${folder}/${name}` : name);

      const buf = img.data.buffer.byteLength === img.data.byteLength
        ? img.data.buffer
        : img.data.buffer.slice(img.data.byteOffset, img.data.byteOffset + img.data.byteLength);

      await this.app.vault.createBinary(path, buf);
      lines.push(`![[${name}]]`);
    }

    editor.replaceRange(lines.join("\n") + "\n", editor.getCursor());
    new Notice(`${images.length} slides pasted!`);
    console.log("[PPT Paste] Done:", images.length, "images");
  }

  private async getAttachmentFolder(file: TFile): Promise<string> {
    // @ts-ignore
    const p: string = this.app.vault.getConfig("attachmentFolderPath") || "/";
    if (p === "/") return "";
    if (p === "./") return file.parent?.path || "";
    if (p.startsWith("./")) {
      const parent = file.parent?.path || "";
      const sub = p.slice(2);
      const full = parent ? `${parent}/${sub}` : sub;
      await this.ensureFolder(full);
      return full;
    }
    await this.ensureFolder(p);
    return p;
  }

  private async ensureFolder(path: string) {
    const p = normalizePath(path);
    if (!this.app.vault.getAbstractFileByPath(p)) {
      await this.app.vault.createFolder(p);
    }
  }
}
