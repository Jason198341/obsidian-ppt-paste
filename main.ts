import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    console.log("[PPT Paste] Plugin loaded v1.4.0");

    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const types = Array.from(cd.types);
        const isPpt = types.includes("ppt/slides");

        // ── Diagnostic logging ──
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

        // ── Collect File objects synchronously (persist after event) ──
        const collectedFiles: File[] = [];
        for (let i = 0; i < cd.items.length; i++) {
          if (cd.items[i].kind === "file") {
            const f = cd.items[i].getAsFile();
            if (f) collectedFiles.push(f);
          }
        }

        // ── Decide whether to intercept ──
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

  // ─── Main extraction flow ──────────────────────────────────

  private async handlePaste(
    html: string,
    rtf: string,
    collectedFiles: File[],
    isPpt: boolean,
    editor: any
  ) {
    const candidates: SlideImage[][] = [];

    // ── Phase 1: Get SVG data (PPT primary path) ──
    let svg = "";
    if (isPpt) {
      svg = await this.obtainSvg(collectedFiles);
    }

    // ── Phase 2: Try all extraction strategies ──

    // S1: SVG embedded base64 images
    if (svg) {
      const s1 = this.fromSvgBase64(svg);
      console.log("[PPT Paste] S1 SVG base64:", s1.length);
      candidates.push(s1);
    }

    // S2: SVG href URLs (file:/// paths)
    if (svg) {
      const s2 = await this.fromSvgUrls(svg);
      console.log("[PPT Paste] S2 SVG URLs:", s2.length);
      candidates.push(s2);
    }

    // S3: SVG render individual slides (if SVG is vector, not raster)
    if (svg && candidates.every(c => c.length <= 1)) {
      const s3 = await this.fromSvgRender(svg);
      console.log("[PPT Paste] S3 SVG render:", s3.length);
      candidates.push(s3);
    }

    // S4: Collected image files
    const imageFiles = await this.fromCollectedFiles(collectedFiles);
    console.log("[PPT Paste] S4 files:", imageFiles.length);
    candidates.push(imageFiles);

    // S5: HTML base64
    if (html) {
      const s5 = this.fromGenericBase64(html);
      console.log("[PPT Paste] S5 HTML base64:", s5.length);
      candidates.push(s5);
    }

    // S6: HTML URLs
    if (html) {
      const s6 = await this.fromHtmlUrls(html);
      console.log("[PPT Paste] S6 HTML URLs:", s6.length);
      candidates.push(s6);
    }

    // S7: RTF
    if (rtf) {
      const s7 = this.fromRtf(rtf);
      console.log("[PPT Paste] S7 RTF:", s7.length);
      candidates.push(s7);
    }

    // ── Phase 3: Pick best result ──
    let images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    // Fallback: at least paste the single file
    if (images.length === 0 && imageFiles.length > 0) {
      console.log("[PPT Paste] Fallback: single file");
      images = imageFiles;
    }

    console.log("[PPT Paste] Final:", images.length, "images");

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice("Could not extract slides.\nCheck console (Ctrl+Shift+I).");
      console.log("[PPT Paste] FAILED — all strategies returned 0");
    }
  }

  // ─── SVG Data Acquisition ─────────────────────────────────

  /**
   * Try multiple methods to read SVG from clipboard.
   * DataTransfer.getData("image/svg+xml") returns "" because
   * the web API only supports text/plain, text/html, etc.
   */
  private async obtainSvg(collectedFiles: File[]): Promise<string> {
    let svg = "";

    // Method 1: Electron clipboard API (most reliable)
    svg = this.readSvgViaElectron();
    if (svg) return svg;

    // Method 2: SVG file from collected items
    for (const file of collectedFiles) {
      if (file.type === "image/svg+xml") {
        try {
          svg = await file.text();
          console.log("[PPT Paste] SVG from File.text():", svg.length, "chars");
          if (svg) return svg;
        } catch (e) {
          console.log("[PPT Paste] File.text() error:", e);
        }
      }
    }

    // Method 3: navigator.clipboard.read()
    svg = await this.readSvgViaNavigator();
    if (svg) return svg;

    console.log("[PPT Paste] SVG: all methods failed");
    return "";
  }

  private readSvgViaElectron(): string {
    try {
      const electron = require("electron");
      const clipboard =
        electron.clipboard ||
        (electron.remote && electron.remote.clipboard);
      if (!clipboard) {
        console.log("[PPT Paste] Electron clipboard not available");
        return "";
      }

      const formats: string[] = clipboard.availableFormats();
      console.log("[PPT Paste] Electron formats:", formats.join(", "));

      // Find SVG format
      const svgFmt = formats.find(
        (f: string) => f.includes("svg") || f === "image/svg+xml"
      );
      if (svgFmt) {
        const buf: Buffer = clipboard.readBuffer(svgFmt);
        if (buf && buf.length > 0) {
          const svg = buf.toString("utf-8");
          console.log("[PPT Paste] Electron SVG (%s):", svgFmt, svg.length, "chars");
          if (svg.length > 0 && svg.length < 50000) {
            console.log("[PPT Paste] SVG content:", svg);
          } else {
            console.log("[PPT Paste] SVG preview:", svg.substring(0, 5000));
          }
          return svg;
        }
      }

      // Log all format sizes for debugging
      for (const fmt of formats) {
        try {
          const buf: Buffer = clipboard.readBuffer(fmt);
          console.log(`[PPT Paste] format '${fmt}': ${buf.length} bytes`);
        } catch {}
      }
    } catch (e) {
      console.log("[PPT Paste] Electron error:", e);
    }
    return "";
  }

  private async readSvgViaNavigator(): Promise<string> {
    try {
      const items = await (navigator as any).clipboard.read();
      for (const item of items) {
        console.log("[PPT Paste] navigator.clipboard types:", item.types.join(", "));
        if (item.types.includes("image/svg+xml")) {
          const blob = await item.getType("image/svg+xml");
          const text = await blob.text();
          console.log("[PPT Paste] navigator.clipboard SVG:", text.length, "chars");
          return text;
        }
      }
    } catch (e) {
      console.log("[PPT Paste] navigator.clipboard error:", e);
    }
    return "";
  }

  // ─── S1: SVG Embedded Base64 ──────────────────────────────

  private fromSvgBase64(svg: string): SlideImage[] {
    const images: SlideImage[] = [];

    // Match href="data:image/..." or xlink:href="data:image/..."
    const regex =
      /(?:xlink:)?href=["'](data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+))["']/g;
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
        console.log("[PPT Paste] S1 error:", e);
      }
    }

    // Fallback: generic data URI search (no href wrapper)
    if (images.length === 0) {
      return this.fromGenericBase64(svg);
    }

    return images;
  }

  // ─── S2: SVG URL References ───────────────────────────────

  private async fromSvgUrls(svg: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    const regex = /(?:xlink:)?href=["']((?:file:\/\/\/)[^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(svg)) !== null) {
      urls.push(match[1]);
    }

    for (const url of urls) {
      try {
        const data = this.readLocalFile(url);
        if (data && data.length >= 3000) {
          images.push({ data, ext: this.extFromPath(url) });
        }
      } catch {}
    }

    return images;
  }

  // ─── S3: SVG Render (vector → raster per slide) ──────────

  /**
   * If SVG contains vector slides (not raster), render each
   * top-level group as a separate PNG using OffscreenCanvas.
   */
  private async fromSvgRender(svg: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];

    try {
      // Parse SVG to find viewBox and top-level groups
      const parser = new DOMParser();
      const doc = parser.parseFromString(svg, "image/svg+xml");
      const svgEl = doc.querySelector("svg");
      if (!svgEl) return images;

      // Get SVG dimensions
      const vb = svgEl.getAttribute("viewBox");
      const width = parseFloat(svgEl.getAttribute("width") || "0");
      const height = parseFloat(svgEl.getAttribute("height") || "0");

      console.log("[PPT Paste] S3 SVG viewBox:", vb, "w:", width, "h:", height);

      // Find top-level groups — each might be a slide
      const groups = svgEl.querySelectorAll(":scope > g");
      console.log("[PPT Paste] S3 top-level <g>:", groups.length);

      if (groups.length <= 1) return images;

      // Try to detect slide layout from group transforms
      // Each slide group should have a transform that positions it
      const slideGroups: Element[] = [];
      groups.forEach((g) => {
        // Only include groups that have visual content
        const hasContent =
          g.querySelector("image, rect, path, text, polygon, polyline, circle, ellipse, line");
        if (hasContent) slideGroups.push(g);
      });

      console.log("[PPT Paste] S3 visual groups:", slideGroups.length);
      if (slideGroups.length <= 1) return images;

      // Render each group as a separate image
      // Calculate per-slide dimensions from viewBox
      let slideW = width || 960;
      let slideH = height || 540;

      if (vb) {
        const parts = vb.split(/[\s,]+/).map(Number);
        if (parts.length === 4) {
          const totalW = parts[2];
          const totalH = parts[3];
          // Heuristic: if height >> width, slides are stacked vertically
          if (totalH / totalW > 1.5) {
            slideW = totalW;
            slideH = totalW * 0.5625; // 16:9 aspect ratio
          } else {
            slideW = totalW;
            slideH = totalH;
          }
        }
      }

      for (let i = 0; i < slideGroups.length; i++) {
        try {
          const groupSvg = this.wrapGroupAsSvg(
            svgEl,
            slideGroups[i],
            slideW,
            slideH
          );
          const png = await this.svgToPng(groupSvg, slideW, slideH);
          if (png && png.length >= 3000) {
            images.push({ data: png, ext: "png" });
          }
        } catch (e) {
          console.log("[PPT Paste] S3 render error for group", i, ":", e);
        }
      }
    } catch (e) {
      console.log("[PPT Paste] S3 error:", e);
    }

    return images;
  }

  private wrapGroupAsSvg(
    originalSvg: SVGSVGElement,
    group: Element,
    w: number,
    h: number
  ): string {
    // Create a standalone SVG with just this group
    const ns = originalSvg.getAttribute("xmlns") || "http://www.w3.org/2000/svg";
    const xlinkNs = "http://www.w3.org/1999/xlink";

    // Clone the defs (gradients, patterns, clip paths)
    const defs = originalSvg.querySelector("defs");
    const defsStr = defs ? defs.outerHTML : "";

    return `<svg xmlns="${ns}" xmlns:xlink="${xlinkNs}" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${defsStr}${group.outerHTML}</svg>`;
  }

  private svgToPng(svgString: string, w: number, h: number): Promise<Uint8Array | null> {
    return new Promise((resolve) => {
      const img = new Image();
      const blob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(blob);

      img.onload = () => {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = Math.min(w, 1920);
          canvas.height = Math.min(h, 1080);
          const ctx = canvas.getContext("2d");
          if (!ctx) { resolve(null); return; }

          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

          canvas.toBlob(
            (pngBlob) => {
              if (!pngBlob) { resolve(null); return; }
              pngBlob.arrayBuffer().then((buf) => {
                resolve(new Uint8Array(buf));
              });
            },
            "image/png"
          );
        } catch (e) {
          console.log("[PPT Paste] canvas error:", e);
          resolve(null);
        } finally {
          URL.revokeObjectURL(url);
        }
      };

      img.onerror = () => {
        URL.revokeObjectURL(url);
        resolve(null);
      };

      img.src = url;
    });
  }

  // ─── S4: Collected Files ──────────────────────────────────

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

  // ─── S5: Generic Base64 Data URIs ─────────────────────────

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
        images.push({ data: bytes, ext: this.extFromMime(`image/${mimeType}`) });
      } catch {}
    }

    return images;
  }

  // ─── S6: HTML src URLs ────────────────────────────────────

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    const regex = /src=["']([^"']+)["']/gi;
    let match;
    const urls: string[] = [];

    while ((match = regex.exec(html)) !== null) {
      if (match[1].startsWith("data:")) continue;
      urls.push(match[1]);
    }

    for (const url of urls) {
      try {
        if (url.startsWith("file:///")) {
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
          images.push({ data, ext: this.extFromMime(resp.headers.get("content-type") || "image/png") });
        }
      } catch {}
    }

    return images;
  }

  private readLocalFile(fileUrl: string): Uint8Array | null {
    try {
      let filePath: string;
      try {
        filePath = require("url").fileURLToPath(fileUrl);
      } catch {
        filePath = decodeURIComponent(fileUrl.replace(/^file:\/\/\//, ""));
      }
      const buffer: Buffer = require("fs").readFileSync(filePath);
      return Uint8Array.from(buffer);
    } catch (e) {
      console.log("[PPT Paste] fs read failed:", e);
      return null;
    }
  }

  // ─── S7: RTF Embedded Images ──────────────────────────────

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
          images.push({ data: bytes, ext: blipType === "jpegblip" ? "jpg" : "png" });
        } catch {}
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
      if (ch === " " || ch === "\r" || ch === "\n" || ch === "\t") { i++; continue; }

      if (ch === "\\") {
        i++;
        while (i < limit && rtf[i] >= "a" && rtf[i] <= "z") i++;
        if (i < limit && (rtf[i] === "-" || (rtf[i] >= "0" && rtf[i] <= "9"))) {
          if (rtf[i] === "-") i++;
          while (i < limit && rtf[i] >= "0" && rtf[i] <= "9") i++;
        }
        if (i < limit && rtf[i] === " ") i++;
        continue;
      }

      if (this.isHex(ch)) {
        let hexCount = 0;
        let j = i;
        while (j < rtf.length) {
          const c = rtf[j];
          if (this.isHex(c)) { hexCount++; j++; }
          else if (c === " " || c === "\r" || c === "\n" || c === "\t") { j++; }
          else break;
        }
        if (hexCount >= 100) {
          const endPos = Math.min(j, rtf.indexOf("}", i) === -1 ? j : rtf.indexOf("}", i));
          return rtf.substring(i, endPos).replace(/[\s\r\n]/g, "");
        }
        i = j;
        continue;
      }

      i++;
    }
    return null;
  }

  private isHex(ch: string): boolean {
    return (ch >= "0" && ch <= "9") || (ch >= "a" && ch <= "f") || (ch >= "A" && ch <= "F");
  }

  // ─── Utilities ────────────────────────────────────────────

  private extFromMime(mime: string): string {
    if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg";
    if (mime.includes("gif")) return "gif";
    if (mime.includes("webp")) return "webp";
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

    const attachFolder = await this.getAttachmentFolder(activeFile);
    const ts = Date.now();
    const lines: string[] = [];

    new Notice(`Pasting ${images.length} slide images...`);

    for (let i = 0; i < images.length; i++) {
      const img = images[i];
      const fileName = `slide-${ts}-${String(i + 1).padStart(2, "0")}.${img.ext}`;
      const filePath = normalizePath(
        attachFolder ? `${attachFolder}/${fileName}` : fileName
      );

      const buf = img.data.buffer.byteLength === img.data.byteLength
        ? img.data.buffer
        : img.data.buffer.slice(img.data.byteOffset, img.data.byteOffset + img.data.byteLength);

      await this.app.vault.createBinary(filePath, buf);
      lines.push(`![[${fileName}]]`);
    }

    editor.replaceRange(lines.join("\n") + "\n", editor.getCursor());
    new Notice(`${images.length} slides pasted!`);
    console.log("[PPT Paste] Done:", images.length, "images");
  }

  private async getAttachmentFolder(activeFile: TFile): Promise<string> {
    // @ts-ignore — internal API
    const p: string = this.app.vault.getConfig("attachmentFolderPath") || "/";
    if (p === "/") return "";
    if (p === "./") return activeFile.parent?.path || "";
    if (p.startsWith("./")) {
      const parent = activeFile.parent?.path || "";
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
