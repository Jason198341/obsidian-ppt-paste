import { Plugin, Notice, TFile, normalizePath } from "obsidian";

interface SlideImage {
  data: Uint8Array;
  ext: string;
}

export default class PPTSlidePaste extends Plugin {
  async onload() {
    console.log("[PPT Paste] Plugin loaded v1.6.0");

    this.registerEvent(
      this.app.workspace.on("editor-paste", (evt: ClipboardEvent, editor) => {
        if (!evt.clipboardData) return;

        const cd = evt.clipboardData;
        const types = Array.from(cd.types);
        const isPpt = types.includes("ppt/slides");

        console.log("[PPT Paste] === Paste ===");
        console.log("[PPT Paste] types:", types.join(", "));
        console.log("[PPT Paste] files:", cd.files.length, "isPpt:", isPpt);

        const html = cd.getData("text/html");
        const rtf = cd.getData("text/rtf");

        // Collect File objects synchronously
        const collectedFiles: File[] = [];
        for (let i = 0; i < cd.items.length; i++) {
          if (cd.items[i].kind === "file") {
            const f = cd.items[i].getAsFile();
            if (f) collectedFiles.push(f);
          }
        }

        const hasMulti = this.hasMultipleImages(cd, html, rtf);

        if (!isPpt && !hasMulti) return;

        evt.preventDefault();
        this.handlePaste(html, rtf, collectedFiles, isPpt, editor);
      })
    );
  }

  // ─── Detection ─────────────────────────────────────────────

  private hasMultipleImages(cd: DataTransfer, html: string, rtf: string): boolean {
    let fc = 0;
    for (let i = 0; i < cd.files.length; i++) {
      if (cd.files[i].type.startsWith("image/") && cd.files[i].size >= 3000) fc++;
    }
    if (fc > 1) return true;

    if (html) {
      if ((html.match(/<img[\s>]/gi) || []).length > 1) return true;
      if ((html.match(/data:image\/[\w+]+;base64,/gi) || []).length > 1) return true;
      if ((html.match(/src=["'][^"']*\.(?:png|jpg|jpeg|gif|bmp|emf|wmf)/gi) || []).length > 1) return true;
      if ((html.match(/<v:imagedata[\s>]/gi) || []).length > 1) return true;
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
    // ── Strategy 1: PowerShell COM automation (PPT only) ──
    if (isPpt) {
      new Notice("Extracting slides from PowerPoint...");
      const s1 = await this.fromPowerPointCOM();
      console.log("[PPT Paste] S1 COM automation:", s1.length);
      if (s1.length > 0) {
        await this.saveAndInsertImages(s1, editor);
        return;
      }
    }

    // ── Fallback strategies ──
    const candidates: SlideImage[][] = [];

    // S2: Collected image files
    const imgFiles = await this.fromCollectedFiles(collectedFiles);
    console.log("[PPT Paste] S2 files:", imgFiles.length);
    candidates.push(imgFiles);

    // S3: HTML base64
    if (html) {
      const s3 = this.fromBase64(html);
      console.log("[PPT Paste] S3 HTML base64:", s3.length);
      candidates.push(s3);
    }

    // S4: HTML URLs
    if (html) {
      const s4 = await this.fromHtmlUrls(html);
      console.log("[PPT Paste] S4 HTML URLs:", s4.length);
      candidates.push(s4);
    }

    // S5: RTF
    if (rtf) {
      const s5 = this.fromRtf(rtf);
      console.log("[PPT Paste] S5 RTF:", s5.length);
      candidates.push(s5);
    }

    let images = candidates.reduce(
      (best, curr) => (curr.length > best.length ? curr : best),
      [] as SlideImage[]
    );

    if (images.length === 0 && imgFiles.length > 0) images = imgFiles;

    console.log("[PPT Paste] Final:", images.length);

    if (images.length > 0) {
      await this.saveAndInsertImages(images, editor);
    } else {
      new Notice("Could not extract slides. Check console (Ctrl+Shift+I).");
    }
  }

  // ─── S1: PowerShell COM Automation ────────────────────────

  /**
   * Use PowerPoint's COM API to paste clipboard slides
   * into a temp presentation and export each as PNG.
   */
  private async fromPowerPointCOM(): Promise<SlideImage[]> {
    const images: SlideImage[] = [];

    try {
      const os = require("os");
      const path = require("path");
      const fs = require("fs");

      const tempDir = path.join(os.tmpdir(), `ppt_slides_${Date.now()}`);
      const scriptPath = path.join(os.tmpdir(), `ppt_export_${Date.now()}.ps1`);

      // PowerShell script: paste slides → export as PNG
      const script = [
        "$ErrorActionPreference = 'Stop'",
        `$tempDir = '${tempDir.replace(/'/g, "''")}'`,
        "try {",
        "  $ppt = New-Object -ComObject PowerPoint.Application",
        "  $ppt.Visible = $true",
        "  $pres = $ppt.Presentations.Add()",
        "  $pres.Slides.Paste() | Out-Null",
        "  $count = $pres.Slides.Count",
        "  New-Item -ItemType Directory -Path $tempDir -Force | Out-Null",
        "  for ($i = 1; $i -le $count; $i++) {",
        "    $p = Join-Path $tempDir ('slide_{0:D2}.png' -f $i)",
        "    $pres.Slides.Item($i).Export($p, 'PNG', 1920, 1080)",
        "  }",
        "  $pres.Close()",
        "  Write-Output \"OK:$count\"",
        "} catch {",
        "  Write-Output \"ERR:$($_.Exception.Message)\"",
        "}",
      ].join("\n");

      fs.writeFileSync(scriptPath, script, "utf-8");

      const result = await this.runPowerShell(scriptPath);
      console.log("[PPT Paste] PowerShell result:", result);

      if (result.startsWith("OK:")) {
        const count = parseInt(result.split(":")[1]);
        for (let i = 1; i <= count; i++) {
          const slidePath = path.join(tempDir, `slide_${String(i).padStart(2, "0")}.png`);
          try {
            if (fs.existsSync(slidePath)) {
              const data: Buffer = fs.readFileSync(slidePath);
              images.push({ data: Uint8Array.from(data), ext: "png" });
            }
          } catch {}
        }
        console.log("[PPT Paste] Exported", images.length, "slides");
      } else {
        console.log("[PPT Paste] PowerShell error:", result);
      }

      // Cleanup
      this.cleanup(fs, path, tempDir, scriptPath);
    } catch (e: any) {
      console.log("[PPT Paste] COM error:", e?.message || e);
    }

    return images;
  }

  private runPowerShell(scriptPath: string): Promise<string> {
    return new Promise((resolve) => {
      const { exec } = require("child_process");
      exec(
        `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${scriptPath}"`,
        { timeout: 30000 },
        (err: any, stdout: string, stderr: string) => {
          if (err) {
            console.log("[PPT Paste] PS error:", err.message);
            console.log("[PPT Paste] PS stderr:", stderr);
            resolve("ERR:" + (err.message || "unknown"));
          } else {
            resolve(stdout.trim());
          }
        }
      );
    });
  }

  private cleanup(fs: any, path: any, tempDir: string, scriptPath: string) {
    try {
      if (fs.existsSync(tempDir)) {
        for (const f of fs.readdirSync(tempDir)) {
          fs.unlinkSync(path.join(tempDir, f));
        }
        fs.rmdirSync(tempDir);
      }
    } catch {}
    try { fs.unlinkSync(scriptPath); } catch {}
  }

  // ─── S2: Collected Files ──────────────────────────────────

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

  // ─── S3: Base64 Data URIs ─────────────────────────────────

  private fromBase64(text: string): SlideImage[] {
    const images: SlideImage[] = [];
    const re = /data:image\/([\w+]+);base64,([A-Za-z0-9+/=\s]+)/g;
    let m;
    while ((m = re.exec(text)) !== null) {
      const b64 = m[2].replace(/\s/g, "");
      try {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        if (bytes.length < 3000) continue;
        images.push({ data: bytes, ext: this.extFromMime(`image/${m[1]}`) });
      } catch {}
    }
    return images;
  }

  // ─── S4: HTML URLs ────────────────────────────────────────

  private async fromHtmlUrls(html: string): Promise<SlideImage[]> {
    const images: SlideImage[] = [];
    const re = /src=["']([^"']+)["']/gi;
    let m;
    while ((m = re.exec(html)) !== null) {
      if (m[1].startsWith("data:")) continue;
      try {
        if (m[1].startsWith("file:///")) {
          let fp: string;
          try { fp = require("url").fileURLToPath(m[1]); }
          catch { fp = decodeURIComponent(m[1].replace(/^file:\/\/\//, "")); }
          const data: Buffer = require("fs").readFileSync(fp);
          if (data.length >= 3000)
            images.push({ data: Uint8Array.from(data), ext: this.extFromPath(m[1]) });
        }
      } catch {}
    }
    return images;
  }

  // ─── S5: RTF ──────────────────────────────────────────────

  private fromRtf(rtf: string): SlideImage[] {
    const images: SlideImage[] = [];
    for (const bt of ["pngblip", "jpegblip"]) {
      const mk = "\\" + bt;
      let from = 0;
      while (true) {
        const pos = rtf.indexOf(mk, from);
        if (pos === -1) break;
        from = pos + mk.length;
        const hex = this.extractHex(rtf, from);
        if (!hex || hex.length < 6000) continue;
        try {
          const bytes = new Uint8Array(hex.length / 2);
          for (let i = 0; i < hex.length; i += 2)
            bytes[i / 2] = parseInt(hex.substring(i, i + 2), 16);
          images.push({ data: bytes, ext: bt === "jpegblip" ? "jpg" : "png" });
        } catch {}
      }
    }
    return images;
  }

  private extractHex(rtf: string, start: number): string | null {
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
      if ((ch >= "0" && ch <= "9") || (ch >= "a" && ch <= "f") || (ch >= "A" && ch <= "F")) {
        let cnt = 0, j = i;
        while (j < rtf.length) {
          const c = rtf[j];
          if ((c >= "0" && c <= "9") || (c >= "a" && c <= "f") || (c >= "A" && c <= "F")) { cnt++; j++; }
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
      const filePath = normalizePath(folder ? `${folder}/${name}` : name);

      const buf = img.data.buffer.byteLength === img.data.byteLength
        ? img.data.buffer
        : img.data.buffer.slice(img.data.byteOffset, img.data.byteOffset + img.data.byteLength);

      await this.app.vault.createBinary(filePath, buf);
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
