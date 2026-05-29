import { fontToEot } from './font-utils.js';

export class PPTXEmbedFonts {
  #zip;
  #fonts = [];

  async loadZip(zip) {
    this.#zip = zip;
  }

  async addFont(name, buffer, type = 'ttf') {
    const eot = await fontToEot(type, buffer);
    this.#fonts.push({ name, eot });
  }

  async updateFiles() {
    if (!this.#zip) throw new Error('PPTX zip has not been loaded.');
    this.#fonts.forEach((font, index) => {
      this.#zip.file(`ppt/fonts/font${index + 1}.fntdata`, font.eot);
    });
  }

  async generateBlob() {
    if (!this.#zip) throw new Error('PPTX zip has not been loaded.');
    return this.#zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }
}
