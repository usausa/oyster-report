# Third-party notices

OysterReport contains source code ported or derived from the following projects.
Each project is distributed under the MIT License reproduced at the end of this file.

| Project | Copyright | Used for |
|---|---|---|
| [PDFKit](https://github.com/foliojs/pdfkit) | Copyright (c) 2014 Devon Govett | PDF object writing, Type0 font embedding and ToUnicode CMap, JPEG/PNG image XObjects (`Internal/Pdf/PdfWriter.cs`, `Internal/Pdf/PdfImageWriter.cs`) |
| [fontkit](https://github.com/foliojs/fontkit) | Copyright (c) 2014 Devon Govett | TrueType and CFF font subsetting (`Internal/Pdf/TrueTypeSubset.cs`, `Internal/Pdf/CffSubset.cs`, `Internal/Pdf/CffFont.cs`) |
| [png.js](https://github.com/foliojs/png.js) | Copyright (c) 2017 Devon Govett | PNG scanline decoding (`Internal/Pdf/RasterDecoders.cs`) |
| [PDFsharp](https://github.com/empira/PDFsharp) 6.2.4 | Copyright (c) 2001-2026 empira Software GmbH, Troisdorf (Cologne Area), Germany | Text placement and line breaking (`XGraphics.DrawString`, `XTextFormatter`) and font line metrics (`Internal/Pdf/PdfTextLayout.cs`, `Internal/Pdf/TrueTypeFont.cs`) |

PDFsharp is used only in the form of code derived from version 6.2.4, which is released under the MIT License.

## MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
