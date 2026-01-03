# PDFNetworker

## 目的

PDFNetworker は、PDF の画像化、ページ操作、埋め込み処理（Embed）などを行うコマンドラインツールです。

PDFNetworker は PDFNET（GUI アプリ）から CLI 経由で呼び出されます。  
両者はプロセス分離されており、リンクや組み込みは行われません。  
そのため、PDFNetworker の AGPL ライセンスは PDFNET（GUI アプリ）には適用されません。

---

## 使い方

PDFNetworker.exe [command] [options]


### コマンド一覧

PDFNETWorker.exe paging <src> <dest> <FreeLicense> <pageInfoJson>
Usage: PDFNETWorker.exe imageonly <pagerawInfoJson> <dest>
Usage: PDFNETWorker.exe stamp <read> <dest> <freeLicense> <xml>
Usage: PDFNETWorker.exe merge <read>
Usage: PDFNETWorker.exe getpagerotation <read> <pageNumber>
Usage: PDFNETWorker.exe makeblankpage
Usage: PDFNETWorker.exe pdf2png <pdfPath> <outputJson>
           

### 出力先について

すべての出力ファイルはテンポラリフォルダに  
**edit_temp.pdf**  
として保存されます。

---

## License

This project is licensed under the **AGPL-3.0 License**.  
See **License(AGPL).txt** for details.

This project also includes third-party libraries:

- **iTextSharp** (LGPL)

Their licenses are included in **License(LGPL).txt**.

---