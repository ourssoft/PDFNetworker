using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;

using iTextSharp.text;
using iTextSharp.text.pdf;

using Newtonsoft.Json;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text;
class Program
{
    static int Main(string[] args)
    {

        // これを最初に呼ぶ
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 以降 iTextSharp が 1252 を使えるようになる


#if DEBUG
        //  args = new string[] { "pdf2png", @"C:\Users\takes\OneDrive\デスクトップ\図面１ - コピー.pdf", @"C:\Users\takes\AppData\Local\Temp\a5947e51-e9a8-4a8a-9dfd-6804df66031d.tmp" };
#endif
        Console.Error.WriteLine("ARGS:");
        for (int i = 0; i < args.Length; i++)
        {
            Console.Error.WriteLine($"[{i}] = '{args[i]}'");
        }
        if (args.Length < 1)
        {
            Console.Error.WriteLine("Usage: PDFNETWorker.exe <command> ");
            return 1;
        }
        switch (args[0].ToLower())
        {
            default:
                Console.Error.WriteLine("Usage: PDFNETWorker.exe <EnableCommand> ");
                return 1;
            case "paging":
                {
                    if (args.Length < 5)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe paging <src> <dest> <FreeLicense> <pageInfoJson>");
                        return 1;
                    }

                    string src = args[1];
                    string dest = args[2];
                    string freeLicense = args[3];
                    string pageInfoJson = args[4];

                    if (freeLicense.ToLower() != "false") freeLicense = "true";
                    try
                    {
                        var pageInfos = JsonConvert.DeserializeObject<List<PageInfo>>(File.ReadAllText(pageInfoJson));
                        var worker = new PdfPagingWorker();
                        worker.savePagingPDF(src, dest, bool.Parse(freeLicense), pageInfos);
                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }

            case "imageonly":
                {
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe imageonly <pagerawInfoJson> <dest>");
                        return 1;
                    }

                    string dest = args[1];
                    string pageRawInfoJson = args[2];

                    try
                    {
                        var pageRawInfos = JsonConvert.DeserializeObject<List<PageRawInfo>>(File.ReadAllText(pageRawInfoJson));
                        var worker = new PdfPagingWorker();
                        worker.SaveAsImageOnlyPdf(pageRawInfos, dest);
                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }

            case "stamp":
                {
                    if (args.Length < 5)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe stamp <read> <dest> <freeLicense> <xml>");
                        return 1;
                    }

                    string read = args[1];
                    string dest = args[2];
                    string freeLicense = args[3];
                    if (freeLicense.ToLower() != "false") freeLicense = "true";

                    string xml = args[4];

                    try
                    {
                        var worker = new PdfPagingWorker();
                        worker.EmbedStamp(read, dest, freeLicense, xml);
                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }
            case "merge":
                {
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe merge <read>");
                        return 1;
                    }

                    string read = args[1];
                    try
                    {
                        var worker = new PdfPagingWorker();
                        var result= worker.MergeIntoTempPdf(read);
                        
                        // ★ JSON で返す
                        Console.WriteLine(JsonConvert.SerializeObject(result));

                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }
            case "getpagerotation":
                {
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe getpagerotation <read> <pageNumber>");
                        return 1;
                    }

                    string read = args[1];
                    string pageNumber = args[2];
                    try
                    {
                        var worker = new PdfPagingWorker();
                        WorkerResult result= worker.GetPageRotation(read, int.Parse( pageNumber));
                        
                        Console.WriteLine(JsonConvert.SerializeObject(result));

                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }
            case "makeblankpage":
                {
                 
                    try
                    {
                        var worker = new PdfPagingWorker();
                        WorkerResult result = worker.MakeBrankPage();

                        Console.WriteLine(JsonConvert.SerializeObject(result));

                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }
            case "pdf2png":
                {
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe pdf2png <pdfPath> <outputJson>");
                        return 1;
                    }

                    string pdfPath = args[1];
                    string outputJson = args[2];

                    try
                    {
                        var worker = new PdfPagingWorker();
                        var pages = worker.LoadPdfPagesRaw(pdfPath);

                        // JSON で返す
                        File.WriteAllText(outputJson, JsonConvert.SerializeObject(pages));

                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);
                        return 2;
                    }
                }

        }



    }

}
public static class Config
{
    public static string tempPdfPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "edit_temp.pdf");
}
public class WorkerResult
{
    public bool Success { get; set; }
    public string Message { get; set; }

    // 追加されたページ番号（差分方式）
    public List<int> AddedPages { get; set; }

    // Worker が生成した PDF のパス（temp）
    public string TempPdfPath { get; set; }

    // ページごとの詳細情報
    public List<PageInfo> PageInfo { get; set; }
}
public class PageInfo
{
    public int Rotation { get; set; }
    public int OriginalPageNumber { get; set; }
    public int PageNumber { get; set; }
    public double HeightPt { get; set; }
    public double WidthPt { get; set; }
}
public class PageRawInfo
{
    // PDFページを画像化した生データ（PNGやJPEGなど）
    public byte[] Bytes { get; set; }

    // ページ番号（元のPDF内のページ番号）
    public int PageNumber { get; set; }
    public int OriginalPageNumber { get; set; }

    // 必要なら追加情報
    public double WidthPt { get; set; }
    public double HeightPt { get; set; }
    public int Rotation { get; set; }
}

public class PdfPagingWorker
{
    public void  savePagingPDF(string srcFile, string destFile,bool bFreeLicense,List<PageInfo>pages)
    {
        //ページングの結果を新しいPDFファイルに保存
        using (var fs = new FileStream(destFile, FileMode.Create))
        {
            var reader = new PdfReader(srcFile);
            using (var ms = new MemoryStream())
            {

                var stamper = new PdfStamper(reader, ms);

                // ここで stamper.GetOverContent(pageNumber) を使って
                // スタンプや文字を描画する

                stamper.Close();// で編集済みPDFが ms に入る



                var editedReader = new PdfReader(ms.ToArray());

                var doc = new Document();
                var writer = new PdfCopy(doc, fs);
                doc.Open();

                if (bFreeLicense)
                {
                    //文書のプロパティをセット
                    // 文書プロパティ設定
                    writer.Info.Put(PdfName.TITLE, new PdfString("ページング結果"));
                    writer.Info.Put(PdfName.AUTHOR, new PdfString("OURS SOFT"));
                    writer.Info.Put(PdfName.SUBJECT, new PdfString("Made by PDFNET -- OURS SOFT"));
                    writer.Info.Put(PdfName.KEYWORDS, new PdfString("Powered by OURS SOFT"));
                }



                foreach (var pageInfo in pages)
                {
                    if (pageInfo.OriginalPageNumber == 0)
                    {
                        //0は新たに追加したページ、元PGなく出力不可能
                        continue;
                    }
                    else
                    {
                        // 元ページを取得
                        PdfImportedPage importedPage = writer.GetImportedPage(editedReader, pageInfo.OriginalPageNumber);

                        // ページ辞書を取得
                        PdfDictionary pageDict = editedReader.GetPageN(pageInfo.OriginalPageNumber);

                        // 元の回転値を取得
                        //PdfNumber rotate = pageDict.GetAsNumber(PdfName.ROTATE);
                        //int currentRotation = rotate != null ? rotate.IntValue : 0;
                        // PageInfo に保持している回転角度を加算
                        //int newRotation = (currentRotation + pageInfo.Rotation) % 360;

                        // 新しい回転値を設定
                        pageDict.Put(PdfName.ROTATE, new PdfNumber(pageInfo.Rotation));

                        // コピー先に追加
                        writer.AddPage(importedPage);

                    }

                }
                writer.Close();
                doc.Close();
            }
            reader.Close();
        }

    }
    public void SaveAsImageOnlyPdf(List<PageRawInfo> pages, string outputPath)
    {
        using (var fs = new FileStream(outputPath, FileMode.Create))

        {
            var doc = new Document();
            var writer = PdfWriter.GetInstance(doc, fs);
            doc.Open();

            foreach (var p in pages)
            {
                // ページサイズを PDF の pt 単位で設定
                doc.SetPageSize(new iTextSharp.text.Rectangle((float)p.WidthPt, (float)p.HeightPt));
                doc.NewPage();

                using (var ms = new MemoryStream(p.Bytes))
                {
                    var img = iTextSharp.text.Image.GetInstance(ms.ToArray());

                    // ページ全体にフィットさせる
                    img.SetAbsolutePosition(0, 0);
                    img.ScaleAbsolute((float)p.WidthPt, (float)p.HeightPt);

                    // 回転がある場合
                    if (p.Rotation != 0)
                    {
                        img.RotationDegrees = p.Rotation;
                    }

                    doc.Add(img);
                }
            }

            doc.Close();
            writer.Close();
        }
    }
    public void EmbedStamp(string inputPdfPath, string outputPdfPath,string freeLicense, string dsXmlPath)
    {
        PdfReader reader = new PdfReader(inputPdfPath);
        PdfStamper stamper = new PdfStamper(reader, new FileStream(outputPdfPath, FileMode.Create));

        DataSet dataSet1 = new DataSet();
        dataSet1.ReadXml(dsXmlPath);

        DataTable stampPlacementTable = dataSet1.Tables["DataTableStampPlacement"];
        DataTable stampTable = dataSet1.Tables["DataTableStamp"];
        DataTable DataTableTextPlacement = dataSet1.Tables["DataTableTextPlacement"];
        DataTable DataTableRectPlacement = dataSet1.Tables["DataTableRectPlacement"];

        for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
        {
            var PageRotation = reader.GetPageRotation(pageNumber);
            // pageSize は iTextSharp の Rectangle (ページサイズ)
            var pageSize = reader.GetPageSize(pageNumber);
            float pageWidth = pageSize.Width;
            float pageHeight = pageSize.Height;

            //指定ページのスタンプ情報を取得
            foreach (DataRow row in stampPlacementTable.Select($"Page={pageNumber}"))
            {


                string stampPath = row["filename"].ToString();
                double left_Rate = Convert.ToDouble(row["Left"]);
                double top_Rate = Convert.ToDouble(row["Top"]);
                int stampRotation = Convert.ToInt16(row["Rotation"]);
                //short style = Convert.ToInt16(row["Style"]);
                string url = row["URL"].ToString();
                string id = row["ID"].ToString();

                //                 stamper.AddAnnotation(stampAnnot, pageNumber);
                DataView v = new DataView(stampTable);
                v.RowFilter = "filename = '" + stampPath + "'";
                if (v.Count > 0)
                {
                    string errstr = "get image size";
                    //値はPDF内寸法
                    float imgWidth = float.Parse(v[0]["width"].ToString());
                    float imgHeight = float.Parse(v[0]["height"].ToString());

                    // 画像読み込み
                    var stmpimg = iTextSharp.text.Image.GetInstance(stampPath);
                    if (stmpimg == null)
                    {
                        continue;
                    }

                    errstr = "calc x,y";
                    Console.WriteLine($"PDF PageSize: {pageWidth} , {pageHeight}");

                    float pdfX = 0f;
                    float pdfY = 0f;

                    pdfY = (float)(pageHeight * top_Rate);
                    pdfX = (float)(pageWidth * left_Rate);



                    switch (PageRotation % 360)
                    {
                        case 90:
                        case -270:
                            // 左下原点 → 右下原点に変換
                            float tmpX = pdfX;
                            pdfX = pageHeight - pdfY;
                            pdfY = pageWidth - tmpX;
                            break;
                        case -180:
                        case 180:
                            pdfX = pageWidth - pdfX;
                            //   pdfY = pageHeight - pdfY;
                            break;
                        case -90:
                        case 270:
                            //float tmpY = pdfY;
                            (pdfX, pdfY) = (pdfY, pdfX);
                            //pdfY= pageWidth - pdfX;
                            //pdfX = pageHeight - tmpY;
                            break;
                        default:
                        case 0:
                            //
                            pdfY = pageHeight - pdfY;
                            break;

                    }

                    //setAbsolutePositionは画像の左下が基準点,保存は中心にしてあるのでスタンプ位置を補正
                    //この補正は回転後に行う事。
                    float dispW = imgWidth;
                    float dispH = imgHeight;
                    int totalDeg = ((stampRotation + PageRotation) % 360 + 360) % 360;
                    if (totalDeg == 90 || totalDeg == 270)
                    {
                        dispW = imgHeight;
                        dispH = imgWidth;
                    }
                    pdfX -= dispW / 2f;
                    pdfY -= dispH / 2f;

                    if ((stampRotation + PageRotation) % 360 != 0)
                    {
                        stmpimg.Rotation = (float)(-totalDeg * Math.PI / 180.0);
                    }


                    stmpimg.SetAbsolutePosition(pdfX, pdfY);

                    stmpimg.ScaleAbsolute(imgWidth, imgHeight);
                    errstr = "newpdfgstate";
                    PdfGState gs1 = new PdfGState();

                    errstr = "setfillopacity";
                    gs1.FillOpacity = 0.7F;
                    errstr = "stamper.getovercontent";
                    // コンテンツ設定
                    // ページに描画
                    PdfContentByte cb = stamper.GetOverContent(pageNumber);


                    errstr = "setgstate";
                    cb.SetGState(gs1);
                    errstr = "addimage overcontent";
                    cb.AddImage(stmpimg);


                    Console.WriteLine($"PDF Location : ({left_Rate} , {top_Rate}");

                    //ハンコの領域を薄いグレーで塗る、ずれを把握するため
#if DEBUG
                    //                    AddColorRectangle(objPDFContByte, iTextSharp.text.Color.LIGHT_GRAY, x[idx], y[idx], width, height, 0.7F, nCount, stamper);
#endif
                    if (url.Length > 0)
                    {
                        // 画像の位置とサイズを矩形にする
                        var linkRect = new iTextSharp.text.Rectangle(pdfX, pdfY, pdfX + dispW, pdfY + dispH);

                        // URLアクションを作成
                        PdfAction action = new PdfAction(url);

                        // リンク注釈を作成
                        PdfAnnotation linkAnnot = PdfAnnotation.CreateLink(
                            stamper.Writer,
                            linkRect,
                            PdfAnnotation.HIGHLIGHT_NONE,
                            action
                        );
                        // タイトルにIDを設定
                        linkAnnot.Title = id;

                        // カスタムキーを埋め込む
                        linkAnnot.Put(new PdfName("ID"), new PdfString(id));

                        // ページに追加
                        stamper.AddAnnotation(linkAnnot, pageNumber);

                    }



                }
            }

            foreach (DataRow row in DataTableTextPlacement.Select($"Page={pageNumber}"))
            {
                string Font = row["Font"].ToString();
                float Size = Convert.ToSingle(row["Size"].ToString());
                long lColor = Convert.ToInt64(row["ColorARGB"].ToString());
                string text = row["text"].ToString();
                double left_Rate = Convert.ToDouble(row["Left"]);
                double top_Rate = Convert.ToDouble(row["Top"]);
                short Style = Convert.ToInt16(row["Style"]);
                //double textWidth = Convert.ToDouble(row["Width"]);
                //double textHeight = Convert.ToDouble(row["Height"]);
                int textRotation = Convert.ToInt16(row["Rotation"]);
                string url = row["URL"].ToString();
                string id = row["ID"].ToString();





                System.Drawing.Color ColorARGB = System.Drawing.Color.FromArgb((int)lColor);

                PdfContentByte cb = stamper.GetOverContent(pageNumber);

                // フォント準備（埋め込み推奨）
                FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), true);

                iTextSharp.text.Font font = FontFactory.GetFont(
                    Font,
                    BaseFont.IDENTITY_H,
                    BaseFont.NOT_EMBEDDED,
                    (float)Size,
                    Style,
                       new iTextSharp.text.Color((int)lColor)
                );
                string fontPath = @"C:\Windows\Fonts\msgothic.ttc,0";
                
                BaseFont bf = font.BaseFont;// nullだった
                // bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                cb.SetFontAndSize(bf, Size);

                // 色設定
                // System.Drawing.Color → iTextSharp.text.Color (RGB)
                System.Drawing.Color gdi = ColorARGB; // your ARGB
                var rgb = new iTextSharp.text.Color(gdi.R, gdi.G, gdi.B); // alpha is not used here
                if (id.Contains("hidetext"))
                {
                    PdfGState gs = new PdfGState();
                    gs.FillOpacity = 0f;   // 完全透明
                    cb.SetGState(gs);
                }
                else
                {
                    // 通常の色設定
                    cb.SetColorFill(rgb);
                }


                //// 透明度設定（必要なら）
                //PdfGState gs = new PdfGState();
                //gs.FillOpacity = 0.7f;   // 不要なら削除
                //cb.SetGState(gs);

                float pdfX = 0f;
                float pdfY = 0f;

                pdfY = (float)(pageHeight * top_Rate);
                pdfX = (float)(pageWidth * left_Rate);

                switch (PageRotation % 360)
                {
                    case 90:
                    case -270:
                        (pdfX, pdfY) = (pdfY, pdfX);
                        pdfX = pageHeight - pdfX;
                        pdfY = pageWidth - pdfY;// 左下原点 → 右下原点に変換
                        break;
                    case -180:
                    case 180:
                        pdfX = pageWidth - pdfX;
                        break;
                    case -90:
                    case 270:
                        (pdfX, pdfY) = (pdfY, pdfX);
                        break;
                    default:
                    case 0:
                        //
                        pdfY = pageHeight - pdfY;// 左下原点 → 右下原点に変換
                        break;

                }
                var sizef = WordsAreaSizeOnPDF(Font, text, (float)Size, Style, cb, true);
                var sizeL = WordsAreaSizeOnPDF(Font, text, (float)Size, Style, cb, false);

                float imgWidth = sizef.Width;
                float imgHeight = sizef.Height;

                //setAbsolutePositionは画像の左下が基準点,保存は中心にしてあるのでスタンプ位置を補正
                //この補正は回転後に行う事。
                //iTextSharp では「理論値」と「実際の描画結果」に差が出ることがあるので、一度 PDF に出力してから座標差分を計測するのが正解です。
                //0.95f=40文字で約2文字のズレから
                float dispW = imgWidth * 0.95f;   //ShowTextAlignedは横座標は中心、ｘ補正不要
                float dispH = imgHeight;
                int totalDeg = ((textRotation + PageRotation) % 360 + 360) % 360;
                if (totalDeg == 90 || totalDeg == 270)
                {
                    (dispW, dispH) = (dispH, dispW);
                }
                switch (PageRotation % 360)
                {
                    case 90:
                    case -270:

                        pdfX += dispW / 2f;
                        pdfY += dispH / 2f;

                        break;
                    case -180:
                    case 180:

                        pdfX += dispW / 2f;
                        pdfY += dispH / 2f;

                        break;
                    case -90:
                    case 270:

                        pdfX += dispW / 2f;
                        pdfY -= dispH / 2f;

                        break;
                    default:
                    case 0:
                        pdfX -= dispW / 2f;
                        pdfY += dispH / 2f;
                        break;

                }


                //itextsharpは反時計回り

                //ラジアンで指定するとき
                //if ((textRotation + PageRotation) % 360 != 0)
                //{
                //    stmpimg.Rotation = (float)(-totalDeg * Math.PI / 180.0);
                //}

                //ShowTextAlignedの場合
                // テキスト描画開始
                cb.BeginText();

                float drawX = pdfX;
                float drawY = pdfY;

                // テキストサイズ
                //1行の高さ
                float ascentSingle = bf.GetAscentPoint("テスト", (float)Size);
                float descentSingle = bf.GetDescentPoint("テスト", (float)Size);
                float _lineHeight = (ascentSingle - descentSingle);

                // 全体の幅・高さ、ただし改行だけやスペースだけは高さがなくなり、ずれてしまう
                float _textWidth = bf.GetWidthPoint(text, (float)Size);
                float ascent = bf.GetAscentPoint(text, (float)Size);
                float descent = bf.GetDescentPoint(text, (float)Size);
                float _textHeight = 0;// ascent - descent;
                                      //ここで補正
                string[] lines = text.Split('\n'); // 改行で分割

                float leading = _lineHeight * 1.2f; // 行送り（フォントサイズの1.2倍など）

                Console.WriteLine($"lineHeight*1.2 : {leading}");
                // WPF の LineHeight を取得して比率を計算したいがGDI+で妥協
                double lineHeightWpf = GetLineHeight_GDI( Font, Size);
                leading = (float)lineHeightWpf; // PDF 側にそのまま適用
                Console.WriteLine($"WPF LineHeight : {leading}");

                //wpfではfontsize*1.2を行送りにした
                leading = Size * 1.2f * 0.96f;
                //実測による補正、２０文字で１文字くらい余計に出力

                _textHeight = leading * lines.Length;
                Console.WriteLine($"Asent/Deceent Point {_textHeight} : Width Point {_textWidth}");

                //ColumnTextで矩形を計る方法
                //高さを正確に出せる
                int maxW = 50000;
                int maxH = 5000;
                ColumnText ct = new ColumnText(null);
                ct.SetSimpleColumn(new Phrase(text, font), 0, 0, maxW, maxH, leading, Element.ALIGN_LEFT);
                int status = ct.Go(true); // true = simulate (描画せず計算だけ)
                _textHeight = maxH - ct.YLine; // 実際に消費した高さ

                float maxWidth = 0;
                foreach (string line in lines)
                {
                    float w = bf.GetWidthPoint(line, (float)Size);
                    if (w > maxWidth) maxWidth = w;
                }
                _textWidth = maxWidth;
                Console.WriteLine($"ColumnText scale {_textHeight} : Width Point(SingleLine) {_textWidth} ");



                switch (totalDeg)
                {
                    case 0:
                        // 正位置はベースライン補正のみ
                        drawY -= _lineHeight;
                        break;

                    case 90:
                        //ベースライン補正
                        drawX -= _lineHeight;
                        break;

                    case 180:
                        // 上下反転なので高さ分補正
                        drawY -= dispH;
                        //ベースライン補正
                        drawY += _lineHeight;
                        //drawX += _textWidth;
                        break;

                    case 270:
                        drawX -= dispW;
                        //ベースライン補正
                        drawX += _lineHeight;
                        //drawY -= _textWidth;
                        break;
                }

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];

                    // 行ごとに Y 座標をずらす
                    float lineX = drawX;
                    float lineY = drawY;
                    switch (totalDeg)
                    {
                        case 0:
                            lineY = drawY - i * leading; // 下方向へ
                            break;
                        case 90:
                            lineX = drawX - i * leading; // 左方向へ
                            break;
                        case 180:
                            lineY = drawY + i * leading; // 上方向へ
                            break;
                        case 270:
                            lineX = drawX + i * leading; // 右方向へ
                            break;
                    }


                    cb.ShowTextAligned(
                        Element.ALIGN_LEFT,
                        line,
                        lineX,
                        lineY,
                        -totalDeg // 角度はそのまま
                    );
                }


                cb.EndText();




            }


            foreach (DataRow row in DataTableRectPlacement.Select($"Page={pageNumber}"))
            {
                // string Font = row["Font"].ToString();
                float Opacity = Convert.ToSingle(row["Size"].ToString());//透明度 0.0～1.0　0は枠のみ
                long lColor = Convert.ToInt64(row["ColorARGB"].ToString());//線色
                                                                           // string text = row["text"].ToString();
                double left_Rate = Convert.ToDouble(row["Left"]);
                double top_Rate = Convert.ToDouble(row["Top"]);
                short LineWidth = Convert.ToInt16(row["Style"]);//線太さ
                double Width_Rate = Convert.ToDouble(row["Width"]);
                double Height_Rate = Convert.ToDouble(row["Height"]);
                int RectRotation = Convert.ToInt16(row["Rotation"]);
                string url = row["URL"].ToString();
                string id = row["ID"].ToString();
                float fillopacity = Opacity;
                if (Opacity == 0)
                {
                    Opacity = 1.0f;
                }

                System.Drawing.Color ColorARGB = System.Drawing.Color.FromArgb((int)lColor);

                PdfContentByte cb = stamper.GetOverContent(pageNumber);
                // 半透明設定
                PdfGState gs = new PdfGState
                {
                    FillOpacity = fillopacity,   // 塗りつぶしの透明度（0.0 完全透明 ～ 1.0 不透明）
                    StrokeOpacity = Opacity  // 枠線の透明度
                };

                cb.SaveState();
                cb.SetGState(gs);

                // 色設定
                // System.Drawing.Color → iTextSharp.text.Color (RGB)
                System.Drawing.Color gdi = ColorARGB; // your ARGB
                var rgb = new iTextSharp.text.Color(gdi.R, gdi.G, gdi.B); // alpha is not used here

                cb.SetColorFill(rgb);
                cb.SetLineWidth(LineWidth);

                cb.SetColorStroke(rgb);

                float pdfX = 0f;
                float pdfY = 0f;
                float pdfH = 0f;
                float pdfW = 0f;

                pdfY = (float)(pageHeight * (top_Rate));
                pdfX = (float)(pageWidth * (left_Rate));
                pdfW = (float)(pageWidth * Width_Rate);
                pdfH = (float)(pageHeight * Height_Rate);

                switch (PageRotation % 360)
                {
                    case 90:
                    case -270:
                        (pdfY, pdfX) = (pdfX, pdfY);
                        (pdfW, pdfH) = (pdfH, pdfW);
                        pdfX = pageHeight - pdfX;//回転補正
                        pdfY = pageWidth - pdfY;//上０下０補正
                        pdfY -= pdfH / 2;
                        pdfX -= pdfW / 2;
                        break;
                    case -180:
                    case 180:
                        pdfX = pageWidth - pdfX;
                        pdfX -= pdfW / 2;
                        pdfY -= pdfH / 2;
                        break;
                    case -90:
                    case 270:
                        (pdfY, pdfX) = (pdfX, pdfY);
                        (pdfW, pdfH) = (pdfH, pdfW);
                        pdfY -= pdfH / 2;
                        pdfX -= pdfW / 2;

                        break;
                    default:
                    case 0:
                        //
                        pdfY = pageHeight - pdfY;
                        pdfX -= pdfW / 2;
                        pdfY -= pdfH / 2;

                        break;

                }

                cb.Rectangle(pdfX, pdfY, pdfW, pdfH);
                cb.FillStroke();//塗りつぶしと枠線描画




            }


            if (freeLicense=="true")
            {
                // コンテンツ設定
                var objPDFContByte = stamper.GetOverContent(pageNumber);


                AddLOGO(objPDFContByte);
                //LOGO

                objPDFContByte = stamper.GetOverContent(pageNumber);

                AddTradeMark(objPDFContByte);
            }


        }

        stamper.Close();
        reader.Close();

    }
    private void AddLOGO(PdfContentByte objPDFContByte)
    {
        iTextSharp.text.Font font =
FontFactory.GetFont("C:\\WINDOWS\\Fonts\\meiryo.ttc,0",
BaseFont.IDENTITY_H,            //横書き
BaseFont.NOT_EMBEDDED,          //フォントを組み込まない
9,
iTextSharp.text.Font.NORMAL,
iTextSharp.text.Color.GRAY);

        ColumnText ct = new ColumnText(objPDFContByte);
        //(x,y,x+width,y+height,fontsize,align)
        ct.SetSimpleColumn(14, 10, 250, 20, 0, Element.ALIGN_LEFT);
        Chunk cnk = new Chunk("            Powered by OURS SOFT", font);
        cnk.SetAnchor("http://ourssoft.cloudfree.jp/pdfnet/");
        ct.AddText(cnk);
        ct.Go();

    }
    private void AddTradeMark(PdfContentByte objPDFContByte)
    {
        //半透明にならない
        //System.Drawing.Image im = System.Drawing.Image.FromFile(Application.StartupPath + @"\tm.gif");

        //System.Drawing.Imaging.ColorMatrix cm = new System.Drawing.Imaging.ColorMatrix();
        //cm.Matrix00 = 1;
        //cm.Matrix11 = 1;
        //cm.Matrix22 = 1;
        //cm.Matrix33 = 0.2F;
        //cm.Matrix44 = 1;
        //var ia = new System.Drawing.Imaging.ImageAttributes();
        //ia.SetColorMatrix(cm);
        //var canvas = new Bitmap(im.Width, im.Height);
        //var g = Graphics.FromImage(canvas);
        //g.DrawImage(im, new System.Drawing.Rectangle(0, 0, im.Width, im.Height), 0, 0, im.Width, im.Height, GraphicsUnit.Pixel, ia);
        //im.Dispose();
        //g.Dispose();

        //var img = iTextSharp.text.Image.GetInstance(canvas, System.Drawing.Imaging.ImageFormat.Gif);
        // 埋め込みリソースから読み込む例
        Assembly asm = Assembly.GetExecutingAssembly();
        using (Stream s = asm.GetManifestResourceStream("PDFNETWorker.tm.gif"))
        {
            var img = iTextSharp.text.Image.GetInstance(s);
            img.SetAbsolutePosition(14, 10);
            img.ScaleAbsolute(32, 32);
            //画像を薄くする
            PdfGState gs1 = new PdfGState();
            gs1.FillOpacity = 0.2F;
            objPDFContByte.SetGState(gs1);
            objPDFContByte.AddImage(img);
        }


    }

    private System.Drawing.SizeF WordsAreaSizeOnPDF(  string Font, string text, float Size, int Style,  PdfContentByte cb, bool isMultiLine)
    {
        FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), true);

        iTextSharp.text.Font font = FontFactory.GetFont(
            Font,
            BaseFont.IDENTITY_H,
            BaseFont.NOT_EMBEDDED,
            Size,
            Style,
            iTextSharp.text.Color.BLACK
        );

        BaseFont bf = font.BaseFont;
        cb.SetFontAndSize(bf, Size);

        float maxWidthPt = 0;
        float totalHeightPt = 0;

        foreach (string s in text.Split('\n'))
        {


            //1もじずつ測定
            float widthchars = 0;
            foreach (char c in s)
            {
                widthchars += bf.GetWidthPoint(c.ToString(), Size);
            }

            //1行を測定
            float widthPoint = bf.GetWidthPoint(s, Size);
            float ascent = bf.GetAscentPoint(s, Size);
            float descent = bf.GetDescentPoint(s, Size);
            float heightPoint = ascent - descent;
            //全く同じだった
            Console.WriteLine($"width chars : {widthchars}  widthPoint : {widthPoint}");

            if (maxWidthPt < widthPoint) maxWidthPt = widthPoint;
            totalHeightPt += heightPoint * 1.2f; // 行間補正
            if (!isMultiLine) break;
        }

        return new System.Drawing.SizeF(maxWidthPt, totalHeightPt);
    }

    private float GetLineHeight_GDI(string fontFamily, float fontSize)
    {
        using (var bmp = new Bitmap(1, 1))
        using (var g = Graphics.FromImage(bmp))
        using (var font = new System.Drawing. Font(fontFamily, fontSize))
        {
            return font.GetHeight(g);
        }
    }
    public WorkerResult MergeIntoTempPdf(string droppedFile)
    {
        string newPdf = "";
        //ショートカットファイルの場合はリンク先を取得
        if (droppedFile.ToLower().EndsWith(".lnk"))
        {
            string target = ShortcutReader.GetShortcutTarget(droppedFile);
            droppedFile = target;
        }
        if (droppedFile.ToLower().EndsWith(".pdf"))
        {
            newPdf = droppedFile;

        }
        else
        {
            if (droppedFile.ToLower().EndsWith(".bmp")
            || droppedFile.ToLower().EndsWith(".jpeg")
             || droppedFile.ToLower().EndsWith(".jpg")
             || droppedFile.ToLower().EndsWith(".png")
            )
            {
                clsPDF cls = new clsPDF();
                newPdf = Path.GetTempPath() + Path.GetRandomFileName() + ".pdf";
                cls.bitmap_to_pdf(new string[] { droppedFile }, newPdf);
            }
            else
            {
                if (droppedFile.ToLower().EndsWith(".tif")
                    || droppedFile.ToLower().EndsWith(".tif")
                    )
                {
                    clsPDF cls = new clsPDF();
                    string tmpTIF = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
                    Directory.CreateDirectory(tmpTIF);
                    cls.ConvertMultipageTifToBmpKeepDpi(droppedFile, tmpTIF);
                    string[] files = Directory.GetFiles(tmpTIF, "*.bmp");
                    newPdf = Path.GetTempPath() + Path.GetRandomFileName() + ".pdf";
                    cls.bitmap_to_pdf(files, newPdf);

                }
            }
        }

        if (newPdf == Config.tempPdfPath)
        {
            //pdffnetが開くファイルはあらかじめtempPdfPathにコピーされているため、
            //スキップ
            return new WorkerResult
            {
                Success = true,
                Message = "tempPdf = droppedFile",
                
            };
        }
        if (!File.Exists(Config.tempPdfPath))
        {
            //ファーストドロップのときはここになる
            File.Copy(newPdf, Config.tempPdfPath);
            return new WorkerResult
            {
                Success = true,
                Message = "First Drop",
                
            };
        }
        try
        {
            var reader1 = new PdfReader(Config.tempPdfPath);
            var reader2 = new PdfReader(newPdf);

            var fs = new FileStream(Config.tempPdfPath + ".tmp", FileMode.Create);
            var doc = new Document();
            var copy = new PdfCopy(doc, fs);

            doc.Open();
            for (int i = 1; i <= reader1.NumberOfPages; i++)
                copy.AddPage(copy.GetImportedPage(reader1, i));
            for (int i = 1; i <= reader2.NumberOfPages; i++)
                copy.AddPage(copy.GetImportedPage(reader2, i));

            copy.Close();      // 明示的に PdfCopy を閉じる
            doc.Close();       // Document を閉じる（PdfCopy.Close() より後）
            fs.Close();        // 最後に FileStream を閉じる

            reader1.Close();
            reader2.Close();



            File.Delete(Config.tempPdfPath);
            File.Move(Config.tempPdfPath + ".tmp", Config.tempPdfPath);
            return new WorkerResult
            {
                Success = true,
                Message = "Merge",

            };

        }
        catch (Exception ex)
        {
            return new WorkerResult
            {
                Success = true,
                Message = $"no PDF file\n{ex.ToString()}",

            };

        }
        finally
        {

        }


    }
    public WorkerResult GetPageRotation(string pdfFilePath, int pageNumber)
    {
        var reader = new PdfReader(pdfFilePath);

        PdfDictionary pageDict = reader.GetPageN(pageNumber);
        PdfNumber rotate = pageDict.GetAsNumber(PdfName.ROTATE);
        reader.Close();
        var pginfo = new PageInfo()
        {
            PageNumber = pageNumber,
            Rotation = rotate != null ? rotate.IntValue : 0
        }            ;
        var result = new WorkerResult
        {
            Success = true,
            AddedPages = null,
            TempPdfPath = null,
            PageInfo = new List<PageInfo>() 
            {
                pginfo,
            }


        };

        return result;
    }

    public WorkerResult MakeBrankPage()
    {
        int addedPageNumber = 1;
        using (var fs = new FileStream(Config.tempPdfPath + ".tmp", FileMode.Create))
        {
            var doc = new Document();
            var copy = new PdfCopy(doc, fs);
            doc.Open();
            if (File.Exists(Config.tempPdfPath))
            {
                var reader = new PdfReader(Config.tempPdfPath);

                // 既存ページをコピー
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    copy.AddPage(copy.GetImportedPage(reader, i));
                }
                addedPageNumber = reader.NumberOfPages + 1;
            }
            // 空白ページを追加
            var blankDoc = new Document(PageSize.A4);
            var blankStream = new MemoryStream();
            var blankWriter = PdfWriter.GetInstance(blankDoc, blankStream);
            blankDoc.Open();
            blankDoc.NewPage();

            // 空白ページに白い矩形を描画（これでページが有効になる）
            var cb = blankWriter.DirectContent;
            cb.Rectangle(0, 0, PageSize.A4.Width, PageSize.A4.Height);
            cb.SetColorFill(iTextSharp.text.Color.WHITE);
            cb.Fill();

            blankDoc.Close();

            var blankReader = new PdfReader(blankStream.ToArray());
            var blankImportedPage = copy.GetImportedPage(blankReader, 1); // ← 修正
            copy.AddPage(blankImportedPage); // ← 修正
            
            doc.Close();
        }
        const double A4WidthPt = 595.0;  // 8.27 inch × 72
        const double A4HeightPt = 842.0; // 11.69 inch × 72

        var blankPage = new PageInfo
        {
            WidthPt = A4WidthPt,
            HeightPt = A4HeightPt,
            PageNumber = addedPageNumber,
            OriginalPageNumber = addedPageNumber,
        };


        File.Delete(Config.tempPdfPath);
        File.Move(Config.tempPdfPath + ".tmp", Config.tempPdfPath);

        var result = new WorkerResult
        {
            Success = true,
            AddedPages = null,
            TempPdfPath = null,
            PageInfo = new List<PageInfo>()
            {
                blankPage,
            }


        };

        return result;

    }
    public List<PageRawInfo> LoadPdfPagesRaw(string pdfFilePath)
    {
        List<PageRawInfo> loadPagesRaw = new();

        try
        {
            string exeDir = Path.GetDirectoryName(
                Process.GetCurrentProcess().MainModule.FileName
            );

            var version = new GhostscriptVersionInfo(
                Path.Combine(exeDir, "gsdll64.dll")
            );

            string tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(pdfFilePath));
            if (pdfFilePath != tempPath) File.Copy(pdfFilePath, tempPath, true);

            using var rasterizer = new GhostscriptRasterizer();
            rasterizer.Open(tempPath, version, false);

            int dpi = rasterizer.PageCount >= 50 ? 96 : 150;

            if (rasterizer.PageCount == 0)
            {
                return loadPagesRaw;
            }

            for (int pageNumber = 1; pageNumber <= rasterizer.PageCount; pageNumber++)
            {
                var sk = rasterizer.GetPage(dpi, pageNumber); // SKBitmap

                using var image = SkiaSharp.SKImage.FromBitmap(sk);
                using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                byte[] pngBytes = data.ToArray();

                double widthPt = sk.Width * 72.0 / dpi;
                double heightPt = sk.Height * 72.0 / dpi;

                int rotation = 0;
                var workerResult = GetPageRotation(tempPath, pageNumber);
                if (workerResult.Success && workerResult.PageInfo != null && workerResult.PageInfo.Count > 0)
                {
                    rotation = workerResult.PageInfo[0].Rotation;
                }

                loadPagesRaw.Add(new PageRawInfo
                {
                    Bytes = pngBytes,
                    WidthPt = widthPt,
                    HeightPt = heightPt,
                    PageNumber = pageNumber,
                    OriginalPageNumber = pageNumber,
                    Rotation = rotation
                });


            }
        }
        catch (Exception ex)
        {
            // worker は例外を JSON で返す
            return new List<PageRawInfo>();
        }

        return loadPagesRaw;
    }
}
