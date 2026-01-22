using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;

using iTextSharp.text;
using iTextSharp.text.exceptions;//5.5
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

using Newtonsoft.Json;
using Org.BouncyCastle.X509.Extension;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Security.Cryptography;
using System.Text;



class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        // ---- ワーキングセット拡張 ----
        try
        {
            var p = Process.GetCurrentProcess();

            // 現在の最大ワーキングセットを取得
            long currentMax = p.MaxWorkingSet.ToInt64();

            Trace.WriteLine("WS(before)=" + p.WorkingSet64);
            Trace.WriteLine("MaxWS(before)=" + currentMax);

            // 最小ワーキングセットを「現在の最大値の範囲内」で設定
            long desiredMin = 256L * 1024L * 1024L; // 256MB

            if (desiredMin > currentMax)
            {
                // OS が Max を小さく設定している場合は、Max に合わせる
                desiredMin = currentMax;
            }

            p.MinWorkingSet = (IntPtr)desiredMin;

            Trace.WriteLine("MinWS(after)=" + p.MinWorkingSet.ToInt64());
            Trace.WriteLine("WS(after)=" + p.WorkingSet64);
        }
        catch (Exception ex)
        {
            Console.WriteLine("WorkingSet set failed: " + ex.Message);
        }




        //string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        //string path = Path.Combine(desktop, "worker_test.txt");
        //string log = "Is64BitProcess=" + Environment.Is64BitProcess + "\r\n" +
        //    "User=" + Environment.UserName + "\r\n" +
        //    "WorkingSet=" + Process.GetCurrentProcess().WorkingSet64;

        //File.WriteAllText(path,log );
        GCSettings.LatencyMode = GCLatencyMode.SustainedLowLatency;

        int result = RunWorker(args);
        Environment.Exit(result);
    }
    static int RunWorker(string[] args)
    {
        
        // これを最初に呼ぶ
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 以降 iTextSharp が 1252 を使えるようになる

        //これは 呼ぶたびにフォントを再スキャンし、FontFactory の内部 Dictionary に登録を追加します。
        //アプリ全体で 1 回だけ呼ぶ
        FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), true);



#if DEBUG
        //          args = new string[] { "pdf2png", @"C:\Users\takes\OneDrive\デスクトップ\図面１ - コピー.pdf", @"C:\Users\takes\AppData\Local\Temp\a5947e51-e9a8-4a8a-9dfd-6804df66031d.tmp" };
        //   args = new string[] { "merge", @"C:\Users\takes\OneDrive\デスクトップ\kyuyofukuro.pdf?", @"C:\Users\takes\AppData\Local\Temp\a5947e51-e9a8-4a8a-9dfd-6804df66031d.tmp" };
        //  args = new string[] { "merge", @"C:\Users\takes\OneDrive\デスクトップ\kyuyofukuro.pdf?"};
        //  args = new string[] { "gettext", @"C:\Users\takes\OneDrive\デスクトップ\b.pdf","1",@"C:\Users\takes\OneDrive\デスクトップ\PDFNET_SAMPLE\sample-multilingual-text.json" };
//        args = new string[] { "gettagtext", @"C:\Users\takes\OneDrive\デスクトップ\b.pdf", "1", @"C:\Users\takes\OneDrive\デスクトップ\PDFNET_SAMPLE\sample-multilingual-text.json" };

        // args = new string[] { "pdf2png", @"C:\Users\takes\OneDrive\デスクトップ\protect.pdf?aaaaaa", @"C:\Users\takes\AppData\Local\Temp\a5947e51-e9a8-4a8a-9dfd-6804df66031d.tmp", @"C:\Users\takes\AppData\Local\Temp\tmp.json" };
        //  args = new string[] { "stamp", @"C:\Users\takes\OneDrive\デスクトップ\_blank.pdf.paging", @"C:\Users\takes\OneDrive\デスクトップ\_blank.pdf", @"C:\Users\takes\AppData\Local\Temp\tmpjusthd.tmp" };

#endif

        //戻り値
        // 0 success
        // 1 Syntax error
        // 2 I/O error
        // 3 PasswordFile
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
            case "gettext":
                {
                    if (args.Length < 4)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe gettext <targetFilePath> <PageNumber> <outputJson>");
                        return 1;
                    }

                    string file = args[1];
                    int page = int.Parse(args[2]);//1 base
                    string outputJson = args[3];

                    var extractor = new WorkerTextExtractor();
                    var doc = extractor.Extract(file);
                    //foreach (var p in doc.Pages)
                    //{
                    //    Trace.WriteLine($"PageNumber={p.PageNumber}, TextLength={p.Items.Count()}");
                    //}
                    
                    var pageInfo = doc.Pages.FirstOrDefault(p => p.PageNumber == page);
                    string fullText = string.Join("", pageInfo.Items.Select(i => i.Text));
                    // JSON で返す
                    //File.WriteAllText(outputJson, JsonConvert.SerializeObject(pageInfo, Formatting.Indented));
                    File.WriteAllText(outputJson, fullText);

                    //Console.WriteLine(JsonConvert.SerializeObject(pageInfo, Formatting.Indented));
                    return 0;
                }
            case "gettagtext":
                {
                    if (args.Length < 3)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe gettext <targetFilePath> <outputJson>");
                        return 1;
                    }

                    string file = args[1];
                    string outputJson = args[2];

                    var extractor = new WorkerTextExtractor();
                    var doc = extractor.Extract(file);

                    // JSON で返す
                    File.WriteAllText(outputJson, JsonConvert.SerializeObject(doc.Pages, Formatting.Indented));

                    return 0;
                }


            case "2in1":
                {
                    if (args.Length < 4)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe 2in1 <sourceFilePath> <1stPageNumber> <2ndPageNumber> <outputPath>");
                        return 1;
                    }
                    string sourceFilePath = args[1];
                    int P1 = int.Parse( args[2]);
                    int P2 = int.Parse(args[3]);
                    string newpg = args[4];// Path.Combine(Path.GetTempPath(), "2in1.pdf");

                    try
                    {
                        var worker = new PdfPagingWorker();
                        worker.page2in1(sourceFilePath, P1, P2,newpg);
                        return 0;
                    }
                    catch (Exception ex)
                    {

                        Console.Error.WriteLine(ex);
                        return 2;
                    }

                }

            case "paging":
                {
                    if (args.Length < 4)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe paging <src> <dest> <pageInfoJson>");
                        return 1;
                    }

                    string src = args[1];
                    string dest = args[2];
                    string pageInfoJson = args[3];

                    try
                    {
                        var pageInfos = JsonConvert.DeserializeObject<List<PageInfo>>(File.ReadAllText(pageInfoJson));
                        var worker = new PdfPagingWorker();
                        worker.savePagingPDF(src, dest, pageInfos);
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
                    if (args.Length < 4)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe stamp <read> <dest> <xml> <EncryptNumber>");
                        return 1;
                    }

                    string read = args[1];
                    string dest = args[2];
                    
                    string xml = args[3];
                    string encryptno = "";
                    if(args.Length > 4) encryptno = args[4];


                    try
                    {
                        var worker = new PdfPagingWorker();
                        worker.EmbedStamp(read, dest, xml,encryptno);
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
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe merge <read?password>");
                        return 1;
                    }

                    string read = args[1].Split('?')[0];
                    string password =  args[1].Contains("?") ? args[1].Split('?')[1] : "";

                    try
                    {
                        var worker = new PdfPagingWorker();
                        var result= worker.MergeIntoTempPdf(read,password);
                        
                        // ★ JSON で返す
                        Console.WriteLine(JsonConvert.SerializeObject(result));

                        if (result.Success == false)
                        {
                            if (result.Message.Contains("Password"))
                            {
                                Console.Error.WriteLine("パスワードが違います");

                                return 3;
                            }
                            else
                            {
                                if (result.Message.Contains("CantOpen"))
                                {
                                    Console.Error.WriteLine("このファイルは編集モードで開けません");

                                    return 2;
                                }
                            }
                        }

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
                    if (args.Length < 2)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe makeblankpage <papersizeName> ");
                        return 1;
                    }
                    string papersize = args[1];


                    try
                    {
                        var worker = new PdfPagingWorker();
                        WorkerResult result = worker.MakeBrankPage(papersize);

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
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe pdf2png <pdfPath?password> <outputJson>");
                        return 1;
                    }

                    string pdfPath = args[1].Split('?')[0];
                    string password = args[1].Split('?')[1];
                    string outputJson = args[2];

                    try
                    {
                        var worker = new PdfPagingWorker();
                        List<PageRawInfo> pages = worker.LoadPdfPagesRaw(pdfPath,password);
                        
                        // JSON で返す
                        File.WriteAllText(outputJson, JsonConvert.SerializeObject(pages));

                        if (pages.Count > 0 && pages[0].Error !=null && pages[0].Error.Contains("PasswordProtected")) return 3;
                        

                        return 0;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex);

                        return 2;
                    }
                }
            case "getpagesizes":
                {
                    if (args.Length < 1)
                    {
                        Console.Error.WriteLine("Usage: PDFNETWorker.exe getpagesizes");
                        return 1;
                    }

                    try
                    {
                        var infos = PageSizeResolver.GetPageSizeInfos();

                        string json = JsonConvert.SerializeObject(
                            infos,
                            Formatting.Indented
                        );

                        Console.WriteLine(json);
                        return 0;


        //                var names = PageSizeResolver.GetPageSizeNames();

        //                string json = JsonConvert.SerializeObject(
        //    names,
        //    Formatting.Indented
        //);

        //                Console.WriteLine(json);
        //                return 0;


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
    public string? Error {  get; set; }
}

public static class PageSizeResolver
{
    private static readonly Dictionary<string, iTextSharp.text.Rectangle> _pageSizes;

    static PageSizeResolver()
    {
        _pageSizes = new Dictionary<string, iTextSharp.text.Rectangle>(StringComparer.OrdinalIgnoreCase);

        var fields = typeof(PageSize).GetFields(BindingFlags.Public | BindingFlags.Static);

        foreach (var field in fields)
        {
            if (field.FieldType == typeof(iTextSharp.text.Rectangle))
            {
                _pageSizes[field.Name] = (iTextSharp.text.Rectangle)field.GetValue(null);
            }
        }
    }
    public static List<PaperInfo> GetPageSizeInfos()
    {
        var list = new List<PaperInfo>();

        var fields = typeof(PageSize).GetFields(BindingFlags.Public | BindingFlags.Static);

        foreach (var field in fields)
        {
            if (field.FieldType == typeof(iTextSharp.text. Rectangle))
            {
                var rect = (iTextSharp.text.Rectangle)field.GetValue(null);

                list.Add(new PaperInfo
                {
                    Name = field.Name,
                    Width = rect.Width,
                    Height = rect.Height
                });
            }
        }

        return list.OrderBy(x => x.Name).ToList();
    }

    // ★ メインアプリに返す API（string のみ）
    public static List<string> GetPageSizeNames()
    {
        return _pageSizes.Keys
            .OrderBy(x => x)
            .ToList();
    }

    // ★ PDF 生成時に使う API
    public static iTextSharp.text.Rectangle GetPageSize(string name)
    {
        if (_pageSizes.TryGetValue(name, out var rect))
            return rect;

        throw new ArgumentException($"Unknown paper size: {name}");
    }
}
public class PdfEncryptionSettings
{
    public string UserPassword { get; set; }
    public string OwnerPassword { get; set; }
    public int Permissions { get; set; }
    public int EncryptionType { get; set; }
}

public class PaperInfo
{
    public string Name { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
}
public class PdfTextItem
{
    public string Text { get; set; }
    public float X { get; set; }        // ページ座標系（左下原点）
    public float Y { get; set; }
    public float Angle { get; set; }    // 度数（0 = 水平, 反時計回りが正）
    public string Tag { get; set; }     // BDCタグ名（なければ null）
}

public class PdfPageTextInfo
{
    public int PageNumber { get; set; }
    public List<PdfTextItem> Items { get; set; } = new List<PdfTextItem>();
}

public class PdfDocumentTextInfo
{
    public string FilePath { get; set; }
    
    public List<PdfPageTextInfo> Pages { get; set; } = new List<PdfPageTextInfo>();
}

public class WorkerTextListener : IExtRenderListener
{
    private readonly PdfPageTextInfo _pageInfo;
    private string _currentTag = null;

    public WorkerTextListener(PdfPageTextInfo pageInfo)
    {
        _pageInfo = pageInfo;
    }

    // ---- IRenderListener ----

    public void RenderText(TextRenderInfo renderInfo)
    {
        string text = renderInfo.GetText();
        if (string.IsNullOrEmpty(text))
            return;

        // 座標（ベースラインの開始点）
        var baseline = renderInfo.GetBaseline();
        var start = baseline.GetStartPoint();
        var end = baseline.GetEndPoint();

        float x = start[Vector.I1];
        float y = start[Vector.I2];

        // 角度を計算（ラジアン -> 度）
        float dx = end[Vector.I1] - start[Vector.I1];
        float dy = end[Vector.I2] - start[Vector.I2];
        float angleRad = (float)Math.Atan2(dy, dx);
        float angleDeg = angleRad * 180f / (float)Math.PI;

        var item = new PdfTextItem
        {
            Text = text,
            X = x,
            Y = y,
            Angle = angleDeg,
            Tag = _currentTag
        };

        _pageInfo.Items.Add(item);
    }

    public void RenderImage(ImageRenderInfo renderInfo)
    {
        // 今回は画像は無視
    }

    public void BeginTextBlock()
    {
    }

    public void EndTextBlock()
    {
    }

    // ---- IExtRenderListener (Marked Content 用) ----

    public void BeginMarkedContent(PdfName tag, PdfDictionary dict)
    {
        // /FieldName → "FieldName" のように加工
        _currentTag = tag != null ? tag.ToString().TrimStart('/') : null;
    }

    public void EndMarkedContent()
    {
        _currentTag = null;
    }

    public void BeginMarkedContentSequence(PdfName tag, PdfDictionary dict)
    {
        BeginMarkedContent(tag, dict);
    }

    public void EndMarkedContentSequence()
    {
        EndMarkedContent();
    }
    public void ModifyPath(PathConstructionRenderInfo renderInfo)
    {
        // パス（線・図形）には興味がないので空実装でOK
    }
    public iTextSharp.text.pdf.parser.Path RenderPath(PathPaintingRenderInfo renderInfo)
    {
        // 図形の描画（線・塗りつぶし）は無視
        return null;
    }
    public void ClipPath(int rule)
    {
        // クリッピングパスも無視
    }


}

public class WorkerTextExtractor
{
    public PdfDocumentTextInfo Extract_OLD(string pdfPath)
    {
        var result = new PdfDocumentTextInfo
        {
            FilePath = pdfPath
        };

        using (var reader = new PdfReader(pdfPath))
        {
            int pageCount = reader.NumberOfPages;

            for (int page = 1; page <= pageCount; page++)
            {
                var pageInfo = new PdfPageTextInfo
                {
                    PageNumber = page
                };

                // リスナーをページ単位で作成
                var listener = new WorkerTextListener(pageInfo);
                var processor = new PdfContentStreamProcessor(listener);

                PdfDictionary pageDic = reader.GetPageN(page);
                PdfDictionary resources = pageDic.GetAsDict(PdfName.RESOURCES);
                byte[] contentBytes = reader.GetPageContent(page);

                processor.ProcessContent(contentBytes, resources);

                result.Pages.Add(pageInfo);
            }
        }

        return result;
    }

    public PdfDocumentTextInfo Extract(string filePath)
    {
        var result = new PdfDocumentTextInfo
        {
            FilePath = filePath
        };

        using (var reader = new PdfReader(filePath))
        {


            int totalPages = reader.NumberOfPages;

            for (int page = 1; page <= totalPages; page++)
            {
                var bytes = ContentByteUtils.GetContentBytesForPage(reader, page);
                Trace.WriteLine(Encoding.UTF8.GetString(bytes));


                var strategy = new TaggedTextExtractionStrategy();
                var processor = new PdfContentStreamProcessor(strategy);

                // BDC / EMC をフック
                processor.RegisterContentOperator("BDC", new BdcOperator(strategy.TagStack));
                processor.RegisterContentOperator("EMC", new EmcOperator(strategy.TagStack));



                PdfDictionary pageDic = reader.GetPageN(page);
                PdfDictionary resourcesDic = pageDic.GetAsDict(PdfName.RESOURCES);

                processor.ProcessContent(
                    ContentByteUtils.GetContentBytesForPage(reader, page),
                    resourcesDic
                );

                var pageInfo = new PdfPageTextInfo
                {
                    PageNumber = page,
                    Items = strategy.Items
                        .OrderByDescending(i => i.Y)
                        .ThenBy(i => i.X)
                        .ToList()
                };

                result.Pages.Add(pageInfo);
            }
        }

        return result;
    }

}

public class PdfPagingWorker
{
    public void page2in1(string sourceFilePath, int p1, int p2, string outputPath)
    {
        using (var reader = new PdfReader(sourceFilePath))
        {

            // 回転後の実サイズ
            var s1 = reader.GetPageSizeWithRotation(p1);
            var s2 = reader.GetPageSizeWithRotation(p2);

            float w1 = s1.Width;
            float h1 = s1.Height;
            float w2 = s2.Width;
            float h2 = s2.Height;

            // 長辺・短辺
            float long1 = Math.Max(w1, h1);
            float short1 = Math.Min(w1, h1);

            float long2 = Math.Max(w2, h2);
            float short2 = Math.Min(w2, h2);

            // 横向き判定（回転後の実サイズで判定 OK）
            bool isLandscape1 = w1 > h1;
            bool isLandscape2 = w2 > h2;

            // 縦横が違う場合は短辺を揃えるために縮小
            float zoom = 1;
            if (long1 != long2)
                zoom = long1 / long2;

            // 新しいページサイズ（横並び）
            var newRect = new iTextSharp.text.Rectangle(short1 + short2 * zoom, long1);

            using (var fs = new FileStream(outputPath, FileMode.Create))
            using (var doc = new Document(newRect))
            using (var writer = PdfWriter.GetInstance(doc, fs))
            {
                doc.Open();
                var cb = writer.DirectContent;

                var page1 = writer.GetImportedPage(reader, p1);
                var page2 = writer.GetImportedPage(reader, p2);

                // 左ページ
                if (isLandscape1)
                {
                    // 90°回転
                    cb.AddTemplate(page1,
                        0, 1,
                        -1, 0,
                        h1, 0
                    );
                }
                else
                {
                    cb.AddTemplate(page1, 0, 0);
                }

                // 右ページ
                if (isLandscape2)
                {
                    cb.AddTemplate(page2,
                        0, zoom,
                        -zoom, 0,
                        h2 * zoom + short1, 0
                    );
                }
                else
                {
                    cb.AddTemplate(page2, zoom, 0, 0, zoom, short1, 0);
                }

                doc.Close();
            }

        }

    }
    string EncryptSettingPath =System.IO. Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"PDFNET_Encrypt_settings.json");

    public void savePagingPDF(string srcFile, string destFile, List<PageInfo> pages)
    {
        using (var reader = new PdfReader(srcFile))
        using (var ms = new MemoryStream())
        {

            //メモリストリームを作ってからusingする事
            //ページングの結果を新しいPDFファイルに保存
            using (var stamper = new PdfStamper(reader, ms))
            {
                stamper.Close();// これでPDFが ms に入る
            }

            using (var fs = new FileStream(destFile, FileMode.Create))
            using (var editedReader = new PdfReader(ms.ToArray()))
            using (var doc = new Document())
            using (var writer = new PdfCopy(doc, fs))
            {

                // ここで stamper.GetOverContent(pageNumber) を使って
                // スタンプや文字を描画する

                

                doc.Open();



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
                reader.Close();
            }

        }


    }
    public void SaveAsImageOnlyPdf(List<PageRawInfo> pages, string outputPath)
    {
        using (var fs = new FileStream(outputPath, FileMode.Create))
        using (var doc = new Document())
        using (var writer = PdfWriter.GetInstance(doc, fs))
        {
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
    public void EmbedStamp(string inputPdfPath, string outputPdfPath, string dsXmlPath,string EncryptNo)
    {

        using (var reader = new PdfReader(inputPdfPath))
        using (PdfStamper stamper = new PdfStamper(reader, new FileStream(outputPdfPath, FileMode.Create)))
        {
            if (EncryptNo.Length > 0 &&            File.Exists(EncryptSettingPath))
                {
                    byte[] encrypted = File.ReadAllBytes(EncryptSettingPath);

                    byte[] decrypted = ProtectedData.Unprotect(
                        encrypted,
                        null,
                        DataProtectionScope.CurrentUser
                    );

                   var json = Encoding.UTF8.GetString(decrypted);
                
                var settingsList = JsonConvert.DeserializeObject<List<PdfEncryptionSettings>>(json)
                               ?? new List<PdfEncryptionSettings>();

                var settings = settingsList[int.Parse(EncryptNo)];

                stamper.SetEncryption(
                  Encoding.UTF8.GetBytes(settings.UserPassword),
                  Encoding.UTF8.GetBytes(settings.OwnerPassword),
                  settings.Permissions,
                  settings.EncryptionType
              );
            }
            DataSet dataSet1 = new DataSet();
            dataSet1.ReadXml(dsXmlPath);

            DataTable stampPlacementTable = dataSet1.Tables["DataTableStampPlacement"];
            DataTable stampTable = dataSet1.Tables["DataTableStamp"];
            DataTable DataTableTextPlacement = dataSet1.Tables["DataTableTextPlacement"];
            DataTable DataTableRectPlacement = dataSet1.Tables["DataTableRectPlacement"];

            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
            {
                //ほぼおまじない ai のアドバイス。
                //注釈を５～６個以上追加して保存を繰り返すとworkerが落ちてしまう。
                //これを AddAnnotation を呼ぶ前に必ず実行します。
                //これだけで iTextSharp が内部で Annots 辞書を作る必要がなくなり、リークが激減します。

                PdfDictionary pageDict = reader.GetPageN(pageNumber);
                PdfArray annots = pageDict.GetAsArray(PdfName.ANNOTS);
                if (annots == null)
                {
                    annots = new PdfArray();
                    pageDict.Put(PdfName.ANNOTS, annots);
                }


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
                        //Console.WriteLine($"PDF PageSize: {pageWidth} , {pageHeight}");
                        
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


                       // Console.WriteLine($"PDF Location : ({left_Rate} , {top_Rate}");

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
                    string FieldName = "";
                    try
                    {
                        FieldName = row["FieldName"] as string ?? "";
                    }
                    catch { }
                    //int Width = Convert.ToInt16(row["Width"]);//Field用
                    //int Height = Convert.ToInt16(row["Height"]);//Field用

                    System.Drawing.Color ColorARGB = System.Drawing.Color.FromArgb((int)lColor);


                    PdfContentByte cb = stamper.GetOverContent(pageNumber);

                    // フォント準備（埋め込み推奨）
                    //FontFactory.RegisterDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), true);

                    iTextSharp.text.Font font = FontFactory.GetFont(
                        Font,
                        BaseFont.IDENTITY_H,
                        BaseFont.NOT_EMBEDDED,
                        (float)Size,
                        Style,
                           new iTextSharp.text.BaseColor((int)lColor)//itext4=color,itext5=basecolor
                    );

                    string fontPath = @"C:\Windows\Fonts\msgothic.ttc,0";

                    BaseFont bf = font.BaseFont;// nullだった
                                                // bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    cb.SetFontAndSize(bf, Size);

                    // 色設定
                    // System.Drawing.Color → iTextSharp.text.Color (RGB)
                    System.Drawing.Color gdi = ColorARGB; // your ARGB
                    var rgb = new iTextSharp.text.BaseColor(gdi.R, gdi.G, gdi.B); // alpha is not used here
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
                    if (!string.IsNullOrEmpty(FieldName))
                    {
                        //BMCが出る。タグが取れない
                        //                        cb.BeginMarkedContentSequence(new PdfName(FieldName)); 
                        cb.InternalBuffer.Append(Encoding.ASCII.GetBytes($"/{FieldName} <</MCID 0>> BDC\n"));

                    }

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

                    //Console.WriteLine($"lineHeight*1.2 : {leading}");
                    // WPF の LineHeight を取得して比率を計算したいがGDI+で妥協
                    double lineHeightWpf = GetLineHeight_GDI(Font, Size);
                    leading = (float)lineHeightWpf; // PDF 側にそのまま適用
                    //Console.WriteLine($"WPF LineHeight : {leading}");

                    //wpfではfontsize*1.2を行送りにした
                    leading = Size * 1.2f * 0.96f;
                    //実測による補正、２０文字で１文字くらい余計に出力

                    _textHeight = leading * lines.Length;
                    //Console.WriteLine($"Asent/Deceent Point {_textHeight} : Width Point {_textWidth}");

                    ////ColumnTextで矩形を計る方法
                    ////高さを正確に出せる
                    //int maxW = 50000;
                    //int maxH = 5000;
                    //ColumnText ct = new ColumnText(null);
                    //ct.SetSimpleColumn(new Phrase(text, font), 0, 0, maxW, maxH, leading, Element.ALIGN_LEFT);
                    //int status = ct.Go(true); // true = simulate (描画せず計算だけ)
                    //_textHeight = maxH - ct.YLine; // 実際に消費した高さ

                    float maxWidth = 0;
                    foreach (string line in lines)
                    {
                        float w = bf.GetWidthPoint(line, (float)Size);
                        if (w > maxWidth) maxWidth = w;
                    }
                    _textWidth = maxWidth;
                    //Console.WriteLine($"ColumnText scale {_textHeight} : Width Point(SingleLine) {_textWidth} ");



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

                        if(totalDeg==0)
                        {
                            cb.SetTextMatrix(lineX, lineY);
                            cb.ShowText(line);

                            //負荷が高いらしい
                            //角度変更なしならテキストを単純出力
                            //cb.ShowTextAligned(
                            //    Element.ALIGN_LEFT,
                            //    line,
                            //    lineX,
                            //    lineY,
                            //    -totalDeg // 角度はそのまま
                            //);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(FieldName)) cb.BeginMarkedContentSequence(new PdfName(FieldName));

                            cb.ShowTextAligned(
                                Element.ALIGN_LEFT,
                                line,
                                lineX,
                                lineY,
                                -totalDeg // 角度はそのまま
                            );
                            if (!string.IsNullOrEmpty(FieldName)) cb.EndMarkedContentSequence();

                        }
                    }


                    cb.EndText();
                    if (!string.IsNullOrEmpty(FieldName))
                    {
                        //BMCのクローズになる
                        //cb.EndMarkedContentSequence();
                        cb.InternalBuffer.Append(Encoding.ASCII.GetBytes("EMC\n"));
                    }



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
                    var rgb = new iTextSharp.text.BaseColor(gdi.R, gdi.G, gdi.B); // alpha is not used here

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

            }

            stamper.Close();
            reader.Close();
        }
            

    }
    
    private System.Drawing.SizeF WordsAreaSizeOnPDF(  string Font, string text, float Size, int Style,  PdfContentByte cb, bool isMultiLine)
    {

        iTextSharp.text.Font font = FontFactory.GetFont(
            Font,
            BaseFont.IDENTITY_H,
            BaseFont.NOT_EMBEDDED,
            Size,
            Style,
            iTextSharp.text.BaseColor.BLACK
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
            //Console.WriteLine($"width chars : {widthchars}  widthPoint : {widthPoint}");

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
    public WorkerResult MergeIntoTempPdf(string droppedFile,string password)
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
                newPdf =System.IO. Path.GetTempPath() + System.IO.Path.GetRandomFileName() + ".pdf";
                cls.bitmap_to_pdf(new string[] { droppedFile }, newPdf);
            }
            else
            {
                if (droppedFile.ToLower().EndsWith(".tif")
                    || droppedFile.ToLower().EndsWith(".tif")
                    )
                {
                    clsPDF cls = new clsPDF();
                    string tmpTIF = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName());
                    Directory.CreateDirectory(tmpTIF);
                    cls.ConvertMultipageTifToBmpKeepDpi(droppedFile, tmpTIF);
                    string[] files = Directory.GetFiles(tmpTIF, "*.bmp");
                    newPdf = System.IO.Path.GetTempPath() + System.IO.Path.GetRandomFileName() + ".pdf";
                    cls.bitmap_to_pdf(files, newPdf);

                }
            }
        }

        if (newPdf == decryptFilePath( Config.tempPdfPath) || newPdf==Config.tempPdfPath)
        {
            //pdffnetが開くファイルはあらかじめtempPdfPathにコピーされているため、
            //スキップ
            return new WorkerResult
            {
                Success = true,
                Message = "tempPdf = droppedFile",
                
            };
        }
        //ファイルが片方だけ残っていると誤動作
        if (!File.Exists(Config.tempPdfPath) && !File.Exists(decryptFilePath( Config.tempPdfPath)))
        {
            //ファーストドロップのときはここになる
            File.Copy(newPdf, Config.tempPdfPath);
            try
            {
                var reader2 = new PdfReader(Config.tempPdfPath, Encoding.UTF8.GetBytes(password));

                //var reader1 = new PdfReader(decryptFilePath(Config.tempPdfPath));
                using (var fs = new FileStream(decryptFilePath(Config.tempPdfPath), FileMode.Create))
                using (var doc = new Document())
                using (var copy = new PdfCopy(doc, fs))
                {
                    doc.Open();
                    //for (int i = 1; i <= reader1.NumberOfPages; i++)
                    //    copy.AddPage(copy.GetImportedPage(reader1, i));
                    for (int i = 1; i <= reader2.NumberOfPages; i++)
                        copy.AddPage(copy.GetImportedPage(reader2, i));

                    copy.Close();      // 明示的に PdfCopy を閉じる
                    doc.Close();       // Document を閉じる（PdfCopy.Close() より後）
                    fs.Close();        // 最後に FileStream を閉じる

                }

                //reader1.Close();
                reader2.Close();
                return new WorkerResult
                {
                    Success = true,
                    Message = "FirstDrop",

                };

            }
            catch (Exception ex)
            {

                //File.Delete(Config.tempPdfPath);
                //"Bad user password"
                if (ex.Message.Contains("password") ||
    ex.Message.Contains("rangecheck"))
                {
                    return new WorkerResult
                    {
                        Success = false,
                        Message = $"PasswordProtected\n{ex.ToString()}",

                    };
                }


                return new WorkerResult
                {
                    Success = false,
                    Message = $"CantOpen\n{ex.ToString()}",

                };

            }
            finally
            {

            }


        }
        try
        {
            using (var reader2 = new PdfReader(newPdf, Encoding.UTF8.GetBytes(password)))
            using (var reader1 = new PdfReader(decryptFilePath(Config.tempPdfPath)))

            using (var fs = new FileStream(decryptFilePath(Config.tempPdfPath) + ".tmp", FileMode.Create))
            using (var doc = new Document())
            using (var copy = new PdfCopy(doc, fs))
            {
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

            }



            File.Delete(decryptFilePath(Config.tempPdfPath));
            File.Move(decryptFilePath(Config.tempPdfPath) + ".tmp", decryptFilePath(Config.tempPdfPath));
            return new WorkerResult
            {
                Success = true,
                Message = "Merge",

            };

        }
        catch (Exception ex)
        {
            //"Bad user password"
            if (ex.Message.Contains("password") ||
ex.Message.Contains("rangecheck"))
            {
                return new WorkerResult
        {
                    Success = false,
                    Message = $"PasswordProtected\n{ex.ToString()}",

                };
            }


            return new WorkerResult
            {
                Success = false,
                Message = $"CantOpen\n{ex.ToString()}",

            };

        }
        finally
        {

        }


    }
    private string decryptFilePath(string targetFilePath)
    {
        return System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(targetFilePath),
            System.IO.Path.GetFileNameWithoutExtension(targetFilePath) + "_d" + System.IO.Path.GetExtension(targetFilePath)
            );
    }
    public WorkerResult GetPageRotation(string pdfFilePath, int pageNumber)
    {
        using (var reader = new PdfReader(pdfFilePath))
        {
            PdfDictionary pageDict = reader.GetPageN(pageNumber);
            PdfNumber rotate = pageDict.GetAsNumber(PdfName.ROTATE);
            reader.Close();

            var pginfo = new PageInfo()
            {
                PageNumber = pageNumber,
                Rotation = rotate != null ? rotate.IntValue : 0
            }
            ;
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
    }

    public WorkerResult MakeBrankPage(string papersize)
    {
        iTextSharp.text. Rectangle pageRect = PageSizeResolver.GetPageSize(papersize);

        int addedPageNumber = 1;
        using (var fs = new FileStream(decryptFilePath(Config.tempPdfPath) + ".tmp", FileMode.Create))
        using (var doc = new Document())
        using (var copy = new PdfCopy(doc, fs))

        {
            doc.Open();
            if (File.Exists(decryptFilePath(Config.tempPdfPath)))
            {
                using (var reader = new PdfReader(decryptFilePath(Config.tempPdfPath)))
                {
                    // 既存ページをコピー
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        copy.AddPage(copy.GetImportedPage(reader, i));
                    }
                    addedPageNumber = reader.NumberOfPages + 1;

                }

            }
            // 空白ページを追加
            var blankDoc = new Document(pageRect);
            var blankStream = new MemoryStream();
            var blankWriter = PdfWriter.GetInstance(blankDoc, blankStream);
            blankDoc.Open();
            blankDoc.NewPage();

            // 空白ページに白い矩形を描画（これでページが有効になる）
            var cb = blankWriter.DirectContent;
            cb.Rectangle(0, 0, pageRect.Width, pageRect.Height);
            cb.SetColorFill(iTextSharp.text.BaseColor.WHITE);
            cb.Fill();

            blankDoc.Close();

            var blankReader = new PdfReader(blankStream.ToArray());
            var blankImportedPage = copy.GetImportedPage(blankReader, 1); // ← 修正
            copy.AddPage(blankImportedPage); // ← 修正

            doc.Close();
        }
        
        var blankPage = new PageInfo
        {
            WidthPt = pageRect.Width,
            HeightPt = pageRect.Height,
            PageNumber = addedPageNumber,
            OriginalPageNumber = addedPageNumber,
        };


        File.Delete(decryptFilePath(Config.tempPdfPath));
        File.Move(decryptFilePath(Config.tempPdfPath) + ".tmp", decryptFilePath(Config.tempPdfPath));

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
    public List<PageRawInfo> LoadPdfPagesRaw(string pdfFilePath, string password)
    {
//        string decryptedPath = Path.Combine(Path.GetDirectoryName(Config.tempPdfPath), Path.GetFileNameWithoutExtension(Config.tempPdfPath) + "_d" + ".pdf");
        List<PageRawInfo> loadPagesRaw = new();

        //decrypt済みファイルであれば作成しない
        if(pdfFilePath != decryptFilePath(Config. tempPdfPath))
        {
            //パスワード判定
            try
            {
                var reader = new PdfReader(pdfFilePath, Encoding.UTF8.GetBytes(password));
                var stamper = new PdfStamper(reader, new FileStream(decryptFilePath(pdfFilePath), FileMode.Create));

                // 読めた → パスワードなし
                reader.Close();
                stamper.Close();
            }
            catch (BadPasswordException)
            {
                // パスワード付き PDF
                return new List<PageRawInfo>
    {
        new PageRawInfo { Error = "PasswordProtected" }
    };
            }


        }


        try
        {
            //パスワード空でここへ来ている＝すでにdecryptFileが渡されている

            string rasterizePath = (password == "") ? pdfFilePath : decryptFilePath(pdfFilePath);
            

            string exeDir = System.IO.Path.GetDirectoryName(
                Process.GetCurrentProcess().MainModule.FileName
            );

            var version = new GhostscriptVersionInfo(
                System.IO.Path.Combine(exeDir, "gsdll64.dll")
            );

            string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetFileName(pdfFilePath));
            if (pdfFilePath != tempPath) File.Copy(pdfFilePath, tempPath, true);

            using var rasterizer = new GhostscriptRasterizer();
            
            rasterizer.Open(rasterizePath, version, false);

            int dpi = rasterizer.PageCount >= 50 ? 96 : 150;
            dpi = 150;

            if (rasterizer.PageCount == 0)
            {
                return loadPagesRaw;
            }
            using (var reader = new PdfReader(pdfFilePath))
            {
                var pageSize_P1 = reader.GetPageSize(1);
                float widthPt_P1 = pageSize_P1.Width;
                float heightPt_P1 = pageSize_P1.Height;


                for (int pageNumber = 1; pageNumber <= rasterizer.PageCount; pageNumber++)
                {
                    // ページごとのサイズを取得
                    var pageSize = reader.GetPageSize(pageNumber);
                    float widthPt = pageSize.Width;
                    float heightPt = pageSize.Height;

                    int widthPx = (int)(pageSize.Width * dpi / 72.0);
                    int heightPx = (int)(pageSize.Height * dpi / 72.0);

                    //全ページが1ページ目のページサイズで描かれるので補正が必要
                    var sk = rasterizer.GetPage(dpi, pageNumber);
                    // sk = GhostscriptRasterizer.GetPage(dpi, pageNumber);

                    int targetWidthPx = (int)(widthPt * dpi / 72.0);
                    int targetHeightPx = (int)(heightPt * dpi / 72.0);

                    // 新しいビットマップを作成
                    var resized = new SkiaSharp.SKBitmap(targetWidthPx, targetHeightPx);

                    // リサイズ処理
                    sk.ScalePixels(resized, SkiaSharp.SKFilterQuality.High);

                    // PNG 化
                    using var image = SkiaSharp.SKImage.FromBitmap(resized);
                    using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                    byte[] pngBytes = data.ToArray();



                    //                using var image = SkiaSharp.SKImage.FromBitmap(sk);
                    //                using var data = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                    //                byte[] pngBytes = data.ToArray();

                    //               double widthPt = sk.Width * 72.0 / dpi;
                    //               double heightPt = sk.Height * 72.0 / dpi;

                    int rotation = 0;
                    var workerResult = GetPageRotation(rasterizePath, pageNumber);
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
                reader.Close();

            }

        }
        catch (Exception ex)
        {
                       throw;
        }


        return loadPagesRaw;
    }
}

//public class TaggedTextExtractionStrategy : IExtRenderListener
//{
//    private Stack<string> tagStack = new Stack<string>();
//    public List<PdfTextItem> Items { get; } = new List<PdfTextItem>();

//    public void BeginMarkedContentSequence(PdfName tag, PdfDictionary dict)
//    {
//        tagStack.Push(tag.ToString().TrimStart('/'));
//    }

//    public void EndMarkedContentSequence()
//    {
//        if (tagStack.Count > 0)
//            tagStack.Pop();
//    }

//    public void RenderText(TextRenderInfo renderInfo)
//    {
//        var baseline = renderInfo.GetBaseline().GetStartPoint();

//        Items.Add(new PdfTextItem
//        {
//            Text = renderInfo.GetText(),
//            X = baseline[0],
//            Y = baseline[1],
//            Angle = 0,
//            Tag = tagStack.Count > 0 ? tagStack.Peek() : null
//        });
//    }

//    public void BeginTextBlock() { }
//    public void EndTextBlock() { }
//    public void RenderImage(ImageRenderInfo renderInfo) { }
//    public void ModifyPath(PathConstructionRenderInfo renderInfo) { }
//    public void ClipPath(int rule) { }
//    public iTextSharp.text.pdf.parser. Path RenderPath(PathPaintingRenderInfo renderInfo) { return null; }
//}

public class BdcOperator : IContentOperator
{
    private readonly Stack<string> tagStack;

    public BdcOperator(Stack<string> tagStack)
    {
        this.tagStack = tagStack;
    }

    public void Invoke(PdfContentStreamProcessor processor, PdfLiteral oper, List<PdfObject> operands)
    {
        // operands[0] = /namae
        var tag = operands[0] as PdfName;
        if (tag != null)
        {
            tagStack.Push(tag.ToString().TrimStart('/'));
        }
    }
}
public class EmcOperator : IContentOperator
{
    private readonly Stack<string> tagStack;

    public EmcOperator(Stack<string> tagStack)
    {
        this.tagStack = tagStack;
    }

    public void Invoke(PdfContentStreamProcessor processor, PdfLiteral oper, List<PdfObject> operands)
    {
        if (tagStack.Count > 0)
            tagStack.Pop();
    }
}
public class TaggedTextExtractionStrategy : IRenderListener
{
    public Stack<string> TagStack { get; } = new Stack<string>();
    public List<PdfTextItem> Items { get; } = new List<PdfTextItem>();

    public void RenderText(TextRenderInfo renderInfo)
    {
        var baseline = renderInfo.GetBaseline().GetStartPoint();

        Items.Add(new PdfTextItem
        {
            Text = renderInfo.GetText(),
            X = baseline[0],
            Y = baseline[1],
            Angle = 0,
            Tag = TagStack.Count > 0 ? TagStack.Peek() : null
        });
    }

    public void BeginTextBlock() { }
    public void EndTextBlock() { }
    public void RenderImage(ImageRenderInfo renderInfo) { }
}
