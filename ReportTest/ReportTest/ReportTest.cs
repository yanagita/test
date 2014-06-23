using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;

//iTextSharp関連の名前空間
using iTextSharp.text;
using iTextSharp.text.pdf;

//iTextSharp.text.FontクラスがSystem.Drawing.Fontクラスと
//混在するためiFontという別名を設定
using iFont = iTextSharp.text.Font;
//test- using sFont = System.Drawing.Font;

//ファイルIO関連の名前空間
using System.IO;
using System.Diagnostics;

namespace ReportTest
{
    public partial class ReportTest : Form
    {
        public ReportTest()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //
            //（注意）pdfを開いたままだとエラーになる。
            //

            //ドキュメントを作成
            Document doc = new Document();

            //ファイルの出力先を設定
            PdfWriter.GetInstance(doc, new FileStream("01_ハロー.pdf", FileMode.Create));

            //ドキュメントを開く
            doc.Open();

            //「Hello iTextSharp」をドキュメントに追加
            doc.Add(new Paragraph("Hello iTextSharp"));

            //ドキュメントを閉じる
            doc.Close();

            //MessageBox.Show("完了", "iTextSharp");

            //pdfファイルを開く
            System.Diagnostics.Process.Start("01_ハロー.pdf");
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //ドキュメントを作成
            Document doc = new Document();
            try
            {
                //ファイル出力先を設定
                PdfWriter.GetInstance(doc, new FileStream("02_JP.pdf", FileMode.Create));
                //ドキュメントを開く
                doc.Open();

                //［1］ MSゴシック
                iFont fnt1 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msgothic.ttc,0", BaseFont.IDENTITY_H, true), 40);

                //［2］ MS Pゴシック-太字
                iFont fnt2 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true),
                    32, iTextSharp.text.Font.BOLD);

                //［3］ MS UI Gothic-斜体-下線
                iFont fnt3 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msgothic.ttc,2", BaseFont.IDENTITY_H, true),
                    20, iTextSharp.text.Font.ITALIC | iTextSharp.text.Font.UNDERLINE);

                //// ※CKJフォントを使う場合、事前にTextAsian.dllをロード（Form_Load参照）
                ////［4］ CJK明朝
                //iFont fnt4 = new iFont(BaseFont.CreateFont
                //    ("HeiseiMin-W3", "UniJIS-UCS2-HW-H", false), 20);

                //フォント(メイリオ)
                iFont fnt10 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\meiryo.ttc,0", BaseFont.IDENTITY_H, true), 20);
                //フォント(メイリオ)
                iFont fnt11 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\meiryo.ttc,0", BaseFont.IDENTITY_H, true), 20, iTextSharp.text.Font.BOLD);
                //フォント(メイリオ)
                iFont fnt12 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\meiryo.ttc,0", BaseFont.IDENTITY_H, true), 20, iTextSharp.text.Font.ITALIC);
                //フォント(メイリオ)
                iFont fnt13 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\meiryo.ttc,0", BaseFont.IDENTITY_H, true), 20, iTextSharp.text.Font.BOLD | iTextSharp.text.Font.ITALIC);

                //フォント(MS明朝)
                iFont fnt14 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msmincho.ttc,0", BaseFont.IDENTITY_H, true), 20);
                //フォント(MS明朝-太字)
                iFont fnt15 = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msmincho.ttc,0", BaseFont.IDENTITY_H, true), 20, iTextSharp.text.Font.BOLD);

                ////［5］ CJKゴシック-赤色
                //iFont fnt5 = new iFont(BaseFont.CreateFont
                //    ("HeiseiKakuGo-W5", "UniJIS-UCS2-HW-H", false), 20);
                //fnt5.SetColor(255, 0, 0);

                //文言とフォントを指定してドキュメントに追加
                doc.Add(new Paragraph("MSゴシックです", fnt1));
                doc.Add(new Paragraph("MS Pゴシックの太字です", fnt2));
                doc.Add(new Paragraph("MS UI Gothicの斜体／下線です", fnt3));
                //doc.Add(new Paragraph("HeiseiMin-W3（明朝）です", fnt4));
                //doc.Add(new Paragraph("HeiseiKakuGo-W5（ゴシック）の赤色です", fnt5));
                doc.Add(new Paragraph("メイリオ　標準", fnt10));
                doc.Add(new Paragraph("メイリオ　太字", fnt11));
                doc.Add(new Paragraph("メイリオ　イタリック", fnt12));
                doc.Add(new Paragraph("メイリオ　太字＋イタリック", fnt13));

                doc.Add(new Paragraph("MS明朝", fnt14));
                doc.Add(new Paragraph("MS明朝　太字", fnt15));
            }
            catch (DocumentException ex)
            {
                Debug.WriteLine("---\r\n" + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show(ex.Message, "DocumentException");
            }
            catch (IOException ex)
            {
                Debug.WriteLine("---\r\n" + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show(ex.Message, "IOException");
            }
            finally
            {
                //ドキュメントを閉じる
                doc.Close();
            }
            //MessageBox.Show("完了", "iTextSharp");
            
            //pdfファイルを開く
            System.Diagnostics.Process.Start("02_JP.pdf");
        }

        private void ReportTest_Load(object sender, EventArgs e)
        {
            ////標準CJKフォント用のリソースをロード
            //if (File.Exists("iTextAsian-1.0.dll"))
            //{
            //    BaseFont.AddToResourceSearch("iTextAsian-1.0.dll");
            //}
            //else
            //{
            //    MessageBox.Show("iTextAsian-1.0.dllが存在しません", "エラー");
            //}
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //PDFドキュメント(ページサイズ,マージン左,右,上,下)
            //Document doc = new Document(PageSize.A4, 50, 200, 50, 500);
            //Document doc = new Document(PageSize.A4.Rotate(), 50, 200, 50, 500); //A4-横
            //Document doc = new Document(PageSize.A5, 50, 50, 50, 50);
            Document doc = new Document(PageSize.A5.Rotate(), 50, 50, 50, 50); //A5-横
            
            try
            {
                //ファイル出力用のストリームを取得
                PdfWriter.GetInstance(doc, new FileStream("03_Layout.pdf", FileMode.Create));

                //フォント(MS Pゴシック)
                iFont fnt = new iFont(BaseFont.CreateFont
                    (@"c:\windows\fonts\msgothic.ttc,1", BaseFont.IDENTITY_H, true), 14);
                //fnt.SetColor(255, 0, 0); //赤字

                //※ヘッダーとフッターはDomumentのOpen前に指定します

                ////ヘッダーの設定をします。
                //HeaderFooter header = new HeaderFooter(new Phrase("ヘッダー（中央揃え）", fnt), false);
                ////test- Header header = new Header(new Phrase("ヘッダー（中央揃え）", fnt), false);
                ////センター寄せ
                //header.SetAlignment(ElementTags.ALIGN_CENTER);
                ////DocumentにHeaderを設定
                //doc.Header = header;

                ////フッターの設定をします。
                //HeaderFooter footer = new HeaderFooter(new Phrase("ここは", fnt), new Phrase("ページ目", fnt));
                ////センター寄せ
                //footer.SetAlignment(ElementTags.ALIGN_CENTER);
                ////上下の線を消す
                //footer.Border = Rectangle.NO_BORDER;
                ////DocumentにFooterを設定
                //doc.Footer = footer;

                //文章の出力を開始します。
                doc.Open();

                //本文１
                doc.Add(new Paragraph("左寄せ（標準）", fnt));
                //本文２
                Paragraph para = new Paragraph("右寄せ", fnt);
                para.Alignment = Element.ALIGN_RIGHT;
                doc.Add(para);

                //改ページ確認用
                for (int i = 1; i <= 5; i++)
                {
                    doc.Add(new Paragraph(i.ToString() + "行目", fnt));
                }



                //2列からなるテーブルを作成
                //Table tbl = new Table(2);
                PdfPTable tbl = new PdfPTable(2);
                //テーブル全体の幅（パーセンテージ）
                //tbl.Width = 100;
                tbl.TotalWidth = 400f;
                //tbl.TotalWidth = doc.Right - doc.Left;

                //テーブル各列の幅
                //tbl.Widths = new float[] { 0.75f, 0.25f };
                tbl.SetTotalWidth(new float[] { 300f, 100f });

                //float[] widths = new float[] { 2f, 4f, 6f };
                //tbl.SetWidths(widths);


                //テーブルのデフォルトの表示位置（横）
                //tbl.DefaultHorizontalAlignment = Element.ALIGN_CENTER;

                //テーブルのデフォルトの表示位置（縦）
                //tbl.DefaultVerticalAlignment = Element.ALIGN_MIDDLE;

                // 左・上・右・下の各辺のパディング量を 5 にする
                tbl.DefaultCell.Padding = 10;
                // 左パディングだけを 3 にする
                tbl.DefaultCell.PaddingLeft = 5;

                //tbl.HorizontalAlignment = 0;

                //テーブルの余白
                //tbl.Padding = 3;
                //テーブルのセル間の間隔
                //tbl.Spacing = 0;
                //テーブルの線の色（RGB:黒）
                //tbl.BorderColor = new iTextSharp.text.Color(0, 0, 0);

                //leave a gap before and after the table
                tbl.SpacingBefore = 20f;
                tbl.SpacingAfter = 30f;

                //タイトルのセルを追加（左の列）
                //Cell cel = new Cell(new Phrase("支払理由", fnt));
                PdfPCell cel = new PdfPCell(new Phrase("支払理由", fnt));

                //セルの網掛け(20%網掛け)
                cel.GrayFill = 0.8f;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                tbl.AddCell(cel);

                //タイトルのセルを追加（右の列）
                //cel = new Cell(new Phrase("金額", fnt));
                cel = new PdfPCell(new Phrase("金額", fnt));
                //セルの網掛け(40%網掛け)
                cel.GrayFill = 0.6f;
                cel.HorizontalAlignment = Element.ALIGN_CENTER;
                tbl.AddCell(cel);

                //表示するデータ
                ArrayList list = new ArrayList();
                list.Add(new string[] { "光熱費", "10,000" });
                list.Add(new string[] { "書籍代", "7,500" });
                list.Add(new string[] { "インターネットプロバイダ代", "5,000" });

                //明細行の追加
                foreach (string[] shiharai in list)
                {
                    //左のセルの追加
                    //cel = new Cell(new Phrase(shiharai[0], fnt));
                    cel = new PdfPCell(new Phrase(shiharai[0], fnt));
                    tbl.AddCell(cel);

                    //右のセルの追加（金額なので右寄せ）
                    //cel = new Cell(new Phrase(shiharai[1], fnt));
                    cel = new PdfPCell(new Phrase(shiharai[1], fnt));
                    cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                    tbl.AddCell(cel);
                }

                //結合したセルの追加
                //cel = new Cell(new Phrase("今月の支出は22,500円", fnt));
                cel = new PdfPCell(new Phrase("今月の支出は22,500円", fnt));
                cel.Colspan = 2;
                cel.HorizontalAlignment = Element.ALIGN_RIGHT;
                tbl.AddCell(cel);

                //テーブルを追加
                doc.Add(tbl);




                // テーブル全体をページの左寄せにする
                tbl.HorizontalAlignment = Element.ALIGN_RIGHT;

                // テーブルの幅をページ幅の 80％ にする
                tbl.WidthPercentage = 80;

                //テーブルを追加
                doc.Add(tbl);



                // テーブル全体をページの左寄せにする
                tbl.HorizontalAlignment = Element.ALIGN_RIGHT;

                // テーブルの幅をページ幅の 80％ にする
                tbl.WidthPercentage = 100;

                //テーブルを追加
                doc.Add(tbl);

                //Paragraph para_t = new Paragraph();
                //para_t.Add(tbl);
                ////para_t.Alignment = Element.ALIGN_CENTER;
                //para_t.Alignment = Element.ALIGN_RIGHT;
                //doc.Add(para_t);

            }
            catch (Exception ex)
            {
                Debug.WriteLine("---\r\n" + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show(ex.Message, "エラー");
            }
            finally
            {
                //ドキュメントを閉じる
                doc.Close();
            }
            
            //MessageBox.Show("完了", "iTextSharp");

            //pdfファイルを開く
            System.Diagnostics.Process.Start("03_Layout.pdf");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            FontFactory.RegisterDirectories();
            iTextSharp.text.Font font = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 12, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontRed = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 12, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.RED);
            iTextSharp.text.Font fontGreen = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 12, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.GREEN);
            iTextSharp.text.Font fontBlue = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 12, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLUE);
            iTextSharp.text.Font fontPink = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 12, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.PINK);

            string sFilename = "Chunk.pdf";
            Document doc = new Document();
            PdfWriter.GetInstance(doc, new FileStream(sFilename, FileMode.Create));
            doc.Open();

            Paragraph p = new Paragraph();

            Chunk c = new Chunk("あ", font);
            p.Add(c);

            c = new Chunk("い", fontRed);
            p.Add(c);

            c = new Chunk("う", fontGreen);
            p.Add(c);

            c = new Chunk("え", fontBlue);
            p.Add(c);

            c = new Chunk("お", fontPink);
            p.Add(c);

            doc.Add(p);
            doc.Close();

            //pdfファイルを開く
            System.Diagnostics.Process.Start(sFilename);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FontFactory.RegisterDirectories();
            iTextSharp.text.Font fontNormal = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 10, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = FontFactory.GetFont("MS-Mincho", BaseFont.IDENTITY_H,
                  BaseFont.NOT_EMBEDDED, 10, iTextSharp.text.Font.BOLD);

            Document doc = new Document(PageSize.A4);

            string sFilename = "TableTest.pdf";
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(sFilename, FileMode.Create));

            // ドキュメントを開く
            doc.Open();

            float[] widths = new float[] { 2, 3 }; //x:yでOK（最後で幅を%指定する時は、）

            // 列数 2 のテーブルを作成する
            PdfPTable table = new PdfPTable(widths);

            // 左・上・右・下の各辺のパディング量を 8 にする
            table.DefaultCell.Padding = 5;

            // 左パディングだけを 4 にする
            table.DefaultCell.PaddingLeft = 3;

            PdfPCell cell;
            cell = new PdfPCell(new Phrase("ヘッダー１", fontBold));
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.GrayFill = 0.8f;
            cell.Padding = 5;
            table.AddCell(cell);

            cell = new PdfPCell(new Phrase("ヘッダー２", fontBold));
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.BackgroundColor = iTextSharp.text.BaseColor.YELLOW;
            cell.Padding = 5;
            table.AddCell(cell);

            // セルを追加する
            table.AddCell(new Phrase("あいうえお", fontNormal));
            table.AddCell(new Phrase("かきくけこ", fontNormal));
            table.AddCell(new Phrase("あいうえおかきくけこさしすせそ", fontNormal));
            table.AddCell(new Phrase("さしすせそ", fontNormal));

            // テーブル全体をページの左寄せにする
            //table.HorizontalAlignment = Element.ALIGN_LEFT;
            table.HorizontalAlignment = Element.ALIGN_CENTER;

            // テーブルの幅をページ幅の 60％ にする
            table.WidthPercentage = 80;

            // テーブルをドキュメントに追加する
            doc.Add(table);

            // ドキュメントを閉じる
            doc.Close();

            //pdfファイルを開く
            System.Diagnostics.Process.Start(sFilename);
        }
    
    }

    ////ヘッダー・フッター設定
    //private class HeaderFooterPage{
    //    inherits PdfPageEventHelper;

    //    private override void OnEndPage (PdfWriter writer, Document document){
    //        MyBase.OnEndPage(writer, document);

    //        // 初期化
    //        PdfContentByte cb;
    //        cb = writer.DirectContent;

    //        string pageNo = writer.PageNumber.ToString();
    //        iTextSharp.text.Rectangle pageSize = document.PageSize;

    //        //ヘッダー出力
    //        cb.BeginText();

    //        cb.SetFontAndSize(HeaderFooterBaseFont, 8);
    //        cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetTop(30));
    //        cb.ShowTextAligned(Element.ALIGN_RIGHT, "ヘッダー", pageSize.Width - 20, pageSize.GetTop(30), 0);

    //        cb.EndText();

    //        // フッター出力
    //        cb.BeginText();

    //        cb.SetFontAndSize(HeaderFooterBaseFont, 8);
    //        cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetBottom(30));
    //        cb.ShowTextAligned(Element.ALIGN_CENTER, pageNo, pageSize.Width / 2, pageSize.GetBottom(30), 0);

    //        cb.EndText();
    //    }
    //}


}
