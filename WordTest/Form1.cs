using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TemplateEngine.Docx;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WordTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Show the dialog that allows user to select a file, the 
            // call will result a value from the DialogResult enum
            // when the dialog is dismissed.
            DialogResult result = this.openFileDialog1.ShowDialog();
            // if a file is selected
            if (result == DialogResult.OK)
            {
                // Set the selected file URL to the textbox
                this.textBox1.Text = this.openFileDialog1.FileName;
                TestDocAndDocTemplatePackage(this.openFileDialog1.FileName);
            }

        }

        private void TestDocAndDocTemplatePackage(string fileName)
        {
            // use docX
            using (DocX document = DocX.Load(fileName))
            {
                var exportDoc = document.Copy();
                var sections = exportDoc.Sections;
                var paragraphs = exportDoc.Paragraphs;
                var tables = exportDoc.Tables;
                //// remove table
                //var factoryMapPart = tables.Where(x => x.Rows[0].Paragraphs[0].Text.Contains("化學品資訊")).FirstOrDefault();
                //factoryMapPart.Remove();

                //// insert table row
                //var custodyTable = tables[4];
                //var headerRows = custodyTable.Rows.Take(3).Select(x=>x).ToList();
                //foreach (var row in headerRows)
                //{
                //    // keepingformat -> 確保格式一樣
                //    custodyTable.InsertRow(row,true);
                //}
                //document.InsertTable(2,3);
                //// 在註解說明前面新增label與table
                //var locationTable = tables[5];
                //var locationLabelParagraph = paragraphs.Where(x => x.Text.Contains("使用／儲存地點")).First();
                var noteParagraph = paragraphs.Where(x => x.Text.Contains("註解說明")).First();
                //var index = paragraphs.IndexOf(noteParagraph);

                #region 將空白圖片替換掉成別的圖片
                var firstHasPicturePara = exportDoc.Paragraphs.Where(x => x.Pictures.Any()).FirstOrDefault();
                ReplacePictureToParagraph(exportDoc, firstHasPicturePara, 
                    Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName + @"\cbimage.png");
                #endregion

                #region 在特定段落連續插入圖片
                var publicItemImgPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName+ @"\公共危險物品.png";
                var hazardIconImgPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName+ @"\危害標示.png";
                var diasterPreventionImgPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName+ @"\防災應變.png";
                var infraImgPath = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName+ @"\設備設施.png";
                var image = System.Drawing.Image.FromFile(infraImgPath);
                var curParagraph = exportDoc.Paragraphs
                .Where(x => x.Text.Contains("Icon圖示說明")).FirstOrDefault();
                // 這邊因為是從圖片insert至段落文字下方
                // 因為這邊是把圖片insert在同一個段落之中，順序不需要倒過來
                var publicItemPicture = AddPictureFittedPageWidth(exportDoc, publicItemImgPath);
                var hazardIconPicture = AddPictureFittedPageWidth(exportDoc, hazardIconImgPath);
                var diasterPreventionPicture = AddPictureFittedPageWidth(exportDoc, diasterPreventionImgPath);
                var infraImgPicture = AddPictureFittedPageWidth(exportDoc, infraImgPath);
                curParagraph.AppendPicture(publicItemPicture);
                curParagraph.AppendPicture(hazardIconPicture);
                curParagraph.AppendPicture(diasterPreventionPicture);
                curParagraph.AppendPicture(infraImgPicture);
                #endregion

               
                #region 在特定段落之間插入文字/table
                var factoryMapPart = paragraphs.Where(x=>x.Text.Contains("化學品資訊")).FirstOrDefault();
                var isLocation = true;
                // 因為table刪除後仍無法調整
                var custodyTable = tables[4];
                var locationTable = tables[5];
                var tableBookmark = exportDoc.Bookmarks["Table"];
                if (tableBookmark != null)
                {
                }

                var tableCol = 15;
                if (isLocation)
                {
                    var locationName = factoryMapPart.InsertParagraphAfterSelf("使用/儲存地點");
                    locationName.Font("Microsoft JhengHei").FontSize(11).Bold(true).Color(Color.Black);
                    var table1 = locationName.InsertTableAfterSelf(locationTable);
                    // 這裡為節省時間不處理row的border，因此先將第一列當作sample row，依照資料列多寡逐筆add進table即可
                    var sampleDataRow = table1.Rows[2];
                    sampleDataRow.Remove();
                    var chemicalNameColorFormat = new Formatting()
                    {
                        FontFamily = new Xceed.Document.NET.Font("Microsoft JhengHei"),
                        Size = 10,
                        FontColor = Color.SkyBlue
                    };
                    var cellFormat = chemicalNameColorFormat;
                    cellFormat.FontColor = Color.SkyBlue;

                    var insertedRow = table1.InsertRow(sampleDataRow);
                    insertedRow.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[0].Paragraphs[0].InsertText("測試用", formatting: chemicalNameColorFormat);
                    insertedRow.Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[1].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[2].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[3].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[3].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    //加入hyper link
                    var uri = new Uri("https://docs.microsoft.com/zh-tw/dotnet/api/system.uri.getleftpart?view=net-6.0");
                    var hyperlink = exportDoc.AddHyperlink("test", uri);
                    var a = insertedRow.Cells[3].Paragraphs[0];
                    var packagePart = a.PackagePart;
                    insertedRow.Cells[3].Paragraphs[0].InsertHyperlink(hyperlink);

                    insertedRow = table1.InsertRow(sampleDataRow);
                    insertedRow.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[0].Paragraphs[0].InsertText("測試用", formatting: chemicalNameColorFormat);
                    insertedRow.Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[1].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[2].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[3].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[3].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    var test = insertedRow.Cells[3].Paragraphs[0];
                    //insertedRow.Cells[3].Paragraphs[0].InsertHyperlink(hyperlink,0);


                    // table2
                    var locationName2 = table1.InsertParagraphAfterSelf("使用/儲存地點");
                    locationName2.Font("Microsoft JhengHei").FontSize(11).Bold(true).Color(Color.Black);
                    var table2 = locationName2.InsertTableAfterSelf(locationTable);
                    //  重要!!!!!複製table時第二次以後packagePart不知為何會變null，需要重新assign 
                    //  不然後面InsertHyperlink時會跳null exception
                    table2.PackagePart = packagePart;
                    var sampleDataRow2 = table2.Rows[2];
                    sampleDataRow2.Remove();
                    insertedRow = table2.InsertRow(sampleDataRow2);
                    insertedRow.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[0].Paragraphs[0].InsertText("測試用", formatting: chemicalNameColorFormat);
                    insertedRow.Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[1].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[2].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[3].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    var uri2 = new Uri("https://docs.microsoft.com/zh-tw/dotnet/api/system.uri.getleftpart?view=net-6.0");
                    var hyperlink2 = exportDoc.AddHyperlink("test2", uri2);
                    insertedRow.Cells[3].Paragraphs[0].AppendHyperlink(hyperlink2).Color(Color.BurlyWood).UnderlineStyle(UnderlineStyle.singleLine);

                    //insertedRow.Cells[3].Paragraphs[0].Font(chemicalNameColorFormat.FontFamily.Name)
                    //    .FontSize(chemicalNameColorFormat.Size.Value)
                    //    .Color(chemicalNameColorFormat.FontColor.Value);

                }
                else
                {
                    // 逐筆複製table
                    var chemicalCustodyTable = factoryMapPart.InsertTableAfterSelf(custodyTable);
                    // table1
                    chemicalCustodyTable.Rows[0].Paragraphs[0].ReplaceText("{idkey}", "CCB");
                    var dataRow = custodyTable.Rows[3];


                    var insertedRow = chemicalCustodyTable.InsertRow(dataRow);
                    var chemicalNameColorFormat = new Formatting()
                    {
                        FontFamily = new Xceed.Document.NET.Font("Microsoft JhengHei"),
                        Size = 10,
                        FontColor = Color.SkyBlue
                    };
                    var cellFormat = chemicalNameColorFormat;
                    cellFormat.FontColor = Color.Black;
                    insertedRow.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[0].Paragraphs[0].InsertText("測試用", formatting: chemicalNameColorFormat);
                    insertedRow.Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[1].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[2].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow.Cells[3].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow.Cells[3].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    chemicalCustodyTable.Rows[3].Remove();

                    // 兩個table中間分行breakline
                    var breaklinePargraph = chemicalCustodyTable.InsertParagraphAfterSelf("\r\n");
                    var chemicalCustodyTable2 = breaklinePargraph.InsertTableAfterSelf(custodyTable);

                    // table2
                    chemicalCustodyTable2.Rows[0].Paragraphs[0].ReplaceText("{idkey}", "CCB");


                    var insertedRow2 = chemicalCustodyTable2.InsertRow(dataRow);

                    insertedRow2.Cells[0].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow2.Cells[0].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow2.Cells[0].Paragraphs[0].InsertText("測試用", formatting: chemicalNameColorFormat);
                    insertedRow2.Cells[1].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow2.Cells[1].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow2.Cells[1].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow2.Cells[2].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow2.Cells[2].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow2.Cells[2].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    insertedRow2.Cells[3].Paragraphs[0].Alignment = Alignment.center;
                    insertedRow2.Cells[3].VerticalAlignment = VerticalAlignment.Center;
                    insertedRow2.Cells[3].Paragraphs[0].InsertText("測試用", formatting: cellFormat);
                    chemicalCustodyTable2.Rows[3].Remove();
                }

                // 將sample table刪除
                custodyTable.Remove();
                locationTable.Remove();
                #endregion


                var memoryStream = new MemoryStream();
                exportDoc.SaveAs(memoryStream);
                memoryStream.Seek(0, SeekOrigin.Begin);

                using (var filestream = new FileStream(@"test.docx", FileMode.OpenOrCreate))
                {
                    memoryStream.Seek(0, SeekOrigin.Begin);
                    memoryStream.CopyTo(filestream);
                }
            }
            // use template engine docX
        }

        private static void ReplacePictureToParagraph(Document docx, Paragraph paragraph,string imagePath)
        {
            // get picture to replaced
            var orgPicture = docx.Pictures[0];
            var img = docx.AddImage(imagePath);
            var picture = img.CreatePicture();
            picture.Width = orgPicture.Width;
            picture.Height = orgPicture.Height;
            paragraph.ReplacePicture(orgPicture, picture);
        }

        private static Picture AddPictureFittedPageWidth(Document docx, string imgFilePath)
        {
            var section = docx.Sections[0];
            // 將超過寬度的圖片縮成不會超過頁面的寬度大小
            // 取得寬度扣掉兩邊的margin後才不會超過邊界
            var sectionPageWidth = docx.PageWidth - (section.MarginLeft + section.MarginRight);
            var image = System.Drawing.Image.FromFile(imgFilePath);
            var reshapeHeight = image.Height * (sectionPageWidth / image.Width);
            var img = docx.AddImage(imgFilePath);
            
            var picture = img.CreatePicture();
            picture.Width = sectionPageWidth;
            picture.Height = reshapeHeight;
            return picture;
        }
    }
}
