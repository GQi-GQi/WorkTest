using NewLife.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using Xunit;

namespace testpdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");


            #region pdf
            ////实例化一个PdfDocument类
            //PdfDocument doc = new PdfDocument();

            ////doc.LoadFromFile("http://192.168.3.132:8080/PDF%E8%A1%A8%E6%A0%BC_1.pdf");

            ////PdfPageBase page = doc.Pages[0];
            //PdfUnitConvertor unitCvtr = new PdfUnitConvertor();
            //PdfMargins margin = new PdfMargins();
            //PdfSection section = doc.Sections.Add();
            //section.PageSettings.Margins = margin;
            //section.PageSettings.Width = 113.38f;
            //section.PageSettings.Height = 56.69f; // 30mm*15mm

            ////// 添加一页，设置文字字体及格式
            //PdfPageBase page = section.Pages.Add();

            ////创建一个PdfGrid对象
            //PdfGrid grid = new PdfGrid();

            ////设置单元格边距和表格默认字体
            //grid.Style.CellPadding = new PdfPaddings(1, 1, 1, 1);

            ////添加一个4行3列表格
            //PdfGridRow row1 = grid.Rows.Add();
            //PdfGridRow row2 = grid.Rows.Add();
            //PdfGridRow row3 = grid.Rows.Add();
            //PdfGridRow row4 = grid.Rows.Add();

            //grid.Columns.Add(2);

            ////设置列宽
            //grid.Columns[0].Width = 100f;
            //grid.Columns[1].Width = 45f;

            ////foreach (PdfGridRow col in grid.Rows)
            ////{
            ////    col.Height = 55f;
            ////}

            ////字体样式
            //PdfFontBase fontInfo = new PdfTrueTypeFont("Microsoft Yahei", 8f, PdfFontStyle.Regular, true);
            //PdfFontBase fontTitle = new PdfTrueTypeFont("Microsoft Yahei", 8f, PdfFontStyle.Bold, true);

            ////写入数据

            //row1.Cells[1].Value = "加工数量";
            //row1.Cells[1].Style.Font = fontTitle;

            //row2.Cells[0].Value = "读取";
            //row2.Cells[0].Style.Font = fontInfo;
            //row2.Cells[1].Value = "读取";
            //row2.Cells[1].Style.Font = fontInfo;

            //row3.Cells[0].Value = "需求方";
            //row3.Cells[0].Style.Font = fontTitle;

            //row4.Cells[0].Value = "读取(客户)";
            //row4.Cells[0].Style.Font = fontInfo;

            ////水平和垂直方向合并单元格
            //row3.Cells[0].ColumnSpan = 2;
            //row4.Cells[0].ColumnSpan = 2;

            ////设置表格边框颜色、粗细
            //PdfBorders borders = new PdfBorders();
            //borders.All = new PdfPen(Color.Black, 0);
            //foreach (PdfGridRow pgr in grid.Rows)
            //{
            //    foreach (PdfGridCell pgc in pgr.Cells)
            //    {
            //        pgc.Style.Borders = borders;
            //        pgc.StringFormat = new PdfStringFormat(PdfTextAlignment.Center);
            //    }
            //}

            ////在指定位置绘入表格
            //grid.Draw(page, new PointF(0, 0));

            ////保存到文档并预览
            //doc.SaveToFile("PDF表格_1.pdf");
            #endregion

            #region web请求测试
            //string ipStr = "";
            //string url = "http://whois.pconline.com.cn/ipJson.jsp?ip=" + ipStr + "&json=true";
            //try
            //{
            //    HttpWebRequest wrt = (HttpWebRequest)WebRequest.Create(url);
            //    wrt.Method = "GET";
            //    wrt.Timeout = 4000;

            //    HttpWebResponse wrp = (HttpWebResponse)wrt.GetResponse();

            //    StreamReader sr = new StreamReader(wrp.GetResponseStream(), Encoding.Default);

            //    Console.WriteLine(wrt.Timeout);


            //    //获取到的是Json数据
            //    string jsonData = sr.ReadToEnd();

            //    //Newtonsoft.Json读取数据
            //    JObject obj = JsonConvert.DeserializeObject<JObject>(jsonData);

            //    string province = obj["pro"].ToString();
            //    string city = obj["city"].ToString();

            //    if (string.IsNullOrEmpty(province))
            //    {
            //        string addr = obj["addr"].ToString();
            //    }
            //    DemoAsync().Wait();
            //    Console.WriteLine(province + city);
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e.Message);
            //}


            #endregion


            #region 文件压缩

            //ZipTest zipTest = new ZipTest();
            //var fielNames = new Dictionary<string, string>();
            //fielNames.Add("C:\\Users\\Admin\\Desktop\\test\\app.json", string.Empty);
            //fielNames.Add("C:\\Users\\Admin\\Desktop\\app.json", string.Empty);
            //fielNames.Add("C:\\Users\\Admin\\Desktop\\test\\test\\app.json", string.Empty);


            //zipTest.Zip(fielNames, "C:\\Users\\Admin\\Desktop\\zips.zip", "");

            #endregion


            #region pdfSharp
            //PdfHelper.Instance.SaveImageAsPdf("底板.PDF");



            #endregion


            #region pdfLibCore
            //Bitmap bm = new Bitmap("115153.png");


            //// 画表格
            //Bitmap bmp = new Bitmap(550, 120);//新建一个图片对象

            //Graphics gra = Graphics.FromImage(bmp);//利用该图片对象生成“画板”

            //Font font = new Font("雅黑", 12.6f, FontStyle.Underline);//设置字体
            //SolidBrush brush = new SolidBrush(Color.Black);//新建一个画刷,到这里为止,我们已经准备好了画板、画刷、和数据


            //Bitmap bm1 = new Bitmap("123.png");
            ////定义一个区域
            //Rectangle rg = new Rectangle(70, bm.Height * 5 / 6, bm1.Width, bm1.Height);

            //// Pen p = new Pen(Color.Black, 1);//定义了一个红色,宽度为的画笔

            ////gra.DrawRectangle(p, 1, 1, 180, 10);//在画板上画矩形,起始坐标为(10,10),宽为80,高为20
            ////gra.DrawRectangle(p, 1, 1, 80, 10);//在画板上画矩形,起始坐标为(90,10),宽为80,高为20
            ////gra.DrawRectangle(p, 170, 10, 80, 10);//
            ////gra.DrawRectangle(p, 250, 10, 80, 10);//
            //gra.DrawString("加工单号：", font, brush, 1, 1);//
            //gra.DrawString("加工数量：1.27 * 3.4PF单排卧贴2P - 20PW1.8自动机带装管 + ", font, brush, 1, 32);
            //gra.DrawString("产品名称：", font, brush, 1, 63);//进行绘制文字。起始坐标为(172, 12)
            //gra.DrawString("制图人：", font, brush, 1, 94);//关键的一步，进行绘制文字。

            //Rectangle rg1 = new Rectangle(70 + bm1.Width, bm.Height * 5 / 6, bmp.Width, bmp.Height);

            //////要绘制到的位图
            ////Graphics g = Graphics.FromImage(bm);
            //////将bm1内rg所指定的区域绘制到bm
            ////g.DrawImage(bm1, rg);
            ////要绘制到的位图
            //Graphics g1 = Graphics.FromImage(bm);
            ////将bm1内rg所指定的区域绘制到bm
            //g1.DrawImage(bm1, rg);
            //g1.DrawImage(bmp, rg1);

            //bm.Save("999.png", ImageFormat.Png);
            //Bitmap bm = new Bitmap(600, 400);//定义位图实例，并初始化大小
            //Graphics g = Graphics.FromImage(bm);//定义绘图画面，并封装上面的位图实例
            //g.Clear(Color.GreenYellow);//定义绘图画面背景色
            //Pen p = new Pen(Color.Blue, 2);//定义一个2像素大小，蓝色的铅笔
            //                               //用铅笔在画面中间绘制一条x轴线
            //g.DrawLine(p, new Point(0, 200), new Point(bm.Width, 200));
            ////用铅笔在画面中间绘制一条y轴线
            //g.DrawLine(p, new Point(300, 0), new Point(300, bm.Height));
            ////定义矩形区域，其参数分别表示一个矩形的位置和大小
            //Rectangle rect1 = new Rectangle(100, 50, 100, 100);
            //Rectangle rect2 = new Rectangle(400, 50, 100, 100);
            //Rectangle rect3 = new Rectangle(250, 250, 100, 100);
            ////分别用上面定义的矩形区域画圆
            //g.DrawEllipse(p, rect1);
            //g.DrawEllipse(p, rect2);
            //g.DrawEllipse(p, rect3);
            ////绘制一条直线链接第1，2个圆的圆点
            //g.DrawLine(p, new Point(rect1.X + rect1.Width / 2, rect1.Y + rect1.Height / 2), new Point(rect2.X + rect2.Width / 2, rect2.Y + rect2.Height / 2));
            //Brush b = new SolidBrush(Color.Red);//定义一个红色的笔刷
            //Font drawFont = new Font("Arial", 12);//定义一个字体实例
            //                                      //定义坐标值，其中DrawString()方法的参数为：显示值，字体属性，笔刷实例，坐标点
            //for (int i = 0; i <= 6; i++)//循环绘制x轴坐标值
            //{
            //    g.DrawString(Convert.ToString(i * 100), drawFont, b, new PointF(i * 50 + 300, 200));
            //}
            //for (int j = 0; j <= 6; j++)//循环绘制y轴坐标值
            //{
            //    g.DrawString(Convert.ToString(j * 100), drawFont, b, new PointF(300, 200 - j * 50));
            //}
            //bm.Save(Response.OutputStream, System.Drawing.Imaging.ImageFormat.Jpeg);
            #endregion


            #region EPPlus
            ////指定EPPlus使用非商业证书
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;



            //FileInfo existingFile = new FileInfo("C:\\Users\\Admin\\Desktop\\送货单模板.xlsx");


            //using (ExcelPackage package = new ExcelPackage(existingFile))
            //{
            //    ExcelWorksheet workSheet = package.Workbook.Worksheets[0];


            //    workSheet.Cells[5, 16].Value = "日期输入";
            //    workSheet.Cells[6, 2].Value = "客户名称";
            //    workSheet.Cells[6, 9].Value = "联系电话";
            //    workSheet.Cells[6, 14].Value = "客户联系人";
            //    workSheet.Cells[6, 16].Value = "单号填写";
            //    workSheet.Cells[7, 2].Value = "所在地址";


            //    List<List<string>> testdata = new List<List<string>>();

            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "产品名称", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });
            //    testdata.Add(new List<string>() { "客户订单号", "产品编号", "欧式弯插B型四排公头90度4*32P折弯机", "订货数量", "本次送货数", "单位", "单价", "金额", "备注" });

            //    // 位移增量
            //    int increment = testdata.Count - 5 <= 0 ? 0 : testdata.Count - 5;

            //    // 尾部位移
            //    if (testdata.Count > 5)
            //    {
            //        workSheet.Cells[14, 1, 17, 19].Copy(workSheet.Cells[14 + increment, 1, 18, 19]);
            //        workSheet.Row(14 + increment).Height = 23;
            //        workSheet.Row(15 + increment).Height = 23;
            //    }

            //    // 金额合计（大写）
            //    workSheet.Cells[14 + increment, 7].Value = "金额大写";

            //    // 金额合计（小写）
            //    workSheet.Cells[14 + increment, 15].Value = "金额小写";

            //    // 制单人
            //    workSheet.Cells[16 + increment, 2].Value = "制单人";

            //    // 表格内容
            //    for (int i = 0; i < testdata.Count; i++)
            //    {
            //        // 如果超过5行，则要进行画图
            //        if (i >= 4 && i + 1 < testdata.Count)
            //        {
            //            workSheet.Cells[9 + i, 1, 9 + i, 19].Copy(workSheet.Cells[10 + i, 1, 10 + i, 19]);
            //            workSheet.Row(10 + i).Height = 23;
            //        }
            //        workSheet.Cells[9 + i, 1].Value = i + 1;
            //        workSheet.Cells[9 + i, 2].Value = testdata[i][0];
            //        workSheet.Cells[9 + i, 7].Value = testdata[i][1];
            //        workSheet.Cells[9 + i, 8].Value = testdata[i][2];
            //        workSheet.Cells[9 + i, 14].Value = testdata[i][3];
            //        workSheet.Cells[9 + i, 15].Value = testdata[i][4];
            //        workSheet.Cells[9 + i, 16].Value = testdata[i][5];
            //        workSheet.Cells[9 + i, 17].Value = testdata[i][6];
            //        workSheet.Cells[9 + i, 18].Value = testdata[i][7];
            //        workSheet.Cells[9 + i, 19].Value = testdata[i][8];
            //    }


            //    //workSheet.Cells[15, 1, 18, 19].Copy(workSheet.Cells[14, 1, 17, 19]);


            //    //// 合并表头
            //    //workSheet.Cells[1, 1, 2, 9].Merge = true; // 不是从零开始
            //    //workSheet.Cells[1, 1].Value = "领料组装单";
            //    //workSheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    //workSheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            //    //workSheet.Cells[1, 1].Style.Font.Size = 14;

            //    //workSheet.Cells[3, 1, 3, 2].Merge = true;
            //    //workSheet.Cells[3, 1].Value = "总金额：";
            //    //workSheet.Cells[3, 1].Style.Font.Size = 12;
            //    //workSheet.Cells[3, 3, 3, 5].Merge = true;
            //    //workSheet.Cells[3, 3].Value = "总图编号：";
            //    //workSheet.Cells[3, 3].Style.Font.Size = 12;
            //    //workSheet.Cells[3, 6, 3, 9].Merge = true;
            //    //workSheet.Cells[3, 6].Value = "生产需求单号：";
            //    //workSheet.Cells[3, 6].Style.Font.Size = 12;

            //    //workSheet.Cells[4, 1, 4, 9].Merge = true;
            //    //workSheet.Cells[4, 1].Value = "零件列表";
            //    //workSheet.Cells[4, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //    //workSheet.Cells[4, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //    //workSheet.Cells[4, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(143, 188, 143));
            //    //workSheet.Cells[4, 1].Style.Font.Size = 12;

            //    //List<string> titleList = new List<string>() { "零件编号", "零件名称", "外形尺寸", "需求数量","已完成数量","已领取","本次领取", "单价", "金额" };

            //    //workSheet.DrawTitleLine(titleList, 5);

            //    //workSheet.Cells[10, 1, 10, 9].Merge = true;

            //    //workSheet.Cells[10, 1].StyleID = workSheet.Cells[4, 1].StyleID; // 复制样式


            //    //workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();  // 自动列宽
            //    //workSheet.Cells[5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.FromArgb(191, 191, 191));//设置单元格所有边框

            //    MemoryStream ms = new MemoryStream(package.GetAsByteArray());
            //    FileStream fs = new FileStream("C:\\Users\\Admin\\Desktop\\app.xlsx", FileMode.OpenOrCreate);
            //    BinaryWriter w = new BinaryWriter(fs);
            //    w.Write(ms.ToArray());
            //    fs.Close();
            //    ms.Close();

            //}



            #endregion

            #region ThreadPool.QueueUserWorkItem

            //ThreadPool.SetMaxThreads(8, 8);

            //for (int i = 0; i < 100; i++)
            //{
            //    ThreadPool.QueueUserWorkItem(new WaitCallback(DoSomething), "线程" + i);
            //}
            //Console.ReadKey();

            //void DoSomething(object srt)
            //{
            //    Console.WriteLine("传入参数：" + srt);
            //    Thread.Sleep(1000);
            //    Console.WriteLine(srt + "线程结束");
            //}

            #endregion


            #region 批量生成图片
            //var reader = new ExcelReader("C:\\Users\\Admin\\Desktop\\标签制作2\\标签数据规则.xlsx"); // 快速读取excel数据
            //var rows = reader.ReadRows("标签").ToList();

            //for (int i = 1; i < rows.Count; i++)
            //{
            //    using Bitmap bmp = new Bitmap("C:\\Users\\Admin\\Desktop\\标签制作2\\标签模板2.png");//新建一个图片对象
            //    Console.WriteLine(rows[i][4].ToString() + rows[i][8].ToString());

            //    Graphics gra = Graphics.FromImage(bmp);//利用该图片对象生成“画板”

            //    Font font = new Font("Microsoft YaHei UI", 4.2f, FontStyle.Bold);//设置字体
            //    Font font2 = new Font("Microsoft YaHei UI", 5.45f, FontStyle.Bold);//设置字体
            //    SolidBrush brush = new SolidBrush(Color.Black);//新建一个画刷,到这里为止,我们已经准备好了画板、画刷、和数据
            //    gra.DrawString(rows[i][4].ToString(), font, brush, 220, 257);
            //    gra.DrawString(rows[i][8].ToString().Trim('_'), font2, brush, 363, 420);

            //    // 读取二维码
            //    using Bitmap bmpQrCode = new Bitmap("C:\\Users\\Admin\\Desktop\\标签制作2\\mes_script1.0\\" + "tp"+ i.ToString().TrimStart('0') + ".jpeg");

            //    //要绘制到的位图
            //    Graphics bmpGra = Graphics.FromImage(bmp);
            //    Rectangle reQrCode = new Rectangle(315, 130, 280, 280);
            //    bmpGra.DrawImage(bmpQrCode, reQrCode);

            //    bmp.Save("C:\\Users\\Admin\\Desktop\\标签制作2\\生成\\" + rows[i][4].ToString() + ".png", ImageFormat.Png);
            //}
            #endregion
          
        }


    }


}
