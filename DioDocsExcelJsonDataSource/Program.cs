// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using System.IO;

Console.WriteLine("JSON形式のデータをワークシートにバインドする");

// 新規ワークブックの作成
var workbook = new GrapeCity.Documents.Excel.Workbook();

#region ワークシートにJSONデータをバインドする

var worksheet = workbook.Worksheets[0];
worksheet.Name = "ワークシートにバインド";

// JSONファイルからデータを取得
var stream = new FileStream("Json\\data.json", FileMode.Open, FileAccess.Read);
var reader = new StreamReader(stream);
var jsonText = reader.ReadToEnd();

// データソースに設定
worksheet.DataSource = new JsonDataSource(jsonText);

// ヘッダー行を追加
worksheet.Range["A1"].EntireRow.Insert(InsertShiftDirection.Down);
var header = new object[] { "地域", "都市", "カテゴリ", "商品名", "売上" };
worksheet.Range["A1:E1"].Value = header;
worksheet.Range["A1:E1"].Style = workbook.Styles["アクセント 1"];
worksheet.Range["A1:E1"].Font.Bold = true;

// Excelファイルに保存
workbook.Save("JsonDataSource.xlsx");

#endregion

#region テーブルにJSONデータをバインドする

var worksheet1 = workbook.Worksheets.Add();
worksheet1.Name = "テーブルにバインド";

// JSONファイルからデータを取得
var stream1 = new FileStream("Json\\data1.json", FileMode.Open, FileAccess.Read);
var reader1 = new StreamReader(stream1);
var jsonText1 = reader1.ReadToEnd();

// データソースに設定
worksheet1.DataSource = new JsonDataSource(jsonText1);

// テーブルを追加
ITable table = worksheet1.Tables.Add(worksheet.Range["A1:E10"], false);
table.AutoGenerateColumns = true;
table.BindingPath = "data";

// テーブルにヘッダーを追加
var header1 = new object[] { "地域", "都市", "カテゴリ", "商品名", "売上" };
table.HeaderRange.Value = header1;

// Excelファイルに保存
workbook.Save("JsonDataSource.xlsx");

#endregion

#region 帳票テンプレートにJSONデータをバインドする

// 帳票テンプレートを読み込む
var workbookTemplate = new GrapeCity.Documents.Excel.Workbook();
var template = new FileStream("Template\\reporttemplate.xlsx", FileMode.Open, FileAccess.Read);
workbookTemplate.Open(template);

// JSONファイルからデータを取得
var stream2 = new FileStream("Json\\data1.json", FileMode.Open, FileAccess.Read);
var reader2 = new StreamReader(stream2);
var jsonText2 = reader2.ReadToEnd();

// データソースに設定
var datasource1 = new JsonDataSource(jsonText2);

// ワークブックに追加
workbookTemplate.AddDataSource("ds", datasource1);

// 帳票テンプレートの処理を実施
workbookTemplate.ProcessTemplate();

// Excelファイルに保存
workbookTemplate.Save("JsonDataReportTemplate.xlsx");

#endregion