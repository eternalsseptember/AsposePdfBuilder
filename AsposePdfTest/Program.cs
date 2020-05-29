using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using Document = Aspose.Pdf.Document;
using SaveFormat = Aspose.Pdf.SaveFormat;

namespace AsposePdfTest
{
    public class SkuPdfModel
    {
        public string Name { get; set; } = "Бумага офисная Svetocopy";

        public decimal? MinPrice { get; set; } = 208;

        public decimal? MaxPrice { get; set; } = 306.32M;

        public decimal? ReferencePrice { get; set; } = 245.88M;

        public bool? IsDemanded { get; set; } = true;

        public decimal? ContractsCount { get; set; } = 759;

        public IList<Characteristics> Characteristics { get; set; }

        public string SkuModel { get; set; } = "Classic";

        public string Producer { get; set; } = "SVETOCOPY";
    }

    public class Characteristics
    {
        public string Name { get; set; }

        public string Value { get; set; }

        public Characteristics(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            foreach (var process in Process.GetProcessesByName("AcroRd32"))
            {
                process.Kill();
            }

            using var fileStream = File.Open("D:\\WorkAsposeFiles\\License.txt", FileMode.Open);

            var pdf = new Aspose.Pdf.License();
            pdf.SetLicense(fileStream);
            fileStream.Seek(0, SeekOrigin.Begin);

            var word = new Aspose.Words.License();
            word.SetLicense(fileStream);
            fileStream.Seek(0, SeekOrigin.Begin);

            var cells = new Aspose.Cells.License();
            cells.SetLicense(fileStream);

            var model = new SkuPdfModel()
            {
                Characteristics = new List<Characteristics>
                {
                    new Characteristics("Количество листов в пачке", "45 шт"),
                    new Characteristics("Класс бумаги", "C"),
                    new Characteristics("Плотность бумаги", "80 гр/м2"),
                    new Characteristics("Страна происхождения", "Россия")
                }
            };

            CreateTemplate(model);

            Process.Start("D:\\WorkAsposeFiles\\template.pdf");
        }

        private static void CreateTemplate(SkuPdfModel model)
        {
            var pdfDocument = new Document();
            using var imageStore = new ImageStore();
            var page = pdfDocument.Pages.Add();

            // Header Table
            var headerTable = CreateHeaderTable(model, imageStore);
            page.Paragraphs.Add(headerTable);

            // Characteristic Table
            var characteristicTable = CreateCharacteristicTable(model);
            page.Paragraphs.Add(characteristicTable);

            
            using var resultStream = File.Open("D:\\WorkAsposeFiles\\template.pdf", FileMode.Create);
            pdfDocument.Save(resultStream, SaveFormat.Pdf);
        }

        private static Table CreateCharacteristicTable(SkuPdfModel model)
        {
            return TableBuilder.Create(builder =>
            {
                builder.SetCellPaddings(5f);
                builder.SetColumnWidths(2, "200 300");
                builder.SetDefaultTextStyle();
                builder.AddBorders();

                builder.AddSingleCellRow(cellBuilder =>
                {
                    cellBuilder.AddHtml(fragmentBuilder => fragmentBuilder.AppendHeader("Характеристики"));
                });

                builder.AddSingleCellRow(cellBuilder =>
                {
                    cellBuilder.AddHtml(fragmentBuilder =>
                        {
                            fragmentBuilder
                                .AppendText("Модель: " + model.SkuModel)
                                .AppendText("Производитель: " + model.Producer);
                        });
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.Row.BackgroundColor = Color.FromRgb(231 / 255f, 239 / 255f, 247 / 255f);

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Наименование характеристики");
                    }));

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Значение");
                    }));
                });

                foreach (var characteristic in model.Characteristics)
                {
                    builder.AddRow(rowBuilder =>
                    {
                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder.AppendText(characteristic.Name);
                        }));

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder.AppendText(characteristic.Value);
                        }));
                    });
                }
            });
        }

        private static Table CreateHeaderTable(SkuPdfModel model, ImageStore imageStore)
        {
            var table = TableBuilder.Create(builder =>
            {
                builder.SetCellPaddings(5f);
                builder.SetColumnWidths(3, "150 150 150");
                builder.SetDefaultTextStyle();
                builder.AddBorders();

                builder.AddRow(rowBuilder =>
                {
                    var image = imageStore.GetImage("D:\\WorkAsposeFiles\\2.png", image => image.ImageScale = 0.6);

                    rowBuilder.AddCell(3, cellBuilder =>
                    {
                        cellBuilder.AddImage(image);
                    });
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.AddCell(2, cellBuilder =>
                    {
                        cellBuilder.Cell.VerticalAlignment = VerticalAlignment.Center;
                        cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder.AppendHighlighted(model.Name).Build();
                        });
                    });

                    rowBuilder.AddCell(cellBuilder =>
                    {
                        var qr = imageStore.GetImage(
                            "D:\\WorkAsposeFiles\\3.gif",
                            image =>
                            {
                                image.FixHeight = 80;
                                image.FixWidth = 80;
                                image.VerticalAlignment = VerticalAlignment.Top;
                                image.HorizontalAlignment = HorizontalAlignment.Right;
                            });

                        cellBuilder.AddImage(qr);
                    });
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.AddCell(cellBuilder =>
                    {
                        var sku = imageStore.GetImage("D:\\WorkAsposeFiles\\1.jpg", image =>
                        {
                            image.FixHeight = 150;
                            image.FixWidth = 150;
                        });

                        cellBuilder.Cell.RowSpan = 3;
                        cellBuilder.AddImage(sku);
                    });

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder
                            .AppendText("Референтная цена")
                            .AppendHighlighted(model.ReferencePrice)
                            .AppendText("штука");
                    }));

                    rowBuilder.AddCell(cellBuilder =>
                    {
                        cellBuilder.Cell.Alignment = HorizontalAlignment.Right;
                        cellBuilder.Cell.VerticalAlignment = VerticalAlignment.Top;

                        cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder
                                .AppendText("Контрактов")
                                .AppendHighlighted(model.ContractsCount);
                        });
                    });
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.AddCell(2, cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder
                            .AppendHighlighted($"от {model.MinPrice} до {model.MaxPrice}")
                            .AppendText("на основе 24 предложений поставщиков");
                    }));
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.AddCell(2, cellBuilder =>
                    {
                        cellBuilder.Cell.VerticalAlignment = VerticalAlignment.Top;

                        cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder
                                .AppendText("Востребованная продукция");
                        });
                    });
                });
            });

            return table;
        }
    }

    public class ImageStore : IDisposable
    {
        private readonly List<Stream> _openStreams = new List<Stream>();

        public Image GetImage(string path, Action<Image> configure)
        {
            var stream = new FileStream(path, FileMode.Open);
            _openStreams.Add(stream);

            var image = new Image
            {
                ImageStream = stream,
            };

            configure(image);
            
            return image;
        }

        public void Dispose()
        {
            foreach (var openStream in _openStreams)
            {
                openStream.Dispose();
            }
        }
    }

    public class TableBuilder
    {
        private int _columnsCount;

        public static Table Create(Action<TableBuilder> configure)
        {
            var builder = new TableBuilder();
            configure(builder);

            return builder.Table;
        }

        public Table Table { get; }

        public TableBuilder()
        {
            Table = new Table();
        }

        public void AddRow(Action<RowBuilder> configure)
        {
            var builder = new RowBuilder(Table.Rows.Add());

            configure(builder);
        }

        public void AddSingleCellRow(Action<CellBuilder> configure)
        {
            var builder = new RowBuilder(Table.Rows.Add());

            builder.AddCell(_columnsCount, configure);
        }

        public void AddBorders()
        {
            Table.IsBordersIncluded = true;
            Table.Border = new BorderInfo(BorderSide.All);
            Table.DefaultCellBorder = new BorderInfo(BorderSide.All);
        }

        public void SetCellPaddings(double padding)
        {
            Table.DefaultCellPadding = new MarginInfo
            {
                Top = padding, 
                Left = padding, 
                Right = padding, 
                Bottom = padding
            };
        }

        public void SetColumnWidths(int columnsCount, string widths)
        {
            Table.ColumnWidths = widths;
            _columnsCount = columnsCount;
        }

        public void SetDefaultTextStyle(TextState customState = null)
        {
            Table.DefaultCellTextState = customState ?? new TextState
            {
                FontSize = 12,
                Font = FontRepository.FindFont("Arial")
            };
        }
    }

    public class RowBuilder
    {
        public Row Row { get; }

        public RowBuilder(Row row)
        {
            Row = row;
        }

        public void AddCell(Action<CellBuilder> configure)
        {
            var builder = new CellBuilder(Row.Cells.Add());

            configure(builder);
        }

        public void AddCell(int colSpan, Action<CellBuilder> configure)
        {
            AddCell(cellBuilder =>
            {
                cellBuilder.Cell.ColSpan = colSpan;
                configure(cellBuilder);
            });
        }
    }

    public class CellBuilder
    {
        public Cell Cell { get; }

        public CellBuilder(Cell cell)
        {
            Cell = cell;
        }

        public void AddHtml(Action<HtmlFragmentBuilder> configure)
        {
            var htmlBuilder = new HtmlFragmentBuilder();

            configure(htmlBuilder);
            
            Cell.Paragraphs.Add(htmlBuilder.Build());
        }

        public void AddImage(Image image)
        {
            Cell.Paragraphs.Add(image);
        }
    }

    public class HtmlFragmentBuilder
    {
        private readonly List<string> _lines = new List<string>();
        
        public HtmlFragmentBuilder AppendText(object text) => AppendText(text.ToString());

        public HtmlFragmentBuilder AppendText(string text)
        {
            _lines.Add(text);

            return this;
        }

        public HtmlFragmentBuilder AppendHighlighted(object text) => AppendHighlighted(text.ToString());

        public HtmlFragmentBuilder AppendHighlighted(string text)
        {
            _lines.Add($"<span style=\"font-size: 1.5em; font-weight: bold;\">{text}</span>");

            return this;
        }

        public HtmlFragmentBuilder AppendHeader(object text) => AppendHeader(text.ToString());

        public HtmlFragmentBuilder AppendHeader(string text)
        {
            _lines.Add($"<span style=\"font-size: 1.4em; font-weight: bold; color: #254b82;\">{text}</span>");

            return this;
        }

        public HtmlFragment Build() => new HtmlFragment(string.Join("<br />", _lines));
    }
}

