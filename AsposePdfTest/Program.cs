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

        public IList<Characteristic> Characteristics { get; set; }

        public IList<Classification> Classifications { get; set; }

        public IList<SupplierOffer> SupplierOffers { get; set; }

        public string SkuModel { get; set; } = "Classic";

        public string Producer { get; set; } = "SVETOCOPY";
    }

    public class Characteristic
    {
        public string Name { get; set; }

        public string Value { get; set; }

        public Characteristic(string name, string value)
        {
            Name = name;
            Value = value;
        }
    }

    public class Classification
    {
        public string Dictionary { get; set; }

        public string Code { get; set; }

        public Classification(string dictionary, string code)
        {
            Dictionary = dictionary;
            Code = code;
        }
    }

    public class SupplierOffer
    {
        public string Name { get; }

        public string Inn { get; }

        public string Offer { get; }

        public string DeliveryTime { get; }

        public string DeliveryRegion { get; }

        public int Price { get; }

        public int Vat { get; }

        public SupplierOffer(string name, string inn, string offer, string deliveryTime, string deliveryRegion, int price, int vat)
        {
            Name = name;
            Inn = inn;
            Offer = offer;
            DeliveryTime = deliveryTime;
            DeliveryRegion = deliveryRegion;
            Price = price;
            Vat = vat;
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
                Characteristics = new List<Characteristic>
                {
                    new Characteristic("Количество листов в пачке", "45 шт"),
                    new Characteristic("Класс бумаги", "C"),
                    new Characteristic("Плотность бумаги", "80 гр/м2"),
                    new Characteristic("Страна происхождения", "Россия")
                },
                Classifications = new List<Classification>
                {
                    new Classification("ОКПД2", "32.99.12.120 Ручки и маркеры с наконечником из фетра и прочих пористых материалов"),
                    new Classification("КПГЗ", "01.15.05.03.04 Принадлежности для досок и флипчартов"),
                    new Classification("КТРУ", "32.99.12.120-00000001 Маркер")
                },
                SupplierOffers = new List<SupplierOffer>
                {
                    new SupplierOffer("ООО \"ЯСТРЕБ\"", "5427713792", "1219511-21", "5-10 дней", "г Москва, обл Московская", 47, 20),
                    new SupplierOffer("ИП ВЛАДИМИР", "4767898456", "4578548-45", "2-3 часа", "г Казань, Татарстан", 30, 20),
                    new SupplierOffer("ООО \"ЖЕНЕРИК\"", "7894598723", "6578459-12", "1-2 недели", "г Москва, обл Московская", 50, 20)
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

            // Classification Table
            var classificationTable = CreateClassificationTable(model);
            page.Paragraphs.Add(classificationTable);

            // Classification Table
            var supplierOffersTable = CreateSupplierOffersTable(model);
            page.Paragraphs.Add(supplierOffersTable);

            
            using var resultStream = File.Open("D:\\WorkAsposeFiles\\template.pdf", FileMode.Create);
            pdfDocument.Save(resultStream, SaveFormat.Pdf);
        }

        private static Table CreateSupplierOffersTable(SkuPdfModel model)
        {
            return TableBuilder.Create(builder =>
            {
                builder.SetCellPaddings(5f);
                builder.SetColumnWidths(3, "150 200 100");
                builder.SetDefaultTextStyle();

                builder.AddSingleCellRow(cellBuilder =>
                {
                    cellBuilder.Cell.Margin = new MarginInfo(0, 12, 0, 15);
                    cellBuilder.AddHtml(fragmentBuilder => fragmentBuilder.AppendHeader("Предложения поставщиков"));
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.Row.BackgroundColor = Color.FromRgb(231 / 255f, 239 / 255f, 247 / 255f);

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Поставщик");
                    }));

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Условия поставки");
                    }));

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Стоимость с НДС");
                    }));
                });

                foreach (var supplierOffer in model.SupplierOffers)
                {
                    builder.AddRow(rowBuilder =>
                    {
                        rowBuilder.Row.Border = new BorderInfo(BorderSide.Top,
                            new GraphInfo {Color = Color.Gray, LineWidth = 0.5f});

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder
                                .AppendBoldText(supplierOffer.Name)
                                .AppendText($"ИНН: {supplierOffer.Inn}")
                                .AppendText($"Оферта: №: {supplierOffer.Offer}");
                        }));

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder
                                .AppendText($"Срок поставки: {supplierOffer.DeliveryTime}")
                                .AppendText($"Регион поставки: {supplierOffer.DeliveryRegion}");
                        }));

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder
                                .AppendHighlighted($"{supplierOffer.Price} ₽")
                                .AppendText($"НДС: {supplierOffer.Vat}%");
                        }));
                    });
                }
            });
        }

        private static Table CreateClassificationTable(SkuPdfModel model)
        {
            return TableBuilder.Create(builder =>
            {
                builder.SetCellPaddings(5f);
                builder.SetColumnWidths(2, "200 300");
                builder.SetDefaultTextStyle();
                //builder.AddBorders();

                builder.AddSingleCellRow(cellBuilder =>
                {
                    cellBuilder.Cell.Margin = new MarginInfo(0, 12, 0, 15);
                    cellBuilder.AddHtml(fragmentBuilder => fragmentBuilder.AppendHeader("Классификация"));
                });

                builder.AddRow(rowBuilder =>
                {
                    rowBuilder.Row.BackgroundColor = Color.FromRgb(231 / 255f, 239 / 255f, 247 / 255f);

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Справочник");
                    }));

                    rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                    {
                        htmlBuilder.AppendText("Код");
                    }));
                });

                foreach (var classification in model.Classifications)
                {
                    builder.AddRow(rowBuilder =>
                    {
                        rowBuilder.Row.Border = new BorderInfo(BorderSide.Top,
                            new GraphInfo {Color = Color.Gray, LineWidth = 0.5f});

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder.AppendText(classification.Dictionary);
                        }));

                        rowBuilder.AddCell(cellBuilder => cellBuilder.AddHtml(htmlBuilder =>
                        {
                            htmlBuilder.AppendText(classification.Code);
                        }));
                    });
                }
            });
        }

        private static Table CreateCharacteristicTable(SkuPdfModel model)
        {
            return TableBuilder.Create(builder =>
            {
                builder.SetCellPaddings(5f);
                builder.SetColumnWidths(2, "200 300");
                builder.SetDefaultTextStyle();
                //builder.AddBorders();

                builder.AddSingleCellRow(cellBuilder =>
                {
                    cellBuilder.Cell.Margin = new MarginInfo(0, 0, 0, 10);
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
                        rowBuilder.Row.Border = new BorderInfo(BorderSide.Top,
                            new GraphInfo {Color = Color.Gray, LineWidth = 0.5f});

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
                //builder.AddBorders();

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

        public HtmlFragmentBuilder AppendEmptyLine()
        {
            _lines.Add(string.Empty);

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

        public HtmlFragmentBuilder AppendBoldText(string text)
        {
            _lines.Add($"<span style=\"font-weight: bold;\">{text}</span>");

            return this;
        }

        public HtmlFragment Build() => new HtmlFragment(string.Join("<br />", _lines));
    }
}

