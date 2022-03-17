using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace UP01._01
{
    /// <summary>
    /// Логика взаимодействия для StockPage.xaml
    /// </summary>
    public partial class StockPage : Page
    {
        List<Material> MaterialStart = BaseClass.Base.Material.ToList();
        PageChanges pc = new PageChanges();
        List<Material> MaterialFilterSort;

        public StockPage()
        {
            InitializeComponent();
            LVStock.ItemsSource = BaseClass.Base.Material.ToList();
            List<MaterialType> MT = BaseClass.Base.MaterialType.ToList();
            CBFilter.Items.Add("Все записи");
            for(int i =0;i<MT.Count();i++)
            {
                CBFilter.Items.Add(MT[i].Title);
            }
            CBFilter.SelectedIndex = 0;
            CBSorting.Items.Add("Все записи");
            CBSorting.Items.Add("Наименование");
            CBSorting.Items.Add("Остаток на складе");
            CBSorting.Items.Add("Стоимость");
            CBSorting.SelectedIndex = 0;
            DataContext = pc;
            pc.CountPage = 15;
            pc.Countlist = MaterialFilters.Count;
            LVStock.ItemsSource = MaterialFilters.Skip(0).Take(pc.CountPage).ToList();
        }

        private void TextBlock_Loaded(object sender, RoutedEventArgs e)
        {
            TextBlock tb = (TextBlock)sender;
            int index = Convert.ToInt32(tb.Uid);
            string typename = "";
            List<Material> ML = BaseClass.Base.Material.Where(x => x.ID == index).ToList();
            foreach (Material s in ML)
            {
                    typename += s.MaterialType.Title + " | " + s.Title;
            }
            tb.Text = typename;
        }

        private void TextBlock_Loaded_1(object sender, RoutedEventArgs e)
        {
            TextBlock tb = (TextBlock)sender;
            int index = Convert.ToInt32(tb.Uid);
            string suppliers = "Поставщики:  ";
            List<MaterialSupplier> MS = BaseClass.Base.MaterialSupplier.Where(x => x.MaterialID == index).ToList();
            List<Supplier> S = BaseClass.Base.Supplier.Where(x => x.ID == index).ToList();
            foreach(MaterialSupplier s in MS)
            {
                foreach (Supplier t in S)
                {
                    suppliers += s.Supplier.Title+", ";
                }
            }
            if(suppliers!="Поставщики:  ")
            {
               tb.Text = suppliers.Substring(0,suppliers.Length-2);
            }
            else
            {
                suppliers += "-";
                tb.Text = suppliers;
            }
           
        }

        List<Material> MaterialFilters;

        private void Filters()
        {
            int index = CBFilter.SelectedIndex;
            if(index!=0)
            {
                MaterialFilters = MaterialStart.Where(x => x.MaterialTypeID == index).ToList();
            }
            else
            {
                MaterialFilters = MaterialStart;
            }

            if(!string.IsNullOrWhiteSpace(TBFilter.Text))
            {
                MaterialFilters = MaterialFilters.Where(x => x.Title.ToLower().Contains(TBFilter.Text.ToLower())).ToList();
            }
            LVStock.ItemsSource = MaterialFilters;
            TBlCount.Text = "Количество записей - " + MaterialFilters.Count() + " из " + MaterialStart.Count() ;
        }

        private void TBFilter_TextChanged(object sender, TextChangedEventArgs e)
        {
            Filters();
            LVStock.ItemsSource = MaterialFilters.Skip(0).Take(pc.CountPage).ToList();
        }

        private void CBFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Filters();
            LVStock.ItemsSource = MaterialFilters.Skip(0).Take(pc.CountPage).ToList();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if(CBSorting.SelectedIndex==0)
            {
                MaterialFilters.Sort((x, y) => x.ID.CompareTo(y.ID));
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 1)
            {
                MaterialFilters.Sort((x, y) => x.Title.CompareTo(y.Title));
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 2)
            {
                MaterialFilters.Sort((x, y) => Convert.ToInt32(x.CountInStock).CompareTo(Convert.ToInt32(y.CountInStock)));
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 3)
            {
                MaterialFilters.Sort((x, y) => x.Cost.CompareTo(y.Cost));
                LVStock.Items.Refresh();
            }
            LVStock.ItemsSource = MaterialFilters.Skip(0).Take(pc.CountPage).ToList();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (CBSorting.SelectedIndex == 0)
            {
                MaterialFilters.Sort((x, y) => x.ID.CompareTo(y.ID));
                MaterialFilters.Reverse();
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 1)
            {
                MaterialFilters.Sort((x, y) => x.Title.CompareTo(y.Title));
                MaterialFilters.Reverse();
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 2)
            {
                MaterialFilters.Sort((x, y) => Convert.ToInt32(x.CountInStock).CompareTo(Convert.ToInt32(y.CountInStock)));
                MaterialFilters.Reverse();
                LVStock.Items.Refresh();
            }
            if (CBSorting.SelectedIndex == 3)
            {
                MaterialFilters.Sort((x, y) => x.Cost.CompareTo(y.Cost));
                MaterialFilters.Reverse();
                LVStock.Items.Refresh();
            }
            LVStock.ItemsSource = MaterialFilters.Skip(0).Take(pc.CountPage).ToList();
        }

        private void AddNew_Click(object sender, RoutedEventArgs e)
        {
            FrameClass.MainFrame.Navigate(new CreateOrUpdate());
        }

        private void Update_Click(object sender, RoutedEventArgs e)
        {
            Button B = (Button)sender;
            int id = Convert.ToInt32(B.Uid);
            Material MaterialUpdate = BaseClass.Base.Material.FirstOrDefault(x => x.ID == id);
            FrameClass.MainFrame.Navigate(new CreateOrUpdate(MaterialUpdate));
        }

        private void LVStock_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(LVStock.SelectedIndex>-1)
            {
                ChangeCount.Visibility = Visibility.Visible;
            }
            else
            {
                ChangeCount.Visibility = Visibility.Collapsed;
            }
                  
        }

        private void ChangeCount_Click(object sender, RoutedEventArgs e)
        {
            var list = LVStock.SelectedItems;
            double max = 1;
            foreach (Material m in list)
            {
                if (m.MinCount > max)
                {
                    max = m.MinCount;
                }
            }
            ChangeMinCountWindow window = new ChangeMinCountWindow(max);
           

            window.ShowDialog();
            if(window.Count>0)
            {
                foreach (Material m in list)
                {
                    m.MinCount = window.Count;
                }
                LVStock.Items.Refresh();
            }
        }

        private void GoPage_MouseDown(object sender, MouseButtonEventArgs e)  
        {
            TextBlock tb = (TextBlock)sender;
            switch (tb.Uid)  
            {
                case "prev":
                    pc.CurrentPage--;
                    break;
                case "next":
                    pc.CurrentPage++;
                    break;
                default:
                    pc.CurrentPage = Convert.ToInt32(tb.Text);
                    break;
            }
            LVStock.ItemsSource = MaterialFilters.Skip(pc.CurrentPage * pc.CountPage - pc.CountPage).Take(pc.CountPage).ToList();  
        }

        private void btnCreateReport_Click(object sender, RoutedEventArgs e)
        {
            SelectReportType window = new SelectReportType();
            window.ShowDialog();
            int type=window.Type;

            switch(type)
            {
                case 1:
                    {
                        List<Material> Material = BaseClass.Base.Material.ToList();
                        var application = new Word.Application();
                        Word.Document document = application.Documents.Add();
                        foreach(Material mat in Material)
                        {
                            Word.Paragraph materialParagraph = document.Paragraphs.Add();
                            Word.Range materialRange = materialParagraph.Range;
                            materialRange.Text = mat.Title;
                            materialRange.InsertParagraphAfter();

                            Word.Paragraph tableParagraph = document.Paragraphs.Add();
                            Word.Range tableRange = tableParagraph.Range;
                            Word.Table materialTable = document.Tables.Add(tableRange, 2, 5);
                            materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            Word.Range cellRange;

                            cellRange = materialTable.Cell(1, 1).Range;
                            cellRange.Text = "Количество в упаковке";
                            cellRange = materialTable.Cell(1, 2).Range;
                            cellRange.Text = "Минимальное количество";
                            cellRange = materialTable.Cell(1, 3).Range;
                            cellRange.Text = "Количество на складе";
                            cellRange = materialTable.Cell(1, 4).Range;
                            cellRange.Text = "Единица измерения";
                            cellRange = materialTable.Cell(1, 5).Range;
                            cellRange.Text = "Стоимость";

                            materialTable.Rows[1].Range.Bold = 1;
                            materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            cellRange = materialTable.Cell(2, 1).Range;
                            cellRange.Text = Convert.ToString(mat.CountInPack);
                            cellRange = materialTable.Cell(2, 2).Range;
                            cellRange.Text = Convert.ToString(mat.MinCount);
                            cellRange = materialTable.Cell(2, 3).Range;
                            cellRange.Text = Convert.ToString(mat.CountInStock);
                            cellRange = materialTable.Cell(2, 4).Range;
                            cellRange.Text = mat.Unit;
                            cellRange = materialTable.Cell(2, 5).Range;
                            cellRange.Text = Convert.ToString(mat.Cost);

                            if (mat != Material.LastOrDefault()) document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                        }

                        string path = Environment.CurrentDirectory;
                        document.SaveAs2(path+"Report1.docx");
                        document.SaveAs2(path + "Report1.pdf",Word.WdExportFormat.wdExportFormatPDF);
                        break;
                    }
                case 2:
                    {
                        List<Material> Materials = BaseClass.Base.Material.Where(x=>x.CountInStock<x.MinCount).ToList();
                        var application = new Word.Application();
                        Word.Document document = application.Documents.Add(); 
                        Word.Paragraph materialParagraph = document.Paragraphs.Add();
                        Word.Range materialRange = materialParagraph.Range;
                        materialRange.Text = "Метериалы количество на складе которых меньше минимального количества";
                        materialRange.InsertParagraphAfter();

                        Word.Paragraph tableParagraph = document.Paragraphs.Add();
                        Word.Range tableRange = tableParagraph.Range;
                        Word.Table materialTable = document.Tables.Add(tableRange, Materials.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        Word.Range cellRange;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование товара";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "Минимальное количество";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Количество на складе";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        for (int i = 0; i < Materials.Count(); i++)
                        {

                            var currentMaterial = Materials[i];

                            cellRange = materialTable.Cell(i + 2, 1).Range;
                            cellRange.Text = currentMaterial.Title;

                            cellRange = materialTable.Cell(i + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentMaterial.MinCount);

                            cellRange = materialTable.Cell(i + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentMaterial.CountInStock);
                        }


                        string path = Environment.CurrentDirectory;
                        document.SaveAs2(path + "Report2.docx");
                        document.SaveAs2(path + "Report2.pdf", Word.WdExportFormat.wdExportFormatPDF);
                        break;
                    }
                case 3:
                    {
                        List<Material> Material = BaseClass.Base.Material.Where(x => x.MaterialTypeID == 1).ToList();
                        List<MaterialType> MatType = BaseClass.Base.MaterialType.ToList();
                        var application = new Word.Application();
                        Word.Document document = application.Documents.Add();
                       
                            Word.Paragraph materialParagraph = document.Paragraphs.Add();
                            Word.Range materialRange = materialParagraph.Range;
                            materialRange.Text = "Гранулы";
                            materialRange.InsertParagraphAfter();

                            Word.Paragraph tableParagraph = document.Paragraphs.Add();
                            Word.Range tableRange = tableParagraph.Range;
                            Word.Table materialTable = document.Tables.Add(tableRange, Material.Count(), 3);
                            materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            Word.Range cellRange;

                            cellRange = materialTable.Cell(1, 1).Range;
                            cellRange.Text = "Наименование товара";
                            cellRange = materialTable.Cell(1, 2).Range;
                            cellRange.Text = "Количество на складе";
                            cellRange = materialTable.Cell(1, 3).Range;
                            cellRange.Text = "Стоимость";
                            materialTable.Rows[1].Range.Bold = 1;
                            materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            for (int i = 0; i < Material.Count(); i++)
                            {
                                
                                var currentMaterial = Material[i];

                                cellRange = materialTable.Cell(i + 2, 1).Range;
                                cellRange.Text = currentMaterial.Title;

                                cellRange = materialTable.Cell(i + 2, 2).Range;
                                cellRange.Text = Convert.ToString(currentMaterial.MinCount);

                                cellRange = materialTable.Cell(i + 2, 3).Range;
                                cellRange.Text = Convert.ToString(currentMaterial.CountInStock);
                            }
                        
                        string path = Environment.CurrentDirectory;
                        document.SaveAs2(path + "Report3.docx");
                        document.SaveAs2(path + "Report3.pdf", Word.WdExportFormat.wdExportFormatPDF);
                        break;
                    }
                case 4:
                    {
                        List<Material> Materials = BaseClass.Base.Material.Where(x => x.CountInStock == x.MinCount*3).ToList();
                        var application = new Word.Application();
                        Word.Document document = application.Documents.Add();
                        Word.Paragraph materialParagraph = document.Paragraphs.Add();
                        Word.Range materialRange = materialParagraph.Range;
                        materialRange.Text = "Метериалы количество на складе которых равно 300% минимального количества";
                        materialRange.InsertParagraphAfter();

                        Word.Paragraph tableParagraph = document.Paragraphs.Add();
                        Word.Range tableRange = tableParagraph.Range;
                        Word.Table materialTable = document.Tables.Add(tableRange, Materials.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        Word.Range cellRange;
                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование товара";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "Минимальное количество";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Количество на складе";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        for (int i = 0; i < Materials.Count(); i++)
                        {

                            var currentMaterial = Materials[i];

                            cellRange = materialTable.Cell(i + 2, 1).Range;
                            cellRange.Text = currentMaterial.Title;

                            cellRange = materialTable.Cell(i + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentMaterial.MinCount);

                            cellRange = materialTable.Cell(i + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentMaterial.CountInStock);
                        }

                        string path = Environment.CurrentDirectory;
                        document.SaveAs2(path + "Report4.docx");
                        document.SaveAs2(path + "Report4.pdf", Word.WdExportFormat.wdExportFormatPDF);
                        break;
                    }
                case 5:
                    {
                        string[] suptype = new string[] { "МКК", "ОАО", "ООО", "МФО", "ПАО", "ЗАО" };
                        List<Supplier> MKK = BaseClass.Base.Supplier.Where(x => x.SupplierType == "МКК").ToList();
                        List<Supplier> OAO = BaseClass.Base.Supplier.Where(x => x.SupplierType == "ОАО").ToList();
                        List<Supplier> OOO = BaseClass.Base.Supplier.Where(x => x.SupplierType == "ООО").ToList();
                        List<Supplier> MFO = BaseClass.Base.Supplier.Where(x => x.SupplierType == "МФО").ToList();
                        List<Supplier> PAO = BaseClass.Base.Supplier.Where(x => x.SupplierType == "ПАО").ToList();
                        List<Supplier> ZAO = BaseClass.Base.Supplier.Where(x => x.SupplierType == "ЗАО").ToList();
                        
                        
                            var application = new Word.Application();
                            Word.Document document = application.Documents.Add();
                            Word.Paragraph materialParagraph = document.Paragraphs.Add();
                            Word.Range materialRange = materialParagraph.Range;
                            materialRange.Text = suptype[0] ;
                            materialRange.InsertParagraphAfter();

                            Word.Paragraph tableParagraph = document.Paragraphs.Add();
                            Word.Range tableRange = tableParagraph.Range;
                            Word.Table materialTable = document.Tables.Add(tableRange, MKK.Count() + 1, 3);
                            materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                            materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            Word.Range cellRange;
                            cellRange = materialTable.Cell(1, 1).Range;
                            cellRange.Text = "Наименование поставщика";
                            cellRange = materialTable.Cell(1, 2).Range;
                            cellRange.Text = "ИНН";
                            cellRange = materialTable.Cell(1, 3).Range;
                            cellRange.Text = "Рейтинг";
                            materialTable.Rows[1].Range.Bold = 1;
                            materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                            for (int j = 0; j < MKK.Count(); j++)
                            {

                                var currentsup = MKK[j];

                                cellRange = materialTable.Cell(j + 2, 1).Range;
                                cellRange.Text = currentsup.Title;

                                cellRange = materialTable.Cell(j + 2, 2).Range;
                                cellRange.Text = Convert.ToString(currentsup.INN);

                                cellRange = materialTable.Cell(j + 2, 3).Range;
                                cellRange.Text = Convert.ToString(currentsup.QualityRating);
                            }
                            document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);


                        materialParagraph = document.Paragraphs.Add();
                        materialRange = materialParagraph.Range;
                        materialRange.Text = suptype[1];
                        materialRange.InsertParagraphAfter();

                        tableParagraph = document.Paragraphs.Add();
                        tableRange = tableParagraph.Range;
                        materialTable = document.Tables.Add(tableRange, OAO.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование поставщика";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "ИНН";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Рейтинг";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (int j = 0; j < OAO.Count(); j++)
                        {

                            var currentsup = OAO[j];

                            cellRange = materialTable.Cell(j + 2, 1).Range;
                            cellRange.Text = currentsup.Title;

                            cellRange = materialTable.Cell(j + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentsup.INN);

                            cellRange = materialTable.Cell(j + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentsup.QualityRating);
                        }
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                        materialParagraph = document.Paragraphs.Add();
                        materialRange = materialParagraph.Range;
                        materialRange.Text = suptype[2];
                        materialRange.InsertParagraphAfter();

                        tableParagraph = document.Paragraphs.Add();
                        tableRange = tableParagraph.Range;
                        materialTable = document.Tables.Add(tableRange, OOO.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование поставщика";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "ИНН";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Рейтинг";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (int j = 0; j < OOO.Count(); j++)
                        {

                            var currentsup = OOO[j];

                            cellRange = materialTable.Cell(j + 2, 1).Range;
                            cellRange.Text = currentsup.Title;

                            cellRange = materialTable.Cell(j + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentsup.INN);

                            cellRange = materialTable.Cell(j + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentsup.QualityRating);
                        }
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                        materialParagraph = document.Paragraphs.Add();
                        materialRange = materialParagraph.Range;
                        materialRange.Text = suptype[3];
                        materialRange.InsertParagraphAfter();

                        tableParagraph = document.Paragraphs.Add();
                        tableRange = tableParagraph.Range;
                        materialTable = document.Tables.Add(tableRange, MFO.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование поставщика";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "ИНН";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Рейтинг";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (int j = 0; j < MFO.Count(); j++)
                        {

                            var currentsup = MFO[j];

                            cellRange = materialTable.Cell(j + 2, 1).Range;
                            cellRange.Text = currentsup.Title;

                            cellRange = materialTable.Cell(j + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentsup.INN);

                            cellRange = materialTable.Cell(j + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentsup.QualityRating);
                        }
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                        materialParagraph = document.Paragraphs.Add();
                        materialRange = materialParagraph.Range;
                        materialRange.Text = suptype[4];
                        materialRange.InsertParagraphAfter();

                        tableParagraph = document.Paragraphs.Add();
                        tableRange = tableParagraph.Range;
                        materialTable = document.Tables.Add(tableRange, PAO.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование поставщика";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "ИНН";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Рейтинг";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (int j = 0; j < PAO.Count(); j++)
                        {

                            var currentsup = PAO[j];

                            cellRange = materialTable.Cell(j + 2, 1).Range;
                            cellRange.Text = currentsup.Title;

                            cellRange = materialTable.Cell(j + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentsup.INN);

                            cellRange = materialTable.Cell(j + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentsup.QualityRating);
                        }
                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                        materialParagraph = document.Paragraphs.Add();
                        materialRange = materialParagraph.Range;
                        materialRange.Text = suptype[5];
                        materialRange.InsertParagraphAfter();

                        tableParagraph = document.Paragraphs.Add();
                        tableRange = tableParagraph.Range;
                        materialTable = document.Tables.Add(tableRange, ZAO.Count() + 1, 3);
                        materialTable.Borders.InsideLineStyle = materialTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                        materialTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                        cellRange = materialTable.Cell(1, 1).Range;
                        cellRange.Text = "Наименование поставщика";
                        cellRange = materialTable.Cell(1, 2).Range;
                        cellRange.Text = "ИНН";
                        cellRange = materialTable.Cell(1, 3).Range;
                        cellRange.Text = "Рейтинг";
                        materialTable.Rows[1].Range.Bold = 1;
                        materialTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        for (int j = 0; j < ZAO.Count(); j++)
                        {

                            var currentsup = ZAO[j];

                            cellRange = materialTable.Cell(j + 2, 1).Range;
                            cellRange.Text = currentsup.Title;

                            cellRange = materialTable.Cell(j + 2, 2).Range;
                            cellRange.Text = Convert.ToString(currentsup.INN);

                            cellRange = materialTable.Cell(j + 2, 3).Range;
                            cellRange.Text = Convert.ToString(currentsup.QualityRating);
                        }

                        string path = Environment.CurrentDirectory;
                        document.SaveAs2(path + "Report5.docx");
                        document.SaveAs2(path + "Report5.pdf", Word.WdExportFormat.wdExportFormatPDF);
                        break;
                    }
                default:
                    {
                        MessageBox.Show("Возникла ошибка при создании отчёта", "Ошибка");
                        break;
                    }
            }
        }
    }
}
