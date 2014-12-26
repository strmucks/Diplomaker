using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Diplomaker
{
    public partial class MainWindow : Window
    {
        #region Fields and Properties

        Separator separator;
        DirectoryInfo directory;

        Typeface TextTypeFace
        {
            get
            {
                return new Typeface(textBox.FontFamily, textBox.FontStyle, textBox.FontWeight, textBox.FontStretch);
            }
        }
        #endregion

        #region Methods

        public MainWindow()
        {
            separator = new Separator() { Width = 435, Height = 10, BorderBrush = new SolidColorBrush() { Color = Color.FromArgb(255, 0, 0, 255) } };
            InitializeComponent();
        }

        //show image that user choose
        //return width for separator
        int ShowBackground()
        {
            var backgroundImage = InitImage(bcgdTextBox.Text, (int)preResultImage.Width, (int)preResultImage.Height);
            if (backgroundImage.UriSource != null)
            {
                preResultImage.Source = backgroundImage;
                return backgroundImage.PixelWidth;
            }

            return (int)separator.Width;
        }


        private BitmapImage InitImage(string backgroundUrl, int imageWidth = 0, int imageHeight = 0)
        {
            var tempBackground = new BitmapImage();
            var fileInfo = new FileInfo(backgroundUrl);
            if (!fileInfo.Exists)
            {
                MessageBox.Show("Verify that the specified file is still there!", "Error!");
                return tempBackground;
            }

            try
            {
                tempBackground.BeginInit();
                tempBackground.UriSource = new Uri(backgroundUrl);
                if (imageHeight != 0 && imageWidth != 0)
                {
                    tempBackground.DecodePixelWidth = imageWidth;
                    tempBackground.DecodePixelHeight = imageHeight;
                }

                tempBackground.EndInit();
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show("Unexpected error " + ex.Message, "Error!");
            }

            return tempBackground;
        }

        void ShowRangeElements()
        {
            columnComboBox.Visibility = Visibility.Visible;
            rangeTextBlock.Visibility = Visibility.Visible;
            rangeColumnTextBlock.Visibility = Visibility.Visible;
            rangeRowTextBlock.Visibility = Visibility.Visible;
            rowComboBox.Visibility = Visibility.Visible;
            rangeSeparator.Visibility = Visibility.Visible;
            fontButton.Visibility = Visibility.Visible;
        }

        //when the user has given a picture and text show preresult with random
        bool ShowPreResult()
        {
            var randomCellValue = GetDataFromExcel(true);
            if (randomCellValue == null)
            {
                return false;
            }

            var rightText = GetRightText(randomCellValue.First());
            var backgroundImage = InitImage(bcgdTextBox.Text, (int)preResultImage.Width, (int)preResultImage.Height);
            if (backgroundImage.UriSource != null)
            {
                preResultImage.Source = DrawImageAndText(backgroundImage, TextTypeFace, rightText, textBox.FontSize, separator.Margin.Top + 90);
                return true;
            }

            return false;
        }

        private RenderTargetBitmap DrawImageAndText(BitmapImage myBitmapImage, Typeface textTypeFace, string text, double size, double yOfText)
        {
            var visual = new DrawingVisual();
            var renderTarget = new RenderTargetBitmap(myBitmapImage.PixelWidth, myBitmapImage.PixelHeight, 96, 96, PixelFormats.Pbgra32);
            using (var drawingContext = visual.RenderOpen())
            {
                drawingContext.DrawImage(myBitmapImage, new Rect(0, 0, myBitmapImage.PixelWidth, myBitmapImage.PixelHeight));

                foreach (var txt in GetTextToDraw(myBitmapImage, textTypeFace, text, size))
                {
                    var xOfText = (myBitmapImage.PixelWidth - txt.Width) / 2;
                    drawingContext.DrawText(txt, new Point(xOfText, yOfText));
                    yOfText += txt.Height;
                }
            }

            renderTarget.Render(visual);
            return renderTarget;
        }

        //if surname+name+middlename is bigger than width of picture then move to another line
        //todo: I think it is not the best way to do this
        private IList<FormattedText> GetTextToDraw(BitmapImage myBitmapImage, Typeface textTypeFace, string text, double size)
        {
            var list = new List<FormattedText>();
            var fioFormatText = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, textTypeFace, size, Brushes.Black);
            var textArray = text.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            if (myBitmapImage.PixelWidth - fioFormatText.Width > 0)
            {
                list.Add(fioFormatText);
            }

            else if (textArray.Count() > 2)
            {
                var surFText = new FormattedText(textArray[0], CultureInfo.InvariantCulture, FlowDirection.LeftToRight, textTypeFace, size, Brushes.Black);
                var naMiFText = new FormattedText(textArray[1] + " " + textArray[2], CultureInfo.InvariantCulture, FlowDirection.LeftToRight, textTypeFace, size, Brushes.Black);
                if (myBitmapImage.PixelWidth - naMiFText.Width > 0)
                {
                    list.Add(surFText);
                    list.Add(naMiFText);
                }
                else
                {
                    var nFText = new FormattedText(textArray[1], CultureInfo.InvariantCulture, FlowDirection.LeftToRight, textTypeFace, size, Brushes.Black);
                    var mFText = new FormattedText(textArray[2], CultureInfo.InvariantCulture, FlowDirection.LeftToRight, textTypeFace, size, Brushes.Black);
                    list.Add(surFText);
                    list.Add(nFText);
                    list.Add(mFText);
                }
            }

            return list;
        }

        private void ShowSourceElements()
        {
            sourceButton.Visibility = Visibility.Visible;
            sourceTextBox.Visibility = Visibility.Visible;
        }

        public async Task MakeDiplomsAsync()
        {
            var textsForDiploms = GetDataFromExcel();
            if (textsForDiploms != null)
            {
                //create directory for results
                var dateNow = DateTime.Now;
                directory = Directory.CreateDirectory("Generated diploms " + dateNow.Day + "." + dateNow.Month + "." + dateNow.Year + " " + dateNow.Hour + "." + dateNow.Minute);
                var bitmapImage = InitImage(bcgdTextBox.Text);
                //for progressbar
                var step = (float)100 / textsForDiploms.Count();
                var percent = step;
                foreach (var text in textsForDiploms)
                {
                    progressBar.Progress = (int)percent;
                    var typeface = TextTypeFace;
                    var yOfText = separator.Margin.Top + 90;
                    var koeff = bitmapImage.PixelHeight / preResultImage.Height;
                    yOfText *= koeff;
                    var size = koeff * textBox.FontSize;
                    //another thread
                    bitmapImage.Freeze();
                    await Task.Run(() =>
                    {
                        var rightText = GetRightText(text);
                        var renderTarget = DrawImageAndText(bitmapImage, typeface, rightText, size, yOfText);
                        var jpeg = new JpegBitmapEncoder();
                        jpeg.Frames.Add(BitmapFrame.Create(renderTarget));
                        using (var resultImage = File.Create(directory + "/" + text + ".jpeg"))
                        {
                            jpeg.Save(resultImage);
                        }
                    });

                    if (percent < 100) percent += step;
                    else percent = 100;
                    progressBar.Progress = (int)percent;
                }
            }
        }

        private IEnumerable<string> GetDataFromExcel(bool isRandom = false)
        {
            var excelApplication = new Excel.Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };
            var excelWorkbooks = excelApplication.Workbooks;
            var fileInfo = new FileInfo(sourceTextBox.Text);
            if (!fileInfo.Exists)
            {
                DisposeExcelInstance(excelApplication, excelWorkbooks);
                MessageBox.Show("Verify that the specified file is still there!", "Error!");
                return null;
            }

            //working with COM-objects (Excel). For disposing instance of application
            var excelWorkbook = excelWorkbooks.Open(sourceTextBox.Text);
            var excelWorksheet = excelWorkbook.ActiveSheet;
            var usedRange = excelWorksheet.UsedRange;
            var usedRangeRows = usedRange.Rows;
            var excelCells = excelWorksheet.Cells;
            if (usedRangeRows.Count <= 1 || rowComboBox.SelectedIndex + 1 > usedRangeRows.Count)
            {
                DisposeExcelInstance(excelApplication, excelWorkbooks, excelWorkbook, excelWorksheet, usedRange, usedRangeRows, excelCells);
                MessageBox.Show("Specified file Microsoft Excel is empty or you select the wrong range!", "Error!");
                return null;
            }

            var result = new List<string>();
            if (isRandom)
            {
                var randomValue = new Random();
                var randomRow = randomValue.Next(rowComboBox.SelectedIndex + 1, usedRangeRows.Count);
                var randomCell = excelCells[randomRow, columnComboBox.SelectedIndex + 1];
                var randomCellValue = randomCell.Value;
                if (randomCellValue == null)
                {
                    DisposeExcelInstance(excelApplication, excelWorkbooks, excelWorkbook, excelWorksheet, usedRange, usedRangeRows, excelCells, randomCell, randomCellValue);
                    MessageBox.Show("Make sure the range does not contain blank cells!", "Error!");
                    return null;
                }
                else
                {
                    result.Add(randomCellValue);
                }

                randomCellValue = null;
                DisposeExcelInstance(excelApplication, excelWorkbooks, excelWorkbook, excelWorksheet, usedRange, usedRangeRows, excelCells, randomCell, randomCellValue);
            }
            else
            {
                var nextRowCellValue = string.Empty;
                var nextRowCell = usedRange;
                for (var row = rowComboBox.SelectedIndex + 1; row <= usedRangeRows.Count; row++)
                {
                    nextRowCell = excelCells[row, columnComboBox.SelectedIndex + 1];
                    nextRowCellValue = nextRowCell.Value;
                    result.Add(nextRowCellValue);
                }

                nextRowCellValue = null;
                DisposeExcelInstance(excelApplication, excelWorkbooks, excelWorkbook, excelWorksheet, usedRange, usedRangeRows, excelCells, nextRowCell, nextRowCellValue);
            }

            return result;
        }

        //working with morpher API for to determine the names and surnames from source string
        //see http://morpher.ru/WebService.aspx for more details
        string GetRightText(string sourceText)
        {
            try
            {
                var document = XElement.Load(@"http://api.morpher.ru/WebService.asmx/GetXml?s=" + sourceText);
                var fioNode = document.Elements().First(x => x.Name.LocalName.Contains("ФИО"));
                var surname = fioNode.Elements().First(x => x.Name.LocalName.Contains("Ф")).Value;
                var name = fioNode.Elements().First(x => x.Name.LocalName.Contains("И")).Value;
                var middle = fioNode.Elements().First(x => x.Name.LocalName.Contains("О")).Value;
                return surname + " " + name + " " + middle;

            }
            catch (WebException)
            {
                return sourceText;
            }
            catch (ArgumentNullException)
            {
                return sourceText;
            }
            catch (InvalidOperationException)
            {
                return sourceText;
            }
        }

        void DisposeExcelInstance(Excel.Application excelApplication, Excel.Workbooks excelWorkbooks, Excel.Workbook excelWorkbook = null, params object[] excelElements)
        {
            if (excelWorkbook != null)
            {
                excelWorkbook.Close(false, Type.Missing, Type.Missing);
            }

            excelWorkbooks.Close();
            excelApplication.Quit();
            foreach (var element in excelElements)
            {
                if (element != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(element);
                }
            }

            excelElements = null;
            if (excelWorkbook != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                excelWorkbook = null;
            }
            if (excelWorkbooks != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                excelWorkbooks = null;
            }
            if (excelApplication != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                excelApplication = null;
            }

            GC.Collect();
        }
        #endregion

        #region EventHandlers
        private void backgroundButton_Click(object sender, RoutedEventArgs e)
        {
            var backgroundDialog = new OpenFileDialog()
            {
                DefaultExt = ".jpg",
                Filter = "JPG|*.jpg;*.jpeg|BMP|*.bmp|GIF|*.gif|PNG|*.png|TIFF|*.tif;*.tiff|"
                    + "All Graphics Types|*.bmp;*.jpg;*.jpeg;*.png;*.tif;*.tiff",
                CheckPathExists = true,
                CheckFileExists = true
            };
            var dialogResult = backgroundDialog.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value)
            {
                var fileInfo = new FileInfo(backgroundDialog.FileName);
                if (!backgroundDialog.Filter.Contains(fileInfo.Extension.ToLower()))
                {
                    MessageBox.Show("Make sure that the selected image file!", "Error!");
                    return;
                }

                bcgdTextBox.Text = backgroundDialog.FileName;
                separator.Width = ShowBackground();
                preResultImage.MouseEnter += preResultImage_MouseEnter;
                preResultImage.MouseLeave += preResultImage_MouseLeave;
                ShowSourceElements();
            }
        }

        private void sourceButton_Click(object sender, RoutedEventArgs e)
        {
            var sourceDialog = new OpenFileDialog()
            {
                DefaultExt = ".xls",
                Filter = "Microsoft Excel Documents|*.xls;*.xlsx",
                CheckPathExists = true,
                CheckFileExists = true
            };

            var dialogResult = sourceDialog.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value)
            {
                var fileInfo = new FileInfo(sourceDialog.FileName);
                if (!sourceDialog.Filter.Contains(fileInfo.Extension.ToLower()))
                {
                    MessageBox.Show("Make sure that the selected Microsoft Excel file!", "Error!");
                    return;
                }

                sourceTextBox.Text = sourceDialog.FileName;
                ShowRangeElements();
            }
        }

        private void preResultImage_MouseEnter(object sender, MouseEventArgs e)
        {
            mainGrid.Children.Remove(separator);
            preResultImage.MouseMove += preResultImage_MouseMove;
            separator.SetValue(Grid.ColumnProperty, 1);
            separator.SetValue(Grid.RowProperty, 1);
            separator.SetValue(Grid.RowSpanProperty, 8);
            mainGrid.Children.Add(separator);

        }

        void preResultImage_MouseMove(object sender, MouseEventArgs e)
        {
            var y = e.GetPosition(mainGrid).Y;
            separator.Margin = new Thickness(0, y - 100, 0, mainGrid.Height - y - 25);
            preResultImage.PreviewMouseLeftButtonDown += preResultImage_PreviewMouseLeftButtonDown;
        }

        private void preResultImage_MouseLeave(object sender, MouseEventArgs e)
        {
            preResultImage.MouseMove -= preResultImage_MouseMove;
            mainGrid.Children.Remove(separator);
        }

        private void preResultImage_PreviewMouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            if (!string.IsNullOrEmpty(sourceTextBox.Text))
            {
                e.Handled = true;
                if (ShowPreResult())
                {
                    makeButton.Visibility = Visibility.Visible;
                    resetButton.Visibility = Visibility.Visible;
                    preResultImage.MouseMove -= preResultImage_MouseMove;
                    preResultImage.MouseEnter -= preResultImage_MouseEnter;
                    preResultImage.MouseLeave -= preResultImage_MouseLeave;
                    preResultImage.PreviewMouseLeftButtonDown -= preResultImage_PreviewMouseLeftButtonDown;
                    mainGrid.Children.Remove(separator);
                }
            }
        }

        private void fontButton_Click(object sender, RoutedEventArgs e)
        {
            var fontChooser = new FontChooser();
            fontChooser.Owner = this;

            fontChooser.SetPropertiesFromObject(textBox);
            fontChooser.PreviewSampleText = textBox.SelectedText;

            if (fontChooser.ShowDialog().Value)
            {
                fontChooser.ApplyPropertiesToObject(textBox);
                ShowPreResult();
            }
        }

        private async void makeButton_Click(object sender, RoutedEventArgs e)
        {
            mainWindow.IsEnabled = false;
            progressBar.Visibility = Visibility.Visible;
            await MakeDiplomsAsync();
            progressBar.Visibility = Visibility.Hidden;
            var messageResult = MessageBox.Show("Success! Open folder with generated diploms?", "What you want to do now?", MessageBoxButton.YesNo);
            if (messageResult == MessageBoxResult.Yes)
            {
                Process.Start(directory.FullName);
                mainWindow.Close();
            }

            mainWindow.IsEnabled = true;
        }

        private void resetButton_Click(object sender, RoutedEventArgs e)
        {
            ShowBackground();
            preResultImage.MouseEnter += preResultImage_MouseEnter;
            preResultImage.MouseLeave += preResultImage_MouseLeave;
        }
        #endregion
    }
}