using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media.Imaging;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Wordprocessing;


namespace DragDropTreeApp
{
    public class DocItem
    {
        public string Title { get; set; }
        public string ImagePath { get; set; }
        public ObservableCollection<DocItem> Children { get; set; } = new();
    }

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Word Documents|*.docx"
            };

            if (dlg.ShowDialog() == true)
            {
                string path = dlg.FileName;

                var items = new ObservableCollection<DocItem>();

                using var wordDoc = WordprocessingDocument.Open(path, false);

                // Thêm các paragraph
                var paras = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>();
                int paraIndex = 1;
                foreach (var para in paras)
                {
                    string text = string.Join("", para.Descendants<Text>().Select(t => t.Text).ToArray()).Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        items.Add(new DocItem
                        {
                            //Title = $"📝 {text}"
                            Title = $"{text}"
                        });
                        paraIndex++;
                    }
                }

                // Thêm các ảnh
                int count = 1;
                foreach (var imgPart in wordDoc.MainDocumentPart.ImageParts)
                {
                    using var stream = imgPart.GetStream();
                    using var image = System.Drawing.Image.FromStream(stream);

                    string tempFolder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MyDocImages");
                    Directory.CreateDirectory(tempFolder);

                    string imgPath = System.IO.Path.Combine(tempFolder, $"img_{count}.png");
                    image.Save(imgPath, System.Drawing.Imaging.ImageFormat.Png);

                    items.Add(new DocItem
                    {
                        // Title = $"🖼️ Ảnh {count}",
                        ImagePath = imgPath
                    });

                    count++;
                }

                TreeDoc.ItemsSource = items;
            }
        }


    }

    public class ImagePathToImageSourceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string path = value as string;
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.UriSource = new Uri(path, UriKind.Absolute);
                    bitmap.EndInit();
                    return bitmap;
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class NullToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return string.IsNullOrEmpty(value as string) ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}



