using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// myDocumentViewer.xaml 的互動邏輯
    /// </summary>
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;
        Color backgroundColor = Colors.Transparent;
        public MyDocumentViewer()
        {
            InitializeComponent();
            foreach (FontFamily fontfamily in Fonts.SystemFontFamilies)
            {
                comFontFamily.Items.Add(fontfamily);
            }
            comFontFamily.SelectedIndex = 0;

            comFontSize.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 32, 40, 50, 60, 80, 100 };

            comFontSize.SelectedIndex = 4;

            fontColorPicker.SelectedColor = fontColor;

            backgroundColorPicker.SelectedColor = backgroundColor;
        }

        private void RtbEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object property = rtbEditor.Selection.GetPropertyValue(Inline.FontWeightProperty);
            btnBold.IsChecked = (property != DependencyProperty.UnsetValue) && (property.Equals(FontWeights.Bold));

            property = rtbEditor.Selection.GetPropertyValue(Inline.FontStyleProperty);
            btnItalic.IsChecked = (property != DependencyProperty.UnsetValue) && (property.Equals(FontStyles.Italic));

            property = rtbEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            btnUnderline.IsChecked = (property != DependencyProperty.UnsetValue) && (property.Equals(TextDecorations.Underline));

            property = rtbEditor.Selection.GetPropertyValue(Inline.FontFamilyProperty);
            comFontFamily.SelectedItem = property;

            property = rtbEditor.Selection.GetPropertyValue(Inline.FontSizeProperty);
            comFontSize.SelectedItem = property;

            SolidColorBrush? foregroundProperty = rtbEditor.Selection.GetPropertyValue(Inline.ForegroundProperty) as SolidColorBrush;
            fontColorPicker.SelectedColor = foregroundProperty.Color;

            SolidColorBrush? backgroundProperty = rtbEditor.Selection.GetPropertyValue(Inline.BackgroundProperty) as SolidColorBrush;
            if (backgroundProperty != null)
                backgroundColorPicker.SelectedColor = backgroundProperty.Color;
            else
                backgroundColorPicker.SelectedColor = Colors.Transparent;

        }

        private void ComFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comFontFamily.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, comFontFamily.SelectedItem);
            }
        }

        private void ComFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comFontSize.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, comFontSize.SelectedItem);
            }
        }

        private void FontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = (Color)e.NewValue;
            SolidColorBrush colorBrush = new SolidColorBrush(fontColor);
            rtbEditor.Selection.ApplyPropertyValue(Inline.ForegroundProperty, colorBrush);
        }

        private void BackgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            backgroundColor = (Color)e.NewValue;
            SolidColorBrush colorBrush = new SolidColorBrush(backgroundColor);
            rtbEditor.Selection.ApplyPropertyValue(Inline.BackgroundProperty, colorBrush);
        }

        private void Open_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Rich Text Format file|*.ref|HTML Text Format file|*.html|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                FileStream fs = new FileStream(fileDialog.FileName, FileMode.Open);
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Load(fs, DataFormats.Rtf);
            }
        }

        private void Save_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Rich Text Format file|*.ref|HTML Text Format file|*.html|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                FileStream fs = new FileStream(fileDialog.FileName, FileMode.Create);
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);
                range.Save(fs, DataFormats.Rtf);
            }
        }

        private void New_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            rtbEditor.Document.Blocks.Clear();
        }
    }
}
