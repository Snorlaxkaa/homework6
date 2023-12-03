using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

namespace _20231124
{
    public partial class MyDocumentViewer : Window
    {
        Color fontColor = Colors.Black;

        public MyDocumentViewer()
        {
            InitializeComponent();
            fontColorPicker.SelectedColor = fontColor;

            foreach (FontFamily fontFamily in Fonts.SystemFontFamilies)
            {
                fontFamilyComboBox.Items.Add(fontFamily);
            }
            fontFamilyComboBox.SelectedIndex = 8;

            fontSizeComboBox.ItemsSource = new List<Double>() { 8, 9, 10, 12, 20, 24, 32, 40, 50, 60, 80, 90 };
            fontSizeComboBox.SelectedIndex = 3;
            backgroundColorPicker.SelectedColor = ((SolidColorBrush)rtbEditor.Background).Color;
        }

        private void NewCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            MyDocumentViewer myDocumentViewer = new MyDocumentViewer();
            myDocumentViewer.Show();
        }

        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Rich Text Format檔案|*.rtf|HTML檔案|*.html|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);

                using (FileStream fileStream = new FileStream(fileDialog.FileName, FileMode.Open))
                {
                    if (Path.GetExtension(fileDialog.FileName).Equals(".rtf", StringComparison.OrdinalIgnoreCase))
                    {
                        range.Load(fileStream, DataFormats.Rtf);
                    }
                    else if (Path.GetExtension(fileDialog.FileName).Equals(".html", StringComparison.OrdinalIgnoreCase))
                    {
                        range.Load(fileStream, DataFormats.Html);
                    }
                }
            }
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Filter = "Rich Text Format檔案|*.rtf|HTML檔案|*.html|所有檔案|*.*";
            if (fileDialog.ShowDialog() == true)
            {
                TextRange range = new TextRange(rtbEditor.Document.ContentStart, rtbEditor.Document.ContentEnd);

                if (Path.GetExtension(fileDialog.FileName).Equals(".rtf", StringComparison.OrdinalIgnoreCase))
                {
                    using (FileStream fileStream = new FileStream(fileDialog.FileName, FileMode.Create))
                    {
                        range.Save(fileStream, DataFormats.Rtf);
                    }
                }
                else if (Path.GetExtension(fileDialog.FileName).Equals(".html", StringComparison.OrdinalIgnoreCase))
                {
                    using (FileStream fileStream = new FileStream(fileDialog.FileName, FileMode.Create))
                    {
                        range.Save(fileStream, DataFormats.Html);
                    }
                }
            }
        }

        private void rtbEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object property = rtbEditor.Selection.GetPropertyValue(TextElement.FontWeightProperty);
            boldButton.IsChecked = (property is FontWeight && (FontWeight)property == FontWeights.Bold);

            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontStyleProperty);
            italicButton.IsChecked = (property is FontStyle && (FontStyle)property == FontStyles.Italic);

            property = rtbEditor.Selection.GetPropertyValue(Inline.TextDecorationsProperty);
            underlineButton.IsChecked = (property != DependencyProperty.UnsetValue && property == TextDecorations.Underline);

            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontFamilyProperty);
            fontFamilyComboBox.SelectedItem = property;

            property = rtbEditor.Selection.GetPropertyValue(TextElement.FontSizeProperty);
            fontSizeComboBox.SelectedItem = property;

            SolidColorBrush? foregroundProperty = rtbEditor.Selection.GetPropertyValue(TextElement.ForegroundProperty) as SolidColorBrush;
            if (foregroundProperty != null)
            {
                fontColorPicker.SelectedColor = foregroundProperty.Color;
            }
        }

        private void fontColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
            fontColor = (Color)e.NewValue;
            SolidColorBrush fontBrush = new SolidColorBrush(fontColor);
            rtbEditor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, fontBrush);
        }

        private void fontFamilyComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontFamilyComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontFamilyProperty, fontFamilyComboBox.SelectedItem);
            }
        }

        private void fontSizeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (fontSizeComboBox.SelectedItem != null)
            {
                rtbEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeComboBox.SelectedItem);
            }
        }

        private void clearButton_Click(object sender, RoutedEventArgs e)
        {
            rtbEditor.Document.Blocks.Clear();
        }

        private void backgroundColorPicker_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<Color?> e)
        {
           Color? selectedColor = e.NewValue;
            if (selectedColor != null)
            {
                Color backgroundColor = (Color)selectedColor;
                SolidColorBrush backgroundBrush = new SolidColorBrush(backgroundColor);
                rtbEditor.Background = backgroundBrush;
            }
        }
    }
}
