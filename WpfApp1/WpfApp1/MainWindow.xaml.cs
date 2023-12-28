using ExportExcelFile;
using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // ファイル保存ダイアログを生成します。
            var dialog = new SaveFileDialog();

            // フィルターを設定します。
            // この設定は任意です。
            dialog.Filter = "Excelファイル(*.xlsx)|*.xlsx";

            // ファイル保存ダイアログを表示します。
            var result = dialog.ShowDialog() ?? false;

            // 保存ボタン以外が押下された場合
            if (!result)
            {
                // 終了します。
                return;
            }

            ExportXlsx(dialog.FileName);
        }

        private void ExportXlsx(string path)
        {
            NameCardParams nameCardParams = new(NameCardType.AAA, System.Drawing.Color.Blue, "AAAAAAAA AAAAAAAAA");

            NameCard nameCard = new(nameCardParams, path);
            var result = nameCard.ExportXlsx();
            if(result != ExportErrorType.None)
            {
                // Error
            }
        }
    }
}