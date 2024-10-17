using Microsoft.Win32;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace Receipts_Distribution
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow:Window
    {
        private string csvFilePath = string.Empty;
        private string masterFilePath = string.Empty;
        private string correctionFilePath = string.Empty;
        private string orderFilePath = string.Empty;

        // DispatcherTimer to simulate progress
        private DispatcherTimer _progressTimer;
        private int _progressValue = 0;

        private bool isCsvFileValid = false;           //validation flag for 点検結果CSV
        private bool isMasterFileValid = false;        //validation flag for 依頼マスタ
        private bool isCorrectionFileValid = false;    //validation flag for 修正指示書
        private bool isOrderFileValid = false;         //validation flag for 点発注書

        //private bool isMasterUploadedOnce = false;      // Tracks if 依頼マスタ has been uploaded at least once
        // Track if files have been uploaded at least once
        private bool isCsvFileUploadedOnce = false;
        private bool isMasterFileUploadedOnce = false;
        private bool isCorrectionFileUploadedOnce = false;
        private bool isOrderFileUploadedOnce = false;

        private CheckFilePath _checkFilePath;
        private CsvToExcel _csvToExcel;
        private RequestMaster _requestMaster;
        private ReceiptsDistributionTask _receiptsDistributionTask;
        private Errors _errors;
        private Logs _logs;

        private BackgroundWorker excelWorker;
        private BackgroundWorker progressWorker;
        private bool hasError = false;

        public MainWindow()
        {
            InitializeComponent();

            ResetUIElements();
            // Initialize the status text as empty
            TextBlock_Status.Text = "";  　　　　　　　　// This ensures no status is shown at the start

            // Initializing classes
            _checkFilePath = new CheckFilePath();
            _csvToExcel = new CsvToExcel();
            _requestMaster = new RequestMaster();
            _receiptsDistributionTask = new ReceiptsDistributionTask();
            _errors = new Errors();
            _logs = new Logs();

            // Initialize progress timer (for simulating progress)
            _progressTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100) // Set timer interval
            };
            _progressTimer.Tick += UpdateProgressSimulation;

            excelWorker = new BackgroundWorker(); // BackgroundWorkerのインスタンスを作成
            excelWorker.WorkerReportsProgress = true; // 進捗報告を有効にする
            excelWorker.WorkerSupportsCancellation = true; // キャンセルをサポートする
                                                           // イベントの設定
            excelWorker.DoWork += ExcelWorker_DoWork;
            excelWorker.RunWorkerCompleted += Worker_RunWorkerCompleted;

            progressWorker = new BackgroundWorker();
            progressWorker.WorkerReportsProgress = true;
            progressWorker.DoWork += ProgressWorker_DoWork;
            progressWorker.ProgressChanged += ProgressWorker_ProgressChanged;
            progressWorker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        public void Log(string message)
        {
            string timestampedMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}";

            TextBox_LogOutput.AppendText($"{timestampedMessage}\n");
            TextBox_LogOutput.ScrollToEnd();  // Auto-scroll to the bottom of the log
            _logs.Log(message);  // Call Logs method to save to file
        }

        // Method to reset the UI elements like Status, Progress Bar, etc.
        private void ResetUIElements()
        {
            // Reset the TextBlock_Status, ProgressBar, and ProgressPercentage to initial values
            TextBlock_Status.Text = "";  // Status will be empty initially
            ProgressBar_Color.Value = 0;  // Reset progress bar
            TextBlock_ProgressPercentage.Text = "進捗率: 0%";  // Reset progress percentage
        }

        // File Selection for 点検結果(CSV)
        private void Button_UploadCSVClick(object sender, RoutedEventArgs e)
        {
            // Check if this is a re-upload
            if( isCsvFileUploadedOnce )
            {
                ResetUIElements();  // Reset UI elements when re-uploading
            }

            // Before selecting the file, show "設定中" status
            TextBlock_Status.Text = "設定中";
            Log("点検結果ファイル 設定中...");


            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            //openFileDialog.Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx";

            try
            {

                if( openFileDialog.ShowDialog() == true )
                {
                    // Set this flag when the user re-uploads 点検結果
                    //isSecondUpload = true;

                    // Clear the textboxes for 点検結果(CSV) if this is the second upload
                    TextBox_CSVFilePath.Text = string.Empty;

                    // Get the selected CSV file path
                    csvFilePath = openFileDialog.FileName;

                    // Validate the CSV file using CheckFilePat //（仮コード：CSVファイルではありませんのエラー表示チェックため）
                    Log("CSVファイルかどうかのチェックを開始します...");
                    string validationError = _checkFilePath.CheckCsvFilePath(csvFilePath);
                    Log("CSVファイルかどうかのチェックが完了しました。");


                    if( !string.IsNullOrEmpty(validationError) )
                    {
                        // Log and update UI for the validation error
                        Log(validationError);  // Log the validation error
                        TextBox_CSVFilePath.Foreground = Brushes.Red;  // Set text color to red for error
                        TextBox_CSVFilePath.Text = validationError;  // Display the validation error
                        HandleFileError("点検結果(CSV)", validationError);
                        Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226));   //light blue
                        return;
                    }

                    Log("CSVファイルのExcelへの変換を開始します...");
                    //_csvToExcel.ChangeInspectionResultCsvToExcel(csvFilePath);                                                      // This will execute the CSV to Excel conversion process.
                    string CSVtoExcelChange = "C:\\Users\\24-N10102Zuser\\Documents\\ReceiptsDistributionSetting\\0902点検結果";
                    Log($"{CSVtoExcelChange}");                                                                 // Log the file path update success
                    Log("CSVファイルのExcelへの変換が完了しました。");

                    // Change button color to gray after successful upload
                    //Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                    Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));   // Light yellow

                    //点検結果(CSV)ファイル設定完了
                    TextBox_CSVFilePath.Text = csvFilePath; // Update the correct text box for CSV file path
                    TextBox_CSVFilePath.Foreground = Brushes.Black;
                    Log("点検結果(CSV)ファイル 設定完了（ファイルパス設定完了）");
                    Log($"{CSVtoExcelChange}");

                    // Set validation flag
                    isCsvFileValid = true; // Set the validation flag to true if valid
                    EnableDistributionButtonIfValid(); // Re-check if all files are valid

                }
            }
            catch( FileNotFoundException )
            {
                Log(string.Format(_errors.GetErrorMessage("E04"), "点検結果(CSV)"));
            }
            catch( PathTooLongException )
            {
                Log(string.Format(_errors.GetErrorMessage("E05"), "点検結果(CSV)"));
            }
        }

        // File Selection for 依頼マスタ
        private void Button_UploadMasterRequestClick(object sender, RoutedEventArgs e)
        {
            if( isMasterFileUploadedOnce )
            {
                ResetUIElements();  // Reset UI elements when re-uploading
            }

            TextBlock_Status.Text = "設定中";
            Log("依頼マスタファイル 設定中...");

            OpenFileDialog openFileDialog = new OpenFileDialog();

            try
            {
                if( openFileDialog.ShowDialog() == true )
                {
                    // Clear the textboxes for 依頼マスタ if this is the second upload
                    TextBox_MasterFilePath.Text = string.Empty;

                    // Get the selected 依頼マスタ file path
                    masterFilePath = openFileDialog.FileName;

                    Log("依頼マスタファイルかどうかチェックを開始します...");
                    // Validate the Master Request file using CheckFilePath
                    string validationError = _checkFilePath.CheckMasterFilePath(masterFilePath);
                    Log("依頼マスタファイのチェックが完了しました。");

                    if( !string.IsNullOrEmpty(validationError) )
                    {
                        Log(validationError);  // Log the validation error
                        TextBox_MasterFilePath.Foreground = Brushes.Red;  // Set text color to red for error
                        TextBox_MasterFilePath.Text = validationError;  // Display the validation error
                        HandleFileError("依頼マスタ", validationError);
                        Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226));   //light blue
                        return;
                    }

                    Log("依頼マスタファイルパスが選択されました。");
                    // Log the file path update success
                    Log($"{masterFilePath}");
                    Log("依頼マスタファイルパスが完了です。");

                    // Change button color to gray after successful upload
                    //Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                    Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));

                    //依頼マスタファイル設定完了                    
                    TextBox_MasterFilePath.Text = masterFilePath; // Update the correct text box for 依頼マスタ file path
                    TextBox_MasterFilePath.Foreground = Brushes.Black;
                    Log("依頼マスタファイル 設定完了（ファイルパス設定完了）");

                    // Set validation flag
                    isMasterFileValid = true; // Set the validation flag to true if valid
                    EnableDistributionButtonIfValid(); // Re-check if all files are valid

                    // Check if this is the second upload (ファイルアップロードが2回目以降の場合)
                    if( isMasterFileUploadedOnce )
                    {
                        // If it's a second upload, change 修正指示書 and 発注書 buttons to gray
                        //Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                        //Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                        Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));　　//light yellow
                        Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));　  　//light yellow
                    }
                    else
                    {
                        // Mark as first upload done
                        isMasterFileUploadedOnce = true;
                    }
                }
            }
            catch( FileNotFoundException )
            {
                Log(string.Format(_errors.GetErrorMessage("E04"), "依頼マスタ"));
            }
            catch( PathTooLongException )
            {
                Log(string.Format(_errors.GetErrorMessage("E05"), "依頼マスタ"));
            }
        }


        // File Selection for 修正指示書
        private void Button_UploadCorrectionClick(object sender, RoutedEventArgs e)
        {
            if( isCorrectionFileUploadedOnce )
            {
                ResetUIElements();  // Reset UI elements when re-uploading
            }

            TextBlock_Status.Text = "設定中";
            Log("修正指示書ファイル 設定中...");

            OpenFileDialog openFileDialog = new OpenFileDialog();

            try
            {
                if( openFileDialog.ShowDialog() == true )
                {
                    correctionFilePath = openFileDialog.FileName;

                    Log("修正指示書ファイルかどうかチェックを開始します...");
                    // Validate the Correction file using CheckFilePath
                    string validationError = _checkFilePath.checkCorrectionFilePath(correctionFilePath);
                    Log("修正指示書ファイルのチェックが完了しました。");


                    if( !string.IsNullOrEmpty(validationError) )
                    {
                        Log(validationError);  // Log the validation error
                        TextBox_CorrectionFilePath.Foreground = Brushes.Red;  // Set text color to red for error
                        TextBox_CorrectionFilePath.Text = validationError;  // Display the validation error
                        HandleFileError("修正指示書", validationError);
                        Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226));   //light blue
                        return;
                    }

                    Log("修正指示書ファイルパスが選択されました。");
                    // Log the file path update success
                    Log($"{correctionFilePath}");
                    Log("修正指示書ファイルパスが完了です。");

                    // Change button color to gray after successful upload
                    //Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                    Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));

                    //修正指示書ファイル設定完了
                    TextBox_CorrectionFilePath.Text = correctionFilePath; // Update the correct text box for 修正指示書 file path
                    TextBox_CorrectionFilePath.Foreground = Brushes.Black;
                    Log("修正指示書ファイルが選択されました。");

                    // Set validation flag
                    isCorrectionFileValid = true; // Set the validation flag to true if valid
                    EnableDistributionButtonIfValid(); // Re-check if all files are valid
                }
            }
            catch( FileNotFoundException )
            {
                Log(string.Format(_errors.GetErrorMessage("E04"), "修正指示書"));
            }
            catch( PathTooLongException )
            {
                Log(string.Format(_errors.GetErrorMessage("E05"), "修正指示書"));
            }
        }

        // File Selection for 発注書
        private void Button_UploadPurchaseOrderClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //CheckFilePath checkFilePath = new CheckFilePath();
            //Errors errors = new Errors();  // Error handling
            //Logs logs = new Logs();  // Assuming P's log class

            try
            {
                if( isOrderFileUploadedOnce )
                {
                    ResetUIElements();  // Reset UI elements when re-uploading
                }

                TextBlock_Status.Text = "設定中";
                Log("発注書ファイル 設定中...");

                if( openFileDialog.ShowDialog() == true )
                {
                    orderFilePath = openFileDialog.FileName;

                    // Validate the Purchase Order file using CheckFilePath
                    string validationError = _checkFilePath.checkPurchaseOrderFilePath(orderFilePath);

                    if( !string.IsNullOrEmpty(validationError) )
                    {
                        Log(validationError);  // Log the validation error
                        TextBox_PuchaseOrderFilePath.Foreground = Brushes.Red;  // Set text color to red for error
                        TextBox_PuchaseOrderFilePath.Text = validationError;  // Display the validation error
                        HandleFileError("発注書", validationError);
                        Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226));     //light blue
                        return;
                    }

                    Log("発注書ファイルパスが選択されました。");
                    // Log the file path update success
                    Log($"{orderFilePath}");
                    Log("発注書ファイルパスが完了です。");

                    // Change button color to gray after successful upload
                    //Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179));
                    Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));

                    //発注書ファイル設定完了
                    TextBox_PuchaseOrderFilePath.Text = orderFilePath; // Update the correct text box for 発注書 file path
                    TextBox_PuchaseOrderFilePath.Foreground = Brushes.Black;
                    Log("発注書ファイルが選択されました。");

                    // Set validation flag
                    isCorrectionFileValid = true; // Set the validation flag to true if valid
                    EnableDistributionButtonIfValid(); // Re-check if all files are valid

                }
            }
            catch( FileNotFoundException )
            {
                Log(string.Format(_errors.GetErrorMessage("E04"), "発注書"));
            }
            catch( PathTooLongException )
            {
                Log(string.Format(_errors.GetErrorMessage("E05"), "発注書"));
            }
        }

        private void HandleFileError(string fileType, string errorMessage)
        {
            TextBlock_Status.Text = "エラー";
            Log($"{fileType} エラー: {errorMessage}");

            // Stop progress bar if an error occurs
            _progressTimer.Stop();
            UpdateProgress(0); // Reset the progress bar
        }


        // Enable the Execute button only if all necessary files are valid　（振り分け実行ボタン）
        private void EnableDistributionButtonIfValid()
        {
            // Check if all required files (点検結果, 依頼マスタ, 修正指示書, 発注書) are uploaded
            bool areAllFilesUploaded = !string.IsNullOrEmpty(csvFilePath) &&
                                       !string.IsNullOrEmpty(masterFilePath) &&
                                       !string.IsNullOrEmpty(correctionFilePath) &&
                                       !string.IsNullOrEmpty(orderFilePath);

            // Check if any TextBox is showing an error (for example, if the text is red or contains an error message)
            bool hasErrors = HasErrorsInTextBoxes();

            // Check if all required files (点検結果, 依頼マスタ, 修正指示書, 発注書) are uploaded
            if( areAllFilesUploaded && !hasErrors )
            {
                // Set status to 設定完了 only when all files are uploaded
                TextBlock_Status.Text = "設定完了";
                _logs.Log("設定完了（ファイルパス設定完了）");

                // Enable the 振り分け実行 button only if all files are uploaded
                Button_Distribution.IsEnabled = true;
                Button_Distribution.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204)); // Light yellow
                Log("すべてのファイルが選択されました。振り分け実行ボタンが有効化されています。");

            }
            else
            {
                // Keep the 振り分け実行 button disabled when not all required files are uploaded
                Button_Distribution.IsEnabled = false;
                Button_Distribution.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Gray
            }

        }

        // Method to check if any of the TextBox elements contain an error (e.g., red text)
        private bool HasErrorsInTextBoxes()
        {
            // Assume errors are displayed with red text in the TextBoxes
            return TextBox_CSVFilePath.Foreground == Brushes.Red ||
                   TextBox_MasterFilePath.Foreground == Brushes.Red ||
                   TextBox_CorrectionFilePath.Foreground == Brushes.Red ||
                   TextBox_PuchaseOrderFilePath.Foreground == Brushes.Red;
        }

        // Handle the Execute button click and start the progress simulation 振り分け実行ボタン
        private void OnExecuteClick(object sender, RoutedEventArgs e)
        {

            //if( !backgroundWorker.IsBusy ) // 処理中でない場合
            //{
            //    TextBlock_ProgressPercentage.Text = "進捗率: 0%";
            //    TextBlock_Status.Text = "実行中";
            //    _logs.Log("振り分け実行開始...");
            //    backgroundWorker.RunWorkerAsync();// 非同期処理を開始   
            //}


            // Excel処理を開始
            excelWorker.RunWorkerAsync();

            // 進捗バーを更新する
            progressWorker.RunWorkerAsync();


            //ExcelWorker_DoWorkに移動
            //TextBlock_ProgressPercentage.Text = "進捗率: 0%";
            //TextBlock_Status.Text = "実行中";
            //_logs.Log("振り分け実行開始...");

            //_progressValue = 0;
            //ProgressBar_Color.Value = 0;
            //_progressTimer.Start(); // Start simulating progress

            //try
            //{
            //    if( string.IsNullOrEmpty(csvFilePath) ||
            //        string.IsNullOrEmpty(masterFilePath) ||
            //        string.IsNullOrEmpty(correctionFilePath) ||
            //        string.IsNullOrEmpty(orderFilePath) )
            //    {
            //        throw new FileNotFoundException(_errors.GetErrorMessage("E04"));
            //    }

            //    // Change the background color of Button_Distribution when clicked
            //    Button_Distribution.Background = new SolidColorBrush(Color.FromRgb(173, 216, 230)); // Light blue
            //    Button_Distribution.IsEnabled = false; // Disable the 振り分け実行 button after clicked

            //    // Disable the file upload buttons after 振り分け実行 is clicked
            //    Button_UploadCSV.IsEnabled = false;
            //    Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

            //    Button_UploadMasterRequest.IsEnabled = false;
            //    Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

            //    Button_UploadCorrection.IsEnabled = false;
            //    Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

            //    Button_UploadPurchase.IsEnabled = false;
            //    Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

            //    //string inspectionResultExcelFilePath = _csvToExcel.ChangeInspectionResultCsvToExcel(csvFilePath);
            //    string inspectionResultExcelFilePath = "C:\\Users\\24-N10100Zuser\\Documents\\ReceiptsDistributionSetting\\1016\\1016点検結果.xlsx";
            //    string requestMasterFilePath = _requestMaster.CalculateNumberOfRequests(masterFilePath, inspectionResultExcelFilePath);
            //    _receiptsDistributionTask.CreateReceiptsDistributionTask(inspectionResultExcelFilePath, requestMasterFilePath, correctionFilePath, orderFilePath);
            //}
            //catch( FileNotFoundException ex )
            //{
            //    _logs.Log(ex.Message);
            //}
        }
        private void ExcelWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //_progressValue = 0;
            //ProgressBar_Color.Value = 0;
            //_progressTimer.Start(); // Start simulating progress

            try
            {
                if( string.IsNullOrEmpty(csvFilePath) ||
                    string.IsNullOrEmpty(masterFilePath) ||
                    string.IsNullOrEmpty(correctionFilePath) ||
                    string.IsNullOrEmpty(orderFilePath) )
                {
                    throw new FileNotFoundException(_errors.GetErrorMessage("E04"));
                }

                Application.Current.Dispatcher.Invoke(() =>
                {
                    // Change the background color of Button_Distribution when clicked
                    Button_Distribution.Background = new SolidColorBrush(Color.FromRgb(173, 216, 230)); // Light blue
                    Button_Distribution.IsEnabled = false; // Disable the 振り分け実行 button after clicked

                    // Disable the file upload buttons after 振り分け実行 is clicked
                    Button_UploadCSV.IsEnabled = false;
                    Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

                    Button_UploadMasterRequest.IsEnabled = false;
                    Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

                    Button_UploadCorrection.IsEnabled = false;
                    Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray

                    Button_UploadPurchase.IsEnabled = false;
                    Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)); // Change to gray
                });
                //string inspectionResultExcelFilePath = _csvToExcel.ChangeInspectionResultCsvToExcel(csvFilePath);

                DateTime today = DateTime.Now;
                string folderName = today.ToString("MMdd");
                string defaultExcelFilePath = Properties.Settings.Default.excelfilePath;
                //同日で二回処理を実行の場合はディレクトリの中の全てのファイルの削除してファイルを新規作成
                //余裕ファイル・シート発生しないように
                DirectoryInfo di = new DirectoryInfo(string.Format("{0}{1}", defaultExcelFilePath, folderName));
                //ファイル消す
                foreach( FileInfo file in di.GetFiles() )
                {
                    file.Delete();
                }
                //フォルダも消す
                foreach( DirectoryInfo dir in di.GetDirectories() )
                {
                    dir.Delete(true);
                }
                string inspectionResultExcelFilePath = "C:\\Users\\24-N10100Zuser\\Documents\\ReceiptsDistributionSetting\\1016\\1016点検結果.xlsx";

                string requestMasterFilePath = _requestMaster.CalculateNumberOfRequests(masterFilePath, inspectionResultExcelFilePath);
                _receiptsDistributionTask.CreateReceiptsDistributionTask(inspectionResultExcelFilePath, requestMasterFilePath, correctionFilePath, orderFilePath);
            }
            catch( FileNotFoundException ex )
            {
                _logs.Log(ex.Message);
                hasError = true;
            }
        }

        private void ProgressWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            for( int i = 0; i <= 100; i++ )
            {
                System.Threading.Thread.Sleep(1800); // 模擬的な処理時間
                progressWorker.ReportProgress(i);
            }
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if( hasError )
            {
                // エラーが発生した場合、すべての処理を停止
                //Application.Current.Shutdown(); // プログラムを終了
                return;
            }
        }

        private void ProgressWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 進捗状況を取得
            int progressPercentage = e.ProgressPercentage;

            // プログレスバーの更新
            _progressValue = e.ProgressPercentage; // プログレスバーを更新
                                                   // 進捗メッセージの表示
            TextBlock_Status.Text = $"進捗: {e.ProgressPercentage}%"; // ステータスラベルを更新
            TextBlock_ProgressPercentage.Text = $"進捗: {e.ProgressPercentage}%";
        }

        // Simulate the progress updates in the progress bar
        private void UpdateProgressSimulation(object sender, EventArgs e)
        {
            _progressValue += 5; // Increment progress by 5% for demonstration

            if( _progressValue >= 100 )
            {
                _progressValue = 100; // Cap progress at 100%
                UpdateProgress(_progressValue); // Ensure the progress is visually updated

                TextBlock_Status.Text = "実行完了";
                Log("実行完了: 全ての処理が完了しました。");
                _progressTimer.Stop(); // Stop the progress timer

                // Re-enable the buttons after completion
                Button_UploadCSV.IsEnabled = true;
                Button_UploadCSV.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226)); // Back to light blue

                Button_UploadMasterRequest.IsEnabled = true;
                Button_UploadMasterRequest.Background = new SolidColorBrush(Color.FromRgb(165, 187, 226)); // Back to light blue

                Button_UploadCorrection.IsEnabled = true;
                Button_UploadCorrection.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));   //  remains yellow

                Button_UploadPurchase.IsEnabled = true;
                Button_UploadPurchase.Background = new SolidColorBrush(Color.FromRgb(255, 255, 204));     //  remains yellow

                // Clear 点検結果(CSV) and 依頼マスタ textboxes
                TextBox_CSVFilePath.Text = "点検結果csvファイルをアップロードしてください。";
                TextBox_CSVFilePath.Foreground = Brushes.Gray;
                csvFilePath = string.Empty;

                TextBox_MasterFilePath.Text = "依頼マスタファイルをアップロードしてください。";
                TextBox_MasterFilePath.Foreground = Brushes.Gray;
                masterFilePath = string.Empty;

                // Reset validation flags for 点検結果 and 依頼マスタ
                isCsvFileValid = false;
                isMasterFileValid = false;

                // Disable 振り分け実行 button
                Button_Distribution.IsEnabled = false;
                Button_Distribution.Background = new SolidColorBrush(Color.FromRgb(179, 179, 179)); // Gray
            }

            UpdateProgress(_progressValue);
        }

        // Update progress bar and progress text dynamically
        private void UpdateProgress(int progress)
        {
            ProgressBar_Color.Value = progress;
            TextBlock_ProgressPercentage.Text = $"進捗率: {progress}%";

            // If progress is 100%, update the status and log
            if( progress == 100 )
            {
                TextBlock_Status.Text = "実行完了";
                Log("実行完了（全て完了）");

                // Set a timer to delay resetting the UI so that the user can see the "実行完了" message
                DispatcherTimer resetTimer = new DispatcherTimer();
                resetTimer.Interval = TimeSpan.FromSeconds(1);  // Delay for 3 seconds
                resetTimer.Tick += (s, e) =>
                {
                    ResetUIElements();  // Reset UI elements after the delay
                    resetTimer.Stop();  // Stop the timer after resetting
                };
                resetTimer.Start();  // Start the timer
            }
        }

    }
}