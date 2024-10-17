using ClosedXML.Excel;
using System.IO;

namespace Receipts_Distribution
{
    public class ReceiptsDistributionTask
    {
        private Logs _logs;
        private Errors _errors;

        public ReceiptsDistributionTask()
        {
            _logs = new Logs();
            _errors = new Errors();
        }

        /// <summary>
        /// 
        /// </summary>
        public class OrderDetails
        {
            /// <summary>
            /// 受託者
            /// </summary>
            public string? Contractor { get; set; }
            /// <summary>
            /// 注文年月日
            /// </summary>
            public string? OrderDate { get; set; }
            /// <summary>
            /// 納期
            /// </summary>
            public string? DelieryDate { get; set; }
            /// <summary>
            /// 委託期間（日）
            /// </summary>
            public string? Duration { get; set; }
            /// <summary>
            /// 委託件数(件)
            /// </summary>
            public int? OutsourcingCount { get; set; }
        }

        bool forAllInspector = false; //各個人別表の合体Excelの為 ↳Excel【別表】※１シート一人分の別表

        /// <summary>
        /// リセ振り分けメイン処理
        /// </summary>
        /// <param name="inspectionResultExcelFilePath"></param>
        /// <param name="masterFilePath"></param>
        /// <param name="correctionFilePath"></param>
        /// <param name="orderFilePath"></param>
        public void CreateReceiptsDistributionTask(string inspectionResultExcelFilePath, string masterFilePath, string correctionFilePath, string orderFilePath)
        {
            try
            {
                var inspectionResultWorkbook = new XLWorkbook(inspectionResultExcelFilePath);
                var inspectionResultSheet = inspectionResultWorkbook.Worksheet(1);
                var sheetB = inspectionResultWorkbook.Worksheet("0902点検結果B"); //別表(原本）

                var masterWorkbook = new XLWorkbook(masterFilePath);
                var masterSheet = masterWorkbook.Worksheet(1);

                // 別表データの全件をリストとして取得（1行目をヘッダーと仮定）
                var requestTotalList = new List<string[ ]>();
                int requestTotal = sheetB.LastRowUsed().RowNumber(); //ヘッダを抜く
                for( int row = 2; row <= requestTotal; row++ )
                {
                    var requestItem = sheetB.Row(row).Cells().Select(c => c.GetString()).ToArray();
                    requestTotalList.Add(requestItem);
                }

                string defaultExcelFilePath = Properties.Settings.Default.excelfilePath;
                DateTime today = DateTime.Now;
                string folderName = today.ToString("MMdd"); //∟フォルダ【0902】

                //各点検者に依頼件数分を振分
                //依頼マスタの最終行を取得
                int lastRow = masterSheet.LastRowUsed().RowNumber() - 1;

                int currentIndex = 0;

                for( int row = 7; row <= lastRow; row++ )
                {
                    string surname = masterSheet.Cell(row, 1).Value.ToString(); // 姓を取得
                    string givenName = masterSheet.Cell(row, 2).Value.ToString();
                    string contractName = masterSheet.Cell(row, 10).Value.ToString();
                    string subFolderName = string.Format("{0}-{1}{2}", contractName, folderName, surname); //∟フォルダ【F診療所-0902舘野】

                    int maxCount = (int) (masterSheet.Cell(row, 7).Value); //最大可能数
                    int requestCount = (int) (masterSheet.Cell(row, 9).Value); //依頼件数

                    //最大件数分が依頼件数分より小さいの場合は最大件数分を依頼件数分に設定
                    requestCount = maxCount < requestCount ? maxCount : requestCount;

                    // 分割するデータを取得
                    var subsetData = requestTotalList.Skip(currentIndex).Take(requestCount).ToList();
                    currentIndex += requestCount;

                    //発注書詳細データ
                    var orderdetails = new OrderDetails
                    {
                        Contractor = string.Format("{0} {1}", surname, givenName), //受託者
                        OrderDate = DateTime.Now.ToString("yyyy/MM/dd"),//注文年月日
                        DelieryDate = string.Format("{0}　終日", masterSheet.Cell(row, 6).GetValue<DateTime>().ToString("MM/dd")), //納期
                        Duration = string.Format("{0}日", masterSheet.Cell(row, 5).Value.ToString()),//委託期間（日）
                        OutsourcingCount = requestCount//委託件数(件)
                    };

                    // 分割したデータを別のファイルとして保存
                    _saveSubsetToFile(subsetData, surname, defaultExcelFilePath, folderName, subFolderName, forAllInspector, contractName);
                    //発注書作成
                    _saveOrderToFile(orderdetails, orderFilePath, surname, defaultExcelFilePath, folderName, subFolderName, contractName);
                    //修正指示書作成
                    _saveCorrectionDataToFile(inspectionResultExcelFilePath, subsetData, correctionFilePath, surname, defaultExcelFilePath, folderName, subFolderName, contractName);
                    _logs.Log($"振分処理が正常に完了しました。");
                }

                //最大可能数より、今回チェック数が増えた場合は最大可能数を超えた分を「(例) 仮名 保留」に保存
                int possibleTotalCount = (int) (masterSheet.Cell("G5").Value);//最大可能数合計 2720
                int maxRequestCount = (int) (masterSheet.Cell("H5").Value);//最大依頼件数 2761
                if( possibleTotalCount < maxRequestCount )
                {
                    int surplus = maxRequestCount - possibleTotalCount; //余ってる分 41
                    // 分割するデータを取得
                    var subsetData = requestTotalList.Skip(possibleTotalCount).Take(surplus).ToList();

                    string subFolderName = string.Format("{0}{1}", folderName, "仮名　保留"); //∟フォルダ【0902仮名　保留】

                    //発注書詳細データ
                    var orderdetails = new OrderDetails
                    {
                        Contractor = "仮名",
                        OrderDate = DateTime.Now.ToString("yyyy/MM/dd"),
                        DelieryDate = "",
                        Duration = "",
                        OutsourcingCount = surplus
                    };

                    // 分割したデータを別のファイルとして保存
                    _saveSubsetToFile(subsetData, "仮名", defaultExcelFilePath, folderName, subFolderName);
                    _saveOrderToFile(orderdetails, orderFilePath, "仮名", defaultExcelFilePath, folderName, subFolderName);
                    _saveCorrectionDataToFile(inspectionResultExcelFilePath, subsetData, correctionFilePath, "仮名", defaultExcelFilePath, folderName, subFolderName);
                }
            }
            catch( Exception ex )
            {
                string errorMessage = string.Format(_errors.GetErrorMessage("E01"), "分割したデータを別のファイルとして保存");
                _logs.Log($"{errorMessage}:{ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// 別表作成処理
        /// </summary>
        /// <param name="subsetData"></param>
        /// <param name="surname"></param>
        /// <param name="defaultExcelFilePath"></param>
        /// <param name="folderName"></param>
        /// <param name="subFolderName"></param>
        // 分割されたデータをExcelファイルに保存するメソッド
        private void _saveSubsetToFile(List<string[ ]> subsetData, string surname, string defaultExcelFilePath, string folderName, string subFolderName, bool forAllInspector = false, string contractName = "")
        {
            try
            {
                string fileNameForAllInspector = "別表.xlsx";
                string subsetFilePathForAllInspector = string.Format("{0}{1}\\{2}", defaultExcelFilePath, folderName, fileNameForAllInspector);
                using( var workbook = new XLWorkbook() )
                {
                    XLWorkbook workbookForAllInspector;
                    if( !File.Exists(subsetFilePathForAllInspector) )
                    {
                        workbookForAllInspector = new XLWorkbook();
                    }
                    else
                    {
                        workbookForAllInspector = new XLWorkbook(subsetFilePathForAllInspector);
                    }

                    var sheet = workbook.AddWorksheet(surname);
                    var sheetForAllInspector = workbookForAllInspector.Worksheets.Add(string.Format("{0}{1}", surname, subsetData.Count));

                    // データを書き込む（1行目にヘッダーがあると仮定）
                    for( int row = 1; row <= subsetData.Count; row++ )
                    {
                        for( int col = 0; col < subsetData[ row - 1 ].Length; col++ )
                        {
                            sheet.Cell(row, col + 1).Value = subsetData[ row - 1 ][ col ];
                            sheet.Cell(row, col + 2).Value = surname;
                            sheetForAllInspector.Cell(row, col + 1).Value = subsetData[ row - 1 ][ col ];
                            sheetForAllInspector.Cell(row, col + 2).Value = surname;
                        }
                    }
                    string tmp = contractName == "" ? "" : "-";
                    string fileName = string.Format("{0}{1}{2}：{3}", contractName, surname, "別表.xlsx"); //∟Excel【F診療所-小沼：別表】
                    string subsetFilePath = string.Format("{0}{1}\\{2}\\{3}", defaultExcelFilePath, folderName, subFolderName, fileName);

                    // ファイルを保存
                    workbook.SaveAs(subsetFilePath);
                    _logs.Log($"別表作成処理が正常に終了しました。");

                    workbookForAllInspector.SaveAs(subsetFilePathForAllInspector);
                    _logs.Log($"subsetFilePathForAllInspectorが正常に終了しました。");
                    throw new InvalidOperationException("A処理でエラーが発生しました。");
                }
            }
            catch( Exception ex )
            {
                string errorMessage = string.Format(_errors.GetErrorMessage("E01"), "別表作成");
                _logs.Log($"{errorMessage}:{ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 発注書作成処理
        /// </summary>
        /// <param name="masterFilePath"></param>
        /// <param name="orderFilePath"></param>
        /// <param name="surname"></param>
        /// <param name="defaultExcelFilePath"></param>
        /// <param name="folderName"></param>
        /// <param name="subFolderName"></param>
        private void _saveOrderToFile(OrderDetails orderDetails, string orderFilePath, string surname, string defaultExcelFilePath, string folderName, string subFolderName, string contractName = "")
        {
            try
            {
                using( var workbook = new XLWorkbook(orderFilePath) )
                {
                    var sheet = workbook.Worksheet("テンプレ");
                    sheet.Cell("C3").Value = orderDetails.Contractor; //受託者
                    sheet.Cell("C4").Value = orderDetails.OrderDate; // "注文年月日";                 
                    sheet.Cell("C8").Value = orderDetails.DelieryDate; //納期
                    sheet.Cell("C9").Value = orderDetails.Duration; //委託期間（日）
                    sheet.Cell("C11").Value = orderDetails.OutsourcingCount; //委託件数(件)
                    string tmp = contractName == "" ? "" : "-";
                    string fileName = string.Format("{0}{1}{2}：{3}", contractName, tmp, surname, "発注書.xlsx"); //∟Excel【F診療所-小沼：発注書】
                    string saveOrderFilePath = string.Format("{0}{1}\\{2}\\{3}", defaultExcelFilePath, folderName, subFolderName, fileName);
                    // ファイルを保存
                    workbook.SaveAs(saveOrderFilePath);
                    _logs.Log($"発注書作成処理が正常に終了しました。");
                }
            }
            catch( Exception ex )
            {
                string errorMessage = string.Format(_errors.GetErrorMessage("E01"), "発注書作成");
                _logs.Log($"{errorMessage}:{ex.Message}");
            }
        }

        /// <summary>
        /// 修正指示書処理
        /// </summary>
        /// <param name="inspectionResultExcelFilePath"></param>
        /// <param name="subsetData"></param>
        /// <param name="correctionFilePath"></param>
        /// <param name="surname"></param>
        /// <param name="defaultExcelFilePath"></param>
        /// <param name="folderName"></param>
        /// <param name="subFolderName"></param>
        private void _saveCorrectionDataToFile(string inspectionResultExcelFilePath, List<string[ ]> subsetData, string correctionFilePath, string surname, string defaultExcelFilePath, string folderName, string subFolderName, string contractName = "")
        {
            try
            {
                // 点検結果データのリストを初期化
                var inspectionData = new List<(string patientId, string[ ] data)>();
                using( var inspectionResultWorkbook = new XLWorkbook(inspectionResultExcelFilePath) )
                using( var correctionWorkbook = new XLWorkbook(correctionFilePath) )
                {
                    var inspectionSheet = inspectionResultWorkbook.Worksheet("0902点検結果C"); //修正指示書(原本）

                    var correctionSheet = correctionWorkbook.Worksheet("点検リスト");

                    int lastRow = inspectionSheet.LastRowUsed().RowNumber();
                    for( int row = 2; row <= lastRow; row++ )
                    {
                        string patientId = inspectionSheet.Cell(row, 2).GetString(); // 患者番号
                        var rowData = inspectionSheet.Row(row).Cells().Select(c => c.GetString()).ToArray();
                        inspectionData.Add((patientId, rowData));
                    }

                    var joinedData = new List<string[ ]>();

                    int rowIndex = 2;
                    foreach( var item in subsetData )
                    {
                        string patientId = item[ 0 ]; // 患者番号
                        var matchingInspections = inspectionData.Where(id => id.patientId == patientId).ToList();
                        foreach( var inspection in matchingInspections )
                        {
                            // 別表のデータと点検結果のデータを結合
                            //var combinedRow = new string[ item.Length + inspection.data.Length ];
                            //item.CopyTo(combinedRow, 0);
                            //inspection.data.CopyTo(combinedRow, item.Length);

                            //joinedData.Add(combinedRow);
                            correctionSheet.Cell(rowIndex, 1).Value = inspection.data[ 1 ];
                            correctionSheet.Cell(rowIndex, 2).Value = inspection.data[ 2 ];
                            correctionSheet.Cell(rowIndex, 3).Value = inspection.data[ 3 ];
                            correctionSheet.Cell(rowIndex, 4).Value = inspection.data[ 4 ];
                            rowIndex++;
                        }
                    }
                    string tmp = contractName == "" ? "" : "-";
                    string fileName = string.Format("{0}{1}{2}：{3}", contractName, tmp, surname, "修正指示書.xlsx"); //∟Excel【F診療所-小沼：修正指示書】
                    string saveCorrectionFilePath = string.Format("{0}{1}\\{2}\\{3}", defaultExcelFilePath, folderName, subFolderName, fileName);
                    // ファイルを保存
                    correctionWorkbook.SaveAs(saveCorrectionFilePath);
                    _logs.Log($"修正指示書処理が正常に終了しました。");
                }
            }
            catch( Exception ex )
            {
                string errorMessage = string.Format(_errors.GetErrorMessage("E01"), "修正指示書作成");
                _logs.Log($"{errorMessage}:{ex.Message}");
            }
        }
    }
}