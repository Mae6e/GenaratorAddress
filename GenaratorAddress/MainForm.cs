using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NBitcoin;
using NBitcoin.Altcoins;
using Nethereum.Contracts.Comparers;
using Nethereum.Hex.HexConvertors.Extensions;
using Nethereum.Web3.Accounts;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace GenaratorAddress
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                btnGenerate.Enabled = false;
                btnGenerate.Cursor = Cursors.WaitCursor;

                if (string.IsNullOrEmpty(cmbCurrency.Text))
                {
                    MessageBox.Show("Please Select Currency");
                    ResetButton();
                    return;
                }

                if (string.IsNullOrEmpty(txtNumber.Text))
                {
                    MessageBox.Show("Please Enter Number");
                    ResetButton();
                    return;
                }

                var generateAddressResponse = GenerateAddress(int.Parse(txtNumber.Text), cmbCurrency.Text);
                if (generateAddressResponse.IsSuccess)
                {
                    var exportExcelResponse = ExportExcel(generateAddressResponse.List, cmbCurrency.Text );

                    MessageBox.Show($"ExportExcel:{exportExcelResponse.Message}");
                    ResetButton();
                }
                else
                {
                    MessageBox.Show($"GenerateAddress:{generateAddressResponse.Message}");
                    ResetButton();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                ResetButton();
            }

        }


        public  ResponseGenerateAddressVm GenerateAddress(int limit , string currency)
        {
            try
            {
                List<(string, string)> list = new List<(string, string)>();

                if (cmbCurrency.Text == "ETH" || cmbCurrency.Text == "BNB")
                {
                    for (int i = 0; i < limit; i++)
                    {
                        var ecKey = Nethereum.Signer.EthECKey.GenerateKey();
                        var privateKey = ecKey.GetPrivateKeyAsBytes().ToHex();
                        var account = new Account(privateKey);

                        list.Add((account.Address, account.PrivateKey));
                    }
                }
                else if (cmbCurrency.Text == "BTC")
                {
                    for (int i = 0; i < limit; i++)
                    {
                        Key privateKey = new Key();
                        BitcoinSecret netPrivateKey = privateKey.GetBitcoinSecret(Network.Main);
                        var address = netPrivateKey.GetAddress(ScriptPubKeyType.SegwitP2SH);

                        list.Add((address.ToString(), netPrivateKey.ToString()));
                    }
                }
                else if (cmbCurrency.Text == "DASH")
                {
                    for (int i = 0; i < limit; i++)
                    {
                        Key privateKey = new Key();
                        BitcoinSecret netPrivateKey = privateKey.GetBitcoinSecret(Dash.Instance.Mainnet);
                        var address = netPrivateKey.GetAddress(ScriptPubKeyType.Legacy);

                        list.Add((address.ToString(), netPrivateKey.ToString()));
                    }
                }
                else if (cmbCurrency.Text == "DOGE")
                {
                    for (int i = 0; i < limit; i++)
                    {
                        Key privateKey = new Key();
                        BitcoinSecret netPrivateKey = privateKey.GetBitcoinSecret(Dogecoin.Instance.Mainnet);
                        var address = netPrivateKey.GetAddress(ScriptPubKeyType.Legacy);

                        list.Add((address.ToString(), netPrivateKey.ToString()));
                    }
                }
                else if (cmbCurrency.Text == "LTC")
                {
                    for (int i = 0; i < limit; i++)
                    {
                        Key privateKey = new Key();
                        BitcoinSecret netPrivateKey = privateKey.GetBitcoinSecret(Litecoin.Instance.Mainnet);
                        var address = netPrivateKey.GetAddress(ScriptPubKeyType.Legacy);

                        list.Add((address.ToString(), netPrivateKey.ToString()));
                    }
                }

                return new ResponseGenerateAddressVm
                {
                    IsSuccess = true,
                    List = list,
                    Message = "Success"
                };
            }
            catch (Exception ex)
            {
                return new ResponseGenerateAddressVm
                {
                    IsSuccess = false,
                    Message = ex.Message,
                    List = new List<(string, string)>()
                };
            }
        }

        public ResponseExcelVm ExportExcel(List<(string, string)> list, string currency)
        {
            try
            {
                var headers = new List<string>
                {
                   "row",
                   "PublicKey",
                   "PrivateKey"
                };

                var workbook = new HSSFWorkbook();
                var sheet = workbook.CreateSheet(currency);
                sheet.IsRightToLeft = true;

                var headerRow = sheet.CreateRow(0);

                //header
                HSSFFont headerFont = (HSSFFont)workbook.CreateFont();
                headerFont.FontHeightInPoints = (short)10;
                headerFont.Color = IndexedColors.Black.Index;
                headerFont.IsItalic = false;
                HSSFCellStyle headerStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                headerStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
                headerStyle.FillPattern = FillPattern.SolidForeground;
                headerStyle.ShrinkToFit = true;

                headerStyle.SetFont(headerFont);

                //content
                HSSFFont rowFont = (HSSFFont)workbook.CreateFont();
                rowFont.FontHeightInPoints = (short)10;
                rowFont.Color = IndexedColors.Black.Index;
                rowFont.IsItalic = false;
                HSSFCellStyle rowStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                rowStyle.FillForegroundColor = IndexedColors.White.Index;
                rowStyle.ShrinkToFit = true;
                rowStyle.SetFont(rowFont);

                HSSFCellStyle foterStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                foterStyle.FillPattern = FillPattern.SolidForeground;
                foterStyle.ShrinkToFit = true;
                foterStyle.SetFont(rowFont);

                int i = 0;
                foreach (var header in headers)
                {
                    headerRow.CreateCell(i).SetCellValue(header);
                    headerRow.Cells[i].CellStyle = headerStyle;
                    i++;
                }

                int rowNumber = 1;
                foreach (var item in list)
                {
                    var row = sheet.CreateRow(rowNumber);
                    row.CreateCell(0).SetCellValue(rowNumber);
                    row.CreateCell(1).SetCellValue(item.Item1);
                    row.CreateCell(2).SetCellValue(item.Item2);

                    row.Cells[0].CellStyle = rowStyle;
                    row.Cells[1].CellStyle = rowStyle;
                    row.Cells[2].CellStyle = rowStyle;

                    rowNumber++;
                }

                for (int w = 0; w < headers.Count; w++)
                {
                    sheet.AutoSizeColumn(w);
                }

                using (var fs = new FileStream($"GenerateAddress_{currency}.xls", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                return new ResponseExcelVm
                {
                    IsSuccess = true,
                    Message = "Success"
                };

            }
            catch (Exception ex)
            {
                return new ResponseExcelVm
                {
                    IsSuccess = true,
                    Message = ex.Message
                };
            }

        }

        public void ResetButton()
        {
            btnGenerate.Enabled = true;
            btnGenerate.Cursor = Cursors.Default;
        }
    }



    public class ResponseExcelVm
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }

    }
    public class ResponseGenerateAddressVm
    {
        public List<(string, string)> List { get; set; }
        public bool IsSuccess { get; set; }
        public string Message { get; set; }

    }
}
