using Microsoft.Win32;
using SetUnitPriceByExcel;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
/*
 23.02.01 업데이트
 --------------------
  사정율에 음의 값 입력 가능하도록 수정
  사정율, 사업자등록번호 값 검토 추가
 --------------------
 23.01.31 업데이트
 --------------------
  UI 수정
  사업자등록번호 입력받기&Data에 저장
  BusinessChangedHandler() 추가
  SetBusinessInfoBtnClick() 추가
 --------------------
*/
namespace CMC_Project.Views
{
    /// <summary>
    /// Interaction logic for AdjustmentPage.xaml
    /// </summary>
    public partial class AdjustmentPage : Page
    {
        private static bool isCalculate = false;
        public static bool isConfirm = false;
        public AdjustmentPage()
        {
            InitializeComponent();
            this.businessInfo.TextChanged += BusinessChangedHandler;
            this.averageRating.TextChanged += AverageChangedHandler;
            this.estimateRating.TextChanged += EstimateChangedHandler;

            // 사정율 초기화
            if (Data.BalanceRateNum != null && Data.PersonalRateNum != null)
            {
                averageRating.Text = ((double)Data.BalanceRateNum).ToString();
                estimateRating.Text = ((double)Data.PersonalRateNum).ToString();
            }
            // 라디오 버튼 초기화
            Data.UnitPriceTrimming = "1";
            // 표준시장 단가 체크
            if (Data.StandardMarketDeduction == "1")
                CheckStandardPrice.IsChecked = true;
            else
                CheckStandardPrice.IsChecked = false;
            // 공종 가중치 체크
            if (Data.ZeroWeightDeduction == "1")
                CheckWeightValue.IsChecked = true;
            else
                CheckWeightValue.IsChecked = false;
            // 법정 요율 체크
            if (Data.CostAccountDeduction == "1")
                CheckCAD.IsChecked = true;
            else
                CheckCAD.IsChecked = false;
            // 원단위 체크
            if (Data.BidPriceRaise == "1")
                CheckCeiling.IsChecked = true;
            else
                CheckCeiling.IsChecked = false;


        }

        //sender: 이벤트 발생자, args: 이벤트 인자

        private void BusinessChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox BusinessInfo = sender as TextBox;
            int selectionStart = BusinessInfo.SelectionStart;
            string result = string.Empty;
            //Data.CompanyRegistrationNum = (Double.Parse(BusinessInfo.GetLineText(0)));

            foreach (char character in BusinessInfo.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character))
                {
                    result += character;

                }
            }
            BusinessInfo.Text = result;
            BusinessInfo.SelectionStart = selectionStart <= BusinessInfo.Text.Length ? selectionStart : BusinessInfo.Text.Length;
        }
        private void AverageChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox averageRating = sender as TextBox;
            int selectionStart = averageRating.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.BalanceRateNum = (Double.Parse(averageRating.GetLineText(0)));

            foreach (char character in averageRating.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0) 
                    || (character =='-' && mCount == 0 ))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount +=1;
                    }

                }
            }
            averageRating.Text = result;
            averageRating.SelectionStart = selectionStart <= averageRating.Text.Length ? selectionStart : averageRating.Text.Length;
        }

        //sender: 이벤트 발생자, args: 이벤트 인자
        private void EstimateChangedHandler(object sender, TextChangedEventArgs args)
        {
            TextBox estimateRating = sender as TextBox;
            int selectionStart = estimateRating.SelectionStart;
            string result = string.Empty;
            int dCount = 0;
            int mCount = 0;
            //Data.PersonalRateNum = (Double.Parse(estimateRating.GetLineText(0)));


            foreach (char character in estimateRating.Text.ToCharArray())
            {
                if (char.IsDigit(character) || char.IsControl(character) || (character == '.' && dCount == 0)
                    || (character == '-' && mCount == 0))
                {
                    result += character;
                    if (character == '.')
                    {
                        dCount += 1;
                    }
                    else if (character == '-')
                    {
                        mCount += 1;
                    }
                }
            }
            
            estimateRating.Text = result;
            estimateRating.SelectionStart = selectionStart <= estimateRating.Text.Length ? selectionStart : estimateRating.Text.Length;
        }

        private void UpBtnClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Clicked");
        }


        // ------------------------- 옵션 입력 버튼 ------------------------------------------------------------------------------------------------------------------------------------------- //
        //소수 1자리 체크
        private void RadioDecimal_Checked(object sender, RoutedEventArgs e)
        {
            Data.UnitPriceTrimming = "1";
        }
        // 정수 체크
        private void RadioInteger_Checked(object sender, RoutedEventArgs e)
        {
            Data.UnitPriceTrimming = "2";
        }

        // 표준시장 단가 체크
        private void CheckStandardPrice_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckStandardPrice.IsChecked)
            {
                Data.StandardMarketDeduction = "1";
            }
            else
            {
                Data.StandardMarketDeduction = "2";
            }
        }

        // 공종 가중치 체크
        private void CheckWeightValue_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckWeightValue.IsChecked)
            {
                Data.ZeroWeightDeduction = "1";
            }
            else
            {
                Data.ZeroWeightDeduction = "2";
            }
        }

        // 법정요율 체크
        private void CheckCAD_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckCAD.IsChecked)
            {
                Data.CostAccountDeduction = "1";
            }
            else
            {
                Data.CostAccountDeduction = "2";
            }
        }

        // 원단위 체크
        private void CheckCeiling_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckCeiling.IsChecked)
            {
                Data.BidPriceRaise = "1";
            }
            else
            {
                Data.BidPriceRaise = "2";
            }
        }

        //노무비 하한율 체크
        private void CheckLaborCost_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)CheckCeiling.IsChecked)
            {
                Data.LaborCostLowBound = "1";
            }
            else
            {
                Data.LaborCostLowBound = "2";
            }
        }

        private void SetBusinessInfoBtnClick(object sender, RoutedEventArgs e)
        {
            if(businessInfo.Text.Length != 10)
            {
                DisplayDialog($"사업자등록번호를 10자리로 입력해주세요.", "Error");
            }
            else
            {
                Data.CompanyRegistrationNum = (businessInfo.Text);
                DisplayDialog($"사업자등록번호({Data.CompanyRegistrationNum})를\n저장했습니다.", "Success");
            }
        }

        private void CalBtnClick(object sender, RoutedEventArgs e)
        {
            if (averageRating.Text == string.Empty || estimateRating.Text == string.Empty)
            {
                DisplayDialog("사정율을 입력해주세요.", "Error");
                return;
            }
            else
            {
                try
                {
                    Data.BalanceRateNum = (Double.Parse(averageRating.Text));
                    if (Data.BalanceRateNum < -2 || Data.BalanceRateNum > 2)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("업체 평균 사정율이 올바르지 않습니다.", "Error");
                    return;
                }
                try
                {
                    Data.PersonalRateNum = (Double.Parse(estimateRating.Text));
                    if (Data.PersonalRateNum < -2 || Data.PersonalRateNum > 2)
                        throw new Exception();
                }
                catch (Exception)
                {
                    DisplayDialog("나의 사정율이 올바르지 않습니다.", "Error");
                    return;
                }
            }
            if (Data.CompanyRegistrationNum == "")
            {
                DisplayDialog("사업자등록번호를 입력해주세요.", "Error");
                return;
            }



            // 단가를 불러온 경우
            if (isConfirm)
            {
                //입찰금액 심사 점수 계산 및 단가 조정
                CalculatePrice.Calculation();

                FixedPercentPrice.Text = Data.FixedPricePercent + " %";
                MyPercent.Text = "(+/-) " + CalculatePrice.myPercent * 100.0m + " %";
                TargetRate.Text = Data.Bidding["도급비계"] + " 원 " + "(" + FillCostAccount.GetRate("도급비계") + " %)"; // 도급비계
                isCalculate = true;

                //OutputTextBlock.Text = "사정율 적용 완료!";
                DisplayDialog("사정율 적용을 완료하였습니다", "Success");
            }

            // 단가를 불러오지 않은 경우
            else
            {
                DisplayDialog("단가를 먼저 세팅해주세요.", "Error");
            }
        }


        // ------------------------- 세부 결과 확인 버튼 ------------------------------------------------------------------------------------------------------------------------------------------------- //
        private void ShowResult_Click(object sender, RoutedEventArgs e)
        {
            if (isCalculate)
            {
                CMC_Project.Views.ResultPage rw = new();

                rw.Show();
            }
            else
            {
                DisplayDialog("계산 후 확인해주세요", "Fail");
            }
        }



        // 메세지 창
        static public void DisplayDialog(String dialog, String title)
        {
            MessageBox.Show(dialog, title, MessageBoxButton.OK, MessageBoxImage.Information);
        }


        // ------------------------- BID파일 저장 버튼 ---------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveBidBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                //단가 세팅 완료한 xml 파일을 다시 BID 파일로 변환
                BidHandling.XmlToBid();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "BID Files (*.BID)|*.BID|All files (*.*)|*.*";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = BidHandling.filename.Substring(0, 16);
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string bidFolder = Data.work_path; //Result Bid 경로
                    string finalBidFile = Path.Combine(bidFolder, BidHandling.filename.Substring(0, 16) + ".BID");

                    File.Move(finalBidFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }
            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("입찰점수를 계산해주세요.", "Error");
            }
        }


        // ------------------------- 원가계산서 저장 버튼 ------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveCostBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                //가격 조정 후 원가계산서 엑셀파일 생성
                FillCostAccount.FillBiddingCosts();

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "Microsoft Excel (*.xlsx)|*.xlsx";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = "원가계산서_세부결과";
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string xlsxFolder = Data.work_path;
                    string costFile = Path.Combine(xlsxFolder, "원가계산서_세부결과.xlsx");

                    File.Move(costFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }
            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("계산을 먼저 실행해주세요.", "Error");
            }
        }


        // ------------------------- 입찰 내역 저장 버튼 -------------------------------------------------------------------------------------------------------------------------------------- //
        private void SaveBiddingZipBtnClick(object sender, System.EventArgs e)
        {
            // TargetRate가 계산 되어 있을 경우
            if (isCalculate)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog.Filter = "Zip 압축 파일 (*.zip)|*.zip";
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.FileName = "입찰내역";
                saveFileDialog.OverwritePrompt = true;


                if (saveFileDialog.ShowDialog() == true)
                {
                    string file = saveFileDialog.FileName.ToString(); //경로와 파일명 저장
                    string biddingFolder = Data.work_path; //입찰 내역 경로
                    string biddingZipFile = Path.Combine(biddingFolder, "입찰내역.zip");

                    Directory.Move(biddingZipFile, file);
                    DisplayDialog("저장되었습니다.", "Save");
                }
                else
                {
                    DisplayDialog("취소되었습니다.", "Error");
                }


            }

            // 계산 안되어 있을 경우
            else
            {
                DisplayDialog("계산을 먼저 실행해주세요.", "Error");
            }

        }
    }
}