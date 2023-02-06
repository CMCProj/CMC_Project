using System;
using System.IO;
using System.IO.Compression;
using System.Text;

//예외 처리가 필요함 (23.01.31)
//예외 처리는 ConvertionPage.xaml.cs의 ConvertButtonClick()에서 처리됨 (23.02.02)
/*
 23.02.02 업데이트
 --------------------
 작업 폴더 경로 수정
 --------------------
*/
/*
 23.02.02 업데이트2
 --------------------
 기존 폴더가 존재해도 제대로 작동되도록 수정
 --------------------
*/
/*
 23.02.06 업데이트
 --------------------
 같은 이름의 BID 파일이 있다면 덮어쓰도록 수정
 --------------------
*/

namespace SetUnitPriceByExcel
{
    class BidHandling
    {
        //String folder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); 사용 안함 (23.02.02)
        public static string filename;

        public static void BidToXml()
        {
            string copiedFolder = Data.folder + "\\EmptyBid"; // EmptyBid폴더 주소 저장 / 폴더 경로 수정 (23.02.02)
            string[] bidFile = Directory.GetFiles(copiedFolder, "*.BID");
            string myfile = bidFile[0];
            filename = Path.GetFileNameWithoutExtension(bidFile[0]);
            File.Move(myfile, Path.ChangeExtension(myfile, ".zip"));

            ZipFile.ExtractToDirectory(Path.Combine(copiedFolder, filename + ".zip"), copiedFolder);
            string[] files = Directory.GetFiles(copiedFolder, "*.BID");
            string text = File.ReadAllText(files[0]); // 텍스트 읽기
            byte[] decodeValue = Convert.FromBase64String(text);  // base64 변환
            text = Encoding.UTF8.GetString(decodeValue);   // UTF-8로 디코딩

            if(File.Exists(Data.folder + "\\OutputDataFromBID.xml"))    //기존 OutputDataFromBID.xml은 삭제한다. (23.02.02)
            {
                File.Delete(Data.folder + "\\OutputDataFromBID.xml");
            }

            File.WriteAllText(Path.Combine(Data.folder, "OutputDataFromBID.xml"), text, Encoding.UTF8); //폴더 경로 수정 (23.02.02)

            //실내역 데이터 복사 및 단가 세팅 & 직공비 고정금액 비중 계산
            Setting.GetData();
        }

        public static void XmlToBid()
        {
            string myfile = Path.Combine(Data.work_path, "Result_Xml.xml");
            byte[] bytes = File.ReadAllBytes(myfile);
            string encodeValue = Convert.ToBase64String(bytes);
            if (File.Exists(Data.work_path + "\\XmlToBID.BID"))    //기존 XmlToBID.BID은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\XmlToBID.BID");
            }
            File.WriteAllText(Path.Combine(Data.work_path, "XmlToBID.BID"), encodeValue);

            string resultFileName = filename.Substring(0, 16) + ".zip";
            if (File.Exists(Data.work_path + "\\" + resultFileName))    //기존 공내역파일명.zip은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\"+ resultFileName);
            }
            using (ZipArchive zip = ZipFile.Open(Path.Combine(Data.work_path, resultFileName), ZipArchiveMode.Create))
            {
                zip.CreateEntryFromFile(Path.Combine(Data.work_path, "XmlToBID.BID"), "XmlToBid.BID");
            }

            string resultBidPath = Path.ChangeExtension(Path.Combine(Data.work_path, resultFileName), ".BID");
            if (File.Exists(resultBidPath))    //기존 공내역파일명.BID은 삭제한다. (23.02.02)
            {
                File.Delete(resultBidPath);
            }
            File.Move(Path.Combine(Data.work_path, resultFileName), Path.ChangeExtension(Path.Combine(Data.work_path, resultFileName), ".BID"));
        }
    }
}