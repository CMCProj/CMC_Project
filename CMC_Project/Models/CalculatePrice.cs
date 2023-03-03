using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;

/*
 23.01.31 업데이트
 ------------------
 새로운 Xml 구조에 맞게 수정
 ==================
 T3
     C24 -> C9
     C4  -> C5
     C15 -> C16
     C16 -> C17
     C17 -> C18
     C18 -> C19
     C19 -> C20
     C20 -> C21
     C21 -> C22
     C22 -> C23
 ==================
 T5
     C9  -> C4
     C22 -> C8
 ==================
*/
/*
 23.01.31 업데이트2
 --------------------
  사업자등록번호 <T1></C17><T1>에 추가
  SetBusinessInfo() 추가
  Calculation() 내부에서 SetBusinessInfo() 호출 추가
 --------------------
 */
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
 고정금액 소수점 5자리에서 절사되도록 수정
 --------------------
*/
/*
 23.02.07 업데이트
 --------------------
 공종 합계 저장 메소드 (SetPriceOfSuperConstruction) 추가
 --------------------
*/
/*
 --------------------
 노무비 80% 미만시 단가조정 메소드 추가 (CheckLaborLimit80)
 --------------------
 */


namespace SetUnitPriceByExcel
{
    class CalculatePrice
    {
        static XDocument docBID;
        static IEnumerable<XElement> eleBID;
        static XElement maxBid = null;  //제요율적용제외공종 항목 중 단가가 가장 높은 항목
        public static decimal myPercent;   //최저네고단가율
        public static decimal balancedUnitPriceRate;   //균형단가율
        public static decimal targetRate;  //타겟율
        static decimal exSum = 0;   //단가조정된 제요율적용제외공종 항목 합계
        static decimal exCount = 0; //제요율적용제외공종 항목 수량 합계

        static void ApplyStandardPriceOption()
        {

            foreach (var bid in eleBID)
            {
                //표준시장단가 항목인경우 99.7% 적용
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    if (curObject.Item.Equals("표준시장단가"))
                    {
                        //직공비, 고정금액, 표준시장단가 금액 재계산
                        Data.RealDirectMaterial -= Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor -= Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense -= Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                        Data.FixedPriceDirectMaterial -= Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.FixedPriceDirectLabor -= Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.FixedPriceOutputExpense -= Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                        Data.StandardMaterial -= Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.StandardLabor -= Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.StandardExpense -= Convert.ToDecimal(string.Concat(bid.Element("C22").Value));

                        //표준시장단가 99.7% 적용
                        if (curObject.MaterialUnit!=0)
                            curObject.MaterialUnit = (Math.Truncate(curObject.MaterialUnit * 0.997m*10m)/10m)+0.1m;
                        if (curObject.LaborUnit != 0)
                            curObject.LaborUnit = (Math.Truncate(curObject.LaborUnit * 0.997m*10m)/10m)+0.1m;
                        if (curObject.ExpenseUnit!=0)
                            curObject.ExpenseUnit =( Math.Truncate(curObject.ExpenseUnit * 0.997m * 10m) / 10m)+0.1m;

                        //단가 변경사항 xml 파일에 적용
                        bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                        bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                        bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                        bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                        bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                        bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                        bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계

                        //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                        Data.RealDirectMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                        Data.FixedPriceDirectMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.FixedPriceDirectLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.FixedPriceOutputExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                        Data.StandardMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.StandardLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.StandardExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                    }
                }
            }
        }

        static void GetFixedPriceRate()
        {
            //고정금액 비율 계산
            var directConstPrice = Data.Investigation["직공비"];
            var fixCostSum = Data.InvestigateFixedPriceDirectMaterial + Data.InvestigateFixedPriceDirectLabor + Data.InvestigateFixedPriceOutputExpense;

            Data.FixedPricePercent = Math.Truncate(((fixCostSum / directConstPrice) * 100) * 10000) / 10000; // 고정금액 비중 계산 / 고정금액 소수점 5자리 수에서 절사 (23.02.06)
        }

        static void FindMyPercent() //고정금액 비중에 따른 최저네고단가율 계산
        {
            if (Data.FixedPricePercent < 20.0m)       //고정금액 < 20%
                myPercent = 0.85m;
            else if (Data.FixedPricePercent < 25.0m)  //고정금액 < 25%
                myPercent = 0.84m;
            else if (Data.FixedPricePercent < 30.0m)  //고정금액 < 30%
                myPercent = 0.83m;
            else myPercent = 0.82m;                   //고정금액 > 30%
        }
        static void GetWeight() //가중치 계산
        {
            var varCostSum = Data.RealPriceDirectMaterial + Data.RealPriceDirectLabor + Data.RealPriceOutputExpense; //총 합계금액(-,PS,표준시장단가 제외)
            decimal weight;
            decimal maxWeight = 0;
            decimal weightSum = 0;
            Data max = new Data();

            foreach (KeyValuePair<string, List<Data>> dic in Data.Dic)
            {
                foreach (var item in dic.Value)
                {
                    if (item.Item.Equals("일반"))
                    {
                        var material = item.Material;
                        var labor = item.Labor;
                        var expense = item.Expense;
                        weight = Math.Round((material + labor + expense) / varCostSum, 6);  //소숫점 일곱 자리 반올림
                        weightSum += weight;        //가중치를 더함
                        if (maxWeight < weight)     //최대 가중치 갱신
                        {
                            maxWeight = weight;
                            max = item;
                        }
                        item.Weight = weight;
                    }
                }
            }

            if (weightSum != 1.0m)    //가중치의 합이 1이 되지 않으면 가중치가 가장 큰 항목에 부족한 양을 더한다
            {
                decimal lack = 1.0m - weightSum;
                max.Weight += lack;
            }
        }
        static void CalculateRate(decimal personalRate, decimal balancedRate)
        {    //Target Rate 계산
            decimal unitPrice = 100;
            balancedUnitPriceRate = ((0.9m * unitPrice * (1.0m + balancedRate / 100.0m) * myPercent) / (1.0m - 0.1m * myPercent)) / 100;   //균형단가율
            targetRate = ((unitPrice * (1.0m + personalRate / 100.0m) * 0.9m + unitPrice * balancedUnitPriceRate * 0.1m) * myPercent) / 100;    //Target_Rate
            targetRate = Math.Truncate(targetRate * 1000000) / 1000000;
        }
        static void RoundOrTruncate(decimal Rate, Data Object, ref decimal myMaterialUnit, ref decimal myLaborUnit, ref decimal myExpenseUnit)
        { //절사,반올림 옵션
            if (Data.UnitPriceTrimming.Equals("1"))
            {
                myMaterialUnit = Math.Truncate(Object.MaterialUnit * Rate * 10) / 10;
                myLaborUnit = Math.Truncate(Object.LaborUnit * Rate * 10) / 10;
                myExpenseUnit = Math.Truncate(Object.ExpenseUnit * Rate * 10) / 10;
            }
            else if (Data.UnitPriceTrimming.Equals("2"))
            {
                myMaterialUnit = Math.Ceiling(Object.MaterialUnit * Rate);
                myLaborUnit = Math.Ceiling(Object.LaborUnit * Rate);
                myExpenseUnit = Math.Ceiling(Object.ExpenseUnit * Rate);
            }
        }
        
        static void CheckLaborLimit80(Data Object, ref decimal myMaterialUnit, ref decimal myLaborUnit, ref decimal myExpenseUnit)
        { //2.8 노무비 80%미만일 경우 조정하는 메소드
            if (Object.LaborUnit * 0.8m > myLaborUnit)
            {
                decimal deficiency = Object.LaborUnit * 0.8m- myLaborUnit;
                if (myMaterialUnit!=0)
                    myMaterialUnit-=deficiency;
                else if (myExpenseUnit!=0)
                    myExpenseUnit-=deficiency;
                myLaborUnit = Object.LaborUnit * 0.8m;                
            }
        }
        static void Recalculation() //사정율에 따라 재계산된 가격을 비드파일에 복사
        {
            exCount = 0;
            exSum = 0;
            foreach (var bid in eleBID)
            {
                //일반 항목인 경우
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    if (curObject.Item.Equals("일반"))
                    {
                        //직접공사비 재계산
                        Data.RealDirectMaterial -= Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor -= Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense -= Convert.ToDecimal(string.Concat(bid.Element("C22").Value));

                        var targetPrice = (curObject.MaterialUnit + curObject.LaborUnit + curObject.ExpenseUnit) * targetRate;  //Target 단가 합계

                        //my 단가를 구하는 과정도 사용자의 옵션에 따라 소수 첫째 자리 아래로 절사(1) / 정수(2)로 나뉜다.
                        decimal myMaterialUnit = 0;
                        decimal myLaborUnit = 0;
                        decimal myExpenseUnit = 0;
                        decimal myPrice;

                        if (Data.ZeroWeightDeduction.Equals("1"))
                        {   //최소단가율 50% 적용 O
                            if (curObject.Weight == 0 && curObject.LaborUnit == 0)
                            {
                                //공종 가중치 0%이고 노무비 단가가 0원인 경우 사용자의 소수처리 옵션과 상관없이 50% 적용후 소수첫째자리에서 올림 (23.2.23)
                                curObject.MaterialUnit = Math.Ceiling(curObject.MaterialUnit * 0.5m);
                                //curObject.LaborUnit = Math.Ceiling(curObject.LaborUnit * 0.5m);
                                curObject.ExpenseUnit = Math.Ceiling(curObject.ExpenseUnit * 0.5m);

                                //최종 단가 및 합계 계산
                                bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                                bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                                bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                                bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                                bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                                bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                                bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                                bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계

                                //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                                Data.RealDirectMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                                Data.RealDirectLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                                Data.RealOutputExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));

                                continue;
                            }
                            else
                            {
                                RoundOrTruncate(targetRate, curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);
                                CheckLaborLimit80(curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);
                            }
                        }
                        else if (Data.ZeroWeightDeduction.Equals("2"))
                        {  //최소단가율 50% 적용 X
                            RoundOrTruncate(targetRate, curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);
                            CheckLaborLimit80(curObject, ref myMaterialUnit, ref myLaborUnit, ref myExpenseUnit);

                        }

                        myPrice = myMaterialUnit + myLaborUnit + myExpenseUnit;

                        if (Data.LaborCostLowBound.Equals("1"))  //노무비 하한 80% 적용 O
                        {
                          //  if (myPrice > targetPrice)
                            {  //여유분 조정 가능(조사노무비 대비 My노무비 비율에 따라 조정)
                                var Excess = myPrice - targetPrice;
                                var laborExcess = myLaborUnit - curObject.LaborUnit * 0.8m;
                                laborExcess = Math.Truncate(laborExcess * 10) / 10;
                                if (laborExcess > 0)
                                {
                                    if (myExpenseUnit != 0)
                                    {
                                        myLaborUnit -= laborExcess;
                                        myExpenseUnit += laborExcess + Excess;
                                    }
                                    else
                                    {
                                        if (myMaterialUnit != 0)
                                        {
                                            myLaborUnit -= laborExcess;
                                            myMaterialUnit += laborExcess + Excess;
                                        }
                                        else
                                        {
                                            myLaborUnit -= laborExcess;
                                            myExpenseUnit += laborExcess + Excess;
                                        }
                                    }
                                }
                                else if (laborExcess < 0)
                                {
                                    myLaborUnit = curObject.LaborUnit * 0.8m;
                                    if (myMaterialUnit != 0)
                                    {
                                        myMaterialUnit += laborExcess + Excess;
                                    }
                                    else
                                    {
                                        myExpenseUnit += laborExcess + Excess;
                                    }
                                }
                            }
                        }

                        curObject.MaterialUnit = myMaterialUnit;
                        curObject.LaborUnit = myLaborUnit;
                        curObject.ExpenseUnit = myExpenseUnit;
                        //최종 단가 및 합계 계산
                        bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                        bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                        bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                        bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                        bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                        bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                        bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계

                        //붙여넣기한 각 객체의 재료비, 노무비, 경비를 직접재료비, 직접노무비, 산출 경비에 더해나감
                        Data.RealDirectMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.RealDirectLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.RealOutputExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                    }
                    //제요율적용제외공종 단가 재세팅
                    else if (curObject.Item == "제요율적용제외")
                    {
                        curObject.MaterialUnit = Math.Truncate(curObject.MaterialUnit * targetRate * 10) / 10;
                        curObject.LaborUnit = Math.Truncate(curObject.LaborUnit * targetRate * 10) / 10;
                        curObject.ExpenseUnit = Math.Truncate(curObject.ExpenseUnit * targetRate * 10) / 10;

                        exSum += curObject.PriceSum;    //사정율을 적용한 제요율적용제외공종 항목의 합계
                        exCount++; //제요율적용제외공종 항목 수
                    }
                }
            }
        }
        static void SetExcludingPrice()
        {    //제요율적용제외공종 단가 처리 및 재세팅
            long TempInvestDirectSum = Data.Investigation["직공비"];    //조사직공비
            long TempRealDirectSum = FillCostAccount.ToLong(Data.RealDirectMaterial + Data.RealDirectLabor + Data.RealOutputExpense);   //사정율적용 직공비
            decimal InvestExSum = Data.ExcludingMaterial + Data.ExcludingLabor + Data.ExcludingExpense; //조사 제요율적용제외공종
            decimal TempExRate = Math.Round(InvestExSum / TempInvestDirectSum, 5); //조사 직공비 대비 조사 제요율적용제외공종 비율
            long TempExPrice = Convert.ToInt64(Math.Ceiling(TempRealDirectSum * TempExRate));  //사정율적용 제요율적용제외공종
            decimal keyFound = 0;   //금액이 가장 높은 항목에 부족분을 더하는 방법과 모든 항목에 분배해서 더하는 방법 분기 점

            if (Data.CostAccountDeduction.Equals("1"))
            {
                TempExPrice = Convert.ToInt64(Math.Ceiling(Math.Ceiling(Convert.ToDecimal(TempExPrice * 0.997m))));
            }   //제경비 99.7% 옵션 적용시 TempExPrice 업데이트

            foreach (var bid in eleBID)
            {
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    //제요율적용제외공종 단가 재세팅
                    if (curObject.Item == "제요율적용제외")
                    {
                        if (maxBid == null)
                        { //maxBid 초기화
                            maxBid = bid;
                        }
                        if (String.Concat(bid.Element("C15").Value) == "1" && String.Concat(bid.Element("C19").Value) != "0")
                        {
                            if (Convert.ToDecimal(bid.Element("C19").Value) > Convert.ToDecimal(maxBid.Element("C19").Value))
                            {
                                if ((Convert.ToDecimal(string.Concat(bid.Element("C19").Value)) * 1.5m) > curObject.PriceSum + (TempExPrice - exSum))
                                {
                                    keyFound = 1;
                                    maxBid = bid;
                                }
                            }
                        }   //수량이 1이고 합계단가가 0이 아닐 때, 조정된 금액이 조사금액의 150% 미만이면 maxBid 업데이트

                        bid.Element("C16").Value = curObject.MaterialUnit.ToString();    //재료비 단가
                        bid.Element("C17").Value = curObject.LaborUnit.ToString();       //노무비 단가
                        bid.Element("C18").Value = curObject.ExpenseUnit.ToString();     //경비 단가
                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();    //합계 단가
                        bid.Element("C20").Value = curObject.Material.ToString();    //재료비
                        bid.Element("C21").Value = curObject.Labor.ToString();       //노무비
                        bid.Element("C22").Value = curObject.Expense.ToString();     //경비
                        bid.Element("C23").Value = curObject.PriceSum.ToString();    //합계
                    }
                }
            }

            if (keyFound == 0)
            {  //조건에 부합하는 maxBid를 찾지 못하면 모든 제요율적용제외공종 항목에 값을 분배하여 적용
                decimal divisionPrice = Math.Truncate((TempExPrice - exSum) / exCount);   //항목의 수에 따라 분배한 금액
                decimal deficiency = Math.Ceiling((TempExPrice - exSum) - (divisionPrice * exCount)); //절사, 반올림에 따른 부족분
                decimal count = 0;
                while (count != exCount)
                {
                    foreach (var bid in eleBID)
                    {
                        if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                        {
                            var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                            var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                            var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                            var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                            if (curObject.Item == "제요율적용제외" && curObject.Quantity == 1)
                            {
                                if (curObject.LaborUnit != 0)
                                {
                                    if ((Convert.ToDecimal(string.Concat(bid.Element("C19").Value)) * 1.5m) > (curObject.LaborUnit + divisionPrice))
                                    {
                                        curObject.LaborUnit += divisionPrice;
                                        bid.Element("C17").Value = curObject.LaborUnit.ToString();        //노무비 단가
                                        bid.Element("C19").Value = curObject.UnitPriceSum.ToString();     //합계 단가
                                        bid.Element("C21").Value = curObject.Labor.ToString();            //노무비
                                        bid.Element("C23").Value = curObject.PriceSum.ToString();         //합계
                                        count++;
                                    }

                                    if (count == exCount)
                                    {   //절사, 반올림에 따른 부족분 조정
                                        bid.Element("C17").Value = (deficiency + curObject.LaborUnit).ToString();       //노무비 단가
                                        bid.Element("C19").Value = (deficiency + curObject.UnitPriceSum).ToString();    //합계 단가
                                        bid.Element("C21").Value = (deficiency + curObject.Labor).ToString();           //노무비
                                        bid.Element("C23").Value = (deficiency + curObject.PriceSum).ToString();        //합계
                                        break;
                                    }
                                }
                                else
                                {
                                    if (curObject.ExpenseUnit != 0)
                                    {
                                        if ((Convert.ToDecimal(string.Concat(bid.Element("C19").Value)) * 1.5m) > (curObject.ExpenseUnit + divisionPrice))
                                        {
                                            curObject.ExpenseUnit += divisionPrice;
                                            bid.Element("C18").Value = curObject.ExpenseUnit.ToString();      //경비 단가
                                            bid.Element("C19").Value = curObject.UnitPriceSum.ToString();     //합계 단가
                                            bid.Element("C22").Value = curObject.Expense.ToString();          //경비
                                            bid.Element("C23").Value = curObject.PriceSum.ToString();         //합계
                                            count++;
                                        }

                                        if (count == exCount)
                                        {   //절사, 반올림에 따른 부족분 조정
                                            bid.Element("C18").Value = (deficiency + curObject.ExpenseUnit).ToString();     //경비 단가
                                            bid.Element("C19").Value = (deficiency + curObject.UnitPriceSum).ToString();    //합계 단가
                                            bid.Element("C22").Value = (deficiency + curObject.Expense).ToString();         //경비
                                            bid.Element("C23").Value = (deficiency + curObject.PriceSum).ToString();        //합계
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if ((Convert.ToDecimal(string.Concat(bid.Element("C19").Value)) * 1.5m) > (curObject.MaterialUnit + divisionPrice))
                                        {
                                            curObject.MaterialUnit += divisionPrice;
                                            bid.Element("C16").Value = curObject.MaterialUnit.ToString();     //재료비 단가
                                            bid.Element("C19").Value = curObject.UnitPriceSum.ToString();     //합계 단가
                                            bid.Element("C20").Value = curObject.Material.ToString();         //재료비
                                            bid.Element("C23").Value = curObject.PriceSum.ToString();         //합계
                                            count++;
                                        }

                                        if (count == exCount)
                                        {   //절사, 반올림에 따른 부족분 조정
                                            bid.Element("C16").Value = (deficiency + curObject.MaterialUnit).ToString();    //재료비 단가
                                            bid.Element("C19").Value = (deficiency + curObject.UnitPriceSum).ToString();    //합계 단가
                                            bid.Element("C20").Value = (deficiency + curObject.Material).ToString();        //재료비
                                            bid.Element("C23").Value = (deficiency + curObject.PriceSum).ToString();        //합계
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if ((keyFound == 1) && (exSum < TempExPrice))
            {
                maxBid.Element("C17").Value = (Convert.ToDecimal(maxBid.Element("C17").Value) + TempExPrice - exSum).ToString();
                maxBid.Element("C19").Value = (Convert.ToDecimal(maxBid.Element("C19").Value) + TempExPrice - exSum).ToString();
                maxBid.Element("C21").Value = (Convert.ToDecimal(maxBid.Element("C21").Value) + TempExPrice - exSum).ToString();
                maxBid.Element("C23").Value = (Convert.ToDecimal(maxBid.Element("C23").Value) + TempExPrice - exSum).ToString();
                //소수부분 차이에 의한 99.7% 이하 위반 문제에 대한 처리 (노무비에 보정)
            }
        }

        static void GetAdjustedExcludePrice()
        {  //사정율 적용한 제요율적용제외 금액 저장
            foreach (var bid in eleBID)
            {
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    if (curObject.Item.Equals("제요율적용제외"))
                    {
                        Data.AdjustedExMaterial += Convert.ToDecimal(string.Concat(bid.Element("C20").Value));
                        Data.AdjustedExLabor += Convert.ToDecimal(string.Concat(bid.Element("C21").Value));
                        Data.AdjustedExExpense += Convert.ToDecimal(string.Concat(bid.Element("C22").Value));
                    }
                }
            }
        }

        public static void SetBusinessInfo()
        {
            
            foreach (var bid in eleBID)
            {
                if (bid.Name == "T1")
                {
                    bid.Element("C17").Value = Data.CompanyRegistrationNum;
                    bid.Element("C18").Value = Data.CompanyRegistrationName;
                }
            }

        }

        public static void SetPriceOfSuperConstruction()    //상위 공종의 각 단가 합 및 합계 세팅 (23.02.07)
        {
            XElement? firstConstruction = null;     //가장 상위 공종
            XElement? secondConstruction = null;    //중간 상위 공종
            XElement? thirdConstruction = null;     //마지막 상위 공종

            foreach (var bid in eleBID)
            {
                if(bid.Name == "T3")
                {
                    if (string.Concat(bid.Element("C5").Value) == "G")  //공종이면
                    {
                        if (bid.Element("C23").Value == "0")    //이미 합계가 세팅되어 있는지 확인 (중복 계산을 막기 위함)
                        {
                            if (firstConstruction == null || string.Concat(bid.Element("C3").Value) == "0") //C3이 0이면 가장 상위 공종
                            {
                                firstConstruction = bid;    //현재 보고있는 object가 가장 상위 공종
                                secondConstruction = null;  //중간 상위 공종 초기화
                                thirdConstruction = null;   //마지막 상위 공종 초기화
                            }
                            else if (string.Concat(bid.Element("C3").Value) == string.Concat(firstConstruction.Element("C2").Value) && firstConstruction != null)   //C3이 가장 상위 공종의 C2와 같다면 중간 상위 공종
                            {
                                secondConstruction = bid;   //현재 보고있는 object가 중간 상위 공종
                                thirdConstruction = null;   //마지막 상위 공종 초기화
                            }
                            else if (string.Concat(bid.Element("C3").Value) == string.Concat(secondConstruction.Element("C2").Value) && secondConstruction != null) // C3이 중간 상위 공종의 C2와 같다면 마지막 상위 공종
                                thirdConstruction = bid;    //현재 보고있는 object가 마지막 상위 공종
                        }
                        else   //공종에 합계가 이미 세팅되어 있다면 전부 초기화
                        {
                            firstConstruction = null;
                            secondConstruction = null;
                            thirdConstruction = null;
                        }
                    }
                    else if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")    //공종이 아니면
                    {
                        if (firstConstruction != null)  //현재 보는 object가 가장 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            firstConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));    //재료비
                            firstConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));    //노무비
                            firstConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));    //경비
                            firstConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(firstConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));    //합계
                        }
                        if (secondConstruction != null) //현재 보는 object가 중간 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            secondConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));  //재료비
                            secondConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));  //노무비
                            secondConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));  //경비
                            secondConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(secondConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));  //합계
                        }
                        if (thirdConstruction != null)  //현재 보는 object가 마지막 상위 공종에 포함되어 있다면 단가별 합과 합계를 더해나감
                        {
                            thirdConstruction.Element("C20").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C20").Value) + Convert.ToDecimal(bid.Element("C20").Value));    //재료비
                            thirdConstruction.Element("C21").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C21").Value) + Convert.ToDecimal(bid.Element("C21").Value));    //노무비
                            thirdConstruction.Element("C22").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C22").Value) + Convert.ToDecimal(bid.Element("C22").Value));    //경비
                            thirdConstruction.Element("C23").Value = string.Concat(Convert.ToDecimal(thirdConstruction.Element("C23").Value) + Convert.ToDecimal(bid.Element("C23").Value));    //합계 
                        }
                    }
                }
            }
        }

        static void SubstitutePrice()
        {  //BID 파일 내 원가계산서 관련 금액 세팅
            foreach (var bid in eleBID)
            {
                if (bid.Name == "T5")   //bid.Name이 T5인지를 확인함으로 간단하게 원가 계산서부분의 element 인지를 판별. Tag는 T3가 아닌 T5 기준을 따른다. (23.01.31 수정)
                {
                    if (string.Concat(bid.Element("C4").Value) != "이윤" && Data.Bidding.ContainsKey(string.Concat(bid.Element("C4").Value)))
                    {
                        bid.Element("C8").Value = Data.Bidding[string.Concat(bid.Element("C4").Value)].ToString();
                    }
                    else if (Data.Rate1.ContainsKey(string.Concat(bid.Element("C4").Value)))
                    {
                        bid.Element("C8").Value = Data.Bidding[string.Concat(bid.Element("C4").Value)].ToString();
                    }
                }
            }

            if(File.Exists(Data.work_path + "\\Result_Xml.xml"))  //기존 Result_Xml 파일은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\Result_Xml.xml");
            }

            //작업후 xml 파일 저장
            StringBuilder sb = new StringBuilder();
            XmlWriterSettings xws = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                Indent = true
            };
            using (XmlWriter xw = XmlWriter.Create(sb, xws))
            {
                docBID.WriteTo(xw);
            }
            File.WriteAllText(Path.Combine(Data.work_path, "Result_Xml.xml"), sb.ToString());
        }

        public static void CreateZipFile(IEnumerable<string> files)
        {
            if (File.Exists(Data.work_path + "\\입찰내역.zip"))  //기존 입찰내역.zip 파일은 삭제한다. (23.02.02)
            {
                File.Delete(Data.work_path + "\\입찰내역.zip");
            }

            var Zip = ZipFile.Open(Path.Combine(Data.work_path, "입찰내역.zip"), ZipArchiveMode.Create);
            foreach (var file in files)
            {
                Zip.CreateEntryFromFile(file, Path.GetFileName(file), CompressionLevel.Optimal);
            }
            Zip.Dispose();
        }
        static void CreateFile()
        {
            //최종 입찰내역 파일 세부공사별로 생성 
            CreateResultFile.Create();
            //생성된 입찰내역 파일 압축 
            string[] files = Directory.GetFiles(Data.folder, "*.xls");  //폴더 경로 수정 (23.02.02)
            CreateZipFile(files);
        }
        static void Reset()
        {
            Data.ExecuteReset = "1";    //Reset 함수 사용 여부

            var DM = Data.Investigation["직접재료비"];
            var DL = Data.Investigation["직접노무비"];
            var OE = Data.Investigation["산출경비"];
            var FM = Data.InvestigateFixedPriceDirectMaterial;
            var FL = Data.InvestigateFixedPriceDirectLabor;
            var FOE = Data.InvestigateFixedPriceOutputExpense;
            var SM = Data.InvestigateStandardMaterial;
            var SL = Data.InvestigateStandardLabor;
            var SOE = Data.InvestigateStandardExpense;
            //조사 내역서 정보 백업

            Data.RealDirectMaterial = DM;
            Data.RealDirectLabor = DL;
            Data.RealOutputExpense = OE;
            Data.FixedPriceDirectMaterial = FM;
            Data.FixedPriceDirectLabor = FL;
            Data.FixedPriceOutputExpense = FOE;
            Data.StandardMaterial = SM;
            Data.StandardLabor = SL;
            Data.StandardExpense = SOE;
            //사정율 재적용을 위한 초기화

            foreach (var bid in eleBID) //Dictionary 초기화
            {
                //일반 항목인 경우
                if (bid.Element("C9") != null && string.Concat(bid.Element("C5").Value) == "S")
                {
                    var constNum = string.Concat(bid.Element("C1").Value);      //세부공사 번호
                    var numVal = string.Concat(bid.Element("C2").Value);        //세부공종 번호
                    var detailVal = string.Concat(bid.Element("C3").Value);     //세부 공종 번호

                    //현재 탐색 공종
                    var curObject = Data.Dic[constNum].Find(x => x.WorkNum == numVal && x.DetailWorkNum == detailVal);
                    curObject.MaterialUnit = Convert.ToDecimal(string.Concat(bid.Element("C16").Value));
                    curObject.LaborUnit = Convert.ToDecimal(string.Concat(bid.Element("C17").Value));
                    curObject.ExpenseUnit = Convert.ToDecimal(string.Concat(bid.Element("C18").Value));
                }
            }
            Data.ExecuteReset = "0";    //Reset 함수 사용이 끝나면 다시 0으로 초기화
        }
        public static void Calculation()
        {
            docBID = XDocument.Load(Path.Combine(Data.folder, "Setting_Xml.xml"));  //폴더 경로 수정 (23.02.02)
            eleBID = docBID.Root.Elements();
            //가격 재세팅 후 리셋 함수 실행 횟수 증가
            Reset();

            //최저네고단가율 계산 전, 표준시장단가 99.7% 적용옵션에 따른 분기처리
            if (Data.StandardMarketDeduction.Equals("1"))
                ApplyStandardPriceOption();

            GetFixedPriceRate();    //직공비 대비 고정금액 비중 계산
            FindMyPercent();        //최저네고단가율 계산
            GetWeight();            //가중치 계산
            CalculateRate(Data.PersonalRate, Data.BalancedRate);    //Target Rate 계산
            Recalculation();    //사정율에 따른 재계산

            if (exCount != 0)
            {
                SetExcludingPrice();        //제요율적용제외공종 항목 Target Rate 적용
                GetAdjustedExcludePrice();  //사정율 적용한 제요율적용제외 금액 저장
            }

            SetPriceOfSuperConstruction();  //공종 합계 bid에 저장 (23.02.07)

            FillCostAccount.CalculateBiddingCosts();    //원가계산서 사정율적용(입찰) 금액 계산 및 저장
            SetBusinessInfo();      //사업자등록번호 <T1></C17></T1>에 추가
            SubstitutePrice();      //원가계산서 사정율 적용하여 계산한 금액들 BID 파일에도 반영
            CreateFile();           //입찰내역 파일 생성
        }
    }
}