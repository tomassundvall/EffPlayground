module DayCount.Tests

open Xunit
open System

[<Theory>]
[<InlineData(4)>]
let ``Test freq2months`` (freq) =
  Assert.Equal(3, Excel.FinancialFunctions.DayCount.freq2months freq)


[<Fact>]
let ``Test findPrevAndNextCoupDays`` () =
  let settl = DateTime.Parse("2023-01-14")
  let mat = DateTime.Parse("2029-04-20")
  let numMonths = 3
  let returnLastMonth = false

  let prevCouponDate, nextCoponDate = Excel.FinancialFunctions.DayCount.findPrevAndNextCoupDays settl mat numMonths returnLastMonth

  Assert.Equal<DateTime>(DateTime.Parse("2022-10-20"), prevCouponDate)
  Assert.Equal<DateTime>(DateTime.Parse("2023-01-20"), nextCoponDate)

[<Fact>]
let ``Test datesAggregate1`` () =
  let startDate = DateTime.Parse("2029-04-20")
  let endDate = DateTime.Parse("2023-01-14")
  let numMonths = -3
  let basis = Excel.FinancialFunctions.DayCountBasis.UsPsa30_360
  let f d1 d2 = 0.
  let acc = 0.
  let returnLastMonth = false

  let frontDate, trailingDate, _ = Excel.FinancialFunctions.DayCount.datesAggregate1 startDate endDate numMonths basis f acc returnLastMonth
  Assert.Equal<DateTime>(DateTime.Parse("2022-10-20"), frontDate)
  Assert.Equal<DateTime>(DateTime.Parse("2023-01-20"), trailingDate)


[<Fact>]
let ``Test numberOfCoupons with UsPsa30_300 basis`` () =
  let settl = DateTime.Parse("2023-01-14")
  let mat = DateTime.Parse("2029-04-20")
  let freq = 4.
  let basis = Excel.FinancialFunctions.DayCountBasis.UsPsa30_360

  let numOfCoupons = Excel.FinancialFunctions.DayCount.numberOfCoupons settl mat freq basis
  Assert.Equal<float>(26, numOfCoupons) 


[<Fact>]
let ``Test findCouponDates for UsPsa30_360 basis`` () =
  let settl = DateTime.Parse("2023-01-14")
  let mat = DateTime.Parse("2029-04-20")
  let freq = 4.
  let basis = Excel.FinancialFunctions.DayCountBasis.UsPsa30_360

  let prevCouponDate, nextCopuonDate = Excel.FinancialFunctions.DayCount.findCouponDates settl mat freq basis
  Assert.Equal(DateTime.Parse("2022-10-20"), prevCouponDate)
  Assert.Equal(DateTime.Parse("2023-01-20"), nextCopuonDate)


//[<Theory>]
//[<InlineData("", "", -4)>]
//let ``Test datesAggregate1`` (startDate, endDate, numOfMonths) =
//  0