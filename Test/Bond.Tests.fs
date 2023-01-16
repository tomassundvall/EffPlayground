module Tests

open System
open Xunit
open Excel.FinancialFunctions.Bond

let toFiveDec (x : float) =
  Math.Round (x, 5)

let toFourDec (x : float) =
  Math.Round (x, 4)

[<Theory>]
//           result            settl         mat            rate        yld            redemption    freq
[<InlineData(88.18822,        "2023-01-14", "2029-04-20",   0.07294,    0.098415429,   100.,         4.)>]
[<InlineData(90.13965,        "2023-01-15", "2024-01-01",   0.07294,    0.187362464,   100.,         4.)>]
[<InlineData(89.78170,        "2023-01-01", "2024-01-01",   0.07294,    0.187362464,   100.,         4.)>]
[<InlineData(90.36331,        "2023-01-01", "2024-01-01",   0.07294,    0.187362464,   100.,         1.)>]
let ``Test valid inputs for Bond.price function for UsPsa30_360 basis`` (r, settl, mat, rate, yld, redemption, freq) =
  let settl' = DateTime.Parse(settl)
  let mat' = DateTime.Parse(mat)
  
  let price' = 
    price settl' mat' rate yld redemption freq Excel.FinancialFunctions.DayCountBasis.UsPsa30_360
    |> toFiveDec

  Assert.Equal<float>(r, price')


[<Theory>]
//           result    settl         maturity     rate      price   redemption  freq
[<InlineData(0.0920, "2023-01-15", "2029-04-20", 0.07294,  91.0,   100.,       4.)>]
let ``Test valid inputs for Yield.price function for UsPsa30_300`` (r, settl, mat, rate, pr, redemption, freq) =
  let settl' = DateTime.Parse(settl)
  let mat' = DateTime.Parse(mat)

  let yield' = 
    yieldFunc settl' mat' rate pr redemption freq Excel.FinancialFunctions.DayCountBasis.UsPsa30_360
    |> toFourDec

  Assert.Equal<float>(r, yield')


[<Fact>]
let ``Test getPriceYieldFactors with UsPsa30_300 basis`` () =
  let settl = DateTime.Parse("2023-01-14")
  let mat = DateTime.Parse("2029-04-20")
  let freq = 4.
  let basis = Excel.FinancialFunctions.DayCountBasis.UsPsa30_360

  let numOfCoupons, _, daysBetween, couponDays, dsc = getPriceYieldFactors settl mat freq basis
  Assert.Equal<float>(26, numOfCoupons)
  Assert.Equal<float>(84, daysBetween)
  Assert.Equal<float>(90, couponDays)
  Assert.Equal<float>(6, dsc)