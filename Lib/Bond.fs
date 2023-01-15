namespace Excel.FinancialFunctions

open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module Bond =
  open System
  
  let getPriceYieldFactors settl mat freq basis =
    let dc = getDayCountImpl basis
    let numOfCoup = dc.CoupNum settl mat freq
    let prevCoupDate = dc.CoupPCD settl mat freq
    let daysBetween = dc.DaysBetween prevCoupDate settl Numerator
    let coupDays = dc.CoupDays settl mat freq
    let dsc = coupDays - daysBetween
    numOfCoup, prevCoupDate, daysBetween, coupDays, dsc


  let price (settl : DateTime) (mat : DateTime) (rate : float) (yld : float) (redemption : float) (freq : float) (basis : DayCountBasis) =
    let dayCountImpl = getDayCountImpl basis

    let numOfCoupons = dayCountImpl.CoupNum settl mat freq |> int
    let prevCouponDate = dayCountImpl.CoupPCD settl mat freq
    let daysBetween = dayCountImpl.DaysBetween prevCouponDate settl Numerator
    let couponDays = dayCountImpl.CoupDays settl mat freq
    let dsc = couponDays - daysBetween

    let coupon = 100. * rate / freq
    let accrInt = 100. * rate / freq * daysBetween / couponDays

    let pvFactor k = pow (1. + yld / freq) (k - 1. + dsc / couponDays)
    let pvOfRedemption = redemption / pvFactor numOfCoupons

    let pvOfCoupons = 
      seq { 1 .. numOfCoupons}
      |> Seq.map (fun k -> coupon / pvFactor (float k))
      |> Seq.sum

    match numOfCoupons with
    | 1 -> (redemption + coupon) / (1. + dsc / couponDays * yld / freq) - accrInt
    | _ -> pvOfRedemption + pvOfCoupons - accrInt