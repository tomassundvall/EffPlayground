namespace Excel.FinancialFunctions

open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module Bond =
  open System

  let formatDate (dt : DateTime) =
    dt.ToString("yyy-MM-dd")

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

    printfn "---"
    printfn "Number of Coupons:\t%i" numOfCoupons
    printfn "Previous Coupon Date\t%s" (prevCouponDate |> formatDate)
    printfn "DSC:\t\t\t%f" dsc 
    //printfn "Days Between\t\t%i" 

    let pvFactor k = pow (1. + yld / freq) (k - 1. + dsc / couponDays)
    let pvOfRedemption = redemption / pvFactor numOfCoupons

    let pvOfCoupons = 
      seq { 1 .. numOfCoupons}
      |> Seq.map (fun k -> coupon / pvFactor (float k))
      |> Seq.sum

    match numOfCoupons with
    | 1 -> (redemption + coupon) / (1. + dsc / couponDays * yld / freq) - accrInt
    | _ -> pvOfRedemption + pvOfCoupons - accrInt


  let yieldFunc settlement maturity rate pr redemption frequency basis =
    let n, pcd, a, e, dsr = getPriceYieldFactors settlement maturity frequency basis
    if n <= 1. then
      let k = (redemption / 100. + rate / frequency) / (pr / 100. + (a / e * rate /frequency)) - 1.0
      k * frequency * e / dsr
    else
      findRoot (fun yld -> price settlement maturity rate yld redemption frequency basis - pr) 0.05