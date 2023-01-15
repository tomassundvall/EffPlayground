namespace Excel.FinancialFunctions

open Excel.FinancialFunctions.Common
open Excel.FinancialFunctions.DayCount

module Bond =
  
  let getPriceYieldFactors settl mat freq basis =
    let dc = dayCount basis
    let numOfCoup = dc.CoupNum settl mat freq
    let prevCoupDate = dc.CoupPCD settl mat freq
    let daysBetween = dc.DaysBetween prevCoupDate settl Numerator
    let coupDays = dc.CoupDays settl mat freq
    let dsc = coupDays - daysBetween
    numOfCoup, prevCoupDate, daysBetween, coupDays, dsc


  let price settl mat (rate : float) (yld : float) redemption (freq : float) basis =
    let numOfCoup, _, daysBetween, coupDays, dsc = 
      getPriceYieldFactors settl mat freq basis
    
    let coupon = 100. * rate / freq
    let accrInt = 100. * rate / freq * daysBetween / coupDays

    let pvFactor k = pow (1. + yld / freq) (k - 1. + dsc / coupDays)
    let pvOfRedemption = redemption / pvFactor numOfCoup

    let mutable pvOfCoupons = 0.
    for k = 1 to int numOfCoup do pvOfCoupons <- pvOfCoupons + coupon / pvFactor (float k)

    if numOfCoup = 1. then
      (redemption + coupon) / (1. + dsc / coupDays * yld / freq) - accrInt
    else
      pvOfRedemption + pvOfCoupons - accrInt