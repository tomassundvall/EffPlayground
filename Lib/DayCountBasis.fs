namespace Excel.FinancialFunctions

open System
open Common

module DayCount =
  type Method360Us =
  | ModifyStartDate
  | ModifyBothDates

  type NumDenumPosition =
  | Denumerator
  | Numerator

  type IDayCount =
    abstract CoupDays: DateTime -> DateTime -> float -> float
    abstract CoupPCD: DateTime -> DateTime -> float -> DateTime
    abstract CoupNCD: DateTime -> DateTime -> float -> DateTime
    abstract CoupNum: DateTime -> DateTime -> float -> float
    abstract CoupDaysBS: DateTime -> DateTime -> float -> float
    abstract CoupDaysNC: DateTime -> DateTime -> float -> float
    abstract DaysBetween: DateTime -> DateTime -> NumDenumPosition -> float
    abstract DaysInYear: DateTime -> DateTime -> float
    abstract ChangeMonth: DateTime -> int -> bool -> DateTime

  let dateDiff360 sd sm sy ed em ey =
    (ey-sy) * 360 + (em - sm) * 30 + (ed - sd)

  let dateDiff365 (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate) =
    let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
    if sd1 > 28 && sm1 = 2 then sd1 <- 28
    if ed1 > 28 && em1 = 2 then ed1 <- 28
    let startd, endd = date sy1 sm1 sd1, date ey1 em1 ed1
    (ey1 - sy1) * 365 + days endd startd       

  let dateDiff360Us (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate)  method360 =
    let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
    if lastDayOfFebruary endDate1 && (lastDayOfFebruary startDate1 || method360 = ModifyBothDates)
        then ed1 <- 30
    if ed1 = 31 && (sd1 >= 30 || method360 = ModifyBothDates) then ed1 <- 30
    if sd1 = 31 then sd1 <- 30
    if lastDayOfFebruary startDate1 then sd1 <- 30
    dateDiff360 sd1 sm1 sy1 ed1 em1 ey1 

  let freq2months (freq : float) = 12 / int freq

  let isLastDayOfMonthBasis y m d basis = lastDayOfMonth y m d || (d = 30 && basis = DayCountBasis.UsPsa30_360)

  let changeMonth (Date(y, m, d) as orgDate) numMonths _ returnLastDay =
    let getLastDay y m = DateTime.DaysInMonth(y, m)
    let (Date(year, month, _) as newDate) = orgDate.AddMonths(numMonths)
    if returnLastDay then date year month (getLastDay year month) else newDate

  let noActionDates (d1 : DateTime) (d2 : DateTime) = 0.

  // need to understand what the purpose of this function is
  let datesAggregate1 startDate endDate numMonths basis f (acc : float) returnLastMonth =
    let rec iter frontDate trailingDate acc =
      let stop = if numMonths > 0 then frontDate >= endDate else frontDate <= endDate
      if stop then frontDate, trailingDate, acc
      else
        let trailingDate = frontDate
        let frontDate = changeMonth frontDate numMonths basis returnLastMonth
        let acc = acc + f frontDate trailingDate
        iter frontDate trailingDate acc
    iter startDate endDate acc

  let findPcdNcd startDate endDate numMonths basis returnsLastMonth =
    let pcd, ncd, _ = datesAggregate1 startDate endDate numMonths basis noActionDates 0. returnsLastMonth
    pcd, ncd

  let findCouponDates settl (Date(my, mm, md) as mat) freq basis =
    let endMonth = lastDayOfMonth my mm md
    let numMonths = - freq2months freq
    findPcdNcd mat settl numMonths basis endMonth

  let findPreviousCouponDate settl mat freq basis = 
    findCouponDates settl mat freq basis |> fst

  let findNextCouponDate settl mat freq basis =
    findCouponDates settl mat freq basis |> snd 

  let numberOfCoupons settl (Date(my, mm, md) as mat) freq basis =
    let (Date(pcy, pcm, _) as _) = findPreviousCouponDate settl mat freq basis
    let months = float ((my - pcy)*12 + (mm - pcm))
    months * freq / 12.
  
  let UsPsa30_360 () =
    {
      new IDayCount with
        member _.CoupDays _ _ freq = 
          360. / freq
                
        member _.CoupPCD settl mat freq = 
          findPreviousCouponDate settl mat freq DayCountBasis.UsPsa30_360
                    
        member _.CoupNCD settl mat freq =
          findNextCouponDate settl mat freq DayCountBasis.UsPsa30_360
        
        member _.CoupNum settl mat freq =
          numberOfCoupons settl mat freq DayCountBasis.UsPsa30_360

        member this.CoupDaysBS settl mat freq =
          dateDiff360Us (this.CoupPCD settl mat freq) settl ModifyStartDate
          |> float
        
        member _.CoupDaysNC settl mat freq =
          let pdc = findPreviousCouponDate settl mat freq DayCountBasis.UsPsa30_360
          let ndc = findNextCouponDate settl mat freq DayCountBasis.UsPsa30_360
          let totDaysInCoup = dateDiff360Us pdc ndc Method360Us.ModifyBothDates 
          let daysToSettl =  dateDiff360Us pdc settl Method360Us.ModifyStartDate
          float (totDaysInCoup - daysToSettl)

        member _.DaysBetween issue settl _ =
          float (dateDiff360Us issue settl Method360Us.ModifyStartDate)

        member _.DaysInYear _ _ =
          360.

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.UsPsa30_360 returnLastDay
    }


  let dayCount = memorize (function
    | DayCountBasis.UsPsa30_360                 -> UsPsa30_360 ()
    // | DayCountBasis.ActualActual                -> ActualActual ()
    // | DayCountBasis.Actual360                   -> Actual360 ()
    // | DayCountBasis.Actual365                   -> Actual365 ()
    // | DayCountBasis.Europ30_360                 -> Europ30_360 ()
    | _                                         -> throw "dayCount: it should never get here")


 // let dayCount = memorize (function
 //     | DayCountBasis.UsPsa30_360   -> UsPsa30_360 ()
 //     | DayCountBasis.
 //     | _ ->                        -> throw "f")