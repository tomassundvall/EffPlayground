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

  
  let isFeb29BetweenConsecutiveYears (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
    if y1 = y2 && isLeapYear date1 then if m1 <= 2 && m2 > 2 then true else false
    elif y1 = y2 then false
    elif y2 = y1 + 1 then
        if isLeapYear date1 then if m1 <= 2 then true else false
        elif isLeapYear date2 then if m2 > 2 then true else false
        else false
    else throw "isFeb29BetweenConsecutiveYears: function called with non consecutive years"


  let considerAsBisestile (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
      (y1 = y2 && isLeapYear date1) || (m2 = 2 && d2 = 29) || isFeb29BetweenConsecutiveYears date1 date2


  let dateDiff360 sd sm sy ed em ey =
    (ey-sy) * 360 + (em - sm) * 30 + (ed - sd)


  let dateDiff365 (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate) =
    let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
    if sd1 > 28 && sm1 = 2 then sd1 <- 28
    if ed1 > 28 && em1 = 2 then ed1 <- 28
    let startd, endd = date sy1 sm1 sd1, date ey1 em1 ed1
    (ey1 - sy1) * 365 + days endd startd       


  let dateDiff360Eu (Date(sy,sm,sd) as startDate) (Date(ey,em,ed) as endDate) =
    let mutable sd1, sm1, sy1, ed1, em1, ey1, startDate1, endDate1 = sd, sm, sy, ed, em, ey, startDate, endDate
    sd1 <- if sd1 = 31 then 30 else sd1
    ed1 <- if ed1 = 31 then 30 else ed1
    dateDiff360 sd1 sm1 sy1 ed1 em1 ey1


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


  let lessOrEqualToAYearApart (Date(y1, m1, d1) as date1) (Date(y2, m2, d2) as date2) =
      y1 = y2 || (y2 = y1 + 1 && (m1 > m2 || (m1 = m2 && d1 >= d2)))


  let actualCoupDays settl mat freq =
    let pcd = findPreviousCouponDate settl mat freq DayCountBasis.ActualActual
    let ncd = findNextCouponDate settl mat freq DayCountBasis.ActualActual
    float (days ncd pcd)


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
          let totDaysInCoup = dateDiff360Us pdc ndc ModifyBothDates 
          let daysToSettl =  dateDiff360Us pdc settl ModifyStartDate
          float (totDaysInCoup - daysToSettl)

        member _.DaysBetween issue settl _ =
          float (dateDiff360Us issue settl ModifyStartDate)

        member _.DaysInYear _ _ =
          360.

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.UsPsa30_360 returnLastDay
    }

  let Europ30_360 () =
    {
      new IDayCount with
        member _.CoupDays _ _ freq =
          360. / freq

        member _.CoupPCD settl mat freq =
          findPreviousCouponDate settl mat freq DayCountBasis.Europ30_360

        member _.CoupNCD settl mat freq =
          findNextCouponDate settl mat freq DayCountBasis.Europ30_360

        member _.CoupNum settl mat freq =
          numberOfCoupons settl mat freq DayCountBasis.Europ30_360

        member this.CoupDaysBS settl mat freq =
          float(dateDiff360Eu (this.CoupPCD settl mat freq) settl)

        member this.CoupDaysNC settl mat freq =
          float(dateDiff360Eu settl (this.CoupNCD settl mat freq))

        member _.DaysBetween issue settl _ =
          float (dateDiff360Eu issue settl)

        member _.DaysInYear _ _ =
          360.

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.Europ30_360 returnLastDay
    }

  let Actual360 () =
    {
      new IDayCount with
        member _.CoupDays _ _ freq =
          360. / freq

        member _.CoupPCD settl mat freq =
          findPreviousCouponDate settl mat freq DayCountBasis.Actual360

        member _.CoupNCD settl mat freq =
          findNextCouponDate settl mat freq DayCountBasis.Actual360

        member _.CoupNum settl mat freq =
          numberOfCoupons settl mat freq DayCountBasis.Actual360

        member this.CoupDaysBS settl mat freq =
          float(days settl (this.CoupPCD settl mat freq))

        member this.CoupDaysNC settl mat freq =
          float (days (this.CoupNCD settl mat freq) settl)

        member _.DaysBetween issue settl position =
          if position = Numerator
          then float (days settl issue)
          else float (dateDiff360Us issue settl ModifyStartDate)

        member _.DaysInYear _ _ =
          360.

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.Actual360 returnLastDay               
    }

  let Actual365 () =
    {
      new IDayCount with
        member _.CoupDays _ _ freq =
          365. / freq

        member _.CoupPCD settl mat freq =
          findPreviousCouponDate settl mat freq DayCountBasis.Actual365

        member _.CoupNCD settl mat freq =
          findNextCouponDate settl mat freq DayCountBasis.Actual365

        member _.CoupNum settl mat freq =
          numberOfCoupons settl mat freq DayCountBasis.Actual365

        member this.CoupDaysBS settl mat freq =
          float(days settl (this.CoupPCD settl mat freq))

        member this.CoupDaysNC settl mat freq =
          float (days (this.CoupNCD settl mat freq) settl)

        member _.DaysBetween issue settl position =
          if position = Numerator
          then float (days settl issue)
          else float (dateDiff365 issue settl)

        member _.DaysInYear _ _ =
          365.

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.Actual365 returnLastDay              
    }

  let ActualActual () =
    {
      new IDayCount with
        member _.CoupDays settl mat freq =
          actualCoupDays settl mat freq

        member _.CoupPCD settl mat freq =
          findPreviousCouponDate settl mat freq DayCountBasis.ActualActual
        
        member _.CoupNCD settl mat freq =
          findNextCouponDate settl mat freq DayCountBasis.ActualActual

        member _.CoupNum settl mat freq =
          numberOfCoupons settl mat freq DayCountBasis.ActualActual

        member this.CoupDaysBS settl mat freq =
          float(days settl (this.CoupPCD settl mat freq))

        member this.CoupDaysNC settl mat freq =
          float (days (this.CoupNCD settl mat freq) settl)

        member _.DaysBetween startDate endDate _ =
          float (days endDate startDate)

        member _.DaysInYear issue settl =
          if not(lessOrEqualToAYearApart issue settl) then
            let totYears = (settl.Year - issue.Year) + 1
            let totDays = days (date (settl.Year + 1) 1 1) (date issue.Year 1 1)
            float totDays / float totYears
          elif considerAsBisestile issue settl then 366. else 365.                    

        member _.ChangeMonth date months returnLastDay =
          changeMonth date months DayCountBasis.ActualActual returnLastDay     
    }

  let getDayCountImpl = memorize (function
    | DayCountBasis.UsPsa30_360                 -> UsPsa30_360 ()
    | DayCountBasis.ActualActual                -> ActualActual ()
    | DayCountBasis.Actual360                   -> Actual360 ()
    | DayCountBasis.Actual365                   -> Actual365 ()
    | DayCountBasis.Europ30_360                 -> Europ30_360 ()
    | _                                         -> throw "dayCount: it should never get here")