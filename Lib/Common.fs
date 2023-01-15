namespace Excel.FinancialFunctions
open System
open System.Collections.Generic

module Common = 
  let throw s = failwith s

  let pow x y = Math.Pow (x, y)

  let days (after : DateTime) (before : DateTime) = (after-before).Days
  let (|Date|) (d1 : DateTime) = (d1.Year, d1.Month, d1.Day)
  let lastDayOfMonth y m d = DateTime.DaysInMonth(y, m ) = d
  let date y m d = new DateTime(y, m, d)
  let lastDayOfFebruary (Date(y, m, d) as _) = m = 2 && lastDayOfMonth y m d
  let isLeapYear (Date(y,_,_) as d) = DateTime.IsLeapYear(y)

  let memorize f = 
    let m = new Dictionary<_,_> ()
    fun x ->
      lock m (fun () ->
        match m.TryGetValue(x) with
        | true, res -> res
        | false, _ ->
          let r = f x
          m.Add(x, r)
          r
      )