module Common.Tests

open Xunit


[<Theory>]
[<InlineData(2029, 04, 20)>]
let ``Test lastDayOfMonth`` (y, m, d) =
  Assert.Equal(false, Excel.FinancialFunctions.Common.lastDayOfMonth y m d)