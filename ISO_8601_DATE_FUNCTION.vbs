CurrentDateTime = ISO_8601_DATETIME(Now)

Function ISO_8601_DATETIME(dt)
  s = datepart("yyyy",dt) & "-"
  s = s & RIGHT("0" & datepart("m",dt),2) & "-"
  s = s & RIGHT("0" & datepart("d",dt),2) & "-"
  s = s & RIGHT("0" & datepart("h",dt),2)
  s = s & RIGHT("0" & datepart("n",dt),2)
  ISO_8601_DATETIME = s
End Function
