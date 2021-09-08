CurrentDateTime = iso8601DateTime(Now)

Function iso8601DateTime(dt)
  s = datepart("yyyy",dt) & "-"
  s = s & RIGHT("0" & datepart("m",dt),2) & "-"
  s = s & RIGHT("0" & datepart("d",dt),2) & "-"
  s = s & RIGHT("0" & datepart("h",dt),2)
  s = s & RIGHT("0" & datepart("n",dt),2)
  iso8601DateTime = s
End Function
