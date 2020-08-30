Dim speaks, speech
speaks = "The time is " & hour(time) & " AM"
If Minute(Time) > 0 Then speaks = "The time is " & hour(time) & " " & Minute(Time) & " AM"
If hour(time)>12 Then speaks = "The time is " & hour(time) - 12 & " PM"
If hour(time)>12 Then If Minute(Time) > 0 Then speaks = "The time is " & hour(time) - 12 & " " & Minute(Time) & " PM"
Set speech=CreateObject("sapi.spvoice")
Set speech.voice = speech.GetVoices.Item(0)
speech.Volume = 50
speech.Speak speaks