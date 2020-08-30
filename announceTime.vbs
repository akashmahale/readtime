Dim speaks, speech

speaks = hour(time) & "ji desu"
If Minute(Time)>0 Then speaks = hour(time) & "ji " & Minute(Time) & " bun desu."
If Minute(Time)=30 Then speaks = hour(time) & "ji " & Minute(Time) & " jupun desu."
Set speech=CreateObject("sapi.spvoice")
Set speech.voice = speech.GetVoices.Item(2)
speech.Volume = 50
speech.Speak speaks
