Option Explicit

Dim Voice
Dim Text
Set Voice = CreateObject("sapi.spvoice")

Text=InputBox("What to say?","Text to speech")

'Zira's Voice
Set Voice.Voice = Voice.GetVoices.Item(1)
Voice.Rate = 2
Voice.Volume = 70
Voice.Speak "My Name is Zira. It's nice to meet you!"
Voice.Speak "Your message is " & Text

'David's Voice
Set Voice.Voice = Voice.GetVoices.Item(0)
Voice.Rate = 2
Voice.Volume = 100
Voice.Speak "My Name is David. It's nice to meet you!"
Voice.Speak "Your message is " & Text


