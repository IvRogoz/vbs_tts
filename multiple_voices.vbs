dim tts_engine
Dim Text

set tts_engine=createobject("sapi.spvoice")
Text = "Hello there user"
tts_engine.speak Text