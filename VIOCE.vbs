dim speechobject
set speechobject=createobject("sapi.spvoice")
set speechobject.voice=speechobject.GetVoices.Item(1)
speechobject.speak "Attention Co-workers! Code 500 is now Activated!!!"
