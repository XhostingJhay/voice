dim speechobject
set speechobject=createobject("sapi.spvoice")
set speechobject.voice=speechobject.GetVoices.Item(1)
speechobject.speak "WELCOME TO MY SYSTEM!!!"
