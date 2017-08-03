Dim message, sapi
Set sapi=CreateObject("sapi.spvoice")
do
message=InputBox("Speech Bot. Version 1.4")

if message <> "" then sapi.Speak message
loop Until message=""
